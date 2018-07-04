unit Commun;

interface

uses
{$IFDEF EAGLCLIENT}
{$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF VER150}
     Variants,
{$ENDIF}
  Forms {TModalResult}, sysutils, Controls,
  extCtrls, hctrls , Constantes, Hmsgbox, Hent1, UTOB,Classes, UProcGen
  , CpTypeCons
  , Ent1
  ,UFonctionsCBP
  ;

{********* GESTION DES COMPTEURS ET DES SOUCHES ********}
Function GetNum                (Module : string ; Societe : string; Transac : string) : string ;
Function SetNum                (Module : string ; Societe : string; Transac : string; Num : string) : string ;
Function InitNumTransac        (Module : string ; Societe : string ; Transaction : string) : string;
{JP 21/12/04 : établit le lien entre TRECRITURE et TRVENTEOPCVM à travers TE_NUMTRANSAC et TVE_NUMVENTE}
function GetNumTransacVente    (NumTransac : string; FromVente : Boolean) : string;
function TrGetNewNumPiece      (Souche, TypeSouche, Societe, Jal : string ; Ecrit : Boolean) : Integer;
{JP 27/11/03 : appelle TrGetNewNumPiece pour une intégration en comptabilité et met à jour TRECRITURE}
function GetSetNumPiece        (tEcriture : TOB) : Integer;
function TrSetNumPiece         (Souche, TypeSouche : string; Num : LongInt) : LongInt;
function TrGetSoucheDuJal      (Jal : string; Cpt : string = '') : TDblResult;
{17/11/06 : Récupération du numéro de période sur les numéros de transaction pour les écritures de bordereau}
function GetNumPeriodeTransac  (NumTransac, Journal : string) : Integer;
{********* GESTION DES COMPTEURS ET DES SOUCHES ********}

{****** CALCUL DE TAUX ************}
Function CalcNbJourParAnBase	(Date : TDateTime ; BaseCalcul : integer) : integer ;
Function CalcNbJourReelBase 	(DateDeb : TDate ; DateFin : TDate ; BaseCalcul : integer) : double ;
Function CalcMontantAgios   	(Montant : double; Taux : double; Base : integer ; NbJour : integer; tauxPreCompte : boolean ; AgiosPreCompte : Boolean) : double ;
Function CalcTauxAgios      	(MontantAgios : Double; Base : integer ; Montant : integer ;NbJour : integer; tauxPostCompte : boolean; AgiosPostCompte : Boolean) : double ;
Function CalcTauxPrePost    	(TauxDepartPost : boolean; ValTaux : double ; NbJour : integer ; Base : integer) : double ;
{****** CALCUL DE TAUX ************}

{****** FONCTIONS GESTION DEVISE ************}
{Récupère la devise du compte bancaire}
Function RetDeviseCompte     (CompteGeneral : string) : string ;
function RetDeviseOPCVM      (OPCVM : string) : string ;
Function GetCompteAgence     (Compte : string) : string ;
Function CalcDecimaleDevise  (Devise  : string) : integer ;
function CalcDeciDeviseIso   (DevIso  : string) : Integer;
function GetInfoDevise       (Dev : string) : TDblResult;
{Renvoie le montant contrevalorisé en DevContrevalorise (l'€ servant de pivot)}
function RetContreValeur     (aDate: TDateTime; Montant : double; DevMontant : string ; DevContrevalorise : string) : double ;
{Renvoie le taux de conversion d'une devise à une autre à une date donnée (l'€ servant de pivot)}
function RetContreTaux       (Date: TDateTime; DevMontant, DevContrevalorise : string) : double ;
function RetPariteEuro       (Devise : string; DateC : TDateTime; CotationOk : Boolean = False) : Double;
{****** FONCTIONS GESTION DEVISE ************}

{****** FONCTIONS GET ************}
function GetCodeAFB	     (transac : string; TypeFlux : string) : string ;
{Récupère le sens à partir de la transaction}
{03/12/04 : Récupère un enregistrement de la table FluxTreso à partir d'un code transaction et d'une "Jambe"}
function GetFluxCibFromTransac(Transac, Jambe : string; AvecRecherche : Boolean = False) : TOB;
{05/01/05 : Chargement des champs de la table FluxTreso dans une Tob TRECRITURE}
procedure ChargeChpFluxTreso (var T : TOB; Transac, Jambe : string; AvecRecherche : Boolean = False);
function GetSensFlux	     (Transac : string; TypeFlux : string) : string ;
{JP 17/09/03 : Récupère le sens à partir du type flux}
function GetSensFromTypeFlux (TypeFlux : string) : string;
{JP 05/11/03 : Récupère les sens à partir du code rubrique}
function GetSensFromRubrique (CodeRub  : string) : string;
function GetContrepartie     (Transac : string; TypeFlux : string) : string ;
{JP 17/09/03 : Récupération du compte général de contrepartie, du sens et du code CIB à partir d'un code flux}
procedure GetCibSensGeneral  (var Gen, Sens, CIB : string; Flux : string);
{JP 27/11/03 : Récupération du mode de règlement à partir d'un sens, d'un CIB et d'un compte général}
function GetModeReglement    (Sens, CIB, Cpte : string; Dossier : string = '') : string;
{10/01/07 : Fonction inverse de GetModeReglement}
function GetCodeCIB          (Sens, Paie, Cpte : string; Dossier : string = '') : string;
function GetRIB              (var Banque, Guichet, Num, Cle, Iban : string ; CompteGeneral : string) : string ;
{18/09/06 : Retourne le Rib et remplit les champs de la table TRECRITURE concernant les informations bancaires}
function GetRIBTob           (var tEcr : TOB) : string ;
function GetConditionFinPlac (var Condition : TOB; Transac : string ; Banque : string ; Compte : string ; societe : string) : boolean ;
function GetConditionDec     (Flux : string ; Compte : string; var TobC : TOB ): Boolean ;
function GetValTauxRef	     (TauxRef : string; Le : string; var Cotation : double) : Boolean ;
function GetCodeFlux	     (CodeTransac, TypeFlux : string) : string;
{Récupère le libéllé à partir d'un code dans la vue TRLISTEFLUX qui est une union entre RUBRIQUE et FLUXTRESO}
function GetLibFluxRub       (Code : string) : string;
{Récupère la cours d'une opcvm à la date donnée}
function GetCoursOpcvm(aCode : string; aDate : TDateTime) : Double;
{****** FONCTIONS GET ************}

{****** FONCTIONS MULTI SOCIETES ************}
{Récupération d'une banque "factice" pour les comptes courant et titre}
function GetBanqueCourant : string;
{Récupération d'une agence "factice" pour les comptes courant et titre}
function GetAgenceCourant : string;
{Récupération d'un Cib "factice" pour les comptes courant et titre}
function GetCibCourant(BqCode : string = '') : string;
{Donne la correspondance entre NoDossier, NomBase et Societe à partir de la table DOSSIER
 ChpSel : champ dont on possède la valeur
 ValSel : valeur de ChpSel
 ChpRes : Champ dont on veut obtenir la valeur}
function GetInfosFromDossier(ChpSel, ValSel, ChpRes : string) : string;
{Retourne le dossier à partir de BQ_Code (alias le champ _GENERAL)}
function GetDossierFromBQCP (BQ_Code : string) : string;
{Retourne le libellé du dossier à partir de BQ_Code (alias le champ _GENERAL)}
function GetLibDossierFromBQ(BQ_Code : string) : string;
{Renvoie la base, le NoDossier et la société depuis un compte}
function GetInfosSocFromBQ(BQ_Code : string; LibOk : Boolean) : TDblResult;
{Récupère dans la table CLIENSSOC le compte courant entre deux sociétés.
 BqCodeOk si on veut BQ_CODE, sinon le retour sera CLS_GENERAL}
function GetCompteCourant(DosSource, DosDest : string; BqCodeOk : Boolean = True) : string;
{Retourne BQ_CODE à partir d'un compte général et du dossier d'origine}
function GetBqCodeFromGene(Gene, Dossier : string) : string;
{Retourne BQ_GENERAL à partir de BQ_CODE}
function GetGeneFromBqCode(BqCode : string) : string;
{Gère-t-on la Tréso en MultiSociété ?}
function IsTresoMultiSoc : Boolean;
{Remplit les champs société et nodossier d'une Tob TRECRITURE à partir de TE_GENERAL}
procedure InitInfosDossier(var T : TOB);
{Exclut la liste des comptes/agences/banques des natures passées en paramètres et séparées par un ";"}
procedure SetPlusBancaire(Composant : TControl; Pref, AExclure : string);
{Renvoie NoDossier1;NoDossier2; depuis une liste de Bases (Base1;Base2;...)}
function GetNoDossierFromBase(ListeBase : string = '') : string;
{Renvoie la la liste des dossier qui ont la base courante comme base Treso.
 Si TokenStOk, chaque nom de base est séparée par un ";", sinon le Result se présente
 sous la forme d'un requête : IN ("BASE1", "BASE2", ...)}
function  FiltreBaseTreso(TokenStOk : Boolean = True) : string;
{Retourne IN ("NoDossier1", "NoDossier2", ...). Si Dossier = "", Filtre sur le regroupement Treso}
function  FiltreNodossier(NomBase : string = '') : string;
{Filtre sur composant attacher à une tablette sur BANQUECP sur les dossiers du regroupement
 Tréso (NomBase = '') ou éventuellement sur demande (Nature du compte, NomBase)}
function  FiltreBanqueCp (DataType, Nature, NomBase : string) : string;
procedure SetPlusBanqueCp(Control : TControl);
{****** FONCTIONS MULTI SOCIETES ************}



{****** CALCULS DE DATE DE VALEUR / CALENDRIER *********}
function GetCalendrier	     (CompteGeneral: String): String;
function CalcDateValeur	     (CodeAFB, CompteGeneral: String; aDate: TDateTime): TDateTime;
function DateEstOuvre	     (aDate : TDateTime; CodeCal : String): Boolean;
function GetPremierOuvre     (aDate : TDateTime; CodeCal : String): TDateTime;
function GetPrecedentOuvre   (aDate : TDateTime; CodeCal : String): TDateTime;
function DateEstOuvrable     (aDate : TDateTime): Boolean;
function GetPremierOuvrable  (aDate : TDateTime): TDateTime;
function GetPrecedentOuvrable(aDate : TDateTime): TDateTime;
{20/11/06 : Renvoie True, si les deux dates ne sont pas sur le même millésime}
function TestDateEtMillesime (NewDt, OldDt : TDateTime;AvecMessage : Boolean = False) : Boolean;
{20/06/07 : Mise en place du concept de création / suppression des écritures}
function CanCreateEcr        (Bouton : TComponent = nil; Parle : Boolean = False) : Boolean;
{20/06/07 : Mise en place du concept de création / suppression des écritures}
function CanValidateBO       (Bouton : TComponent = nil; Parle : Boolean = False) : Boolean;
{****** CALCULS DE DATE DE VALEUR / CALENDRIER *********}

{****** DIVERS  *********}
function  TrShowMessage	      (Caption: String; Num: Integer; P1, P2: String): TModalResult;
procedure ValidCodeOnKeyPress (var Key: Char);
Procedure AssignDrapeau	      (Source : TImage ; CodeIso : string) ;
{Chargement du drapeau et affichage du code devise}
procedure MajAffichageDevise(Image, Libelle : TControl; Code : string; Source : TCatSource);
function  GetValeursMinMax     (Dev, Chp, TT, WH : string; Deb, Fin : TDateTime): TDblResult;
{JP 09/03/04 : création automatique des agences correspondant aux différents comptes}
function  CreerAgence : Boolean;
{JP 23/04/04 : Test des droits d'accès à certaines fonctionnalités par popup}
function  AutoriseFonction(MenuATester : Integer) : Boolean;
{JP 01/10/03 : Formate les clefs de traitement sur les dates comptables et de valeur}
function  RetourneCleEcriture(Dt : TDateTime; NumP : Integer) : string;
{JP 19/11/04 : FQ 10183 : Consulte l'indicateur "appli activée" dans la base commune}
function  IsFlagAppliOk : Boolean;
{09/12/04 : Renvoie le type de vente des opcvm}
function  GetTypeVente(OldValue, Value : string) : string;
{20/12/06 : Retourne le cout d'achat moyen pondéré des OPCVM}
function GetCAMP(Compte, OPCVM : string; DateRef : TDateTime) : Double;
{15/12/04 : Remplit la tob avec le detail des opcvm pour un Portefeuille, un code OPCVM et un général}
procedure ChargeDetailOPCVM(var aTob : TOB; Portef, Opcvm, General : string);
{21/12/04 : Remplit la tob avec le detail des ventes pour un Portefeuille, un code OPCVM et un général}
procedure ChargeDetailVente(var aTob : TOB; NumVente : string; PourEcr : Boolean = False);
{Charge la Variable globale de Treso}
procedure ChargeVarTreso;
{JP 25/01/08 : FQ 10547 : Renvoie les dossiers / Bases autorisés}
function GetTrFiltreDossiers : string;
{Recherche dans toute la hiérarchie de la TOB, les champ "Field" et leur affecte Value}
procedure PutValueWholeTob(MyTob : TOB; FieldName : string; Value : Variant);
{Fonction pour savoir si on gère les nouveaux soldes}
function  IsNewSoldes : Boolean;
{}
{****** DIVERS  *********}

{********* Fonctions pour le moment inutilisées *********}
function calcCMVT	     (nombre : double ; TobCond : TOB) : double ;
function calcInteretDEB	     (nb : integer ; taux : double; base : integer; date : TdateTime) : double ;
{********* Fonctions pour le moment inutilisées *********}


implementation

uses
  {$IFNDEF EAGLSERVER}
  AglInit, Menus,
  {$ENDIF EAGLSERVER}
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  HStatus, HDB, ParamSoc, UtilPgi;

{---------------------------------------------------------------------------------------}
{----------------------- FONCTIONS POUR LE MOMENT INUTILISÉES---------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
function CalcCMVT(Nombre : Double ; TobCond : TOB) : Double ;
{---------------------------------------------------------------------------------------}
var
  Taux : Double ;
begin
  if Nombre <= TobCond.GetValue('TCN_PLAFOND1') then
    GetValTauxRef(TobCond.GetValue('TCN_TAUXREF1'), DateToStr(V_PGI.DateEntree), Taux)
  else if Nombre <= TobCond.GetValue('TCN_PLAFOND2') then
    GetValTauxRef(TobCond.GetValue('TCN_TAUXREF2'), DateToStr(V_PGI.DateEntree), Taux)
  else if Nombre > TobCond.GetValue('TCN_PLAFOND2') then
    GetValTauxRef(TobCond.GetValue('TCN_TAUXREF3'), DateToStr(V_PGI.DateEntree), Taux);

  Result := Nombre * (Taux /100);
end ;

{---------------------------------------------------------------------------------------}
function calcInteretDEB(nb : integer ; taux : double; base : integer; date : TdateTime) : double ;
{---------------------------------------------------------------------------------------}
var
  Num : Double;
begin
  Result := 0;
  Num := CalcNbJourParAnBase(Date, Base) * 100;
  if Num <> 0 then Result := (Nb * Taux) / Num;
end ;

{---------------------------------------------------------------------------------------}
{--------------------- FIN DE FONCTIONS POUR LE MOMENT INUTILISÉES ---------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
procedure AssignDrapeau(Source : TImage; CodeIso : string) ;
{---------------------------------------------------------------------------------------}
var
  {$IFDEF EAGLCLIENT}
  s   : TStringStream ;
  {$ENDIF}
  SQL : string ;
  Q   : TQuery ;
  w   : string;
begin
  if Trim(CodeIso) = '' then begin
    Source.Picture.Assign(nil);
    Exit;
  end;
  {01/10/04 : Gestion du cas particulier de l'euro : en effet, les pays de l'UE ont tous l'euro
              comme devise => la requête sur PY_DEVISE risque de retourner le drapeau belge, italien ...}
  w := 'PY_DEVISE = "' + CodeIso + '"';
  if CodeIso = V_PGI.DevisePivot then
    w := w + ' AND PY_PAYS = "' + CodeIso + '"';

  {JP 25/07/03 : Maintenant on va chercher les drapeaux dans la table pays}
  SQL := 'SELECT PY_DRAPEAUX  FROM PAYS WHERE ' + w;
  Q:= OpenSQL(SQL, true);
  {$IFDEF EAGLCLIENT}
  if not Q.Eof then begin
    s := TStringStream.Create(Q.Fields[0].AsString);
    s.Seek(0,0);
    Source.Picture.Bitmap.LoadFromStream(s);
    s.Free ;
  end
  else
    Source.Picture.Assign(nil);
  {$ELSE}
  if not Q.Eof then Source.Picture.Assign(Q.FindField('PY_DRAPEAUX'))
               else Source.Picture.Assign(nil);
  {$ENDIF}
  Ferme(Q);
End ;

{Chargement du drapeau et affichage du code devise
{---------------------------------------------------------------------------------------}
procedure MajAffichageDevise(Image, Libelle : TControl; Code : string; Source : TCatSource);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  {Une petite précaution bienvenue !!}
  if not Assigned(Image) and not Assigned(Libelle) then Exit;

  {Si source = sd_Aucun, c'est que l'on passe directement l.e code devise}
  if (Code <> '') and (Source <> sd_Aucun) then begin
    {Récupération de la devise dans la table demandée}
    case Ord(Source) of
      Ord(sd_OPCVM ) : Q := OpenSql('SELECT TOF_DEVISE FROM TROPCVMREF WHERE TOF_CODEOPCVM = "' + Code + '"', True);
      Ord(sd_Compte) : Q := OpenSql('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_CODE = "' + Code + '"', True);
      else
        Exit;
    end;

    if not Q.EOF then
      Code := Q.Fields[0].AsString
    else
      Code := '';
    Ferme(Q);
  end;

  {Mise à jour du label}
  if Assigned(Libelle) then THLabel(Libelle).Caption := Code;
  {Récupération du drapeau}
  if Assigned(Image  ) then AssignDrapeau(TImage(Image), Code);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 02/10/2001
Modifié le ... :   /  /
Description .. : Affiche le message passé en paramètre, stocké dans
Suite ........ : Constantes.pas
Mots clefs ... : MESSAGES
*****************************************************************}
function TrShowMessage(Caption: String; Num: Integer; P1, P2: String): TModalResult;
var	S: String;
begin
  S := IntToStr(Num)+';'+Caption+';'+TrMessage[Num]; // N°;Titre;Message
  Result := HShowMessage(S, P1, P2);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 05/12/2001
Modifié le ... :   /  /    
Description .. : Filtre les caractères saisis pour un code afin de respecter les
Suite ........ : règles d'ergonomie (règle N°7)
Suite ........ : Le caractère est changé en code 0 s'il n'est pas accepté
Mots clefs ... : CODE;TOUCHE;FILTRE
*****************************************************************}
procedure ValidCodeOnKeyPress(var Key: Char);
begin
  if (Key in ['0'..'9', 'A'..'Z', 'a'..'z']) or (Key < ' ') then
  begin	// Caractères de contrôle, lettres majuscules et chiffres seuls autorisés
	if Key >= 'a' then
		Key := Chr(Ord(Key) and $DF); // Passe en majuscule
  end
  else
	Key := #0;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 31/10/2001
Modifié le ... :   /  /
Description .. : Rend le code calendrier à partir du compte général, via
Suite ........ : l'agence
Mots clefs ... : CALENDRIER
*****************************************************************}
function GetCalendrier(CompteGeneral: String) : String; //CodeCalendrier
var Q: TQuery;
begin
  Result := '';
  Q := OpenSQL('SELECT TRA_CODECAL FROM AGENCE LEFT JOIN BANQUECP ON TRA_AGENCE = BQ_AGENCE '+
				'WHERE BQ_CODE = "' + CompteGeneral + '"', True);
  if not Q.Eof then
	Result := Q.Fields[0].AsString;
  Ferme(Q);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 31/10/2001
Modifié le ... :   /  /
Description .. : Rend la date de valeur à partir d'une date et des conditions
Suite ........ : de valeur
Mots clefs ... : VALEUR;GLISSEMENT
*****************************************************************}
function CalcDateValeur(CodeAFB, CompteGeneral: String; aDate: TDateTime): TDateTime;
var
  Q       : TQuery;
  CodeCal : string;
  N       : Integer;

    {-----------------------------------------------------------------------}
    procedure ExecGlissement(TypeGlissement, NbJours : String);
    {-----------------------------------------------------------------------}
    begin
      N := Q.FindField(NbJours).AsInteger;
      if N > 0 then
        case Q.FindField(TypeGlissement).AsInteger of
          0 : Result := Result + N; {Calendaire}
          1 : while N > 0 do begin {Ouvré}
                Result := GetPremierOuvre(Result+1, CodeCal);
                Dec(N);
              end;
          2 : while N > 0 do begin {Ouvrable}
                Result := GetPremierOuvrable(Result+1);
                Dec(N);
              end;
        end
      else if N < 0 then
        case Q.FindField(TypeGlissement).AsInteger of
          0 : Result := Result + N; {Calendaire}
          1 : while N < 0 do begin {Ouvré}
                Result := GetPrecedentOuvre(Result-1, CodeCal);
                Inc(N);
              end;
          2 : while N < 0 do begin {Ouvrable}
                Result := GetPrecedentOuvrable(Result-1);
                Inc(N);
              end;
        end;
    end;

begin
  Result := aDate;

  Q := OpenSQL('SELECT TCV_TYPEGLISSEMENT,TCV_NBJPREMIER,TCV_TYPEPREMIER,TCV_NBJDEUXIEME,TCV_TYPEDEUXIEME '+
               'FROM CONDITIONVAL WHERE TCV_CODECIB = "' + CodeAFB + '" AND TCV_GENERAL = "' + CompteGeneral + '"', True);
  if not Q.Eof then begin
    CodeCal := GetCalendrier(CompteGeneral);
    {Glissement initial}
    N := Q.FindField('TCV_TYPEGLISSEMENT').AsInteger;
    {0 : Pas de report}
    if N > 0 then begin
      Result := GetPremierOuvre(Result, CodeCal); {Premier jour ouvré}
      if N = 2 then {Lendemain du premier jour ouvré}
        Result := Result + 1;
    end;
    {Premier glissement}
    ExecGlissement('TCV_TYPEPREMIER', 'TCV_NBJPREMIER');
    {Deuxieme glissement}
    ExecGlissement('TCV_TYPEDEUXIEME', 'TCV_NBJDEUXIEME');
  end;

  Ferme(Q);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 31/10/2001
Modifié le ... :   /  /
Description .. : Vérifie si une date/Calendrier ouvrée
Mots clefs ... : CALENDRIER;FERME;OUVRE
*****************************************************************}
function DateEstOuvre(aDate: TDateTime; CodeCal: String): Boolean;
var	Week, tmpDate: TDateTime;
	S: String;
	D, M, Y: Word;
begin
  Result := False;
  S := 'Select TJF_DATEFERMEE from JOURFERME where TJF_CODECAL="'+CodeCal+'" and '; // Calendrier concerné
  // Construit date "jour de la semaine"
  Week := CalendarWeekBound + (DayOfWeek(aDate)+5) mod 7; // Dimanche = 1 -> 6, Lundi = 2 -> 0 !
  // Rend l'année neutre
  DecodeDate(aDate, Y, M, D);
  tmpDate := EncodeDate(CalendarClosedYear, M, D);
  // Recherche parmi les dates fermées si existe, ou si jour de la semaine fermé
  if ExisteSQL(S+'(TJF_DATEFERMEE="'+UsDateTime(tmpDate)+'" or TJF_DATEFERMEE="'+UsDateTime(Week)+'")') then
	Exit;

  // Cherche parmi les jours fériés mobiles
  { Optimisation : voir si le décalage est reconnu
  I := Trunc(CalculPaques(Y) - aDate); // Décalage
  if I in [0, 1, 39, 49, 50] then	Voir tablette JFE !
  begin }
	tmpDate := CalendarEasterBound + aDate - CalculPaques(Y);
	if ExisteSQL(S+'TJF_DATEFERMEE="'+UsDateTime(tmpDate)+'"') then
		Exit;

  Result := True; // Ouvré
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 31/10/2001
Modifié le ... :   /  /
Description .. : Rend la prochaine date ouvrée
Mots clefs ... : CALENDRIER;OUVRE
*****************************************************************}
function GetPremierOuvre(aDate: TDateTime; CodeCal: String): TDateTime;
begin	// Pour optimiser, faire une version de DateEstFerme avec Tob
  Result := aDate;
  while not DateEstOuvre(Result, CodeCal) do
	Result := Result+1;
end;

function GetPrecedentOuvre(aDate: TDateTime; CodeCal: String): TDateTime;
begin	// Pour optimiser, faire une version de DateEstFerme avec Tob
  Result := aDate;
  while not DateEstOuvre(Result, CodeCal) do
	Result := Result-1;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 07/11/2001
Modifié le ... :   /  /
Description .. : Vérifie si une date est ouvrable
Suite ........ : Pas dimanche ni jour férié
Mots clefs ... : CALENDRIER;OUVRABLE
*****************************************************************}
function DateEstOuvrable(aDate: TDateTime): Boolean;
var	EasterDate, tmpDate: TDateTime;
	Q: TQuery;
	S: String;
	I: Integer;
	D, M, Y: Word;
begin
  Result := False;
  if DayOfWeek(aDate) = 1 then // Dimanche = 1
	Exit;

  // Cherche parmi les jours fériés (voir TOM_CALENDRIER.LoadJourFerie)
  DecodeDate(aDate, Y, M, D);
  EasterDate := CalculPaques(Y);
  tmpDate := EncodeDate(CalendarClosedYear, M, D); // Rend l'année neutre
  Q := OpenSQL('SELECT CO_LIBRE FROM COMMUN WHERE CO_TYPE = "TJF"' {Ancien type :JFE}, True);
  try
	while not Q.Eof do
	begin
		S := Q.Fields[0].AsString;
		if S[1] = 'P' then	// Date à calculer par rapport à Pâques
		begin	// Extrait le décalage (<200)
			if S = 'P' then
				I := 0
			else	// P+
				I := StrToInt(Copy(S, 3, 9));
			if aDate = (EasterDate+I) then
				Exit;
		end
		else	// Date fixe
			if tmpDate = StrToDate(S) then
				Exit;

		Q.Next;
	end;
  finally
	Ferme(Q);
  end;

  Result := True; // Ouvrable
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 07/11/2001
Modifié le ... :   /  /
Description .. : Rend la prochaine date ouvrable
Mots clefs ... : CALENDRIER;OUVRABLE
*****************************************************************}
function GetPremierOuvrable(aDate: TDateTime): TDateTime;
begin	// Pour optimiser, faire une version avec Tob
  Result := aDate;
  while not DateEstOuvrable(Result) do
	Result := Result+1;
end;

function GetPrecedentOuvrable(aDate: TDateTime): TDateTime;
begin	// Pour optimiser, faire une version avec Tob
  Result := aDate;
  while not DateEstOuvrable(Result) do
	Result := Result-1;
end;

{20/11/06 : Renvoie True, si les deux dates ne sont pas sur le même millésime
{---------------------------------------------------------------------------------------}
function TestDateEtMillesime (NewDt, OldDt : TDateTime;AvecMessage : Boolean = False) : Boolean;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
  y, t, d : Word;
begin
  DecodeDate(NewDt, a, m, j);
  DecodeDate(OldDt, y, t, d);
  Result := a <> y;
  if Result and AvecMessage then
    Result := PGIAsk(TraduireMemoire('La nouvelle date n''appartient pas à la même année que l''ancienne.'#13 +
                                     'Cela va fausser vos soldes de réinitialisation.'#13#13 +
                                     'Souhaitez-vous poursuivre ?')) <> mrYes;
end;

{20/06/07 : Mise en place du concept de création / suppression des écritures}
{---------------------------------------------------------------------------------------}
function CanCreateEcr(Bouton : TComponent = nil; Parle : Boolean = False) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := ExJaiLeDroitConcept(ciCreSupFlux, Parle);
  if Assigned(Bouton) then begin
    {$IFNDEF EAGLSERVER}
         if (Bouton is TControl ) then (Bouton as TControl ).Visible := Result
    else if (Bouton is TMenuItem) then (Bouton as TMenuItem).Visible := Result;
    {$ENDIF EAGLSERVER}
  end;
end;

{20/06/07 : Mise en place du concept de création / suppression des écritures}
{---------------------------------------------------------------------------------------}
function CanValidateBO(Bouton : TComponent = nil; Parle : Boolean = False) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := ExJaiLeDroitConcept(ciValidationBO, Parle);
  if Assigned(Bouton) then begin
    {$IFNDEF EAGLSERVER}
         if (Bouton is TControl ) then (Bouton as TControl ).Visible := Result
    else if (Bouton is TMenuItem) then (Bouton as TMenuItem).Visible := Result;
    {$ENDIF EAGLSERVER}
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Calcul le montant des agios/intérets par rapport au type de
Suite ........ : taux, et du type d'agios/intérets
Mots clefs ... : MONTANT;POSTCOMPTE; PRECOMPTE;AGIOS;INTERET
*****************************************************************}
Function CalcMontantAgios(Montant : Double ; Taux   : Double;
                          Base    : Integer; NbJour : Integer;
                          TauxPreCompte  : Boolean ;
                          AgiosPreCompte : Boolean) : Double ;
var
  vResult : Double;
begin
  vResult := 0;
  Result  := 0;
  Taux := Taux /100 ;
  try
    if not AgiosPreCompte then begin
      if Base = 0 then Exit;
      {Si Agios PostCompté TauxPostCompte}
      vResult := Montant * Taux / Base * NbJour;
    end else begin
      if (nbjour = 0) or (Taux = 0) then Exit;
      {Cas Agios PréCompte TauxPostCompte}
      vResult := Montant / (1 + (Base / (Taux * NbJour)));
    end;
  finally
    Result := vResult;
  end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Calcul le Taux des Agios/intérêt par rapport au type de
Suite ........ : taux, et du type d'agios/intérets
Mots clefs ... : MONTANT;POSTCOMPTE; PRECOMPTE;AGIOS;INTERET
*****************************************************************}
Function CalcTauxAgios(MontantAgios : Double; Base : integer ; Montant : integer ;NbJour : integer; tauxPostCompte : boolean; AgiosPostCompte : Boolean) : double ;
Var vResult : Double ;
begin
	vResult := 0;
	try
	if AgiosPostCompte then
	  // Cas Agios PostCompte TauxPostCompte
	  // -----------------------------------
	  if tauxPostCompte then
		vResult := (montantAgios* base)/(Montant*NbJour)
	  else
	  // Cas Agios PostCompte TauxPréCompte
	  // -----------------------------------
		vResult := (montantAgios* base)/(NbJour*(Montant+MontantAgios))
	else
	  // Cas Agios PréCompte TauxPréCompte
	  // -----------------------------------
	  if Not tauxPostCompte then
        vResult := (MontantAgios*Base)/(Montant*NbJour)
      else
      // Cas Agios PréCompte TauxPostCompte
      // -----------------------------------
        vResult := (montantAgios* base)/(NbJour*(Montant-MontantAgios))
  finally
  	Result := vResult *100;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... : 01/10/2001
Description .. : Calcul le taux pré ou post compté d'un taux avec comme
Suite ........ : crité si c'est un taux postcompé
Mots clefs ... : TAUX PRECOMPTE; TAUX POSTCOMPTE
*****************************************************************}
Function CalcTauxPrePost(TauxDepartPost : boolean; ValTaux : double ;
			NbJour : integer ; Base : integer) : double ;
Begin
	Try
	valTaux := valTaux / 100;
	if TauxDepartPost then
		//Si c'est un taux postcompté on retourne un taux Précompté
		Result := (ValTaux)/(1+(ValTaux  * NbJour / Base)) * 100
	else
		//Si c'est un taux Précompté on retourne un taux Postcompté
		Result := (ValTaux) /(1-(ValTaux * NbJour / Base)) * 100;
	except
	Result := 0 ;
	end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 31/07/2001
Modifié le ... : 19/09/2001
Description .. : Fonction qui génère un N° Transactionnel
Mots clefs ... : GETNUM; NUMERO;TRANSACTION
*****************************************************************}
function GetNum(Module : string; Societe : string; Transac : string) : string ;
var
  Q   : TQuery ;
  SQL : string ;
  Num : integer ;
begin
  Num := 0 ;

  SQL := 'SELECT TCP_COMPTEURTRESO, TCP_MODULECOMPT,' +
             'TCP_NUMCOMPT,TCP_SOCIETECOMPT,TCP_TRANSACCOMPT ' +
             'FROM COMPTEURTRESO WHERE ' ;
  SQL := SQL + 'TCP_MODULECOMPT="'+ Module + '"';
  SQL := SQL + ' AND TCP_SOCIETECOMPT="' + Societe + '"' ;
  SQL := SQL + ' AND TCP_TRANSACCOMPT="' + Transac + '"' ;
  Q := OpenSQL(SQL, True);

  if not Q.Eof then
    Num := Q.Findfield('TCP_NUMCOMPT').Asinteger ;
  Ferme(Q) ;

  {11/08/06 : Gestion du retour à zéro si on atteint la longueur maximale.
              Je laise une marge de 100000 dans l'hypothèse d'un traitement de masse,
              en espérant qu'il n'atteigne pas 100000 écritures importées par dossier
              CODEUNIQUEDEC, concerne les écritures de réinitialisation qui fonctionnent à
              rebours (9999999 - Num) : il n'y a pas besoin de descendre en dessous de 9999000}
  if (Between(Num, 1, 9900000) and (Module <> CODEUNIQUEDEC)) or
     (Between(Num, 1, 1000) and (Module = CODEUNIQUEDEC)) then begin
    Inc(Num) ;
    Result := PadL(IntToStr(Num), '0', 7);
  end
  else {Première fois que l'on crée une transaction sur cette catégorie}
    Result := '0000001';
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 29/10/2001
Modifié le ... :   /  /
Description .. : retourne un numéro de transaction tréso
Mots clefs ... : TRANSACTION;NUMERO;TRESORERIE
*****************************************************************}
function SetNum(Module : string; Societe : string; Transac : string; Num : string) : string ;
var
  SQL : string ;
begin
  SQL := 'UPDATE COMPTEURTRESO SET TCP_NUMCOMPT="' + Num ;
  SQL := SQL + '" WHERE' ;
  SQL := SQL + ' TCP_MODULECOMPT="'+ Module + '"';
  SQL := SQL + ' AND TCP_SOCIETECOMPT="' + Societe + '"' ;
  SQL := SQL + ' AND TCP_TRANSACCOMPT="' + Transac + '"' ;

  {Si aucun enregistrement n'a été traité, c'est que le compteur n'existe pas => on le crée}
  if ExecuteSQL(SQL) <= 0 then begin
    SQL := 'INSERT INTO COMPTEURTRESO (TCP_NUMCOMPT, TCP_MODULECOMPT, TCP_SOCIETECOMPT, TCP_TRANSACCOMPT) ' ;
    SQL := SQL + ' VALUES (' ;
    SQL := SQL + '"' + Num + '", ' ;
    SQL := SQL + '"' + Module + '", ';
    SQL := SQL + '"' + Societe + '", ' ;
    SQL := SQL + '"' + Transac + '")' ;
    ExecuteSQL(SQL);
  end;
end;

{JP 31/07/03 : Recherche de la souche correspondant à un journal
{---------------------------------------------------------------------------------------}
function TrGetSoucheDuJal(Jal : string; Cpt : string = '') : TDblResult;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  w : string;
begin
  if Cpt = '' then
    w := 'WHERE J_JOURNAL = "' + Jal + '"'
  else begin
    {07/08/06 : Avec le Multi Sociétés, on travaille sur BQ_CODE => il faut récupérer BQ_GENERAL}
    Q := OpenSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_CODE = "' + Cpt + '"', True);
    if not Q.EOF then
      Cpt := Q.findField('BQ_GENERAL').AsString;
    Ferme(Q);
    w := 'WHERE J_CONTREPARTIE = "' + Cpt + '"';
  end;

  Q := OpenSQL('SELECT J_COMPTEURNORMAL, J_CONTREPARTIE, J_JOURNAL FROM JOURNAL ' + w, True);
  if not Q.EOF then begin
    Result.RE := Q.Fields[0].AsString;
    Result.RC := Q.Fields[1].AsString;
    Result.RT := Q.Fields[2].AsString;
  end;
  Ferme(Q) ;
end;

{JP 31/07/03 : Découplage de la recherche et de la mise à jour
  Recherche du dernier numéro de pièce pour un journal donné : stockés dans la table SOUCHE (SH)
{---------------------------------------------------------------------------------------}
function TrGetNewNumPiece(Souche, TypeSouche, Societe, Jal : string ; Ecrit : Boolean) : Integer;
{---------------------------------------------------------------------------------------}
var
  Q   : TQuery;
  Num : Longint;
  SQL : string;
  Lib : string;
begin
  Num := 0 ;
  SQL := 'SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="' + TypeSouche + '" AND SH_SOUCHE="'+Souche+'"' ;
  Q   := OpenSQL(SQL,True) ;
  try
    if not Q.EOF then
      Num:=Q.Fields[0].AsInteger
    else begin
      Lib := '';
      if Jal <> '' then begin
        Ferme(Q);
        Q := OpenSQL('SELECT J_LIBELLE FROM JOURNAL WHERE J_JOURNAL = "' + Jal + '"', True);
        if not Q.EOF then Lib := Copy(Q.FindField('J_LIBELLE').AsString, 1, 17);
      end;
      if Lib = '' then Lib := 'SAISIE EN TRESO';

      {Si c'est la première fois qu'on appelle la fonction pour la trésorerie, on crée la souche}
      SQL := 'INSERT INTO SOUCHE (SH_NUMDEPART, SH_TYPE, SH_SOUCHE, SH_LIBELLE, SH_ABREGE, SH_SOCIETE, SH_DATEDEBUT,  ' +
             'SH_DATEFIN, SH_ANALYTIQUE, SH_FERME, SH_NUMDEPARTP, SH_NUMDEPARTS, SH_SIMULATION, SH_SOUCHEEXO) ';

      SQL := SQL + ' VALUES (2, "' + TypeSouche + '", "' + Souche + '", "' + Lib + ' (TRESO)", "' + Lib + '", "' + Societe + '", "' + USDateTime(iDate1900) +
             '", "' + USDateTime(iDate2099) + '", "-", "-", 1, 1, "-", "-")';
      {Si la création ne s'est pas déroulée normalement}
      if ExecuteSQL(SQL) <= 0 then
        V_PGI.IoError := oeUnknown
      else
        Num := 1;
    end;
  finally
    Ferme(Q);
  end;

  if ((V_PGI.IoError <> oeUnknown) and (Ecrit)) then  begin
    {Ecriture du nouveau point de départ}
    if TrSetNumPiece(Souche, TypeSouche, Num + 1) <= 0 then begin
      Num := 0 ;
      V_PGI.IoError := oeSaisie;
    end;
  end;

  Result := Num;
end;

{JP 31/07/03 : MAJ du nouveau compteur pour un journal donné
{---------------------------------------------------------------------------------------}
function TrSetNumPiece(Souche, TypeSouche : string; Num : LongInt) : LongInt;
{---------------------------------------------------------------------------------------}
begin
  Result := ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART = ' + IntToStr(Num) + ' WHERE SH_TYPE = "' + TypeSouche +
                       '" AND SH_SOUCHE = "' + Souche + '"');
end;

{JP 27/11/03 : appelle TrGetNewNumPiece pour une intégration en comptabilité et met à jour TRECRITURE}
{---------------------------------------------------------------------------------------}
function GetSetNumPiece(tEcriture : TOB) : Integer;
{---------------------------------------------------------------------------------------}
var
  sh  : string;
  Num : Integer;
begin
  sh  := TrGetSoucheDuJal(tEcriture.GetValue('E_JOURNAL'), '').RE;
  Num := TrGetNewNumPiece(sh, 'CPT', tEcriture.GetValue('E_SOCIETE'), tEcriture.GetValue('E_JOURNAL'), True);
  Result := Num;
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 01/10/2001
Modifié le ... :   /  /
Description .. : Créer le numéro de transaction initial
Mots clefs ... : TRANSACTION; NUMERO TRANSACTION
*****************************************************************}
function InitNumTransac(Module, Societe, Transaction : string) : string;
var
  Str,
  Num : string;
begin
  Result := '';
  if Length(Transaction) <> 0 then begin
//    Num := GetNum(Module, Societe, Transaction);
    Num := GetNum(Module, V_PGI.CodeSociete, Transaction); {25/09/06 : pour le moment (?!) on force le code société}
    Str := Module + Societe + Transaction + Num;
  end;
  Result := Str;
end;

{17/11/06 : Récupération du numéro de période sur les numéros de transaction pour les écritures de bordereau}
{---------------------------------------------------------------------------------------}
function GetNumPeriodeTransac(NumTransac, Journal : string) : Integer;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  Result := - 1;
  {Si l'on est sur un journal de bordereau}
  if ExisteSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_MODESAISIE IN("BOR", "LIB") AND J_JOURNAL = "' + Journal + '"') then begin
    {Numéro de transaction d'un flux comptable en saisie bordereau / libre : "ICP" + SOC + JAL + PERI + PIEC}
    s := Copy(Numtransac, 10, 4);
    if IsNumeric(s) then Result := 200000 + StrToInt(s);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 07/09/2001
Modifié le ... : 01/10/2001
Description .. : Fonction de retour du nombre de jour d'une  année par
Suite ........ : rapport à la base de calcul
Mots clefs ... : BASE DE CALCUL; NOMBRE JOUR; ANNEE
*****************************************************************}
Function CalcNbJourParAnBase(Date : TDateTime ; BaseCalcul : integer) : integer ;
var
   base : integer ;
   Annee,mm,jj : word ;
begin
  Base := 360;
  try
    DecodeDate(Date, Annee, mm, jj);
    Base := 360;
    case BaseCalcul of
      B30_360      : Base := 360 ;
      B30_365      : Base := 365 ;
      BJUSTE_360   : Base := 360 ;
      BJUSTE_365   : Base := 365 ;
      BJUSTE_JUSTE : begin
                       Base := 365 ;
                       //Vérifie si c'est une année bisextile
                       if IsLeapYear(Annee) then base := 366 ;
                     end ;
      B30E_360     : Base := 360 ;
    end ; //Case
  finally
    Result := base ;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 07/09/2001
Modifié le ... : 19/09/2001
Description .. : Retourne le nombre de jour entre 2 dates par rapport à une
Suite ........ : base de calcul
Mots clefs ... : NOMBRE DE JOUR; BASE DE CALCUL
*****************************************************************}
function CalcNbJourReelBase(DateDeb : TDate ; DateFin : TDate ; BaseCalcul : Integer) : Double ;
var
  jDeb, mDeb,
  aDeb, jFin,
  mFin, aFin : Word;
  Nbj        : Double;
begin
   Nbj := 0 ;
   try
    if (DateDeb <> DateFin) and (BaseCalcul >= 0) then begin
      case BaseCalcul of
        B30_360,
        B30_365,
        B30E_360 : begin
                     DecodeDate(DateDeb, aDeb, mDeb, jDeb);
                     DecodeDate(DateFin, aFin, mfin, jFin);
                     if jDeb >= 30 then jDeb := 30;
                     if jFin >= 30 then jFin := 30;
                     Nbj := ((aFin - aDeb) * 12 + (mFin - mDeb)) * 30 + (jFin - jDeb);
                     {Cas particulier}
                     if BaseCalcul = B30E_360 then begin
                       if mDeb = 2 then
                         case jDeb of
                            28 : Nbj:=Nbj-3 ;
                            29 : Nbj:=Nbj-2 ;
                         end;

                       if mFin = 2 then
                         case jfin of
                            28 : Nbj:=Nbj+2 ;
                            29 : Nbj:=Nbj+1 ;
                         end;
                     end;{if BaseCalcul = B30E_360}
                   end; {Case B30_360, ...}

        BJUSTE_360,
        BJUSTE_365,
        BJUSTE_JUSTE : Nbj := DateFin - DateDeb;
      end; {case}
    end ; {if datedeb = datefin}
  finally
    Result:= Nbj;
  end;
end;

{JP 23/04/04 : Test des droits d'accès à certaines fonctionnalités par popup. Elle évite
               l'ajout du uses HCtrls.
{---------------------------------------------------------------------------------------}
function AutoriseFonction(MenuATester : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := JaiLeDroitTag(MenuATester);
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... :   /  /
Description .. : retourne les conditions de financement et placement
Mots clefs ... :
*****************************************************************}
function GetConditionFinPlac(var Condition : TOB; Transac : string ; Banque : string ; Compte : string ; societe : string) : boolean ;
Var
  Q : TQuery ;
  SQL : string ;
Begin
  Result := false ;
  if not Assigned(Condition) then Exit;
  
  SQL := 'SELECT * FROM CONDITIONFINPLAC WHERE ' ;
  SQL := SQL + ' TCF_CONDITIONFP = "' + Transac + '"' ;
  if Length(Compte) > 0 then
    SQL := SQL + ' AND TCF_GENERAL = "' + compte + '"'
  else
  begin
    SQL := SQL + ' AND TCF_BANQUE = "' + banque + '"' ;
    SQL := SQL + ' AND TCF_SOCIETE = "' + societe + '"' ;
  end ;

  if Length(Transac) > 0 then begin
    Q := OpenSQL(SQL, True);
    if not Q.EOF then
      with  Condition do	begin
	       PutValue('TCF_AGIOSPRECOMPTE',  Q.FindField('TCF_AGIOSPRECOMPTE').AsString);
        PutValue('TCF_AGIOSDEDUIT'   ,  Q.FindField('TCF_AGIOSDEDUIT'   ).AsString);
        PutValue('TCF_TAUXFIXE'      ,  Q.FindField('TCF_TAUXFIXE'      ).AsString);
        PutValue('TCF_BASECALCUL'    ,  Q.FindField('TCF_BASECALCUL'    ).AsString);
        PutValue('TCF_TAUXPRECOMPTE' ,  Q.FindField('TCF_TAUXPRECOMPTE' ).AsString);
        PutValue('TCF_VALTAUX'       ,  Q.FindField('TCF_VALTAUX'       ).AsFloat);
        PutValue('TCF_TAUXVAR'       ,  Q.FindField('TCF_TAUXVAR'       ).AsString);
        PutValue('TCF_VALMAJORATION' ,  Q.FindField('TCF_VALMAJORATION' ).AsFloat);
        PutValue('TCF_VALMULTIPLE'   ,  Q.FindField('TCF_VALMULTIPLE'   ).AsFloat);
        PutValue('TCF_NBJOURBANQUE'  ,  Q.FindField('TCF_NBJOURBANQUE'  ).AsFloat);
        PutValue('TCF_NBJOURENCAISS' ,  Q.FindField('TCF_NBJOURENCAISS' ).AsFloat);
        PutValue('TCF_NBJOURMINAGIOS',  Q.FindField('TCF_NBJOURMINAGIOS').AsFloat);
        Result := True;
      end;
    Ferme(Q);
  end ;
End ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... : 06/03/2002
Description .. : retourne une TOB remplie des conditions de découverts
Suite ........ : correspondant à la banque et au flux demandé ...
Mots clefs ... : CONDITION DECOUVERT
*****************************************************************}
function GetConditionDec (Flux : string ; Compte : string; var TobC : TOB ) : boolean ;
var
  SQL : String ;
  Q : TQuery ;
  TobL: TOB;
begin
  SQL := 'SELECT * FROM CONDITIONDEC WHERE TCN_FLUX="'+ Flux + '"' ;
  SQL := SQL + ' AND TCN_GENERAL = "' + Compte + '"' ;
  Q := OpenSQL(SQL,true) ;
  if not Q.EOF then 
    TobC.LoadDetailDB( 'cond', '', '', Q, False, True)
  else begin
    TobL := TOB.Create('', TobC, -1);
    TobL.AddChampSupValeur('TCN_FLUX',flux);
    TobL.AddChampSupValeur('TCN_GENERAL', Compte);
    TobL.AddChampSupValeur('TCN_AGENCE', '');
    TobL.AddChampSupValeur('TCN_SOCIETE','');
    TobL.AddChampSupValeur('TCN_AUTORISATION', 0);
    TobL.AddChampSupValeur('TCN_TYPECALCFRAIS',0);
    TobL.AddChampSupValeur('TCN_MAJOTAUX1',0);
    TobL.AddChampSupValeur('TCN_MAJOTAUX2',0);
    TobL.AddChampSupValeur('TCN_MAJOTAUX3',0);
    TobL.AddChampSupValeur('TCN_MULTAUX1',0);
    TobL.AddChampSupValeur('TCN_MULTAUX2',0);
    TobL.AddChampSupValeur('TCN_MULTAUX3',0);
    TobL.AddChampSupValeur('TCN_TAUXREF1',0);
    TobL.AddChampSupValeur('TCN_TAUXREF2',0);
    TobL.AddChampSupValeur('TCN_TAUXREF3',0);
    TobL.AddChampSupValeur('TCN_PLAFOND1',0);
    TobL.AddChampSupValeur('TCN_PLAFOND2',0);
    TobL.AddChampSupValeur('TCN_BASECALCUL',0);
    TobL.AddChampSupValeur('TCN_DATECONTRAT',iDate1900); {18/05/06 : FQ 10340}
    TobL.AddChampSupValeur('TCN_PERIODE',0);
    TobL.AddChampSupValeur('TCN_CALCULSOLDE',0);
    TobL.AddChampSupValeur('TCN_SOLDEVALEUR',0);
    TobL.AddChampSupValeur('TCN_NBJOUR',0);
    TobL.AddChampSupValeur('TCN_TYPEPREMIER',0);
    TobL.AddChampSupValeur('TCN_CPFD',0);
    TobL.AddChampSupValeur('TCN_PLAFONNEE',0);
    TobL.AddChampSupValeur('TCN_LIEAUTO','-');
    TobL.AddChampSupValeur('TCN_PLAFONDCPFD',0);
  end ;

  Ferme(Q);
  result := true ;
End ;



{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 29/10/2001
Modifié le ... :   /  /
Description .. : Retourne la devise d'un compte général
Mots clefs ... : DEVISE, COMPTE, GENERAL
*****************************************************************}
Function RetDeviseCompte(CompteGeneral : string) : string ;
var
  SQL : string ;
  Q   : TQuery ;
begin
  Result := '';
  SQL := 'SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_CODE = "';
  SQL := SQL + CompteGeneral + '"';
  Q   := OpenSQL(SQL, True);
  if not Q.EOF then Result := Q.FindField('BQ_DEVISE').AsString;
  Ferme(Q);
end ;

{---------------------------------------------------------------------------------------}
function RetDeviseOPCVM(OPCVM : string) : string ;
{---------------------------------------------------------------------------------------}
var
  SQL : string ;
  Q   : TQuery ;
begin
  Result := V_PGI.DevisePivot;
  SQL := 'SELECT TOF_DEVISE FROM TROPCVMREF WHERE TOF_CODEOPCVM ="' ;
  SQL := SQL + OPCVM + '"';
  Q   := OpenSQL(SQL, True);
  if not Q.EOF then
    Result := Q.FindField('TOF_DEVISE').AsString;
  Ferme(Q);
end ;

{---------------------------------------------------------------------------------------}
function GetCompteAgence(Compte : string) : string ;
{---------------------------------------------------------------------------------------}
var
  SQL : string ;
  Q   : TQuery ;
begin
  Result := '';
  SQL := 'SELECT BQ_AGENCE FROM BANQUECP WHERE BQ_CODE = "' ;
  SQL := SQL + Compte + '"';
  Q   := OpenSQL(SQL, True);
  if not Q.EOF then Result := Q.Fields[0].AsString;
  Ferme(Q);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. : Retourne le nb de décimale d'une devise donnée
Mots clefs ... : DEVICE;DECIMALE
*****************************************************************}
function CalcDecimaleDevise(Devise  : string) : integer ;
var
  SQL : string ;
  Q : TQuery ;
begin
  Result := V_PGI.OkDecV;
  if Devise = V_PGI.DevisePivot then Exit;
  
  SQL := 'SELECT D_DECIMALE FROM DEVISE WHERE D_DEVISE="' ;
  SQL := SQL + Devise + '"';
  Q   := OpenSQL(SQL, True);
  if not Q.EOF then result := Q.FindField('D_DECIMALE').AsInteger;
  Ferme(Q);
end ;

{JP 25/04/07 : Recherche des décimales depuis les devise
{---------------------------------------------------------------------------------------}
function CalcDeciDeviseIso(DevIso  : string) : Integer;
{---------------------------------------------------------------------------------------}
var
  SQL : string ;
  Q : TQuery ;
begin
  Result := V_PGI.OkDecV;
  SQL := 'SELECT D_DECIMALE FROM DEVISE WHERE D_CODEISO = "' ;
  SQL := SQL + DevIso + '"';
  Q   := OpenSQL(SQL, True);
  if not Q.EOF then Result := Q.FindField('D_DECIMALE').AsInteger;
  Ferme(Q);
end;

{JP 21/08/03 : récupère la quotité et le nombre de décimales d'une devise
{---------------------------------------------------------------------------------------}
function GetInfoDevise(Dev : string) : TDblResult;
{---------------------------------------------------------------------------------------}
var
  SQL : string ;
  Q : TQuery ;
begin
  Result.RE := '2';
  Result.RC := '1';
  SQL := 'SELECT D_DECIMALE, D_QUOTITE FROM DEVISE WHERE D_DEVISE="' ;
  SQL := SQL + Dev + '"';
  Q := OpenSQL(SQL, True);
  if not Q.EOF then begin
    Result.RE := Q.Fields[0].AsString;
    Result.RC := Q.Fields[1].AsString;
  end;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. : Retourne le cours contre l'Euro, le plus proche en date,
Suite ........ : d'une devise à une date donnée
Mots clefs ... : DATE;COURS;DEVISE;PARITE;EURO
*****************************************************************}
function RetPariteEuro(Devise : string; DateC : TDateTime; CotationOk : Boolean = False) : Double;
var
  SQL  : string;
  Q    : TQuery;
  Quot : Double;
  {$IFNDEF EAGLSERVER}
  Deci : Integer;{03/12/04}
  {$ENDIF !EAGLSERVER}
  Chp  : string;{03/12/04}
begin
  Result := 1 ;
  {JP 03/12/2004 : A force de conserver cette fonction qui fait le contraire du reste de la
                   Trésorerie, je finis par m'y perdre. J'ajoute donc le paramètre CotationOK}
  if Devise <> V_PGI.DevisePivot then begin
    {03/12/04 : Pendant que j'y suis, gestion du format du résultat}
    SQL := 'SELECT D_QUOTITE, D_DECIMALE FROM DEVISE WHERE D_DEVISE = "' + Devise + '"';
    Q 	:= OpenSQL(SQL,TRue);
    if not Q.EOF then begin
      Quot := Q.FindField('D_QUOTITE').AsFloat;
      {$IFNDEF EAGLSERVER}
      Deci := Q.FindField('D_DECIMALE').AsInteger;
      {$ENDIF EAGLSERVER}
      Ferme(Q) ;
      {$IFNDEF EAGLSERVER}
      {23/05/05 : on mémorise le nombre de décimales, pour uniformiser le formatage des montants dans TRECRITURE}
      if Assigned(TheData) and (TheData is TObjDetailDevise) then
        TObjDetailDevise(TheData).NbDecimal := Deci;
      {$ENDIF EAGLSERVER}

      if CotationOk then Chp := 'H_COTATION'
                    else Chp := 'H_TAUXREEL';

      SQL := 'SELECT H_DATECOURS, H_DEVISE, ' + Chp + ' FROM CHANCELL ';
      SQL := SQL + 'WHERE H_DEVISE = "' + Devise + '" ';
      SQL := SQL + 'AND H_DATECOURS = (SELECT MAX(H_DATECOURS) FROM CHANCELL ';
      SQL := SQL + 'where H_DATECOURS <= "' + UsDateTime(DateC) +'" ' ;
      SQL := SQL + 'AND H_DEVISE = "' + Devise + '" )';
      Q := OpenSQL(SQL,true);
      {On récupère le taux pour un € et on divise par la quotité}
      if Quot = 0 then Quot := 1;
      if not Q.EOF then Result := Q.FindField(Chp).AsFloat;
      {03/12/04 : Gestion de la quotité}
      if CotationOk then Result := Result * Quot
                    else Result := Result / Quot;
      {03/12/04 : Formatage du Résultat}
      Result := Valeur(FormateMontant(Result, NBDECIMALTAUX));
      {Par précaution !!!}
      if Result = 0 then Result := 1;
    end;
    Ferme(Q);
  end
  {$IFNDEF EAGLSERVER}
  else
    {20/07/05 : c'est en corrigeant la FQ 10273 que je me suis aperçu de ce gros oubli de ma part =>
                Tout les montants en devise pivot étaient arrondis à l'unité : NbDecimal vaut 0 par défaut !!!}
    if Assigned(TheData) and (TheData is TObjDetailDevise) then
      TObjDetailDevise(TheData).NbDecimal := V_PGI.OkDecV;
  {$ENDIF EAGLSERVER}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 25/01/2002
Modifié le ... :   /  /
Description .. : Donne le code flux à partir du code transaction
Mots clefs ... : FLUX;TRANSACTION
*****************************************************************}
function GetCodeFlux(CodeTransac, TypeFlux: String): String;
var	Q : TQuery;
begin
	Result := '';
	Q := OpenSQL('Select TTR_'+TypeFlux+' from TRANSAC where TTR_TRANSAC="'+CodeTransac+'"', True);
	if not Q.EOF Then
		Result := Q.Fields[0].AsString;
	Ferme(Q);

end;

{Récupère le libéllé à partir d'un code dans la vue TRLISTEFLUX qui est une union entre RUBRIQUE et FLUXTRESO
{---------------------------------------------------------------------------------------}
function GetLibFluxRub(Code : string) : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT TFT_LIBELLE FROM TRLISTEFLUX WHERE TFT_FLUX = "' + code + '"', True);
  if not Q.EOF then
    Result := Q.Fields[0].AsString;
  Ferme(Q);
end;

{Récupère un enregistrement de la table FluxTreso à partir d'un code transaction et d'une "Jambe"
{---------------------------------------------------------------------------------------}
function GetFluxCibFromTransac(Transac, Jambe : string; AvecRecherche : Boolean = False) : TOB;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Q   : TQuery;
begin
  Result := TOB.Create('FLUXTRESO', nil, -1);
  SQL := 'SELECT * FROM FLUXTRESO WHERE TFT_FLUX = (' ;
  SQL := SQL + 'SELECT TTR_' + Jambe + ' FROM TRANSAC WHERE TTR_TRANSAC = ';
  if AvecRecherche then
    SQL := SQL + '(SELECT TOP_TRANSACTION FROM TROPCVM WHERE TOP_NUMOPCVM = "' + Transac + '"))'
  else
    SQL := SQL + '"' + Transac + '")';
  Q := OpenSQL(SQL, True);
  try
    Result.SelectDB('', Q);
  finally
    Ferme(Q);
  end;
end;

{05/01/05 : Chargement des champs de la table FluxTreso dans une Tob TRECRITURE
{---------------------------------------------------------------------------------------}
procedure ChargeChpFluxTreso(var T : TOB; Transac, Jambe : string; AvecRecherche : Boolean = False);
{---------------------------------------------------------------------------------------}
var
  F : TOB;
begin
  {Récupération des informations depuis la table FluxTreso : Le flux est celui paramétré comme
   étant celui de versement dans la table Transac}
  F := GetFluxCibFromTransac(Transac, Jambe, AvecRecherche);
  try
    T.SetString('TE_CODEFLUX', F.GetString('TFT_FLUX'));
    T.SetString('TE_CONTREPARTIETR', F.GetString('TFT_GENERAL'));
    T.SetString('TE_CODECIB', F.GetString('TFT_CODECIB'));
  finally
    FreeAndNil(F);
  end;
end;

{Récupère la cours d'une opcvm à la date donnée
{---------------------------------------------------------------------------------------}
function GetCoursOpcvm(aCode : string; aDate : TDateTime) : Double;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Result := 0;
  Q := OpenSQL('SELECT TTO_COTATION FROM TRCOTATIONOPCVM WHERE TTO_CODEOPCVM = "' + aCode + '" AND TTO_DATE = ' +
               '(SELECT MAX(TTO_DATE) FROM TRCOTATIONOPCVM WHERE TTO_CODEOPCVM = "' + aCode + '" AND ' +
               'TTO_DATE <= "' + UsDateTime(aDate) + '")', True);
  if not Q.EOF then
    Result := Q.FindField('TTO_COTATION').AsFloat;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 26/12/2001
Modifié le ... :   /  /
Description .. : Retourne le code AFB d'une transaction
Mots clefs ... : TRANSACTION;AFB
*****************************************************************}
function GetCodeAFB(transac : string; TypeFlux : string) : string ;
Var SQL : string ;
	Q : TQuery ;
begin
			result := '' ;
        SQL := 'SELECT TFT_CODECIB,TFT_FLUX FROM FLUXTRESO WHERE TFT_FLUX=(' ;
        SQL := SQL + 'SELECT TTR_'+TypeFlux +' FROM TRANSAC WHERE TTR_TRANSAC="' + Transac + '")';
        Q:=OpenSQL(SQL, True);
        if not Q.EOF Then
        begin
        	result := Q.FindField('TFT_CODECIB').AsString;
				end ;
		Ferme(Q);

end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... :   /  /
Description .. : retourne le sens d'une transac
Mots clefs ... :
*****************************************************************}
function GetSensFlux(Transac : string; TypeFlux : string) : string ;
var
  SQL : string ;
  Q   : TQuery ;
begin
  Result := '' ;

  SQL := 'SELECT TTL_SENS FROM TYPEFLUX WHERE TTL_TYPEFLUX=(' ;
  SQL := SQL + 'SELECT TFT_TYPEFLUX FROM FLUXTRESO WHERE TFT_FLUX=(' +
                  'SELECT TTR_'+TypeFlux+ ' FROM TRANSAC WHERE TTR_TRANSAC="' + Transac + '"))';

  Q := OpenSQL(SQL, True);
  if not Q.EOF then
    Result := Q.FindField('TTL_SENS').AsString;      //D=DEBIT;C=CREDIT

  Ferme(Q);
end;

{JP 17/09/03 : Récupère le sens à partir du type flux
{---------------------------------------------------------------------------------------}
function GetSensFromTypeFlux(TypeFlux : string) : string ;
{---------------------------------------------------------------------------------------}
var
  SQL : string ;
  Q   : TQuery ;
begin
  SQL := 'SELECT TTL_SENS FROM TYPEFLUX WHERE TTL_TYPEFLUX = "' + TypeFlux + '"';
  Q := OpenSQL(SQL, True);
  if not Q.EOF then
    Result := Q.FindField('TTL_SENS').AsString; {D=DEBIT;C=CREDIT}
  Ferme(Q);
end;

{Récupération du sens de la rubrique
{---------------------------------------------------------------------------------------}
function GetSensFromRubrique(CodeRub  : string) : string;
{---------------------------------------------------------------------------------------}
var
  SQL : string ;
  Q   : TQuery ;
begin
  Result := 'C';
  SQL := 'SELECT RB_SIGNERUB FROM RUBRIQUE WHERE RB_RUBRIQUE = "' + CodeRub + '" ' +
         'AND RB_CLASSERUB = "TRE" AND RB_NATRUB = "TRE"';
  Q := OpenSQL(SQL, True);
  if not Q.EOF then
    Result := Q.Fields[0].AsString; {D=DEBIT;C=CREDIT}
  Ferme(Q);
end;

{JP 27/08/03 : Récupération du compte général de contrepartie lors de la génération d'une
               écriture de trésorerie à partie du code de transaction d'une opération
{---------------------------------------------------------------------------------------}
function GetContrepartie(Transac : string; TypeFlux : string) : string ;
{---------------------------------------------------------------------------------------}
var
  SQL : string ;
  Q : TQuery ;
begin
  Result := '';

  SQL := 'SELECT TFT_GENERAL FROM FLUXTRESO WHERE TFT_FLUX=(' +
            'SELECT TTR_' + TypeFlux + ' FROM TRANSAC WHERE TTR_TRANSAC = "' + Transac + '")';

  Q := OpenSQL(SQL, True);
  if not Q.EOF then
    Result := Q.Fields[0].AsString;

  Ferme(Q);
end;

{JP 17/09/03 : Récupération du compte général de contrepartie, du sens et du code CIB à partir d'un code flux}
{---------------------------------------------------------------------------------------}
procedure GetCibSensGeneral(var Gen, Sens, CIB : string; Flux : string);
{---------------------------------------------------------------------------------------}
var
  SQL : string ;
  Q : TQuery ;
begin
  Gen  := '';
  CIB  := '';
  Sens := '';

  SQL := 'SELECT TFT_GENERAL, TFT_CODECIB, TFT_TYPEFLUX FROM FLUXTRESO WHERE TFT_FLUX = "' + Flux + '"';

  Q := OpenSQL(SQL, True);
  try
    if not Q.EOF then begin
      Gen  := Q.Fields[0].AsString;
      CIB  := Q.Fields[1].AsString;
      Sens := Q.Fields[2].AsString;
    end;
  finally
    Ferme(Q);
  end;
end;

{JP 27/11/03 : Récupération du mode de règlement à partir d'un sens, d'un CIB et d'un compte général
{---------------------------------------------------------------------------------------}
function GetModeReglement(Sens, CIB, Cpte : string; Dossier : string = '') : string;
{---------------------------------------------------------------------------------------}
var
  QQ : TQuery;
  CH : string;
begin
  Result := '';

  {18/10/06 :
   05/04/07 : FQ 10429 : Impossible de me souvenir pourquoi j'ai mis cette hérésie ici !!!
   Pour les comptes courants et les comptes titres, on prend le mode de règlement par défaut
  if (CIB = CODECIBTITRE) or (CIB = CODECIBCOURANT) then begin
    Result := GetParamSocDossierSecur('SO_GCMODEREGLEDEFAUT', '@@@', Dossier);
    Exit;
  end;}

  CH := 'SELECT TCI_MODEPAIE, TCI_SENS, BQ_CODE FROM CIB, BANQUECP WHERE BQ_BANQUE = TCI_BANQUE AND ' +
        'BQ_CODE = "' + Cpte + '" AND TCI_CODECIB = "' + CIB + '" AND TCI_SENS IN ("MIX", "' + Sens + '")';
  QQ := OpenSQL(CH, True);

  try
    if QQ.RecordCount = 1 then
      Result := QQ.Fields[0].AsString
    else if QQ.RecordCount = 0 then
      Result := '@@@'
    else begin
      {On cherche à récupérer le règlement correspondant au sens}
      while not QQ.EOF do begin
        if QQ.Fields[1].AsString = Sens then begin
          Result := QQ.Fields[0].AsString;
          Break;
        end;
        QQ.Next;
      end;
      {Si on n'a pas trouver le bon mode de règlement, on prend un mixte}
      if Result = '' then Result := QQ.Fields[0].AsString;
    end;
  finally
    Ferme(QQ);
  end;

  {18/10/06 : Pour les comptes courants et les comptes titres, on prend le mode de
              règlement par défaut, si le cib n'a pas été paramétré
  if ((CIB = CODECIBTITRE) or (CIB = CODECIBCOURANT)) and (Result = '@@@') then
    Result := GetParamSocDossierSecur('SO_GCMODEREGLEDEFAUT', '@@@', Dossier);}

end;

{10/01/07 : Fonction inverse de GetModeReglement}
{---------------------------------------------------------------------------------------}
function GetCodeCIB(Sens, Paie, Cpte : string; Dossier : string = '') : string;
{---------------------------------------------------------------------------------------}
var
  QQ : TQuery;
  CH : string;
begin
  Result := '';
  if Dossier = '' then
    CH := 'SELECT TCI_CODECIB, TCI_SENS FROM CIB, BANQUECP WHERE BQ_BANQUE = TCI_BANQUE AND ' +
          'BQ_CODE = "' + Cpte + '" AND TCI_MODEPAIE = "' + Paie + '" AND TCI_SENS IN ("MIX", "' + Sens + '")'
  else
    CH := 'SELECT TCI_CODECIB, TCI_SENS FROM CIB, BANQUECP WHERE BQ_BANQUE = TCI_BANQUE AND BQ_GENERAL = "' + Cpte + '" AND ' +
          'BQ_NODOSSIER = "' + Dossier + '" AND TCI_MODEPAIE = "' + Paie + '" AND TCI_SENS IN ("MIX", "' + Sens + '")';

  QQ := OpenSQL(CH, True);

  try
    if QQ.RecordCount = 1 then
      Result := QQ.Fields[0].AsString
    else if QQ.RecordCount = 0 then begin
      if ctxTreso in V_PGI.PGIContexte then Result := '@@@'
                                       else Result := '';
    end
    else begin
      {On cherche à récupérer le règlement correspondant au sens}
      while not QQ.EOF do begin
        if QQ.Fields[1].AsString = Sens then begin
          Result := QQ.Fields[0].AsString;
          Break;
        end;
        QQ.Next;
      end;
      {Si on n'a pas trouver le bon mode de règlement, on prend un mixte}
      if Result = '' then Result := QQ.Fields[0].AsString;
    end;
  finally
    Ferme(QQ);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 06/03/2002
Modifié le ... :   /  /
Description .. : retourne le cours d'un taux de référence à un date
Mots clefs ... :
*****************************************************************}
function GetValTauxRef(TauxRef : string; Le : string; var Cotation : Double) : Boolean ;
Var
  SQL : string;
  Q   : TQuery;
begin
  Result := False;
  (* SQL:= 'SELECT MAx(TTA_COTATION) COTATION, Max(TTA_DATE), max(TTA_CODE) FROM COTATIONTAUX WHERE TTA_CODE="'+TauxRef+'"' +
                  ' AND TTA_DATE<="' + USDATETIME(strtodatetime(Le)) + '"' ; *)
  {Récupération du taux le plus récent par rapport à la date donnée}
  SQL := 'SELECT TTA_COTATION FROM COTATIONTAUX WHERE TTA_CODE = "' + TauxRef + '"' +
         ' AND TTA_DATE <= "' + USDATETIME(StrToDateTime(Le)) + '" ORDER BY TTA_DATE DESC';

  Q := OpenSQL(SQL, True);
  if not Q.EOF then begin
    Cotation := Q.Fields[0].AsFloat;
    Result := True ;
  end
  else
    Cotation := 1;

  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 08/01/2002
Modifié le ... :   /  /
Description .. : Retourne dans Banque, Guichet, Num, cle
Suite ........ : les valeurs du RIB correspond
Suite ........ : ET
Suite ........ : retourne dans result la concaténation du RIB ENTIER
Mots clefs ... : RIB;BANQUE;GUICHET;COMPTE;NUMERO;CLE;CLERIB
*****************************************************************}
function GetRIB(var Banque, Guichet, Num, Cle, Iban : string ; CompteGeneral : string) : string ;
var
  SQL : string;
  Q   : TQuery;
begin
  Result := '' ;
  {JP 06/10/2003 : On va chercher l'établissement bancaire dans la table des banques car s'il s'agit d'un
                   compte étranger "BQ_ETABBQ" n'est pas renseigné. Dans la Tréso, on utilise ce champ pour
                   les regroupements
   JP 23/04/04 : Ajout de la recherche de l'Iban}
  SQL := 'SELECT BQ_PAYS, BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, PQ_ETABBQ, BQ_CODEIBAN FROM BANQUECP ' +
         'LEFT JOIN BANQUES ON PQ_BANQUE = BQ_BANQUE ' +
         'WHERE BQ_CODE = "' + CompteGeneral + '" AND PQ_BANQUE IS NOT NULL';
  Q:=OpenSQL(SQL, True);
  if not Q.EOF then begin
    Banque := Q.FindField('BQ_ETABBQ').AsString;
    Guichet := Q.FindField('BQ_GUICHET').AsString ;
    Num := Q.FindField('BQ_NUMEROCOMPTE').AsString ;
    Cle := Q.FindField('BQ_CLERIB').AsString;
    Iban := Q.FindField('BQ_CODEIBAN').AsString;
    Result := Banque ;
    Result := result + Guichet;
    Result := result + Num ;
    Result := result + cle;
    if Trim(Banque) = '' then Banque := Q.FindField('PQ_ETABBQ').AsString;
  end;
  Ferme(Q);
end;

{Retourne le Rib et remplit les champs de la table TRECRITURE concernant les informations bancaires
{---------------------------------------------------------------------------------------}
function GetRIBTob(var tEcr : TOB) : string;
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Q   : TQuery;
begin
  Result := '' ;
  {JP 06/10/2003 : On va chercher l'établissement bancaire dans la table des banques car s'il s'agit d'un
                   compte étranger "BQ_ETABBQ" n'est pas renseigné. Dans la Tréso, on utilise ce champ pour
                   les regroupements
   JP 23/04/04 : Ajout de la recherche de l'Iban}
  SQL := 'SELECT BQ_PAYS, BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, PQ_ETABBQ, BQ_CODEIBAN FROM BANQUECP ' +
         'LEFT JOIN BANQUES ON PQ_BANQUE = BQ_BANQUE ' +
         'WHERE BQ_CODE = "' + tEcr.GetString('TE_GENERAL') + '" AND PQ_BANQUE IS NOT NULL';
  Q := OpenSQL(SQL, True);
  if not Q.EOF then begin
    tEcr.SetString('TE_CODEBANQUE' , Q.FindField('BQ_ETABBQ').AsString);
    tEcr.SetString('TE_CODEGUICHET', Q.FindField('BQ_GUICHET').AsString);
    tEcr.SetString('TE_NUMCOMPTE'  , Q.FindField('BQ_NUMEROCOMPTE').AsString);
    tEcr.SetString('TE_CLERIB'     , Q.FindField('BQ_CLERIB').AsString);
    tEcr.SetString('TE_IBAN'       , Q.FindField('BQ_CODEIBAN').AsString);
    if tEcr.GetString('TE_CODEBANQUE') = '' then tEcr.SetString('TE_CODEBANQUE', Q.FindField('PQ_ETABBQ').AsString);
    Result := tEcr.GetString('TE_CODEBANQUE') ;
    Result := Result + tEcr.GetString('TE_CODEGUICHET');
    Result := Result + tEcr.GetString('TE_NUMCOMPTE');
    Result := Result + tEcr.GetString('TE_CLERIB');
  end;
  Ferme(Q);
end;

{---------------------------------------------------------------------------------------}
function GetBanqueCourant : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT PQ_BANQUE FROM BANQUES WHERE PQ_ETABBQ = "' + CODEATTENTE5 + '"', True);
  try
    if not Q.EOF then
      Result := Q.FindField('PQ_BANQUE').AsString
    else begin
      ExecuteSql('INSERT INTO BANQUES (PQ_BANQUE, PQ_ETABBQ, PQ_LIBELLE, PQ_ABREGE) VALUES ' +
                 '("' + CODECOURANTS + '", "' + CODEATTENTE5 + '", "' + TraduireMemoire('Banque comptes courants')  +
                 '", "' + TraduireMemoire('Banque C/C') + '")');
      Result := CODECOURANTS;
    end;
  finally
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
function GetAgenceCourant : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT TRA_AGENCE FROM AGENCE WHERE TRA_GUICHET = "' + CODEATTENTE5 + '"', True);

  try
    if not Q.EOF then
      Result := Q.FindField('TRA_AGENCE').AsString
    else begin
      ExecuteSql('INSERT INTO AGENCE (TRA_AGENCE, TRA_ETABBQ, TRA_GUICHET, TRA_LIBELLE, TRA_ABREGE, TRA_BANQUE) ' +
                 'VALUES ("' + CODECOURANTS + '", "' + CODEATTENTE5 + '", "' + CODEATTENTE5 + '", "' + TraduireMemoire('Agence comptes courants')  +
                 '", "' + TraduireMemoire('Agence C/C') + '", "' + GetBanqueCourant + '")');
      Result := CODECOURANTS;
    end;
  finally
    Ferme(Q);
  end;
end;

{Récupération d'un Cib "factice" pour les comptes courants et titres
{---------------------------------------------------------------------------------------}
function GetCibCourant(BqCode : string = '') : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  b : string;
begin
  Result := '';

  if BqCode <> '' then begin
    Q := OpenSQL('SELECT BQ_CODECIB FROM BANQUECP WHERE BQ_CODE = "' + BqCode + '"', True);
    if not Q.EOF then
      Result := Q.FindField('BQ_CODECIB').AsString;
    Ferme(Q);
  end;

  if Result = '' then begin
    b := GetBanqueCourant;
    Q := OpenSQL('SELECT TCI_CODECIB FROM CIB WHERE TCI_BANQUE IN ("' + b + '", "' + CODECIBREF + '") ' +
                 'AND TCI_SENS = "MIX" AND TCI_CODECIB = "' + CODECIBCOURANT + '"', True);
    try
      if Q.EOF then begin
        ExecuteSql('INSERT INTO CIB (TCI_ABREGE, TCI_BANQUE, TCI_CODECIB, TCI_LIBELLE, TCI_MODEPAIE, TCI_SENS, TCI_PREDEFINI) ' +
                   'VALUES ("Courant", "' + CODECIBREF + '", "' + CODECIBCOURANT + '", "Courant", "VIR", "MIX", "STD")');
        ExecuteSql('INSERT INTO CIB (TCI_ABREGE, TCI_BANQUE, TCI_CODECIB, TCI_LIBELLE, TCI_MODEPAIE, TCI_SENS, TCI_PREDEFINI) ' +
                   'VALUES ("Courant", "' + b + '", "' + CODECIBCOURANT + '", "Courant", "VIR", "MIX", "STD")');
      end;
      Result := CODECIBCOURANT;
    finally
      Ferme(Q);
    end;
  end;
end;

{Donne la correspondance entre NoDossier, NomBase et Societe à partir de la table DOSSIER
 ChpSel : champ dont on possède la valeur
 ValSel : valeur de ChpSel
 ChpRes : Champ dont on veut obtenir la valeur
{---------------------------------------------------------------------------------------}
function GetInfosFromDossier(ChpSel, ValSel, ChpRes : string) : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Result := '';
  Q := OpenSQL('SELECT ' + ChpRes + ' FROM DOSSIER WHERE ' + ChpSel + '= "' + ValSel + '"', True);
  if not Q.EOF then
    Result := Q.FindField(ChpRes).AsString;
  Ferme(Q);
end;

{Retourne le dossier à partir de BQ_Code (alias le champ _GENERAL)
{---------------------------------------------------------------------------------------}
function GetDossierFromBQCP(BQ_Code : string) : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  if not EstMultiSoc then
    Result := V_PGI.NoDossier
  else begin
    Q := OpenSQL('SELECT BQ_NODOSSIER FROM BANQUECP WHERE BQ_CODE = "' + BQ_Code + '"', True);
    if not Q.EOF then
      Result := Q.FindField('BQ_NODOSSIER').AsString;
    Ferme(Q);
  end;
  if Result = '' then begin
    if V_PGI.SAV then PGIBox(TraduireMemoire('Impossible de récupérer le code dossier du compte.'));
    Result := CODEDOSSIERDEF;
  end;
end;

{Retourne le libellé du dossier à partir de BQ_Code (alias le champ _GENERAL)
{---------------------------------------------------------------------------------------}
function GetLibDossierFromBQ(BQ_Code : string) : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  if not IsTresoMultiSoc then
    Result := VarToStr(GetParamSocSecur('SO_LIBELLE', ''))
  else begin
    Q := OpenSQL('SELECT BQ_NODOSSIER, DOS_LIBELLE FROM BANQUECP ' +
                 'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
                 'WHERE BQ_CODE = "' + BQ_Code + '"', True);
    if not Q.EOF then
      Result := Q.FindField('DOS_LIBELLE').AsString;
    Ferme(Q);
  end;
end;

{Renvoie la base, le NoDossier et la société depuis un compte
{---------------------------------------------------------------------------------------}
function GetInfosSocFromBQ(BQ_Code : string; LibOk : Boolean) : TDblResult;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  if BQ_Code = '' then Exit;
  
  if IsTresoMultiSoc then begin
    Q := OpenSQL('SELECT BQ_NODOSSIER, DOS_NOMBASE, DOS_SOCIETE, DOS_LIBELLE FROM BANQUECP ' +
                 'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
                 'WHERE BQ_CODE = "' + BQ_Code + '"', True);
    if not Q.EOF then begin
      Result.RE := Q.FindField('BQ_NODOSSIER').AsString;
      Result.RC := Q.FindField('DOS_SOCIETE').AsString;
      if LibOk then Result.RT := Q.FindField('DOS_LIBELLE').AsString
               else Result.RT := Q.FindField('DOS_NOMBASE').AsString;
    end;
    Ferme(Q);
  end
  else begin
    Result.RE := V_PGI.NoDossier;
    Result.RC := V_PGI.CodeSociete;
    if LibOk then Result.RT := V_PGI.NomSociete
             else Result.RT := V_PGI.SchemaName;
  end;
end;

{Récupère dans la table CLIENSSOC le compte courant entre deux sociétés.
 BqCodeOk si on veut BQ_CODE, sinon le retour sera CLS_GENERAL}
{---------------------------------------------------------------------------------------}
function GetCompteCourant(DosSource, DosDest : string; BqCodeOk : Boolean = True) : string;
{---------------------------------------------------------------------------------------}
var
  BaseS : string;
  BaseD : string;
  Q     : TQuery;
begin
  Result := '';

  if DosSource = V_PGI.NoDossier then
    BaseS := V_PGI.SchemaName
  else
    BaseS := GetInfosFromDossier('DOS_NODOSSIER', DosSource, 'DOS_NOMBASE');

  if DosDest = V_PGI.NoDossier then
    BaseD := V_PGI.SchemaName
  else
    BaseD := GetInfosFromDossier('DOS_NODOSSIER', DosDest, 'DOS_NOMBASE');

  {Récupération dans la base source du compte groupe vers la base de destination}
  Q := OpenSelect('SELECT CLS_GENERAL FROM CLIENSSOC WHERE CLS_DOSSIER = "' + BaseD + '"', BaseS);
  try
    if not Q.EOF then begin
      Result := Q.FindField('CLS_GENERAL').AsString;

      if BqCodeOk and (Result <> '') then
        Result := GetBqCodeFromGene(Result, DosSource);
    end;
  finally
    Ferme(Q);
  end;
end;

{Retourne BQ_CODE à partir d'un compte général et du dossier d'origine
{---------------------------------------------------------------------------------------}
function GetBqCodeFromGene(Gene, Dossier : string) : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Result := '';
  Q := OpenSQL('SELECT BQ_CODE FROM BANQUECP WHERE BQ_GENERAL = "' + Gene +
               '" AND BQ_NODOSSIER = "' + Dossier + '"', True);
  if not Q.EOF then
    Result := Q.FindField('BQ_CODE').AsString;
  Ferme(Q);
end;

{Retourne BQ_GENERAL à partir de BQ_CODE}
{---------------------------------------------------------------------------------------}
function GetGeneFromBqCode(BqCode : string) : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Result := '';
  Q := OpenSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_CODE = "' + BqCode + '"', True);
  if not Q.EOF then
    Result := Q.FindField('BQ_GENERAL').AsString;
  Ferme(Q);
end;

{Gère-t-on la Tréso en MultiSociété ?}
{---------------------------------------------------------------------------------------}
function IsTresoMultiSoc : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := EstMultiSoc and (GetParamSocSecur('SO_TRBASETRESO', '') <> '') and
            (ctxTreso in V_PGI.PGIContexte);
end;

{Remplit les champs société et nodossier d'une Tob TRECRITURE à partir de TE_GENERAL
{---------------------------------------------------------------------------------------}
procedure InitInfosDossier(var T : TOB);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  T.SetString('TE_SOCIETE', V_PGI.CodeSociete);
  T.SetString('TE_NODOSSIER', V_PGI.NoDossier);
  if not IsTresoMultiSoc then exit;

  Q := OpenSQL('SELECT BQ_CODE, BQ_NODOSSIER, DOS_SOCIETE FROM BANQUECP ' +
               'LEFT JOIN DOSSIER ON DOS_NODOSSIER = BQ_NODOSSIER ' +
               'WHERE BQ_CODE = "' + T.GetString('TE_GENERAL') + '"', True);
  if not Q.EOF then begin
    T.SetString('TE_SOCIETE', Q.FindField('DOS_SOCIETE').AsString);
    T.SetString('TE_NODOSSIER', Q.FindField('BQ_NODOSSIER').AsString);
  end;
  Ferme(Q);
end;

{Exclut la liste des comptes/agences/banques des natures passées en paramètres et séparées par un ";"}
{---------------------------------------------------------------------------------------}
procedure SetPlusBancaire(Composant : TControl; Pref, AExclure : string);
{---------------------------------------------------------------------------------------}
var
  Clause : string;
begin
       if Pref = 'BQ'  then Clause := 'BQ_NATURECPTE NOT IN (' + GetClauseIn(AExclure) + ')'
  else if Pref = 'TRA' then Clause := 'TRA_AGENCE NOT IN (' + GetClauseIn(AExclure) + ')'
  else if Pref = 'PQ'  then Clause := 'PQ_BANQUE NOT IN (' + GetClauseIn(AExclure) + ')';

  if Composant is THMultiValComboBox then (Composant as THMultiValComboBox).Plus := Clause
  else if Composant is THValComboBox then (Composant as THValComboBox     ).Plus := Clause
  else if Composant is THEdit        then (Composant as THEdit            ).Plus := Clause
  else if Composant is THDBEdit      then (Composant as THDBEdit          ).Plus := Clause;
end;

{Renvoie NoDossier1;NoDossier2; depuis une liste de Bases (Base1;Base2;...)}
{---------------------------------------------------------------------------------------}
function GetNoDossierFromBase(ListeBase : string = '') : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  s : string;
begin
  Result := '';
  if not IsTresoMultiSoc then
    Result := V_PGI.NoDossier + ';'
  else begin
    if ListeBase = '' then
      ListeBase := VarTreso.lNomBase;
    if ListeBase = '' then begin
      Result := V_PGI.NoDossier + ';';
      Exit;
    end;

    s := GetClauseIn(ListeBase);
    {... Récupération des NoDossier correspondants}
    Q := OpenSQL('SELECT DOS_NODOSSIER FROM DOSSIER WHERE DOS_NOMBASE IN (' + s + ')', True);
    while not Q.EOF do begin
      Result := Result + Q.FindField('DOS_NODOSSIER').AsString + ';';
      Q.Next;
    end;
    Ferme(Q);
  end;
end;

{Renvoie la la liste des dossier qui ont la base courante comme base Treso.
 Si TokenStOk, chaque nom de base est séparée par un ";", sinon le Result se présente
 sous la forme d'un requête : IN ("BASE1", "BASE2", ...)
{---------------------------------------------------------------------------------------}
function FiltreBaseTreso(TokenStOk : Boolean = True) : string;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  n : Integer;
begin
  Result := '';
  if IsTresoMultiSoc then begin
    Result := '';
    T := RecupInfosSocietes('SO_TRBASETRESO');
    try
      for n := 0 to T.Detail.Count - 1 do begin
        {Si la base traitée a pour base Tréso la base courante}
        if T.Detail[n].GetString('SO_TRBASETRESO') = V_PGI.SchemaName then
          Result := Result + T.Detail[n].GetString('SCHEMANAME') + ';';
      end;
      if (Result <> '') and not TokenStOk then
        Result := ' IN (' + GetClauseIn(Result) + ')';
    finally
      if Assigned(T) then FreeAndNil(T);
    end;
  end
  else if EstMultiSoc then begin
    if TokenStOk then Result := V_PGI.SchemaName + ';'
                 else Result := ' = "' + V_PGI.SchemaName + '" ';

  end;
end;

{Retourne IN ("NoDossier1", "NoDossier2", ...). Si NomBase = "", Filtre sur le regroupement Treso
{---------------------------------------------------------------------------------------}
function FiltreNodossier(NomBase : string = '') : string;
{---------------------------------------------------------------------------------------}
var
  s: string;
begin
  Result := '';
  if IsTresoMultiSoc then begin
    if NomBase = '' then
      Result := 'IN (' + GetClauseIn(VarTreso.lNoDossier) + ')'
    else begin
      if (NomBase <> '') and (NomBase[Length(NomBase)] <> ';') then NomBase := NomBase + ';';
      {NomBase peut contenir soit des NoDossier soit des NomBase en fonction de la tablette sur la combo dossier}
      s := NomBase;
      if Pos(ReadTokenSt(s) + ';', VarTreso.lNomBase) > 0 then
        Result := 'IN (' + GetClauseIn(GetNoDossierFromBase(NomBase)) + ')'
      else
        Result := 'IN (' + GetClauseIn(NomBase) + ')';
    end;
  end
  else
    Result := ' = "' + V_PGI.NoDossier + '"';
end;

{Attention !! : fonction à réserver à la compta à l'exclusion de la Tréso à cause du filtre
                sur le dossier
{---------------------------------------------------------------------------------------}
procedure SetPlusBanqueCp(Control : TControl);
{---------------------------------------------------------------------------------------}

    {----------------------------------------------------------------------}
    function _SetPlus(DataType, Plus : string) : string;
    {----------------------------------------------------------------------}
    var
      SqlConf : string;
    begin
      SqlConf := '';

      Result := Plus;

      {Si le filtre sur BQ_NODOSSIER existe déjà, on sort}
      if Pos('BQ_NODOSSIER', Plus) = 0 then begin
        {En compta, on filtre sur sur le dossier et les comptes bancaires}
        if DataType = tcp_Bancaire then
          Result := 'AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"'
        else
          Result := 'BQ_NODOSSIER = "' + V_PGI.NoDossier + '" AND ' + BQCLAUSEWHERE;
      end;

      {$IFDEF TRCONF}
      SqlConf := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque, GetExeContext);
      if SqlConf <> '' then SqlConf := ' AND (' + SqlConf + ') ';
      Result := Result + SqlConf;
      {$ENDIF TRCONF}
    end;

begin
       if Control is THEdit then (Control as THEdit).Plus := _SetPlus((Control as THEdit).DataType, (Control as THEdit).Plus)
  else if Control is THDBEdit then (Control as THDBEdit).Plus := _SetPlus((Control as THDBEdit).DataType, (Control as THDBEdit).Plus)
  else if Control is THValComboBox then (Control as THValComboBox).Plus := _SetPlus((Control as THValComboBox).DataType, (Control as THValComboBox).Plus)
  else if Control is THDBValComboBox then (Control as THDBValComboBox).Plus := _SetPlus((Control as THDBValComboBox).DataType, (Control as THDBValComboBox).Plus)
  else if Control is THMultiValComboBox then (Control as THMultiValComboBox).Plus := _SetPlus((Control as THMultiValComboBox).DataType, (Control as THMultiValComboBox).Plus)
end;

{Filtre sur composant attacher à une tablette sur BANQUECP sur les dossiers du regroupement
 Tréso (Dossier = '') ou éventuellement sur demande (Nature du compte, dossier)
 REMARQUE !!!! En cas de modification penser à TRSYNCHRONISATION_TOF.PoseFiltreSurCpte
{---------------------------------------------------------------------------------------}
function FiltreBanqueCp(DataType, Nature, NomBase : string) : string;
{---------------------------------------------------------------------------------------}
var
  SqlConf : string;
begin
  Result  := '';
  SqlConf := '';
  {$IFDEF TRCONF}
  SqlConf := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque, GetExeContext);
  {$ENDIF}

  if (Nature <> '' ) and (Nature[Length(Nature)] <> ';') then Nature := Nature + ';';
  if (NomBase <> '') and (NomBase[Length(NomBase)] <> ';') then NomBase := NomBase + ';';

  {$IFNDEF COMPTA}
  {Ne concerne pas le mono dossier "pur". Par contre je ne teste pas si la Treso est Multi sociétés
   car on peut avoir une Tréso autonome dont la base appartient à un regroupement avec partage de
   BanqueCp : dans ce cas, il ne faut afficher que les comptes de la base}
  if not EstMultiSoc then begin
    if SqlConf <> '' then begin
      if DataType = tcp_Bancaire then Result :=  ' AND (' + SqlConf + ') '
                                 else Result := SqlConf;
    end;
    Exit;
  end;
  if SqlConf <> '' then SqlConf := ' AND (' + SqlConf + ') ';
  {$ELSE}
  if SqlConf <> '' then SqlConf := ' AND (' + SqlConf + ') ';
  {En compta, on filtre sur sur le dossier et les comptes bancaires}
  if DataType = tcp_Bancaire then
    Result := 'AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"' + SqlConf
  else
    Result := 'BQ_NODOSSIER = "' + V_PGI.NoDossier + '" AND ' + BQCLAUSEWHERE + SqlConf;
  Exit;
  {$ENDIF COMPTA}

  {Base autonome dans un regroupement}
  if not IsTresoMultiSoc then begin
    if DataType = tcp_Bancaire then
      Result := 'AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"' + SqlConf
    else
      Result := 'BQ_NODOSSIER = "' + V_PGI.NoDossier + '"' + SqlConf;
    Exit;
  end;

  Result := 'BQ_NODOSSIER ' + FiltreNodossier(NomBase);

  if Nature <> '' then begin
    if Nature = (tcb_Bancaire + ';') then begin
      if Result <> '' then Result := Result + ' AND ' + BQCLAUSEWHERE
                      else Result := BQCLAUSEWHERE;
    end
    else
      if Result <> '' then Result := Result + ' AND BQ_NATURECPTE IN (' + GetClauseIn(Nature) + ')'
                      else Result := ' BQ_NATURECPTE IN (' + GetClauseIn(Nature) + ')';
  end;

  if DataType = tcp_Bancaire then Result := 'AND ' + Result;

  Result := Result + SqlConf;
end;

{***********A.G.L.***********************************************
Auteur  ...... : RRO
Créé le ...... : 30/10/2001
Modifié le ... :   /  /
Description .. : Retourne à une date le montant contrevalorisé d'un montant
Suite ........ : en une devise dans une autre devise
Mots clefs ... : MONTANT;DEVISE;CONTREVALORISE
*****************************************************************}
Function RetContreValeur(aDate: TDateTime; Montant : double; DevMontant : string ; DevContrevalorise : string) : double ;
var
  Parite1 : double ;
  Parite2 : Double ;
begin
  Parite1 := RetPariteEuro(DevMontant, aDate);
  Parite2 := RetPariteEuro(DevContrevalorise, aDate);

  if Parite2 <> 0 then
    Result := Montant * Parite1/Parite2
  else
    Result := 0 ;
end ;

{JP 16/09/03 : Renvoie le taux de conversion d'une devise à une autre à une date donnée (l'€ servant de pivot)}
{---------------------------------------------------------------------------------------}
function RetContreTaux(Date: TDateTime; DevMontant, DevContrevalorise : string) : double ;
{---------------------------------------------------------------------------------------}
var
  Parite1 : double ;
  Parite2 : Double ;
begin
  {Si les devises sont identiques, on sort}
  if DevMontant = DevContrevalorise then begin
    Result := 1;
    Exit;
  end;

  Parite1 := RetPariteEuro(DevMontant, Date);
  Parite2 := RetPariteEuro(DevContrevalorise, Date);

  if Parite1 <> 0 then
    Result := Parite1/Parite2
  else
    Result := 1;
end ;

{JP 21/08/03 : Récupère les amplitudes minimale et maximale d'une devise sur une période
{---------------------------------------------------------------------------------------}
function GetValeursMinMax(Dev, Chp, TT, WH : string; Deb, Fin : TDateTime): TDblResult;
{---------------------------------------------------------------------------------------}
var
  SQL : string ;
  Q : TQuery ;
begin
  Result.RE := '0';
  Result.RC := '1';
  SQL := 'SELECT MIN(' + Chp +'), MAX(' + Chp + ') FROM ' + TT + ' WHERE ' + WH;
  Q := OpenSQL(SQL, True);
  if (not Q.EOF) and (Q.Fields[0].AsString <> '') and (Q.Fields[1].AsString <> '') then begin
    Result.RE := Q.Fields[0].AsString;
    Result.RC := Q.Fields[1].AsString;
  end;
  Ferme(Q);
end;

{JP 01/10/03 : Formate les clefs de traitement sur les dates comptables et de valeur
               On formate une chaine la forme  "aammjjNumUnique" ce qui permet d'avoir
               un champ de tri dans la table écriture en fonction des dates et d'un numéro
               unique : cela simplifie les clauses where , diminue les "in select"
               dans les recalculs de solde, les équilibrages, les fiches de suivi ...
{---------------------------------------------------------------------------------------}
function RetourneCleEcriture(Dt : TDateTime; NumP : Integer) : string;
{---------------------------------------------------------------------------------------}
var
  a, j, m : Word;
  Annee   : string;
begin
  DecodeDate(Dt, a, m, j);
  Annee := Copy(IntToStr(a), 3, 2);
  Result := Annee + PadL(IntToStr(m), '0', 2) + PadL(IntToStr(j), '0', 2) + PadL(IntToStr(NumP), '0', 7);
end;

{JP 09/03/04 : création automatique des agences correspondant aux différents comptes
{---------------------------------------------------------------------------------------}
function CreerAgence : Boolean;
{---------------------------------------------------------------------------------------}
var
  T, Q, G, F : TOB;
  n          : Integer;
  s, c, Agc,
  Cal, Tva   : string;
  QQ         : TQuery;
  Ok         : Boolean;

    {---------------------------------------------------------------------}
    function MajCpteAgence : Boolean;
    {---------------------------------------------------------------------}
      var
        TT : TOB;

        {-----------------------------------------------------}
        function ValStr(Chp : string) : string;
        {-----------------------------------------------------}
        begin
          Result := VarToStr(G.GetValue(Chp));
        end;
    begin
      Result := False;
      {On génère le code agence supérieur}
      if not UProcGen.IncCode(Agc) then Exit;
      {On remplit la nouvelle banque}
      TT := TOB.Create('AGENCE', Q, -1);
      TT.PutValue('TRA_ABREGE'    , ValStr('BQ_DOMICILIATION'));
      TT.PutValue('TRA_ADRESSE1'  , ValStr('BQ_ADRESSE1'));
      TT.PutValue('TRA_ADRESSE2'  , ValStr('BQ_ADRESSE2'));
      TT.PutValue('TRA_AGENCE'    , Agc);
      TT.PutValue('TRA_BANQUE'    , ValStr('BQ_BANQUE'));
      TT.PutValue('TRA_CODECAL'   , Cal);
      TT.PutValue('TRA_CODEPOSTAL', ValStr('BQ_CODEPOSTAL'));
      TT.PutValue('TRA_ETABBQ'    , ValStr('BQ_ETABBQ'));
      TT.PutValue('TRA_GUICHET'   , ValStr('BQ_GUICHET'));
      TT.PutValue('TRA_LIBELLE'   , ValStr('BQ_DOMICILIATION'));
      TT.PutValue('TRA_REGIMETVA' , Tva);
      TT.PutValue('TRA_VILLE'     , ValStr('BQ_VILLE'));
      TT.InsertDb(nil);
      {On met à jour la table BanqueCp}
      ExecuteSql('UPDATE BANQUECP SET BQ_AGENCE = "' + Agc + '" WHERE BQ_CODE = "' + ValStr('BQ_CODE') + '"');
      Result := True;
    end;

begin
  Ok := False;
  Result := Ok;
  {Message d'avertissement
   26/07/05 : FQ 10278 : suppression du bouton Cancel dans la Boîte de dialogue}
  if HShowMessage('0;Initialisation des agences bancaires;Vous allez lancer un traitement de création d''agences pour chacun de'#13 +
               'vos comptes bancaires et de mise à jour de la table BanqueCP.'#13#13 +
               'Confirmez-vous le lancement du traitement ?;Q;YNH;N;N;False;False;500002;', '', '') <> mrYes then Exit;

  {Récupération du code agence maximum}
  QQ := OpenSQL('SELECT MAX(TRA_AGENCE), MIN(TRA_AGENCE) FROM AGENCE', True);
  if not QQ.EOF then begin
    if QQ.Fields[1].AsString > '999' then Agc := '000'
                                     else Agc := QQ.Fields[0].AsString
  end
  else
    Agc := '000';
  Ferme(QQ);
  if Trim(Agc) = '' then Agc := '000';

  {Récupération du code calendrier}
  QQ := OpenSQL('SELECT TCL_CODECAL FROM TRCALENDRIER', True);
  if not QQ.EOF then Cal := QQ.Fields[0].AsString
                else Cal := '';
  Ferme(QQ);
  {Récupération du régime de Tva avec une priorité au régime français}
  Tva := '';
  QQ := OpenSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE = "RTV" AND CC_CODE = "FRA"', True);
  if not QQ.EOF then
    Tva := QQ.Fields[0].AsString
  else begin
    Ferme(QQ);
    QQ := OpenSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE = "RTV"', True);
    if not QQ.EOF then Tva := QQ.Fields[0].AsString;
  end;
  Ferme(QQ);


  T := TOB.Create('£BANQUECP', nil, -1);
  Q := TOB.Create('£AGENCE', nil, -1);
  try
    {On récupère les agences}
    Q.LoadDetailFromSQL('SELECT TRA_AGENCE, TRA_GUICHET, TRA_BANQUE FROM AGENCE');
    {On récupère tous les comptes qui n'ont pas d'agence}
    T.LoadDetailFromSQL('SELECT BQ_BANQUE, BQ_ADRESSE1, BQ_ADRESSE2, BQ_CODEPOSTAL, BQ_DOMICILIATION, BQ_GUICHET, ' +
                        'BQ_ETABBQ, BQ_VILLE, BQ_CODE, BQ_CODE FROM BANQUECP WHERE BQ_AGENCE = "" OR BQ_AGENCE IS NULL');

    if T.Detail.Count > 0 then begin
      {On boucle sur le résultat de la requête}
      for n := 0 to T.Detail.Count - 1 do begin
        G := T.Detail[n];
        {On regarde s'il existe déjà une agence avec le code guichet du compte en courant}
        s := VarToStr(G.GetValue('BQ_GUICHET'));
        c := VarToStr(G.GetValue('BQ_BANQUE'));
        F := Q.FindFirst(['TRA_GUICHET', 'TRA_BANQUE'], [s, c], True);

        if F <> nil then ExecuteSQL('UPDATE BANQUECP SET BQ_AGENCE = "' + VarToStr(F.GetValue('TRA_AGENCE')) + '" ' +
                                    {23/02/06 : le Where me semble un incontestable plus !!}
                                    'WHERE BQ_CODE = "' + VarToStr(G.GetValue('BQ_CODE')) + '"')
                    else if not MajCpteAgence then Break;
        Ok := n = T.Detail.Count - 1;
      end;
    end
    else begin
      HShowMessage('1;Initialisation des agences bancaires;Tous les comptes avaient déjà une agence.;I;O;O;O;', '', '');
      Exit;
    end;
  finally
    FreeAndNil(T);
    FreeAndNil(T);
  end;
  {Message d'avertissement}
  if Ok then HShowMessage('1;Initialisation des agences bancaires;Le traitement s''est correctement effectué.;I;O;O;O;', '', '')
        else HShowMessage('2;Initialisation des agences bancaires;Le traitement a été interrompu.;W;O;O;O;', '', '');
  Result := Ok;
end;

{Consulte l'indicateur "appli activée" dans la base commune FQ 10183
{---------------------------------------------------------------------------------------}
function IsFlagAppliOk : Boolean;
{---------------------------------------------------------------------------------------}
var
  Q      : TQuery;
  NoDoss : string;
begin
  Result := False;

  {DOSSAPPLI est liée à des dossiers, donc uniquement gérée en mode DB0}
  Q := OpenSQL('SELECT SO_MODEFONC FROM SOCIETE', True);
  if not Q.EOF then begin
    if Q.FindField('SO_MODEFONC').AsString <> 'DB0' then begin
      Ferme(Q);
      Exit;
    end;
  end
  else
    Ferme(Q);

  if V_PGI.RunFromLanceur then
    {éxécuté depuis un lanceur (donc multi-dossier)}
    NoDoss := V_PGI.NoDossier
  else
    Exit;

  if NoDoss = '' then begin
    PGIInfo('Pas de dossier ou de société en cours.', TitreHalley);
    Exit;
  end;

  Q := OpenSQL('SELECT DAP_NOMEXEC FROM ##DP##.DOSSAPPLI WHERE DAP_NODOSSIER = "' + NoDoss +
               '" AND DAP_NOMEXEC = "' + Uppercase('CCS5.EXE') + '"', True);
  if not Q.Eof then Result := True;
  Ferme(Q);
end;

{Renvoie le type de vente des opcvm : si l'ancienne valeur est vide, on retourne Value.
 Par contre si les deux valeurs sont différentes, on renvoie DIV, pour Divers
{---------------------------------------------------------------------------------------}
function GetTypeVente(OldValue, Value : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := Value;
  if (OldValue <> Value) and (OldValue <> '') then
    Result := vop_Autre;
end;

{Retourne le cout d'achat moyen pondéré des OPCVM
{---------------------------------------------------------------------------------------}
function GetCAMP(Compte, OPCVM : string; DateRef : TDateTime) : Double;
{---------------------------------------------------------------------------------------}
var
  S : string;
  Q : TQuery;
begin
  {On retourne le prix moyen en euro sur les opcvm achetés antérieurement à DateRef et non vendu}
  S := 'SELECT SUM((TOP_NBPARTACHETE - TOP_NBPARTVENDU) * TOP_PRIXUNITAIRE / TOP_COTATIONACH) / ' +
       'SUM(TOP_NBPARTACHETE - TOP_NBPARTVENDU)  FROM TROPCVM ';
  S := S + 'WHERE TOP_DATEACHAT <= "' +UsDateTime(DateRef) + '" AND TOP_CODEOPCVM = "' + OPCVM;
  {Le where sur le nombre de parts est théoriquement inutile, mais on n'est jamais trop prudent}
  S := S + '" AND TOP_GENERAL = "' + Compte + '" AND (TOP_NBPARTACHETE - TOP_NBPARTVENDU) > 0';
  Q := OpenSQL(S, True);
  if not Q.Eof then Result := Q.Fields[0].AsFloat
               else Result := 0;
  Ferme(Q);
end;

{15/12/04 : Remplit la tob avec le detail des opcvm pour un Portefeuille, un code OPCVM et un général
{---------------------------------------------------------------------------------------}
procedure ChargeDetailOPCVM(var aTob : TOB; Portef, Opcvm, General : string);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  T : TOB;
begin
  Q := OpenSQL('SELECT TOP_NUMOPCVM, TOP_CODEOPCVM, TOP_TRANSACTION, TOP_GENERAL, '+
               'TOP_LIBELLE, TOP_BASECALCUL, TOP_MONTANTACH, TOP_FRAISACH, TOP_COMACHAT, ' +
               'TOP_COTATIONACH, TOP_NBPARTACHETE, TOP_NBPARTVENDU, TOP_TVAACHAT FROM TROPCVM ' +
               'WHERE TOP_VALBO = "X" AND TOP_COMPTA = "X" AND TOP_STATUT <> "X" AND ' +
               'TOP_GENERAL = "' + General + '" AND TOP_CODEOPCVM = "' + Opcvm + '" ' +
               {11/06/07 : FQ 10474 : Ajout du ORDER BY TOP_DATEACHAT}
               'AND TOP_PORTEFEUILLE = "' + Portef + '" ORDER BY TOP_DATEACHAT', True);
  try
    while not Q.EOF do begin
      T := TOB.Create('£££', aTob, -1);
      T.AddChampSupValeur('TOP_NUMOPCVM'    , Q.FindField('TOP_NUMOPCVM'    ).AsVariant);
      T.AddChampSupValeur('TOP_CODEOPCVM'   , Q.FindField('TOP_CODEOPCVM'   ).AsVariant);
      T.AddChampSupValeur('TOP_TRANSACTION' , Q.FindField('TOP_TRANSACTION' ).AsVariant);
      T.AddChampSupValeur('TOP_GENERAL'     , Q.FindField('TOP_GENERAL'     ).AsVariant);
      T.AddChampSupValeur('TOP_LIBELLE'     , Q.FindField('TOP_LIBELLE'     ).AsVariant);
      T.AddChampSupValeur('TOP_BASECALCUL'  , Q.FindField('TOP_BASECALCUL'  ).AsVariant);
      T.AddChampSupValeur('TOP_MONTANTACH'  , Q.FindField('TOP_MONTANTACH'  ).AsVariant);
      T.AddChampSupValeur('TOP_FRAISACH'    , Q.FindField('TOP_FRAISACH'    ).AsVariant);
      T.AddChampSupValeur('TOP_TVAACHAT'    , Q.FindField('TOP_TVAACHAT'    ).AsVariant);
      T.AddChampSupValeur('TOP_COMACHAT'    , Q.FindField('TOP_COMACHAT'    ).AsVariant);
      T.AddChampSupValeur('TOP_COTATIONACH' , Q.FindField('TOP_COTATIONACH' ).AsVariant);
      T.AddChampSupValeur('TOP_NBPARTACHETE', Q.FindField('TOP_NBPARTACHETE').AsVariant);
      T.AddChampSupValeur('TOP_NBPARTVENDU' , Q.FindField('TOP_NBPARTVENDU' ).AsVariant);
      {Champs "calculés"}
      T.AddChampSupValeur('MONTANTTOT', T.GetDouble('TOP_MONTANTACH') + T.GetDouble('TOP_FRAISACH') + T.GetDouble('TOP_TVAACHAT') + T.GetDouble('TOP_COMACHAT'));
      T.AddChampSupValeur('DATEVENTE', V_PGI.DateEntree);
      T.AddChampSupValeur('PARTAVENDRE', 0);
      T.AddChampSupValeur('MONTANTVEN', 0.0);
      T.AddChampSupValeur('TOP_DATEACHAT', iDate1900);
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;

end;

{21/12/04 : Remplit la tob avec le detail des ventes pour un Portefeuille, un code OPCVM et un général
{---------------------------------------------------------------------------------------}
procedure ChargeDetailVente(var aTob : TOB; NumVente : string; PourEcr : Boolean = False);
{---------------------------------------------------------------------------------------}
var
  SQL : string;
begin
  {19/12/06 : FQ 10392 : on génère un flux dans TRECRITURE par vente}
  if PourEcr then begin
    SQL := 'SELECT TVE_GENERAL, TVE_CODEOPCVM, TVE_DATEVENTE, MIN(TVE_NUMTRANSAC) TVE_NUMTRANSAC, TVE_NUMVENTE, ' +
           {12/02/07 : MAX est beaucoup mieux que SUM !!}
           'MAX(TVE_COURSEUR) TVE_COURSEUR, SUM(TVE_NBVENDUE) TVE_NBVENDUE FROM TRVENTEOPCVM ';
    SQL := SQL + 'WHERE TVE_NUMVENTE = ' + NumVente + ' GROUP BY ' +
           'TVE_GENERAL, TVE_CODEOPCVM, TVE_DATEVENTE, TVE_NUMVENTE ';
    aTob.LoadDetailFromSQL(SQL);
  end
  else
    aTob.LoadDetailDBFromSQL('TRVENTEOPCVM', 'SELECT * FROM TRVENTEOPCVM WHERE TVE_NUMVENTE = ' + NumVente);
end;

{21/12/04 : Le lien entre TRECRITURE et TRVENTEOPCVM se fait entre TE_NUMTRANSAC et TVE_NUMVENTE :
            TE_NUMTRANSAC : "TRE" + CodeSociete + "VOP" + PadL(TVE_NUMVENTE, '0', 7)
{---------------------------------------------------------------------------------------}
function GetNumTransacVente(NumTransac : string; FromVente : Boolean) : string;
{---------------------------------------------------------------------------------------}
begin
  if FromVente then
    Result := CODEMODULE + V_PGI.CodeSociete + TRANSACVENTE + PadL(NumTransac, '0', 7)
  else
    Result := IntToStr(StrToInt(Copy(NumTransac, Length(NumTransac) - 6, 7)));
end;

{Charge la Variable globale de Treso}
{---------------------------------------------------------------------------------------}
procedure ChargeVarTreso;
{---------------------------------------------------------------------------------------}
begin
  VarTreso.lNomBase   := FiltreBaseTreso;
  VarTreso.lNoDossier := GetNoDossierFromBase;
end;

{JP 25/01/08 : FQ 10547 : Renvoie les dossiers / Bases autorisés}
{---------------------------------------------------------------------------------------}
function GetTrFiltreDossiers : string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';
  if IsTresoMultiSoc then begin
    if (VarTreso.lNomBase = '') then
      ChargeVarTreso;
    Result := VarTreso.lNomBase;
  end
  else
    Result := V_PGI.SchemaName + ';';
end;

{Recherche dans toute la hiérarchie de la TOB, les champ "Field" et leur affecte Value
{---------------------------------------------------------------------------------------}
procedure PutValueWholeTob(MyTob : TOB; FieldName : string; Value : Variant);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;
begin
  if not Assigned(MyTob) then Exit;
  for n := 0 to MyTob.Detail.Count - 1 do begin
    F := MyTob.Detail[n];
    {Appel récursif : cette fonction a pour but de mettre à jour uen TOB dont on ne
     connaîtrait pas la structure}
    PutValueWholeTob(F, FieldName, Value);
    if F.FieldExists(FieldName) then F.PutValue(FieldName, Value);
  end;
end;

{Fonction pour savoir si on gère les nouveaux soldes : Cette fonction qui peut sembler
 inutile est là pour rajouter un paramsoc, si le risque parait trop grand.
 La nouvelle gestion des soldes consiste à ne plus tavailler sur les soldes stockés dans
 TRECRITURE, mais de toujours les calculer à la volée par requête.
{---------------------------------------------------------------------------------------}
function  IsNewSoldes : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
end;

end.

