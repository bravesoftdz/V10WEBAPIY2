{$A+,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O-,P+,Q-,R-,S-,T-,U-,V+,W+,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
{$A+,B-,C+,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O-,P+,Q-,R-,S-,T-,U-,V+,W+,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 14/06/2001
Modifié le ... :   /  /
Description .. : Unit de traitement du fichier extérieur à la paie
Suite ........ : Utile pour traiter des absences, des saisies d'éléments
Suite ........ : de paie provenant d'un autre logiciel,
Suite ........ : de heures de présences et/ou d'absences provenant d'une
Suite ........ : pointeuse
Suite ........ : Utilisé pour récupérer  les absences venant d'une saisie
Suite ........ : déportée.
Mots clefs ... : PAIE;ABSENCES;IMPORT
*****************************************************************}
{
PT1    : 14/06/2001 PH Rajout libellés complémentaires dans le traitement des
                       lignes absences pour prendre en compte les libellés
                       complémentaires lors de la récup des absences déportées
PT2    : 09/10/2001 PH Rajout traitement ligne Analytique
PT3    : 26/11/2001 PH Récupération du libellé saisi de l'absence et non pas un
                       libellé forcé
PT4    : 26/11/2001 PH Incrementation du numéro d'ordre des mvts de type CP
PT5    : 27/11/2001 PH Initialisation de tous les champs de la table
                       absencesalarie
PT6    : 27/11/2001 PH Si on ne gère pas les cp et que l'on a des lignes de cp,
                       on perd les lignes sans que l'utilisateur le sache
PT7    : 18/12/2001 PH remplacement strtofloat par valeur sinon erreur de
                       conversion
PT8    : 15/02/2002 PH test pour les mvts de cp sur mois clos N et N+1 pour
                       prendre les mvts saisis sur le mois precedant (même si
                       clos) au dernier mois
PT9    : 28/10/2002 SB On considère que c'est le salarie qui saisie les mvts
                       absences de l'import
PT10   : 30/10/2002 SB On tronque le libellé de l'absence à 35 caractères
PT11-1 : 02/12/2002 PH Regroupement lecture du fichier de façon à trouver matin
                       et après midi pour le controle que l'absence n'a pas été
                       déjà saisie
PT11-2 : 02/12/2002 PH Si Salarié sorti et absence déjà saisie alors OK
PT12   : 28/01/2003 SB Erreur appli si type separateur différent de celui
                       utilisé
PT13   : 13/03/2003 PH Traitement des lignes import pré-ventilations analytiques
                       salariés
PT14   : 04/06/2003 PH Création automatique des sections analytiques lors de
                       l'import si autorisé
// **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
PT15   : 25/08/2003 PH Correction prise en compte traitement analytique ANA et
                       VEN
PT16   : 28/08/2003 PH Création des sections en récupérant le libellé de la
                       section
PT17   : 08/09/2003 PH Destruction des mouvements ANA
PT18   : 10/09/2003 PH Controle du séparateur du fichier import et creation
                       automatique fichier rapport
PT19   : 16/12/2003 PH OPtimisation de code pour éviter requêtes répétitives +
                       traitement des absences sur périodes closes
PT20   : 06/01/2004 PH Suppression des ventilations analytiques salarie
                       uniquement sur le même axe que celles contenues dans le
                       fichier d'import
PT21   : 12/03/2004 SB V_50 FQ 11162 Encodage de de la date de cloture erroné si
                            fin fevrier
PT22   : 09/04/2004 SB V_50 FQ 11136 Ajout Gestion des congés payés niveau
                            salarié
PT22   : 26/05/2004 PH V_50 Non traitement des absences portant sur des périodes
                            closes mias ne provoquant pas de rejet du fichier
PT23   : 03/06/2004 PH V_50 FQ 11271 Erreur SAV TOB.getvalue sur un champ non
                            existant ==> provoque une erreur ?
PT24   : 08/09/2004 PH V_50 Modif pour traitement automatique fichier NetService
                            PCL
PT25   : 20/09/2004 PH V_50 Source sans uses spécifiques de la paie
PT26   : 02/11/2004 PH V_60 Prise en compte des rubriques de type permanent dans
                            le fichier d'import
PT27   : 02/11/2004 PH V_60 FQ 11409 Traitement Analytique salarié pour un mois
                            + Création de salarié
PT28   : 30/11/2004 PH V_60 FQ 11794 Fin de traitement OK
PT29   : 31/01/2005 SB V_60 Export des proc. communes dans pgcommun
PT30   : 08/08/2005 PH V_60 FQ 11926 Erreur si salarié inexistant
//*** REFONTE COMPLETE DU SOURCE pour automatisation EWS et eaglserveur ***//
PT31   : 31/08/2005 PH V_60 FQ 12517 Ergonomie
PT32   : 02/03/2005 PH V_60 FQ 12959 Traitement cas particulier 1 salarié en
                            création
PT33   : 08/03/2006 MF V_65 FQ 12922 récup. absence : alimentation booléen
                            PCN_GESTIONIJSS et du double PCN_NBJCALEND
PT34   : 14/03/2006 PH V_65 Prise en compte création de salarié sans salarié
                            déjà existant
PT35   : 23:03:2006 GGS V_65 Alimentation table de travail(importpaie) pour
                            édition des données importées
PT36   : 11/04/2006 PH V_65 FQ 12582 Gestion des doublons Ne pas doubler les
                            lignes si import sur les mêmes critères
PT37   : 11/04/2006 PH V_65 Prise en compte des nouveaux champs salariés
PT38   : 13/04/2006 PH V_65 FQ 12490 Initialisation des dates avec iDate1900
PT39   : 13/04/2006 SB V_65 Traitement gestion maximum + FQ 12882 & 12941
                            Traitement calcul duree absence
PT40   : 13/04/2006 SB V_65 Traitement de l'annulation de l'absence
PT41   : 14/04/2006 SB V_65 FQ 11953 Paramsoc absence à cheval sur plusieurs
                            mois
PT42   : 08/06/2006 PH V_65 Recup de fichier de version inférieure
PT43   : 09/06/2006 SB V_65 FQ 13202 Ajout option contrôle absence sur motif
                            d'absence
PT44   : 14/06/2006 PH V_65 FQ 13291 Prise en compte des dates dans le requêtes
PT45   : 07/07/2006 SB V_65 FQ 13362 Import d'absence, traitement des doublons
                            d'absence et des salariés sortis
PT45-2 : 07/07/2006 SB V_65 Anomalie syntaxe
PT46   : 31/08/2006 PH V_65 Compatibilité avec les . ou ,  dans les montants et
                            test existance motif absence
PT47   : 20/09/2006 PH V_65 Numérotation création salarié S1 ou S1_ +
                            Uniformisation versionfic
PT48   : 30/11/2006 PH V_65 FQ 13710 Prise en compte longueur du numéro de
                            dossier pour mutisoc en non PCL
PT49   : 26/12/2006 PH V_70 Non rejet du fichier import si absence en anomalie.
PT50   : 04/01/2007 PH V_70 Recup de du champ condemploi au lieu de code emploi
PT51   : 25/04/2007 GGS V_70 FQ 13255 Ajout nom prénom dans liste de control
                            importation
PT52   : 22/05/2007 VG V_72 Intégration du planning unifié
PT53   : 07/09/2007 PH V_72 FQ 14740 Import uniquement des absences
PT54   : 11/10/2007 PH V_80 FQ 14834 FQ 14833 Import ligne MLB et import de
                            lignes sur salarie en creation
PT55   : 18/10/2007 VG V_80 Mise à jour du planning unifié
PT56   : 30/10/2007 GGU V_80 Gestion des absences en horaire
PT57   : 30/10/2007 GGU V_80 Problème dans la gestion d'anomalies lors de l'import
                             des fichiers d'absence : Gestion de 2 variables
                             d'exclusion de ligne (AbsNonTrait pour gérer au niveau
                             de la ligne et AbsExclue pour gérer au niveau de l'import)
PT58   : 03/01/2008 GGU V_80 Paramètre société permettant de cocher par défaut
                             la case Ignorer les anomalies d'import d'absence
PT59   : 08/01/2008 PH  V_80 Dans le cas d'EWS, on ne rejette plus les enregistrement MHE incomplet
                              (base, taux, ...).
PT60   : 13/10/2008 JS      FQ n°15288 Création automatique du compte tiers lorsque le matricule est alpha-numérique
}


unit UtofPgImportFic;

interface
uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
  SysUtils,
  HCtrls,
  HEnt1,
  Classes,
  StdCtrls,
{$IFNDEF EAGLCLIENT}
  Db,
  HDB,
  Hqry,
  Fiche,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
  DBCtrls,
  QRe, EdtREtat,
{$ELSE}
  eFiche, UtileAgl,
{$ENDIF}
  Controls,
  forms,
  Vierge,
  ComCtrls,
  HMsgBox,
  Math,
  UTOM,
  UTOB,
  UTOF,
  HTB97,
  ParamSoc,
  Dialogs,
  HRichOLE,
  HStatus,
  PgOutils2,
  ed_tools,
  UTOBDebug,
  PGCommun,
  YRessource,
  ULibEditionPaie,
  PgPlanningUnifie;


type
  TOF_PGIMPORTFIC = class(TOF)
  private
    RapportErreur: THRichEditOLE;
    Sep: Thvalcombobox; // PT18
    procedure FicChange(Sender: TObject);
    procedure SepChange(Sender: TObject);
    procedure ImprimeClick(Sender: TObject);
    procedure ConsultErreurClick(Sender: Tobject);
    procedure ValideEcran(Top: boolean);
{$IFNDEF COMSX}
    procedure Traiteligne(chaine, S: string);
    procedure Erreur(Chaine, S: string; NoErreur: integer; WW: string = '');
    procedure Anomalie(Chaine, S: string; NoErreur: integer);
    function ControleEnrDossier(chaine, S: string): integer;
    procedure ControleEnregMHE_MFP_MLB(TypeEnr, chaine, S: string; Ecrit: boolean);
    procedure ControleEnregMAB(chaine, S: string; Ecrit: boolean);
    procedure ControleEnregANA(TypeEnr, chaine, S: string; Ecrit: boolean);
    procedure ControleEnregVSA(TypeEnr, chaine, S: string; Ecrit: boolean); // PT27
    function ControleEnregSAL(TypeEnr, chaine, S: string; Ecrit: boolean): boolean; // PT27
    procedure ControleEnregVEN(TypeEnr, chaine, S: string; Ecrit: boolean);
    function ExisteRubrique(Rubrique: string): boolean;
    function ExisteSection(Axe, Section, LibSect: string): boolean;
    procedure ControleSalarie(TypeEnr, chaine, S: string; var CodeSalarie: string; Ecrit: Boolean);
    procedure ImporteFichier;
    function IsMarqueDebut(S: string): boolean;
    function IsMarqueFin(S: string): boolean;
    procedure FermeTout;
    function IsPeriodecloture(Dated, DateF: TDateTime; CasCp: string = '-'): integer;
    function RendNumOrdre(TypeMvt, CodeSalarie: string): integer;
    function PgImportDef(LeNomduFichier, LeTypImp: string; LeSeparateur: Char = '|'): Boolean;
    procedure ControleFichierImport(Sender: Tobject);
    procedure ConversionJourHeure(chaine, NBj, Nbh: string; var VNbj, VNbh: string);
{$ENDIF}
//PT29    function ImportRendPeriodeEnCours(var ExerPerEncours, DebPer, FinPer: string): Boolean;
//PT29    function ImportRendExerSocialEnCours(var MoisE, AnneeE, ComboExer: string; var DebExer, FinExer: TDateTime): Boolean;
//PT29    function ImportColleZeroDevant(Nombre, LongChaine: integer): string;
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
  end;
  {
  NoErreur :
  5 : Manque marque début fichier
  }
{$IFDEF COMSX}
procedure Traiteligne(chaine, S: string);
procedure Erreur(Chaine, S: string; NoErreur: integer; WW: string = '');
procedure Anomalie(Chaine, S: string; NoErreur: integer);
function ControleEnrDossier(chaine, S: string): integer;
procedure ControleEnregMHE_MFP_MLB(TypeEnr, chaine, S: string; Ecrit: boolean);
procedure ControleEnregMAB(chaine, S: string; Ecrit: boolean);
procedure ControleEnregANA(TypeEnr, chaine, S: string; Ecrit: boolean);
procedure ControleEnregVSA(TypeEnr, chaine, S: string; Ecrit: boolean); // PT27
function ControleEnregSAL(TypeEnr, chaine, S: string; Ecrit: boolean): boolean; // PT27
procedure ControleEnregVEN(TypeEnr, chaine, S: string; Ecrit: boolean);
function ExisteRubrique(Rubrique: string): boolean;
function ExisteSection(Axe, Section, LibSect: string): boolean;
procedure ControleSalarie(TypeEnr, chaine, S: string; var CodeSalarie: string; Ecrit: Boolean);
procedure ImporteFichier;
function IsMarqueDebut(S: string): boolean;
function IsMarqueFin(S: string): boolean;
procedure FermeTout;
function IsPeriodecloture(Dated, DateF: TDateTime; CasCp: string = '-'): integer;
function RendNumOrdre(TypeMvt, CodeSalarie: string): integer;
function PgImportDef(LeNomduFichier, LeTypImp: string; LeSeparateur: Char = '|'): Boolean;
function ControleFichierImport(Sender: Tobject): Boolean;
procedure ConversionJourHeure(chaine, NBj, Nbh: string; var VNbj, VNbh: string);
{$ENDIF}

implementation
uses TiersUtil
{$IFDEF EWS}
  , UtileWS
{$ENDIF}
  , PgCalendrier;

var MarqueDebut, MarqueFin, typebase, typetaux, typecoeff, typemontant, SalariePrec, RubPrec, Libellerubrique: string;
  TFinPer, DateEntree, DateSortie: TDatetime;
  Separateur: char;
  MarqueFinOk, Erreur_10, ExisteErreur, LibErreur, libanomalie, AbsExclue: boolean;
  Tob_sal, Tob_SalATT, tob_rubrique, Tob_section, Tob_Abs, Tob_HSR, Tob_OrdreAbs, Tob_OrdreHSR, Tob_PaieVentil: tob;
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
  Tob_Ventil, Tob_Etab: TOB; // PT27 Rajout traitement creation de salarie vérif etablissement
  Tob_MotifAbs, T_MotifAbs: Tob; { PT39 }
  FR: TextFile;
  F: TextFile;
  taille: integer;
  NumOrdre: Integer;
  NomDuFichier: string;
  TypImport: string;
  Cloture: string; // PT19
  MatSouche: Integer; // Matricule souche pour la création automatique
  TopCreat: Boolean; // Top salariés déjà dejà crées ou modifiés
  DebExer, FinExer: tdatetime;
  Iduniq, Statutenreg, VersionFic: string; //GGR PT35 pour table de travail PT42
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : fonction d'initialisation d'un mouvement congés payés
Suite ........ : A appeler avant tout renseignement de nouveau mvt
Mots clefs ... : PAIE;CP
*****************************************************************}

(* PT29 procedure ImportInitialiseTobAbsenceSalarie(TCP: tob);
begin
  TCP.Putvalue('PCN_TYPEMVT', 'CPA');
  TCP.Putvalue('PCN_SALARIE', '');
  TCP.Putvalue('PCN_DATEDEBUT', 0);
  TCP.Putvalue('PCN_DATEFIN', 0);
  TCP.Putvalue('PCN_ORDRE', 0);
  TCP.Putvalue('PCN_TYPECONGE', '');
  TCP.Putvalue('PCN_SENSABS', '');
  TCP.PutValue('PCN_PERIODEPY', -1);
  TCP.Putvalue('PCN_LIBELLE', '');
  TCP.Putvalue('PCN_DATEMODIF', 0);
  TCP.Putvalue('PCN_DATESOLDE', 0);
  TCP.Putvalue('PCN_DATEVALIDITE', 0);
  TCP.PutValue('PCN_PERIODECP', -1);
  TCP.Putvalue('PCN_DATEDEBUTABS', 0);
  TCP.Putvalue('PCN_DATEFINABS', 0);
  TCP.Putvalue('PCN_DATEPAIEMENT', 0);
  TCP.Putvalue('PCN_CODETAPE', '...'); //PT-13
  TCP.Putvalue('PCN_JOURS', 0);
  TCP.Putvalue('PCN_BASE', 0);
  TCP.Putvalue('PCN_NBREMOIS', 0);
  TCP.Putvalue('PCN_CODERGRPT', 0);
  TCP.PutValue('PCN_MVTDUPLIQUE', '-');
  TCP.Putvalue('PCN_ABSENCE', 0);
  TCP.Putvalue('PCN_MODIFABSENCE', '-');
  TCP.Putvalue('PCN_APAYES', 0);
  TCP.Putvalue('PCN_VALOX', 0);
  TCP.Putvalue('PCN_VALOMS', 0);
  TCP.Putvalue('PCN_VALORETENUE', 0);
  TCP.Putvalue('PCN_MODIFVALO', '-');
  TCP.Putvalue('PCN_PERIODEPAIE', '');
  TCP.Putvalue('PCN_ETABLISSEMENT', '');
  TCP.Putvalue('PCN_TRAVAILN1', '');
  TCP.Putvalue('PCN_TRAVAILN2', '');
  TCP.Putvalue('PCN_TRAVAILN3', '');
  TCP.Putvalue('PCN_TRAVAILN4', '');
  TCP.PutValue('PCN_GENERECLOTURE', '-');
  // a voir   TCP.PutValue('PCN_CONFIDENTIEL','-');
end;  *)
// DEB PT25

(* PT29 function PGImportEncodeDateBissextile(AA, MM, JJ: WORD): TDateTime;
begin
  Result := Idate1900;
  if IsValidDate(IntToStr(JJ) + '/' + IntToStr(MM) + '/' + IntToStr(AA)) then
    Result := encodedate(AA, MM, JJ);
  if (MM = 2) and ((JJ = 28) or (JJ = 29)) then //Année bissextile
  begin
    Result := encodedate(AA, MM, 1);
    Result := FindeMois(Result);
  end;
end;    *)
// FIN PT25

(* PT29 function ImportCalculPeriode(DTClot, DTValidite: TDatetime): integer;
var
  Dtdeb, DtFin, DtFinS: TDATETIME;
  aa, mm, jj: word;
  i: integer;
begin
  result := -1;
  if DTClot <= idate1900 then exit; //PT53
  Decodedate(DTclot, aa, mm, jj);
  DtDeb := PGEncodeDateBissextile(aa - 1, mm, jj) + 1; { PT21 } // PT25
  DtFin := DtClot;
  DtFinS := PGEncodeDateBissextile(aa + 1, mm, jj); { PT21 } // PT25
  if Dtvalidite > Dtfins then
  begin
    result := -9;
    exit;
  end;
  if DtValidite > DtClot then exit;
  result := 0;
  i := 0;
  while not ((DTValidite >= DtDeb) and (DTValidite <= DtFin)) do
  begin
    i := i + 1;
    if i > 50 then exit; // pour ne pas boucler au cas où....
    result := result + 1;
    DtFin := DtDeb - 1;
    Decodedate(DTFin, aa, mm, jj);
    DtDeb := PGEncodeDateBissextile(aa - 1, mm, jj) + 1; { PT21 } // PT25
  end;
end;  *)

(* PT29 function ImportRendNoDossier(): string;
begin
  if V_PGI.NoDossier <> '' then result := V_PGI.NoDossier
  else result := '000000';
end;                        *)

function PgImpnotVide(T: tob; const MulNiveau: boolean): boolean;
var
  tp: tob;
begin
  result := false;
  if t = nil then
    exit;
  tp := t.findfirst([''], [''], MulNiveau);
  result := (tp <> nil);
end;

{$IFNDEF COMSX}

function TOF_PGIMPORTFIC.IsMarqueDebut(S: string): boolean;
{$ELSE}

function IsMarqueDebut(S: string): boolean;
{$ENDIF}
begin
  result := true;
  if S = MarqueDebut then exit;
  S := readtokenPipe(S, Separateur);
  if S = MarqueDebut then exit
  else result := false;
end;
{$IFNDEF COMSX}

function TOF_PGIMPORTFIC.IsMarqueFin(S: string): boolean;
{$ELSE}

function IsMarqueFin(S: string): boolean;
{$ENDIF}
begin
  result := true;
  if S = MarqueFin then exit;
  S := readtokenPipe(S, Separateur);
  if S = MarqueFin then exit
  else result := false;
end;
{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.FermeTout;
{$ELSE}

procedure FermeTout;
{$ENDIF}
begin
  CloseFile(F);

  CloseFile(FR);
  if PgImpnotVide(Tob_sal, true) then
    Tob_sal.Free;
  if PgImpnotVide(Tob_rubrique, true) then
    Tob_rubrique.Free;
  //   PT2 09/10/01 PH Rajout traitement ligne Analytique
  if Tob_section <> nil then
  begin
    Tob_section.free;
    tob_section := nil;
  end;
  // DEB PT27
  if Tob_Etab <> nil then
  begin
    Tob_Etab.free;
    Tob_Etab := nil;
  end;
  if Tob_SalATT <> nil then
  begin
    Tob_SalATT.free;
    Tob_SalATT := nil;
  end;
  // FIN PT27
end;
{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.ImporteFichier;
{$ELSE}

procedure ImporteFichier;
{$ENDIF}
var
  i, ii: integer;
  LaNature, MesAxes: string;
  S, CodeEnr, st, tempo, CodeS, salP: string;
  Tob_S, T: tob;
  Q: TQuery;
  DD, DF: TDateTime;
PGYPL : TypePGYPL;
begin
  SalP := '';
  st := 'SELECT PCN_SALARIE,PCN_TYPEMVT,MAX(PCN_ORDRE) AS NUMORDRE FROM ABSENCESALARIE GROUP BY PCN_SALARIE,PCN_TYPEMVT ORDER BY PCN_SALARIE';
  Q := Opensql(st, true);
  //  if Q.eof then
  //     begin  Ferme(Q);   exit;     end;
  Tob_OrdreABS := Tob.create('Table des ordre absences maxi', nil, -1);
  Tob_ordreAbs.loaddetaildb('LES ABSENCESALARIES', '', '', Q, false);
  Ferme(Q);
  if Tob_ordreABS = nil then
  begin
    exit;
  end;

  st := 'SELECT PSD_SALARIE,MAX(PSD_ORDRE) AS PSD_ORDRE FROM HISTOSAISRUB GROUP BY PSD_SALARIE';
  Q := Opensql(st, true);
  //  if Q.eof then
  //     begin     Ferme(Q);   exit;     end;
  Tob_OrdreHSR := tob.create('Table des ordre maxi', nil, -1);
  Tob_ordreHSR.addchampsup('PSD_SALARIE', True);
  Tob_ordreHSR.addchampsup('PSD_ORDRE', True);
  Tob_ordreHSR.loaddetaildb('HISTOSAISRUB', '', '', Q, false);
  Ferme(Q);
  if Tob_ordreHSR = nil then
  begin
    exit;
  end;
{$IFNDEF COMSX}
  InitMoveProgressForm(nil, 'Ecriture de la base', 'Veuillez patienter SVP ...', taille, False, TRUE);
{$ENDIF}
  Tob_S := Tob.create('Tob lignes import', nil, -1);
  Tob_Abs := tob.create('ABSENCES A INSERER', Tob_S, -1);
  Tob_HSR := tob.create('HISTOSAISRUB A INSERER', Tob_S, -1);
  // PT2 09/10/01 PH Rajout traitement ligne Analytique
  if GetParamSoc('SO_PGANALYTIQUE') then // PT24
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
  begin
    Tob_PaieVentil := TOB.Create('Les Ventils Analytiques', nil, -1);
    Tob_Ventil := TOB.Create('Les Préventils Analytiques', nil, -1);
  end;
  i := 0;
  while not Eof(F) do //Deuxieme passage
  begin
    i := i + 1;
    MoveCurProgressForm(IntToStr(i));
    Readln(F, S);
    if isMarqueDebut(S) then
      continue;
    if isMarqueFin(S) then
      break;
    CodeEnr := readtokenPipe(S, Separateur);
    tempo := S;
    if CodeEnr = '000' then
      continue;
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
    // DEB PT27 Création automatique des salariés en attente
    if (Tob_SalATT <> nil) and (Tob_SalATT.Detail.Count - 1 >= 0) then ControleEnregSAL('', '', '', TRUE); // Creation des salariés
    if CodeEnr = 'SAL' then continue;
    // FIN PT27
    // DEB PT27 Rajout ventilation type salarié uniquement valables dans le mois
    if (CodeEnr = 'VEN') or (CodeEnr = 'VSA') then
    begin
      CodeS := readtokenPipe(tempo, Separateur);
      if CodeS <> SalP then
      begin
        if (CodeEnr = 'VEN') then st := 'Salarie ' + CodeS + ' : enregistrement(s) pré-ventilations analytiques traité(s)'
        else st := 'Salarie ' + CodeS + ' : enregistrement(s) pré-ventilations analytiques salarié pour la période traité(s)';
        Anomalie(st, st, 98);
        SalP := CodeS;
      end;
      if (CodeEnr = 'VEN') then ControleEnregVEN(CodeEnr, S, S, True)
      else ControleEnregVSA(CodeEnr, S, S, True);
      // FIN PT27
    end
    else
    begin
      // FIN PT13
      if CodeEnr = 'MAB' then
      begin // Traitement ligne absences
        CodeS := readtokenPipe(tempo, Separateur);
        if CodeS <> SalP then
        begin
          st := 'Salarie ' + CodeS + ' : enregistrement(s) Absence traité(s)';
          Anomalie(st, st, 98);
          SalP := CodeS;
        end;
        ControleEnregMAB(S, S, True);
      end // Fin traitement ligne absences
      else
      begin
        CodeS := readtokenPipe(tempo, Separateur);
        if CodeS <> SalP then
        begin
          if CodeEnr <> 'ANA' then st := 'Salarie ' + CodeS + ' : enregistrement(s) MHE ou MLB ou MFP traité(s)'
          else st := 'Salarie ' + CodeS + ' : enregistrement(s) Analytique(s) traité(s)';
          Anomalie(st, st, 98);
          SalP := CodeS;
        end;

        if CodeEnr <> 'ANA' then ControleEnregMHE_MFP_MLB(CodeEnr, S, S, True)
        else ControleEnregANA(CodeEnr, S, S, True);
      end;
    end;
  end; // Fin boucle lecture fichier import
  try
    BeginTrans;

    if Tob_HSR <> nil then
    begin
      Tob_HSR.SetAllModifie(TRUE);
      Tob_HSR.InsertOrUpdateDB(false);
    end;
    if Tob_ABS <> nil then
    begin
      Tob_ABS.SetAllModifie(TRUE);
      Tob_ABS.InsertOrUpdateDB(false);
//PT55
      for i:= 0 to Tob_ABS.Detail.Count-1 do
          CreatePGYPL (Tob_ABS.Detail[i], PGYPL);
//FIN PT55
    end;
    DD := IDate1900;
    DF := IDate1900;
    if Tob_PaieVentil <> nil then
    begin
      // PT17   08/09/2003 PH Destruction des mouvements ANA
      SalP := '';
      T := Tob_PaieVentil.FindFirst([''], [''], FALSE);
      // PT20 Suppression des ventilations analytiques salarie sur axe dans le fichier d'import
      MesAxes := '-----';
      while T <> nil do
      begin
        CodeS := T.GetValue('PAV_COMPTE');
        DD := T.GetValue('PAV_DATEDEBUT');
        DF := T.GetValue('PAV_DATEFIN');
        if CodeS <> SalP then
        begin
          st := 'DELETE FROM PAIEVENTIL WHERE PAV_NATURE LIKE "PG%" AND PAV_COMPTE="' + SalP + '"' +
            ' AND PAV_DATEDEBUT >= "' + UsDateTime(DD) + '" AND PAV_DATEFIN <= "' + UsDateTime(DF) + '"';
          ExecuteSQL(St);
          SalP := CodeS;
          MesAxes := '-----';
        end;
        // PT23 V_NATURE remplacé par PAV_NATURE
        ii := ValeurI(Copy(T.GetValue('PAV_NATURE'), 3, 1));
        MesAxes[ii] := 'X';
        T := Tob_PaieVentil.FindNext([''], [''], FALSE);
      end;
      // FIN PT17
      // Traitement du dernier salarié
      st := 'DELETE FROM PAIEVENTIL WHERE PAV_NATURE LIKE "PG%" AND PAV_COMPTE="' + SalP + '"' +
        ' AND PAV_DATEDEBUT >= "' + UsDateTime(DD) + '" AND PAV_DATEFIN <= "' + UsDateTime(DF) + '"';
      ExecuteSQL(St);
      // FIN PT20
      Tob_PaieVentil.SetAllModifie(TRUE);
      Tob_PaieVentil.InsertOrUpdateDB(false);
    end;
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
    if Tob_Ventil <> nil then
    begin
      // PT15   25/08/2003 PH Correction destruction des pré-ventilations analytiques
      SalP := '';
      T := Tob_Ventil.FindFirst([''], [''], FALSE);
      // PT20 Suppression des ventilations analytiques salarie sur axe dans le fichier d'import
      MesAxes := '-----';
      while T <> nil do
      begin
        CodeS := T.GetValue('V_COMPTE');
        if CodeS <> SalP then
        begin
          // PT15   25/08/2003 PH Correction destruction des pré-ventilations analytiques
          for ii := 1 to 5 do
          begin
            if MesAxes[ii] = 'X' then
            begin //  on traite tjrs le salarié en cours cad celui qui traité avt rupture de matricule
              LaNature := 'SA' + IntToStr(ii);
              ExecuteSQL('DELETE FROM VENTIL WHERE V_NATURE LIKE "' + LaNature + '" AND V_COMPTE="' + SalP + '"');
            end;
          end;
          SalP := CodeS;
          MesAxes := '-----';
          // FIN PT20
        end;
        ii := ValeurI(Copy(T.GetValue('V_NATURE'), 3, 1));
        MesAxes[ii] := 'X'; // DEB PT20
        T := Tob_Ventil.FindNext([''], [''], FALSE);
      end;
      // PT20 Suppression des ventilations analytiques salarie sur axe dans le fichier d'import
      for ii := 1 to 5 do
      begin // Pour traiter le cas du dernier salarié
        if MesAxes[ii] = 'X' then
        begin
          LaNature := 'SA' + IntToStr(ii);
          ExecuteSQL('DELETE FROM VENTIL WHERE V_NATURE LIKE "' + LaNature + '" AND V_COMPTE="' + SalP + '"');
        end;
      end;
      // FIN PT20
      Tob_Ventil.SetAllModifie(TRUE);
      Tob_Ventil.InsertOrUpdateDB(false);
    end;
    // FIN PT13
    // FIN PT2
{$IFNDEF COMSX}
    FiniMoveProgressForm();
{$ENDIF}
    if Tob_ordreHSR <> nil then Tob_ordreHSR.free;
    if Tob_ordreABS <> nil then Tob_ordreABS.free;
    if Tob_S <> nil then Tob_S.free;
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
    if Tob_Ventil <> nil then Tob_Ventil.free;
    if Tob_PaieVentil <> nil then Tob_PaieVentil.free;
    // FIN PT13

    CommitTrans;
  except
    Rollback;
{$IFNDEF COMSX}
    PGIBox('Une erreur est survenue lors de l''écriture dans la base', 'Ecriture fichier import');
{$ENDIF}
  end;

end;

{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.ControleFichierImport(Sender: Tobject);
{$ELSE}

function ControleFichierImport(Sender: Tobject): Boolean;
{$ENDIF}
var
  NomFic, FicRapport: ThEdit;
  S, st: string;
  NoErreur, i, ll, Rep: integer;
  ExerPerEncours, DebPer, FinPer: string;
  Q: TQuery;
  LeFic, LeRapport: string;
begin
  ExisteErreur := false;
  AbsExclue := false; // PT22
  MarqueFinOk := false;
  i := 0;
  if TypImport <> 'AUTO' then
  begin
{$IFNDEF COMSX}
    NomFic := THedit(GetControl('EDTFICIMPORT'));
    if NomFic = nil then exit;
    if NomFic.Text = '' then
    begin
      PGIBox('Fichier d''import obligatoire', 'Import Fichier');
      exit;
    end;
    if not FileExists(NomFic.text) then
    begin
      PGIBox('Fichier d''import inexistant', 'Import Fichier');
      exit;
    end;
    FicRapport := THedit(GetControl('FICRAPPORT'));
    if FicRapport = nil then exit;
    if FicRapport.Text = '' then
    begin
      PGIBox('Fichier de rapport obligatoire', 'Import Fichier');
      exit;
    end;
    if not FileExists(FicRapport.text) then
    begin
      // PT18    PGIBox('Fichier de rapport inexistant', 'Import Fichier'); exit ;
    end;
    valideEcran(false);
    LeFic := NomFic.Text;
    LeRapport := FicRapport.Text;
{$ENDIF}
  end
  else
  begin
    LeFic := NomDuFichier;
    ll := Pos('.NSV', LeFic);
    LeRapport := Copy(LeFic, 1, ll - 1) + '.log'; // nom du fichier de rapport = .log au lieu de txt
    if not FileExists(LeRapport) then
    begin
      AssignFile(FR, LeRapport); // PT24
{$I-}
      ReWrite(FR);
{$I+}
      if IoResult <> 0 then
      begin
{$IFNDEF COMSX}
        PGIBox('Fichier rapport inaccessible : ' + LeFic, 'Abandon du traitement');
{$ELSE}
//        ErrComSx := TRUE;
{$ENDIF}
        Exit;
      end;
      closeFile(FR);
    end;
  end;
  AssignFile(F, LeFic);
  Reset(F);
  AssignFile(FR, LeRapport);
  Reset(FR);
  ReWrite(FR);
  liberreur := true;
  S := '';
  while ((not IsMarqueDebut(s)) and (not eof(F))) do
  begin
    Readln(F, S);
  end;
  if eof(F) then
  begin
    Erreur(s, S, 5);
    CloseFile(F);
    CloseFile(FR);
    Exit;
  end;
{$IFNDEF COMSX}
  InitMoveProgressForm(nil, 'Contrôle du fichier d''import', 'Veuillez patienter SVP ...', taille, False, TRUE);
{$ENDIF}
  if not Eof(F) then
  begin
    i := i + 1;
{$IFNDEF COMSX}
    MoveCurProgressForm(IntToStr(i));
{$ENDIF}
    Readln(F, S);
    Noerreur := ControleEnrDossier(S, S);
    if NoErreur <> 0 then
      Erreur(s, S, NoErreur);
  end;
  ExerPerEncours := '';
  DebPer := '';
  FinPer := '';
  TFinPer := 0;
  if not RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer) then { PT29 }
    Erreur('', '', 7);
  if FinPer <> '' then
    if Isvaliddate(FinPer) then
      TFinPer := strtodate(FinPer) else
      Erreur('', '', 7);
  {PT51 Ajout psa libelle et prenom dans TOB
     st := 'SELECT PSA_SALARIE,PSA_DATEENTREE,PSA_DATESORTIE,PSA_SUSPENSIONPAIE,PSA_ETABLISSEMENT, ' +
    'PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_CONFIDENTIEL,PSA_CONGESPAYES,' + { PT22 }
  { 'ETB_DATECLOTURECPN,ETB_CONGESPAYES FROM SALARIES LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT = ETB_ETABLISSEMENT';}
  st := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE,PSA_DATESORTIE,PSA_SUSPENSIONPAIE,PSA_ETABLISSEMENT, ' +
    'PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_CONFIDENTIEL,PSA_CONGESPAYES,' + { PT22 }
    'ETB_DATECLOTURECPN,ETB_CONGESPAYES FROM SALARIES LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT = ETB_ETABLISSEMENT';
  Q := opensql(st, true);
  if (not Q.eof) then
  begin
    Tob_sal := TOB.Create('Table des Salariés', nil, -1);
    if tob_sal <> nil then
      Tob_sal.loaddetaildb('INFOS SALARIES', '', '', Q, false, false, -1, 0);
  end;
  Ferme(Q);
// PT34 Suppression des lignes de tests si pas de salaries
  st := 'SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_RUBRIQUE <>""';
  Q := opensql(st, true);
  if not Q.eof then
  begin
    Tob_rubrique := TOB.Create('Table des rubriques', nil, -1);
    if tob_rubrique <> nil then
      Tob_rubrique.loaddetaildb('REMUNERATION', '', '', Q, false, false, -1, 0);
  end;
  Ferme(Q);
  //   PT2 09/10/01 PH Rajout traitement ligne Analytique
  if GetParamSoc('SO_PGANALYTIQUE') then // PT24
  begin
    st := 'SELECT S_AXE,S_SECTION FROM SECTION ORDER BY S_AXE,S_SECTION';
    Q := opensql(st, true);
    if not Q.eof then
    begin
      Tob_section := TOB.Create('Table des sections', nil, -1);
      if tob_section <> nil then
        Tob_section.loaddetaildb('SECTION', '', '', Q, false, false, -1, 0);
    end;
    Ferme(Q);
  end;
  // FIN PT2
  while not Eof(F) do //Deuxieme passage
  begin
    i := i + 1;
{$IFNDEF COMSX}
    MoveCurProgressForm(IntToStr(i));
{$ENDIF}
    Readln(F, S);
    if MarqueFinOk then
    begin
      Erreur_10 := true;
      Erreur(s, S, 10);
    end;
    St := S;
    St := readtokenPipe(St, Separateur);
    if St = MarqueFin then
    begin
      MarqueFinOk := true;
      continue
    end;
    Traiteligne(S, S);
  end;
  if not MarqueFinOk then Erreur('', '', 15);
{$IFNDEF COMSX}
  FiniMoveProgressForm;
{$ENDIF}
  st := 'SELECT PIE_TYPEDONNEE,PIE_SALARIE,PIE_LIBELLE,PIE_PRENOM,' +
    'PIE_NUMEROSS,PIE_DATEENTREE,PIE_DATESORTIE,PIE_COMMENTAIRE,' +
    'PIE_RUBOUMOTIF,PIE_LIBELLE2,PIE_BASEJOUR,PIE_TAUXHEURES,PIE_COEFF,' +
    'PIE_MONTANT,PIE_DATEDEBUT,PIE_DATEFIN,PIE_DEBUTDJ,PIE_FINDJ,PIE_STATUT ' +
    ' FROM IMPORTPAIE WHERE PIE_GUID = "' + IDUNIQ + '" ORDER BY PIE_TYPEDONNEE,PIE_SALARIE';
  LanceEtat('E', 'PAY', 'PIE', True, False, False, nil, ST, '', False);
  if not ExisteErreur then
    Rep := PgiAsk('Aucune erreur détectée, voulez-vous continuer le traitement ?', 'Import de données')
  else Rep := mrYes;
  if (not ExisteErreur) and (Rep = mrYes) then
  begin
    Reset(F);
    ImporteFichier();
{$IFNDEF COMSX}
    TFVierge(Ecran).Retour := ''; // PT24
//    LanceEtat('E','PAY','PIE',True,False,False,nil,ST,'',False);
    if RapportErreur <> nil then
    begin
      // PT22   26/05/2004 PH V_50 Non traitement des absences portant sur des périodes closes mias ne provoquant pas de rejet du fichier
      if not AbsExclue then Writeln(FR, 'VOTRE IMPORT S''EST BIEN PASSE')
      else
      begin
        Writeln(FR, 'VOTRE IMPORT S''EST BIEN PASSE mais des lignes absences ont été exclues');
        PGIBox('Des lignes absences ont été exclues car elles sont hors période #13#10' +
          'Pour en consulter la liste, cliquez sur le bouton Consultation du rapport d''import', 'Import fichier');
      end;
      // FIN PT22
            //RapportErreur.Lines.add('VOTRE IMPORT S''EST BIEN PASSE');
      PGIBox('Le fichier a été importé avec succès.#13#10' +
        ' Lancez la génération de bulletins pour l''intégration dans les paies.', 'Import fichier'); // PT31
    end;
  end
  else
  begin
    if ExisteErreur then
    begin
      TFVierge(Ecran).Retour := 'ERR;' + LeRapport; // PT24
      PGIBox('Des erreurs détectées n''ont pas permis l''import de votre fichier.#13#10' +
        ' Pour consulter la liste d''erreurs, cliquez sur le bouton Consultation du rapport d''import', 'Import de fichier'); // PT31
    end;
  end;
{$ELSE}
//    LanceEtat('E','PAY','PIE',True,False,False,nil,ST,'',False);
    if not ExisteErreur then
    begin
      if not AbsExclue then Writeln(FR, 'VOTRE IMPORT S''EST BIEN PASSE')
      else Writeln(FR, 'VOTRE IMPORT S''EST BIEN PASSE mais des lignes absences ont été exclues');
      Result := TRUE;
    end
    else
    begin
      Writeln(FR, 'Des erreurs détectées n''ont pas permis l''import de votre fichier.');
      Result := FALSE;
    end;
  end;
{$ENDIF}
  FermeTout;
  ExecuteSql('DELETE FROM IMPORTPAIE WHERE PIE_GUID="' + Iduniq + '"');
{$IFNDEF COMSX}
  valideEcran(true);
{$ENDIF}
end;

{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.Traiteligne(chaine, S: string);
{$ELSE}

procedure Traiteligne(chaine, S: string);
{$ENDIF}
var
  CodeEnr: string;
begin
  CodeEnr := readtokenPipe(S, Separateur);
  if (CodeEnr = 'MHE') or (CodeEnr = 'MFP') or (CodeEnr = 'MLB') then
    ControleEnregMHE_MFP_MLB(CodeEnr, chaine, S, false)
  else
    if CodeEnr = 'MAB' then ControleEnregMAB(chaine, S, false)
    // PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
    else
      if CodeEnr = 'VEN' then ControleEnregVEN(CodeEnr, chaine, S, false)
      else
    // PT15   25/08/2003 PH Correction prise en compte traitement analytique ANA
        if CodeEnr = 'ANA' then ControleEnregANA(CodeEnr, chaine, S, false)
        else
    // DEB PT27
          if CodeEnr = 'VSA' then ControleEnregVSA(CodeEnr, chaine, S, false)
          else
            if CodeEnr = 'SAL' then ControleEnregSAL(CodeEnr, chaine, S, false)
    // FIN PT27
            else Erreur('', S, 20);
end;
{$IFNDEF COMSX}

function TOF_PGIMPORTFIC.ControleEnrDossier(chaine, S: string): integer;
{$ELSE}

function ControleEnrDossier(chaine, S: string): integer;
{$ENDIF}
var
  NoDoss, DossEnCours, ent: string;
  DatedebES, DateFinES, DebExer, FinExer: TDateTime;
  MoisE, AnneeE, ComboExer, Period: string;
  aa, mm, jj: word;
  rep: Integer;
begin
  ent := readtokenPipe(S, Separateur);
  if ent <> '000' then Erreur(chaine, DossEnCours, 220);
  // **** Refonte accès V_PGI_env ***** V_PGI_env.nodossier remplacé par PgRendNoDossier() *****
  DossEnCours := PGRendNoDossier(); { PT29 }
  NoDoss := readtokenPipe(S, Separateur);
  //DEB PT48
  if (Length(DossEnCours) = 8) and (Length(NoDoss) = 6) then NoDoss := '00' + NoDoss;
  if (DossEnCours <> NoDoss) and (DossEnCours <> '000000') and (DossEnCours <> '') then
    if (NoDoss <> '000000') then Erreur(chaine, DossEnCours, 230);
  //FIN PT48
  // format US vers format Français
  {DEB PT12 transtypage incorrect si S=''}
  DecodeDate(idate1900, aa, mm, jj);
  if S <> '' then jj := strtoint(readtokenPipe(S, '/'));
  if S <> '' then mm := strtoint(readtokenPipe(S, '/'));
  if S <> '' then aa := strtoint(readtokenPipe(S, Separateur));
  DatedebES := PGEncodeDateBissextile(aa, mm, jj); { PT21 } // PT25  { PT29 }
  DecodeDate(idate1900, aa, mm, jj);
  if S <> '' then jj := strtoint(readtokenPipe(S, '/'));
  if S <> '' then mm := strtoint(readtokenPipe(S, '/'));
  if S <> '' then aa := strtoint(readtokenPipe(S, Separateur));
  DatefinES := PGEncodeDateBissextile(aa, mm, jj); { PT21 } // PT25  { PT29 }
  {FIN PT12}

  MoisE := '';
  AnneeE := '';
  ComboExer := '';
  DebExer := 0;
  FinExer := 0;
  RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer); { PT29 }
  if ((datedebES <> DebExer) or (dateFinES <> FinExer)) and ((datedebES > Idate1900) and (dateFinES > Idate1900)) then
  begin
{$IFNDEF COMSX}
    Rep := PGIAsk('Les dates d''exercice social du fichier sont différentes de celles de votre entreprise#13#10Voulez-vous quand même prendre en compte votre fichier?',
      'Incohérence des dates d''exercice social.');
    if rep <> mrYes then Erreur(chaine, DossEnCours, 240);
{$ELSE}
    Erreur(chaine, DossEnCours, 240);
{$ENDIF}
  end;
  result := 0;
  Period := readtokenPipe(S, Separateur); // PT42
  Period := readtokenPipe(S, Separateur); // PT42
  VersionFic := readtokenPipe(S, Separateur); // PT42 Récup Numuéro de version du fichier
end;
{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.ControleEnregMHE_MFP_MLB(TypeEnr, chaine, S: string; Ecrit: boolean);
{$ELSE}

procedure ControleEnregMHE_MFP_MLB(TypeEnr, chaine, S: string; Ecrit: boolean);
{$ENDIF}
var
  Codesalarie, DateDebutperiode, DateFinperiode, SNoOrdre: string;
  Rubrique, typeAlim, libelle, sBase, staux, sCoefficient, sMontant, ll: string;
  TDateDeb, TDateFin: tdatetime;
  Base, taux, coeff, montant: double;
  Sal, TS, T: tob;
  Ordre, TopCloture: integer;
  st, dudate, audate, LetypEnr: string; // PT54
  Q: TQuery; // PT36
  Nom, prenom: string; //PT51
  TheSal : TOB;  // PT54
begin
  libErreur := false;
  libanomalie := false;
  Statutenreg := ''; //GGR PT35
  Ordre := 1;
  // contrôle sur le salarié
  Codesalarie := readtokenPipe(S, Separateur);
  if (CodeSalarie <> Salarieprec) or (POS('S1', CodeSalarie) > 0) then // PT32 Pour traiter le cas du 1er // PT47
  begin
    ControleSalarie(TypeEnr, chaine, S, CodeSalarie, Ecrit);
  end;
  Salarieprec := Codesalarie;
  DateDebutperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DatedebutPeriode) then
    Erreur(chaine, DatedebutPeriode, 30);
  DateFinperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DateFinPeriode) then
    Erreur(chaine, DateFinPeriode, 35);

  TDateDeb := 0;
  TDateFin := 0;

  if (Isvaliddate(DatedebutPeriode)) and (Isvaliddate(DatefinPeriode)) then
  begin
    TDateDeb := strtodate(DatedebutPeriode);
    TDateFin := strtodate(DatefinPeriode);
  end;
  Dudate := UsDateTime(iDate1900); // PT44
  audate := UsDateTime(iDate1900); // PT44
  if TDateFin < TDateDeb then
    Erreur(chaine, DateDebutPeriode + '//' + DateFinPeriode, 40);

  Topcloture := IsPeriodecloture(TDatedeb, TDateFin);
  // PT22   26/05/2004 PH V_50 Non traitement des absences portant sur des périodes closes mias ne provoquant pas de rejet du fichier
  if not Ecrit then ll := 'NF' else ll := '';
  if TopCloture = 1 then Erreur(Chaine, '', 44)
  else
    if TopCloture = 2 then Erreur(Chaine, '', 47, ll)
    else
      if TopCloture = 3 then Erreur(Chaine, '', 48, ll);
  // FIN PT22
  if (DateSortie > 10) and (DateSortie < TDateDeb) then
    Erreur(Chaine, '', 46);

  SNoOrdre := readtokenPipe(S, Separateur);
  if SnoOrdre = '' then
    Erreur(chaine, SNoOrdre, 45);

  Rubrique := readtokenPipe(S, Separateur);
  if TypeEnr <> 'MLB' then RubPrec := Rubrique;

  TypeBase := '';
  TypeTaux := '';
  TypeCoeff := '';
  TypeMontant := '';
  Libellerubrique := ''; //GGR PT35
  if (TypeEnr = 'MHE') or (typeEnr = 'MFP') then
  begin
    if not ExisteRubrique(Rubrique) then
      Erreur(chaine, Rubrique, 50);
  end
  else
  begin // Ligne de type MLB
    if (Copy(Rubrique, 1, 4)) <> rubPrec then // PT54
      Erreur(Chaine, Rubrique, 52);
  end;
  libelle := readtokenPipe(S, Separateur);

  TypeAlim := readtokenPipe(S, Separateur);
  if (TypeAlim = '') and (TypeEnr <> 'MLB') then
    Erreur(chaine, Rubrique, 55);
  // DEB PT26
  if ((TypeBase <> '00') and (TypeBase <> '01') and (pos('B', typeAlim) > 0) or
    (((typeBase = '00') or (typeBase = '01')) and (pos('B', typeAlim) = 0)) or
    ((TypeTaux <> '00') and (TypeTaux <> '01') and (pos('T', typeAlim) > 0)) or
    (((typeTaux = '00') or (typeTaux = '01')) and (pos('T', typeAlim) = 0)) or
    ((TypeCoeff <> '00') and (TypeCoeff <> '01') and (pos('C', typeAlim) > 0)) or
    (((typeCoeff = '00') or (typeCoeff = '01')) and (pos('C', typeAlim) = 0)) or
    ((TypeMontant <> '00') and (TypeMontant <> '01') and (pos('M', typeAlim) > 0)) or
    (((typeMontant = '00') or (typeMontant = '01')) and (pos('M', typeAlim) = 0))) and (TypeEnr <> 'MLB') then
    begin
{$IFDEF EWS}
      if (NOT GetParamsocSecur('SO_NE_EWSACTIF', False)) OR (Not IsDossierEws (PGRendNoDossier())) then Erreur(chaine, Rubrique, 57); // PT59
{$ELSE}
      Erreur(chaine, Rubrique, 57);
{$ENDIF}
    end;
  // FIN PT26
  if TypeEnr = 'MLB' then  // DEB PT54
  begin
    sbase := '0';
    staux := '0';
    sCoefficient := '0';
    sMontant := '0';
    Libellerubrique := Libelle;
  end
  else
  begin // Traitement des champs pour type enreg <> MLB // FIN PT54
    sBase := readtokenPipe(S, Separateur);
    if pos('B', typeAlim) > 0
      then
    begin
      if sbase = '' then
        Erreur(chaine, sBase, 60)
      else
        if not Isnumeric(sBase) then
          Erreur(chaine, sBase, 60);
    end
    else
    begin
      if sBase <> '' then
        Erreur(chaine, sBase, 65)
      else sBase := '0';
    end;

    staux := readtokenPipe(S, Separateur);
    if pos('T', typeAlim) > 0
      then
    begin
      if staux = '' then
        Erreur(chaine, staux, 70)
      else
        if not Isnumeric(staux) then
          Erreur(chaine, staux, 70);
    end
    else
    begin
      if staux <> '' then
        Erreur(chaine, staux, 75)
      else
        Staux := '0';
    end;

    sCoefficient := readtokenPipe(S, Separateur);
    if pos('C', typeAlim) > 0
      then
    begin
      if sCoefficient = '' then
        Erreur(chaine, sCoefficient, 80)
      else
        if not Isnumeric(sCoefficient) then
          Erreur(chaine, sCoefficient, 80);
    end
    else
    begin
      if sCoefficient <> '' then
        Erreur(chaine, sCoefficient, 85);
    end;

    sMontant := readtokenPipe(S, Separateur);
    if pos('M', typeAlim) > 0
      then
    begin
      if sMontant = '' then
        Erreur(chaine, sMontant, 90)
      else
        if not Isnumeric(sMontant) then
        begin
          Erreur(chaine, sMontant, 90);
          sMontant := '0';
        end
    end
    else
    begin
      if sMontant <> '' then
        Erreur(chaine, sMontant, 95) else Smontant := '0';
    end;
  end; // Fin de traitement champ enreg de type MHE  PT54
//GGS PT35  if not Ecrit then exit;
  if not ecrit
    then
  begin
    if iduniq = '' then iduniq := AglGetGuid();
    Datedebutperiode := UsDatetime(Tdatedeb);
    DateFinPeriode := UsDateTime(Tdatefin);
    if libanomalie then statutenreg := 'Anomalie';
    if liberreur then statutenreg := 'Erreur';
    NumOrdre := NumOrdre + 1;
//debut PT51
    Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [CodeSalarie], true);
    // DEB PT54
    if Sal = nil then
    begin
      if Tob_SalATT <> nil then TheSal := Tob_SalATT.findfirst(['PIM_SALARIE'], [Codesalarie], true)
      else TheSal := nil;
      if TheSal = nil then
      begin
        Erreur(chaine, Codesalarie, 25);
        exit;
      end;
    end;
    if Sal <> nil then
    begin
      nom := Sal.Getvalue('PSA_LIBELLE');
      prenom := Sal.Getvalue('PSA_PRENOM');
    end
    else
    begin
      nom := TheSal.Getvalue('PIM_LIBELLE');
      prenom := TheSal.Getvalue('PIM_PRENOM');
    end;
// FIN PT54
//fin PT51
    // DEB PT46
{    st := 'insert into IMPORTPAIE(PIE_GUID,PIE_TYPEDONNEE,PIE_SALARIE,PIE_NUMORDRE,PIE_LIBELLE,PIE_PRENOM,PIE_NUMEROSS,' +
      'PIE_DATEENTREE,PIE_DATESORTIE,PIE_COMMENTAIRE,PIE_RUBOUMOTIF,PIE_LIBELLE2,PIE_BASEJOUR,PIE_TAUXHEURES,' +
      'PIE_COEFF,PIE_MONTANT,PIE_DATEDEBUT,PIE_DATEFIN,PIE_DEBUTDJ,PIE_FINDJ,PIE_STATUT) VALUES ("' +
      IDUNIQ + '","VAR","' + Codesalarie + '",' + IntToStr(NumOrdre) + ',"","","","' + Dudate + '","' + audate + '","","' + Rubrique + '","' + Libellerubrique + '",' +
       StrFPoint(Valeur(Sbase)) + ',' +  StrFPoint(Valeur(Staux)) + ',"' +  StrFPoint(Valeur(Scoefficient)) + '",' + StrFPoint(Valeur(Smontant)) + ',"' + DateDebutPeriode + '","' + DateFinPeriode +
      '"," "," ","' + statutenreg + '")';}
// FIN PT46
    LetypEnr := 'VAR'; // PT54
    st := 'insert into IMPORTPAIE(PIE_GUID,PIE_TYPEDONNEE,PIE_SALARIE,PIE_NUMORDRE,PIE_LIBELLE,PIE_PRENOM,PIE_NUMEROSS,' +
      'PIE_DATEENTREE,PIE_DATESORTIE,PIE_COMMENTAIRE,PIE_RUBOUMOTIF,PIE_LIBELLE2,PIE_BASEJOUR,PIE_TAUXHEURES,' +
      'PIE_COEFF,PIE_MONTANT,PIE_DATEDEBUT,PIE_DATEFIN,PIE_DEBUTDJ,PIE_FINDJ,PIE_STATUT) VALUES ("' +
      IDUNIQ + '","' + LeTypEnr + '","' + Codesalarie + '",' + IntToStr(NumOrdre) + ',"' + nom + '","' + prenom + '","","' + Dudate + '","' + audate + '","","' + Copy(Rubrique, 1, 4) + '","' + Libellerubrique + '",' +
      StrFPoint(Valeur(Sbase)) + ',' + StrFPoint(Valeur(Staux)) + ',"' + StrFPoint(Valeur(Scoefficient)) + '",' + StrFPoint(Valeur(Smontant)) + ',"' + DateDebutPeriode + '","' + DateFinPeriode +
      '"," "," ","' + statutenreg + '")';
    ExecuteSQL(St);
    exit;
  end;
//FIN ggs PT35
// DEB PT54
  Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [CodeSalarie], true);
  if Sal = nil then
  begin
    if Tob_SalATT <> nil then TheSal := Tob_SalATT.findfirst(['PIM_SALARIE'], [Codesalarie], true)
    else TheSal := nil;
    if TheSal = nil then
    begin
      Erreur(chaine, Codesalarie, 25);
      exit;
    end
    else
    begin
      Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [TheSal.getValue('CodeS5')], true);
      if Sal = nil then
      begin
        Erreur(chaine, Codesalarie, 25);
        exit;
      end;
    end;
  end;
// FIN PT54
//  TOBDEBUG (Tob_OrdreHSR);
  TS := Tob_OrdreHSR.findfirst(['PSD_SALARIE'], [CodeSalarie], true);
  if TS <> nil then
  begin
  // DEB PT36
    st := 'SELECT PSD_ORDRE FROM HISTOSAISRUB WHERE PSD_SAlARIE="' + CodeSalarie + '" AND PSD_DATEDEBUT="' +
      UsDateTime(TDateDeb) + '" AND PSD_DATEFIN="' + UsDateTime(TDateFin) + '" AND PSD_RUBRIQUE="' + Rubrique +
      '" AND PSD_ORIGINEMVT="' + TypeEnr + '"';
    Q := OpenSql(st, true);
    if not Q.EOF then Ordre := Q.FindField('PSD_ORDRE').asinteger
    else Ordre := TS.getvalue('PSD_ORDRE') + 1;
    Ferme(Q);
    TS.putvalue(('PSD_ORDRE'), Ordre);
    // FIN PT36
  end
  else
  begin
    TS := Tob.create('HISTOSAISRUB', Tob_OrdreHSR, -1);
    if Ts <> nil then
    begin
      Ts.putvalue('PSD_SALARIE', CodeSalarie);
      Ts.putValue('PSD_Ordre', 1);
    end;
  end;
  if not Ecrit then exit;
  Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [CodeSalarie], true);
  if (Sal = nil) then exit;
  T := Tob.create('HISTOSAISRUB', TOB_HSR, -1);
  if T = nil then exit;
  TS := Tob_OrdreHSR.findfirst(['PSD_SALARIE'], [CodeSalarie], true);
  if TS <> nil then
  begin
    // DEB PT36
    st := 'SELECT PSD_ORDRE FROM HISTOSAISRUB WHERE PSD_SAlARIE="' + CodeSalarie + '" AND PSD_DATEDEBUT="' +
      UsDateTime(TDateDeb) + '" AND PSD_DATEFIN="' + UsDateTime(TDateFin) + '" AND PSD_RUBRIQUE="' + Rubrique +
      '" AND PSD_ORIGINEMVT="' + TypeEnr + '"';
    Q := OpenSql(st, true);
    if not Q.EOF then Ordre := Q.FindField('PSD_ORDRE').asinteger
    else Ordre := TS.getvalue('PSD_ORDRE') + 1;
    Ferme(Q);
    TS.putvalue(('PSD_ORDRE'), Ordre);
    // FIN PT36
  end
  else
  begin
    TS := Tob.create('HISTOSAISRUB', Tob_OrdreHSR, -1);
    if Ts <> nil then
    begin
      Ts.putvalue('PSD_SALARIE', CodeSalarie);
      Ts.putValue('PSD_Ordre', 1);
    end;
  end;
  Base := 0;
  Taux := 0;
  Coeff := 0;
  Montant := 0;
  T.PutValue('PSD_ORIGINEMVT', TypeEnr);
  T.PutValue('PSD_SALARIE', CodeSalarie);
  T.PutValue('PSD_DATEDEBUT', TDateDeb);
  T.PutValue('PSD_DATEFIN', TDateFin);
  T.PutValue('PSD_RUBRIQUE', Rubrique);
  T.PutValue('PSD_ORDRE', Ordre);
  T.PutValue('PSD_LIBELLE', Libelle);
  T.PutValue('PSD_RIBSALAIRE', '');
  T.PutValue('PSD_BANQUEEMIS', '');
  T.PutValue('PSD_TOPREGLE', '-');
  T.PutValue('PSD_TYPALIMPAIE', typeAlim);
  T.PutValue('PSD_ETABLISSEMENT', sal.getvalue('PSA_ETABLISSEMENT'));
  if pos('B', typeAlim) > 0 then Base := Valeur(sBase);
  if pos('T', typeAlim) > 0 then Taux := Valeur(sTaux);
  if pos('C', typeAlim) > 0 then Coeff := Valeur(sCoefficient);
  if pos('M', typeAlim) > 0 then Montant := Valeur(sMontant);

  T.PutValue('PSD_BASE', Base);
  T.PutValue('PSD_TAUX', Taux);
  T.PutValue('PSD_COEFF', Coeff);
  T.PutValue('PSD_MONTANT', Montant);
  T.PutValue('PSD_DATEINTEGRAT', IDate1900); // PT38
  T.PutValue('PSD_DATECOMPT', IDate1900); // PT38
  T.PutValue('PSD_AREPORTER', '');
  T.Putvalue('PSD_CONFIDENTIEL', sal.getvalue('PSA_CONFIDENTIEL'));

end;
//   PT2 09/10/01 PH Rajout traitement ligne Analytique Salarie/rubrique
{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.ControleEnregANA(TypeEnr, chaine, S: string; Ecrit: boolean);
{$ELSE}

procedure ControleEnregANA(TypeEnr, chaine, S: string; Ecrit: boolean);
{$ENDIF}
var
  Codesalarie, DateDebutperiode, DateFinperiode, SNoOrdre: string;
  Rubrique, Axe, Section, CodeAxe, sMontant, LibSect: string;
  TDateDeb, TDateFin: tdatetime;
  T: tob;
  TopCloture: integer;
begin
  libErreur := false;
  libanomalie := false;
  if not GetParamSoc('SO_PGANALYTIQUE') then
  begin
    Erreur(chaine, '', 101); // PT24
    exit;
  end;
  // contrôle sur le salarié
  Codesalarie := readtokenPipe(S, Separateur);
  if (CodeSalarie <> Salarieprec) or (POS('S1', CodeSalarie) > 0) then // PT32 // PT47
  begin
    ControleSalarie(TypeEnr, chaine, S, CodeSalarie, Ecrit);
  end;
  Salarieprec := Codesalarie;
  DateDebutperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DatedebutPeriode) then
    Erreur(chaine, DatedebutPeriode, 30);
  DateFinperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DateFinPeriode) then
    Erreur(chaine, DateFinPeriode, 35);
  TDateDeb := 0;
  TDateFin := 0;
  if (Isvaliddate(DatedebutPeriode)) and (Isvaliddate(DatefinPeriode)) then
  begin
    TDateDeb := strtodate(DatedebutPeriode);
    TDateFin := strtodate(DatefinPeriode);
  end;
  if TDateFin < TDateDeb then
    Erreur(chaine, DateDebutPeriode + '//' + DateFinPeriode, 40);
  Topcloture := IsPeriodecloture(TDatedeb, TDateFin);
  if TopCloture = 1 then
    Erreur(Chaine, '', 44);
  if TopCloture = 2 then
    Erreur(Chaine, '', 47);
  if TopCloture = 3 then
    Erreur(Chaine, '', 48);
  if (DateSortie > 10) and (DateSortie < TDateDeb) then
    Erreur(Chaine, '', 46);

  Rubrique := readtokenPipe(S, Separateur);
  RubPrec := Rubrique;
  if not ExisteRubrique(Rubrique) then Erreur(chaine, Rubrique, 50);
  if Rubrique <> rubPrec then Erreur(Chaine, Rubrique, 52);
  Axe := readtokenPipe(S, Separateur);
  if (Axe < 'A1') or (Axe > 'A5') then Erreur(chaine, Axe, 103);
  Section := readtokenPipe(S, Separateur);

  Smontant := readtokenPipe(S, Separateur);
  if sMontant = '' then
    Erreur(chaine, sMontant, 90)
  else
    if not Isnumeric(sMontant) then Erreur(chaine, sMontant, 90);
  SNoOrdre := readtokenPipe(S, Separateur);
  if SNoOrdre = '' then
    Erreur(chaine, SNoOrdre, 102)
  else
    if not Isnumeric(SNoOrdre) then Erreur(chaine, SNoOrdre, 102);

  // PT16   28/08/2003 PH Création des sections en récupérant le libellé de la section
  if S <> '' then LibSect := readtokenPipe(S, Separateur)
  else LibSect := '';
  if not ExisteSection(Axe, Section, LibSect) then Erreur(Chaine, Section, 100);
  // Fin PT16

  if not Ecrit then exit;
  T := Tob.create('PAIEVENTIL', TOB_Paieventil, -1);
  if T = nil then exit;
  CodeAxe := Copy(Axe, 2, 1);
  T.PutValue('PAV_NATURE', 'PG' + CodeAxe);
  T.PutValue('PAV_COMPTE', CodeSalarie + ';' + Rubrique);
  T.PutValue('PAV_DATEDEBUT', strtodate(DateDebutperiode));
  T.PutValue('PAV_DATEFIN', strtodate(DateFinperiode));
  T.PutValue('PAV_SECTION', Section);
  T.PutValue('PAV_TAUXMONTANT', Valeur(sMontant));
  T.PutValue('PAV_TAUXQTE1', 0);
  T.PutValue('PAV_TAUXQTE2', 0);
  T.PutValue('PAV_NUMEROVENTIL', StrToInt(SNoOrdre));
  T.PutValue('PAV_SOCIETE', GetParamSoc('SO_SOCIETE'));
  T.PutValue('PAV_MONTANT', 0);
  T.PutValue('PAV_SOUSPLAN1', '');
  T.PutValue('PAV_SOUSPLAN2', '');
  T.PutValue('PAV_SOUSPLAN3', '');
  T.PutValue('PAV_SOUSPLAN4', '');
  T.PutValue('PAV_SOUSPLAN5', '');
  T.PutValue('PAV_SOUSPLAN6', '');

end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 13/03/2003
Modifié le ... : 13/03/2003
Description .. : Traitement d'une ligne du fichier import qui remplacera les
Suite ........ : ventilations analytiques salarié qui sont dans la table
Suite ........ : VENTIL.
Suite ........ :
Suite ........ : Elles correspondent  aux pré-ventilations analytiques.
Mots clefs ... :
*****************************************************************}
//PT13   13/03/2003 PH Traitement des lignes import pré-ventilations analytiques salariés
{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.ControleEnregVEN(TypeEnr, chaine, S: string; Ecrit: boolean);
{$ELSE}

procedure ControleEnregVEN(TypeEnr, chaine, S: string; Ecrit: boolean);
{$ENDIF}
var
  Codesalarie, SNoOrdre, LibSect: string;
  Axe, Section, CodeAxe, sMontant: string;
  T: tob;
begin
  libErreur := false;
  libanomalie := false;
  if not GetParamSoc('SO_PGANALYTIQUE') then Erreur(chaine, '', 101); // PT24
  // contrôle sur le salarié
  Codesalarie := readtokenPipe(S, Separateur);
  if (CodeSalarie <> Salarieprec) or (POS('S1', CodeSalarie) > 0) then // PT32 // PT47
    ControleSalarie(TypeEnr, chaine, S, CodeSalarie, Ecrit);
  Salarieprec := Codesalarie;
  Axe := readtokenPipe(S, Separateur);
  if (Axe < 'A1') or (Axe > 'A5') then Erreur(chaine, Axe, 103);
  Section := readtokenPipe(S, Separateur);
  Smontant := readtokenPipe(S, Separateur);
  if sMontant = '' then
    Erreur(chaine, sMontant, 90)
  else
    if not Isnumeric(sMontant) then Erreur(chaine, sMontant, 90);
  SNoOrdre := readtokenPipe(S, Separateur);
  if SNoOrdre = '' then
    Erreur(chaine, SNoOrdre, 102)
  else
    if not Isnumeric(SNoOrdre) then Erreur(chaine, SNoOrdre, 102);
  // PT16   28/08/2003 PH Création des sections en récupérant le libellé de la section
  if S <> '' then LibSect := readtokenPipe(S, Separateur)
  else LibSect := '';
  if not ExisteSection(Axe, Section, LibSect) then Erreur(Chaine, Section, 100);
  // Fin PT16
  if not Ecrit then exit;
  T := Tob.create('VENTIL', TOB_Ventil, -1);
  if T = nil then exit;
  CodeAxe := Copy(Axe, 2, 1);
  T.PutValue('V_NATURE', 'SA' + CodeAxe);
  T.PutValue('V_COMPTE', CodeSalarie);
  T.PutValue('V_SECTION', Section);
  T.PutValue('V_TAUXMONTANT', Valeur(sMontant));
  T.PutValue('V_TAUXQTE1', 0);
  T.PutValue('V_TAUXQTE2', 0);
  T.PutValue('V_NUMEROVENTIL', StrToInt(SNoOrdre));
  T.PutValue('V_SOCIETE', GetParamSoc('SO_SOCIETE'));
  T.PutValue('V_MONTANT', 0);
  T.PutValue('V_SOUSPLAN1', '');
  T.PutValue('V_SOUSPLAN2', '');
  T.PutValue('V_SOUSPLAN3', '');
  T.PutValue('V_SOUSPLAN4', '');
  T.PutValue('V_SOUSPLAN5', '');
  T.PutValue('V_SOUSPLAN6', '');
end;

// Test existance section axe analytique
{$IFNDEF COMSX}

function TOF_PGIMPORTFIC.ExisteSection(Axe, Section, LibSect: string): boolean;
{$ELSE}

function ExisteSection(Axe, Section, LibSect: string): boolean;
{$ENDIF}
var
  T, T1: tob;
  ll: Integer;
begin
  result := false;
  //  if not PgImpnotVide(tob_section, true) then exit; // PT24
  T := Tob_section.findfirst(['S_AXE', 'S_SECTION'], [Axe, Section], true);
  if T = nil then
  begin
    // PT14   04/06/2003 PH Création automatique des sections analytiques lors de l'import si autorisé
    if GetParamSoc('SO_PGCREATIONSECTION') then // PT24
    begin
      T := TOB.Create('LES SECTIONS', nil, -1);
      // PT15   25/08/2003 PH Correction creation automatique des sections analytiques
      T1 := TOB.Create('SECTION', T, -1);
      T1.PutValue('S_SECTION', Section);
      //PT16   28/08/2003 PH Création des sections en récupérant le libellé de la section
      if LibSect <> '' then
      begin
        T1.PutValue('S_LIBELLE', LibSect);
        ll := StrLen(PChar(LibSect));
        if ll > 17 then ll := 17;
        if ll > 0 then T1.PutValue('S_ABREGE', Copy(LibSect, 1, ll))
        else T1.PutValue('S_ABREGE', 'Section paie ');
      end
      else
      begin
        T1.PutValue('S_LIBELLE', 'Section paie ' + Section);
        T1.PutValue('S_ABREGE', 'Section paie ');
      end;
      // FIN PT16
      T1.PutValue('S_SENS', 'M');
      T1.PutValue('S_SOLDEPROGRESSIF', 'X');
      T1.PutValue('S_AXE', Axe);
      T.InsertDB(nil, FALSE);
      T1.ChangeParent(Tob_section, -1);
      Tob_section.Detail.Sort('S_AXE;S_SECTION');
      FreeAndNil(T);
      // PT15   25/08/2003 PH Creation section OK
      result := true;
    end
    else // AB 16/10/03
      Writeln(FR, 'Vous n''êtes pas autorisé à créer des sections analytiques dans les paramètres société comptable PAIE.'); // FQ 14483
    // FIN PT14
  end
  else result := true;
end;

// FIN PT2
// Fonction de récupération d'un mvt de type Absence
{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.ControleEnregMAB(chaine, S: string; Ecrit: boolean);
{$ELSE}

procedure ControleEnregMAB(chaine, S: string; Ecrit: boolean);
{$ENDIF}
var
  CodeSalarie, DateDebutAbs, DateFinAbs, st: string;
  TDateDeb, TDateFin, Dateclot: tdatetime;
  NbJ, NbH, MotifAbs, Libelle: string;
  Q: tquery;
  T, Sal: tob;
  Ordre, TopCloture: integer;
  LibCompl1, LibCompl2, Dj, Fj, ll: string;
  OkOk, SalSorti: Boolean;
  TAtt : TOB; //PT51
// d PT33
  GestionIJSS, AbsNonTrait : boolean;    //PT57
  CarenceIjss, WCarence: integer;
  Tob_AbsIJSS, AbsIJSS: TOB;
  WDateDeb, WDateFin: tdatetime;
// f PT33
  dudate, audate, Vnbj, Vnbh: string; //GGR PT35
  Nodjm, Nodjp, l1: integer; { PT39 }
  NB_J, NB_H: Double; { PT39 }
  Etab: string; { PT39 }
  Nom, Prenom: string; //PT51
begin
  AbsNonTrait := FALSE;  //PT57
  libErreur := false;
  libanomalie := false;
  CarenceIjss := 0; // PT33
  Dudate := UsDateTime(iDate1900); // PT44
  audate := UsDateTime(iDate1900); // PT44
  StatutEnreg := '';
  // contrôle sur le salarié
  Codesalarie := readtokenPipe(S, Separateur);
  if (CodeSalarie <> Salarieprec) or (POS('S1', CodeSalarie) > 0) then // PT32 //PT47
  begin
    ControleSalarie('MAB', Chaine, S, CodeSalarie, Ecrit);
  end;
  Salarieprec := Codesalarie;
  DateDebutAbs := readtokenPipe(S, Separateur);
  if not Isvaliddate(DateDebutAbs) then
    Erreur(chaine, DateDebutAbs, 110);
  DateFinAbs := readtokenPipe(S, Separateur);
  if not Isvaliddate(DateFinAbs) then
    Erreur(chaine, DateFinAbs, 120);
  TDateDeb := 0;
  TDateFin := 0;
  if (Isvaliddate(DateDebutAbs)) and (Isvaliddate(DateFinAbs)) then
  begin
    TDateDeb := strtodate(DateDebutAbs);
    TDateFin := strtodate(DateFinAbs);
  end;
  if TDateFin < TDateDeb then
    Erreur(chaine, DateDebutAbs + '//' + DateFinAbs, 130);
  //  PT8 15/02/02 PH test pour les mvts de cp sur mois clos N et N+1
  Topcloture := IsPeriodecloture(TDatedeb, TDateFin, 'X');
  // PT22   26/05/2004 PH V_50 Non traitement des absences portant sur des périodes closes mias ne provoquant pas de rejet du fichier
  if not Ecrit then ll := 'NF' else ll := '';
  if TopCloture = 1 then Erreur(Chaine, '', 44)
  else
    if TopCloture = 2 then Erreur(Chaine, '', 47, ll)
    else
      if TopCloture = 3 then Erreur(Chaine, '', 48, ll)
      else
        if TopCloture = 4 then Erreur(Chaine, '', 49, ll); { PT41 }

  // FIN PT22
  if (DateSortie > 10) and (DateSortie < TDateDeb) then Erreur(Chaine, '', 46);
  //  PT11-1 02/12/2002 PH Regroupement lecture du fichier pour controle existance absence avec
  //  les bornes matin après midi
  Vnbj := '0.00'; //GGS pt35
  Vnbh := '0.00'; //GGS pt35
  NbJ := readtokenPipe(S, Separateur);
  NbH := readtokenPipe(S, Separateur);
  MotifAbs := readtokenPipe(S, Separateur);
  Libelle := readtokenPipe(S, Separateur);
  // PT1 14/06/01 PH Rajout récup libellé complémentaires
  LibCompl1 := readtokenPipe(S, Separateur);
  LibCompl2 := readtokenPipe(S, Separateur);
  Dj := readtokenPipe(S, Separateur);
  Fj := readtokenPipe(S, Separateur);
  if Dj = '' then Dj := 'MAT';
  if Fj = '' then Fj := 'PAM';
{ DEB PT39 }
  if Assigned(Tob_MotifAbs) then
    T_MotifAbs := Tob_MotifAbs.FindFirst(['PMA_MOTIFABSENCE'], [Motifabs], False);

  if not assigned(T_MotifAbs) then Erreur(Chaine, MotifAbs, 170);

    { Calcul du nombre de jour et d'heure de l'absence }
  Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [CodeSalarie], true);
  if Assigned(Sal) then Etab := Sal.getvalue('PSA_ETABLISSEMENT') else Etab := '';
{$IFNDEF COMSX}
  if GetControltext('CALABS') = 'X' then
{$ELSE}
  if GetParamSocSecur('LeParamSocQuiVaBien', TRUE) then
{$ENDIF}
  begin
    AffecteNodemj(Dj, Nodjm);
    AffecteNodemj(Fj, Nodjp);
    CalculNbJourAbsence(TDateDeb, TDateFin, CodeSalarie, Etab,
      MotifAbs, T_MotifAbs, Nb_J, Nb_H, Nodjm, Nodjp);
    NbJ := FloatToStr(Nb_J);
    NbH := FloatToStr(Nb_H);
  end
  else
  begin
    Nb_J := Valeur(NBJ);
    Nb_H := Valeur(NBH);
  end;

  ConversionJourHeure(chaine, NBj, Nbh, VNbj, VNbh);

    { Contôle Gestion maximun }
  if ((MotifAbs <> 'PRI') and (T_MotifAbs <> nil)) then // PT46
  begin
    if ControleGestionMaximum(CodeSalarie, MotifAbs, T_MotifAbs, TDateDeb, Nb_J, Nb_H) then
    begin
      if T_MotifAbs.GetValue('PMA_JOURHEURE') = 'JOU' then
      begin
        Erreur(chaine, '', 155);
        Exit;
      end
      else
        //PT56
        if (T_MotifAbs.GetValue('PMA_JOURHEURE') = 'HEU') or (T_MotifAbs.GetValue('PMA_JOURHEURE') = 'HOR') then
        begin
          Erreur(chaine, '', 156);
          Exit;
        end;
    end;
  end;
{ FIN PT39 }
  //  PT11 02/12/2002 PH Modification de la requete pour prendre en compte les demi journées
  St := PgGetClauseAbsAControler(T_MotifAbs, Tob_MotifAbs, MotifAbs); { PT43 }

  st := 'SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_TYPECONGE,PCN_TYPEMVT FROM ABSENCESALARIE ' +
    ' WHERE PCN_SALARIE = "' + CodeSalarie + '" AND PCN_DEBUTDJ="' + DJ + '" AND PCN_FINDJ="' + FJ + '"' +
   // ' AND (PCN_TYPECONGE = "PRI" OR (PCN_TYPEMVT = "ABS" AND PCN_SENSABS="-"))' + { PT43 }
  St + ' AND PCN_ETATPOSTPAIE <> "NAN" ' + { PT40 }
    ' AND ((PCN_DATEDEBUTABS >="' + usdatetime(TDateDeb) + '" AND PCN_DATEDEBUTABS <= "' +
    usdatetime(TDateFin) + '")' +
    'OR (PCN_DATEFINABS >="' + usdatetime(TDateDeb) + '" AND PCN_DATEDEBUTABS <= "' +
    usdatetime(TDateFin) + '")' +
    'OR (PCN_DATEFINABS >="' + usdatetime(TDateFin) + '" AND PCN_DATEDEBUTABS <= "' +
    usdatetime(TDateDeb) + '")' +
    'OR(PCN_DATEFINABS <="' + usdatetime(TDateFin) + '" AND PCN_DATEDEBUTABS >= "' +
    usdatetime(TDateDeb) + '"))';
  // FIN PT11-1
  OkOk := TRUE;
  SalSorti := FALSE;
  Q := opensql(st, True);
  if not Q.eof then
  begin // Absence déjà existante
    Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [CodeSalarie], true);
    // PT28 Rajout du test si Tob_SalATT <> NIL
    if Tob_SalATT <> nil then TAtt := Tob_SalATT.findfirst(['PIM_SALARIE'], [Codesalarie], true)
    else TAtt := nil;
    //  PT11-2 02/12/2002 PH si les absences sont déjà saisies et que le salarié est sorti alors OK
    if (Sal = nil) and (TAtt = nil) then OkOk := FALSE // PT27
    else
    begin
      if (TAtt = nil) then
      begin
        if ((Sal.GetValue('PSA_DATESORTIE')) <> NULL) and (Sal.GetValue('PSA_DATESORTIE') > Idate1900)
          and (TDateFin > Sal.GetValue('PSA_DATESORTIE')) then { PT45 }
          SalSorti := TRUE else OkOk := FALSE;
      end;
    end;
// DEB PT49
{$IFNDEF COMSX}
    if GetControltext('ANOABS') = 'X' then
    begin // Pas de rejet du fichier mais non traitement des lignes absences
      if not OkOk then
      begin
        OkOk := TRUE;
        AbsExclue := TRUE;    //PT57
        AbsNonTrait := TRUE;  //PT57
        Anomalie(Chaine, 'Absence non intégrée', 140)
      end;
    end;
{$ENDIF}
// FIN PT49
    if not OkOk then Erreur(chaine, '', 140);
    // FIN PT11-2
  end;
  ferme(Q);
// d PT33


  GestionIJSS := False;

  if ((T_MotifAbs <> nil) and (T_MotifAbs.GetValue('PMA_GESTIONIJSS') = 'X')) then // PT46
  begin
    GestionIJSS := true;
    CarenceIjss := T_MotifAbs.GetValue('PMA_CARENCEIJSS');
  end;
// f PT33

//GGS PT35  if not Ecrit then exit;
  if not Ecrit then
  begin
    if iduniq = '' then iduniq := AglGetGuid();
    Datedebutabs := UsDatetime(Tdatedeb);
    Datefinabs := UsDateTime(Tdatefin);
    if libanomalie then StatutEnreg := 'Anomalie';
    if liberreur then StatutEnreg := 'Erreur';
    NumOrdre := NumOrdre + 1;
// DEB PT46
    l1 := StrLen(PChar(Libelle));
    if l1 > 35 then libelle := Copy(Libelle, 1, 35);
//DEB PT51
    Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [CodeSalarie], true);
    if sal <> nil then begin
      nom := Sal.Getvalue('PSA_LIBELLE');
      prenom := Sal.Getvalue('PSA_PRENOM');
    end;
    // DEB PT53
    if Tob_SalATT <> nil then
    begin
      Tatt := Tob_SalATT.findfirst(['PIM_SALARIE'], [Codesalarie], true);
      if Tatt <> nil then begin
        nom := TATT.GetValue('PIM_LIBELLE');
        prenom := Tatt.GetValue('PIM_PRENOM');
      end;
    end;
   // FIN PT53
//FIN PT51
    st := 'insert into IMPORTPAIE(PIE_GUID,PIE_TYPEDONNEE,PIE_SALARIE,PIE_NUMORDRE,PIE_LIBELLE,PIE_PRENOM,PIE_NUMEROSS,' +
      'PIE_DATEENTREE,PIE_DATESORTIE,PIE_COMMENTAIRE,PIE_RUBOUMOTIF,PIE_LIBELLE2,PIE_BASEJOUR,PIE_TAUXHEURES,' +
      'PIE_COEFF,PIE_MONTANT,PIE_DATEDEBUT,PIE_DATEFIN,PIE_DEBUTDJ,PIE_FINDJ,PIE_STATUT) VALUES ("' +
      IDUNIQ + '","VAR","' + Codesalarie + '",' + IntToStr(NumOrdre) + ',"' + nom + '","' + prenom + '","","' + Dudate + '","' + audate + '","","' + MotifAbs + '","' + Libelle + '",' +
      VNbj + ',' + VNbh + ',"",0.00,"' + DateDebutAbs + '","' + DateFinAbs + '","' + Dj + '","' + Fj +
      '","' + StatutEnreg + '")';
// FIN PT46
    ExecuteSQL(St);
    Exit;
  end;
//Fin GGS PT35
  //  PT11-2 02/12/2002 PH si les absences sont déjà saisies et que le salarié est sorti alors OK
  if OkOk and SalSorti then
  begin { DEB PT45 }
    if Ecrit then { Pour indiquer la sortie du salarié }
    begin
      Anomalie(Chaine, 'Salarié sorti. Absence non intégrée', 46)
    end; {FIN PT45 }
    exit; // On ne traite pas la ligne du fichier d'import
  end;

  if AbsNonTrait then exit; // PT49   //PT57

  Sal := Tob_Sal.findfirst(['PSA_SALARIE'], [CodeSalarie], true);
  if Sal = nil then exit;
  // PT6 27/11/01 PH Si on ne gère pas les cp et que l'on a des lignes de cp,on perd les lignes
  if (Sal.getvalue('ETB_CONGESPAYES') <> 'X') and (Sal.getvalue('PSA_CONGESPAYES') <> 'X') and (MotifAbs = 'PRI') then { PT22 }
  begin
    Erreur(chaine, CodeSalarie, 27);
    exit;
  end;

  { Création du mouvement d'absence }
  T := Tob.create('ABSENCESALARIE', TOB_ABS, -1);
  if T = nil then exit;

  //  Ordre := IncrementeSeqNoOrdre('CPA',codeSalarie);
  //   PT4 26/11/01 PH Incrementation du numéro d'ordre des mvts de type CP
  if MotifAbs = 'PRI' then Ordre := RendNumOrdre('CPA', CodeSalarie)
  else Ordre := RendNumOrdre('ABS', CodeSalarie);
  InitialiseTobAbsenceSalarie(T); { PT29 }
  T.PutValue('PCN_SALARIE', Codesalarie);
  T.PutValue ('PCN_GUID', AglGetGUID ()); //PT55
  if MotifAbs = 'PRI' then
    T.PutValue('PCN_TYPEMVT', 'CPA')
  else T.PutValue('PCN_TYPEMVT', 'ABS');
  T.PutValue('PCN_ORDRE', Ordre);
  T.PutValue('PCN_CODERGRPT', Ordre);
  T.PutValue('PCN_DATEDEBUT', TDateDeb);
  if MotifAbs <> 'PRI' then T.PutValue('PCN_DATEFIN', TDateFin)
  else T.PutValue('PCN_DATEFIN', 2); // On force la date de fin pourf les CP à idate1900
  T.PutValue('PCN_DATEDEBUTABS', TDateDeb);
  T.PutValue('PCN_DATEFINABS', TDateFin);
  T.PutValue('PCN_TYPECONGE', MotifAbs);
  if MotifAbs = 'PRI' then
    T.PutValue('PCN_MVTPRIS', MotifAbs);
  T.PutValue('PCN_SENSABS', '-');
  // PT3 26/11/01 PH Récupération du libellé saisi de l'absence
  //  T.PutValue('PCN_LIBELLE'      ,'Mvt d''import');
  T.PutValue('PCN_LIBELLE', Copy(Libelle, 1, 35)); //PT10
  // PT1 14/06/01 PH Rajout récup libellé complémentaires et matin/AP
  T.PutValue('PCN_LIBCOMPL1', LibCompl1);
  T.PutValue('PCN_LIBCOMPL2', LibCompl2);
  T.PutValue('PCN_DEBUTDJ', Dj);
  T.PutValue('PCN_FINDJ', Fj);

  T.PutValue('PCN_DATEMODIF', Date);
  T.PutValue('PCN_DATEVALIDITE', TDateFin);
  Dateclot := Sal.getvalue('ETB_DATECLOTURECPN');
  T.PutValue('PCN_PERIODECP', CalculPeriode(Dateclot, TDatefin)); { PT29 }
  if Nbj = '' then NBj := '0';
  //   PT7 18/12/01 PH remplacement strtofloaat par valeur sinon erreur de conversion
  T.PutValue('PCN_JOURS', VALEUR(Nbj));
  if NbH = '' then NBH := '0';
  T.PutValue('PCN_HEURES', VALEUR(NbH));
  //   PT5 27/11/01 PH Initialisation de tous les champs de la table absencesalarie
  T.PutValue('PCN_SAISIEDEPORTEE', 'X');
  T.PutValue('PCN_VALIDRESP', 'VAL');
  T.PutValue('PCN_VALIDSALARIE', 'SAL'); //PT9
  T.PutValue('PCN_EXPORTOK', 'X');

  T.PutValue('PCN_ETABLISSEMENT', Sal.getvalue('PSA_ETABLISSEMENT'));
  T.Putvalue('PCN_TRAVAILN1', Sal.getvalue('PSA_TRAVAILN1'));
  T.Putvalue('PCN_TRAVAILN2', Sal.getvalue('PSA_TRAVAILN2'));
  T.Putvalue('PCN_TRAVAILN3', Sal.getvalue('PSA_TRAVAILN3'));
  T.Putvalue('PCN_TRAVAILN4', Sal.getvalue('PSA_TRAVAILN4'));
  T.Putvalue('PCN_CODESTAT', Sal.getvalue('PSA_CODESTAT'));
  T.Putvalue('PCN_CONFIDENTIEL', Sal.getvalue('PSA_CONFIDENTIEL'));

// d PT33
  if (GestionIJSS) then
  begin
    T.Putvalue('PCN_GESTIONIJSS', 'X');

    Tob_AbsIJSS := Tob.Create('Absences IJSS', nil, -1);
    st := 'SELECT PCN_SALARIE, PCN_DATEDEBUTABS, PCN_DATEFINABS, ' +
      'PCN_LIBELLE, PCN_NBJCARENCE,PCN_GESTIONIJSS ' +
      'FROM ABSENCESALARIE WHERE PCN_SALARIE = "' + Codesalarie +
      '" AND PCN_GESTIONIJSS="X" AND ' +
      'PCN_DATEDEBUTABS >="' + UsDateTime(PlusMois(TDateDeb, -12)) + '" AND ' +
      'PCN_DATEDEBUTABS <="' + UsDateTime(TDateDeb - 1) + '" ' +
      'ORDER BY PCN_SALARIE, PCN_DATEFINABS';
    Q := OpenSql(st, TRUE);

    if not (Q.eof) then
      Tob_AbsIJSS.LoadDetailDB('ABSENCESALARIE', '', '', Q, False);
    ferme(Q);

    // détermination de la carence cumulée depuis 12 mois
    WDateFin := IDate1900;
    WCarence := 0;

    AbsIJSS := Tob_AbsIJSS.FindFirst([''], [''], false);
    while (AbsIJSS <> nil) do
    begin
      if (AbsIJSS.GetValue('PCN_DATEDEBUTABS') <> WDateFin + 1) then
        WCarence := AbsIJSS.GetValue('PCN_NBJCARENCE')
      else
        WCarence := WCarence + AbsIJSS.GetValue('PCN_NBJCARENCE');

      if (WCarence > CarenceIJSS) then
        WCarence := CarenceIJSS;

      WDateFin := AbsIJSS.GetValue('PCN_DATEFINABS');
      AbsIJSS := Tob_AbsIJSS.FindNext([''], [''], false);
    end;

    // calcul du nombre de jours calendaires d'absence
    T.Putvalue('PCN_NBJCALEND', TDateFin - TDateDeb + 1);

    // calcul du nbre de jours de carence
    WDateDeb := T.Getvalue('PCN_DATEDEBUTABS');
    if (WDateDeb = WDateFin + 1) then
      T.Putvalue('PCN_NBJCARENCE', CarenceIJSS - WCarence)
    else
      T.Putvalue('PCN_NBJCARENCE', CarenceIJSS);

    if (T.Getvalue('PCN_NBJCALEND') < T.Getvalue('PCN_NBJCARENCE')) then
      T.Putvalue('PCN_NBJCARENCE', T.Getvalue('PCN_NBJCALEND'));

    // calcul du nombre de jours d'IJSS
    T.Putvalue('PCN_NBJIJSS',
      T.Getvalue('PCN_NBJCALEND') - T.Getvalue('PCN_NBJCARENCE'));

    FreeAndNil(Tob_AbsIJSS);

  end;
// f PT33
end;
// PT22   26/05/2004 PH V_50 Non traitement des absences portant sur des périodes closes mias ne provoquant pas de rejet du fichier
{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.Erreur(Chaine, S: string; NoErreur: integer; WW: string = '');
{$ELSE}

procedure Erreur(Chaine, S: string; NoErreur: integer; WW: string = '');
{$ENDIF}
var
  st: string;
begin
  if ((NoErreur = 48) or (NoErreur = 47)) and GetParamSocSecur('SO_IFDEFCEGID', FALSE) then
  begin
    if WW = 'NF' then
    begin
      Writeln(FR, '');
      Writeln(FR, 'l''enregistrement absence suivant sera exclus car ces dates sont hors période');
      Writeln(FR, Chaine);
      taille := taille + 3;
      AbsExclue := TRUE;
    end;
    exit;
  end;
  // FIN PT22
  ExisteErreur := true;
  if not liberreur then
    {       RapportErreur.lines.Add('Erreur sur l''enregistrement ||   '+ chaine);}
  begin
    Writeln(FR, '');
    Writeln(FR, 'Erreur sur l''enregistrement ||   ' + chaine);
    taille := taille + 2;
  end;
  Liberreur := true;
  case NoErreur of
    5: st := 'Marque de début de fichier non trouvée';
    7: st := 'Impossible de récupérer l''exercice social en cours - rejet du fichier';
    10: st := 'Des enregistrements ont été détectés après la marque de fin de fichier';
    15: st := 'Marque de fin de fichier non trouvée';
    20: st := 'Code Enregistrement inexistant';
    25: st := 'Code salarié inexistant dans ce dossier';
    26: st := 'Salarie en suspension de paye';
    27: st := 'L''établissement auquel appartient le salarié ne gère pas les CP alors que des CP pris sont trouvés';
    30: st := 'Date début de période non valide';
    35: st := 'Date fin de période non valide';
    40: st := 'la date fin de période est antérieure à la date début de période';
    44: st := 'Il n''y a pas d''exercice social actif correspondant à ces dates';
    45: st := 'Numéro d''ordre obligatoire';
    46: st := 'Attention, salarié sorti';
    47: st := 'Le mois de la date début est une période cloturée';
    48: st := 'Le mois de la date fin est une période cloturée';
    49: st := 'L''absence ne peut être à cheval sur plusieurs mois'; { PT41 }
    50: st := 'Rubrique inexistante';
    52: st := 'La rubrique d''une ligne de commentaire doit être précédée d''un enregistrement faisant référence à cette même rubrique';
    53: st := 'La rubrique d''une ligne analytique doit être précédée d''un enregistrement faisant référence à cette même rubrique';
    55: st := 'Type alimentation obligatoire';
    57: st := 'L''alimentation de la rubrique est différente de celle du plan de paie';
    60: st := 'La base doit être renseignée et être numérique';
    65: st := 'La base ne doit pas être renseignée';
    70: st := 'Le taux doit être renseigné et être numérique';
    75: st := 'Le taux ne doit pas être renseigné';
    80: st := 'Le coefficient doit être renseigné et être numérique';
    85: st := 'Le coefficient ne doit pas être renseigné';
    90: st := 'Le montant doit être renseigné et être numérique';
    95: st := 'Le montant ne doit pas être renseigné';
    100: st := 'Section analytique inconnue';
    101: st := 'Il y a des lignes analytiques alors que vous ne gerez pas l''analytique';
    102: st := 'Le numéro de ventilation analytique doit être renseigné et être numérique';
    110: st := 'Date début absence non valide';
    120: st := 'Date fin absence non valide';
    130: st := 'La date début doit être antérieure à la date de fin';
    140: st := 'il existe déjà une absence pour ce salarié à cette période';
    150: st := 'Nombre de jours non numérique';
    155: st := 'Le nombre de jours maximun octroyés est dépassé'; { PT39 }
    156: st := 'Le nombre d''heures maximun octroyés est dépassé'; { PT39 }
    160: st := 'Nombre d''heures non numérique';
    170: st := 'Code motif inexistant dans la table';
    220: st := 'Code enregistrement de la ligne Dossier-société erroné';
    230: st := 'Le numéro de dossier est erroné';
    240: st := 'Erreur de structure de l''enregistrement';
    250: st := 'Impossible de créer le salarié car le matricule existe déjà';
    251: st := 'Impossible de créer le salarié car il manque la date d''entrée';
    252: st := 'Impossible de créer le salarié car il manque son nom';
    253: st := 'Impossible de créer le salarié car il manque son prénom';
    255: st := 'Impossible de créer le salarié car plusieurs salariés en attente de création ont le même matricule';
    256: st := 'Création des salariés interrompue car il n''existe aucun salarié';
    257: st := 'Création des salariés interrompue car l''établissement indiqué n''existe pas';
    258: st := 'il est impossible de créer des salariés si la condification n''est pas numérique';
  else st := 'une erreur a été détectée dans votre fichier,veuillez le vérifier'
  end;
  Writeln(FR, st);
  taille := taille + 1;
  {   if RapportErreur <> nil then
        RapportErreur.lines.Add(St+' ||  '+S);}
end;
{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.Anomalie(Chaine, S: string; NoErreur: integer);
{$ELSE}

procedure Anomalie(Chaine, S: string; NoErreur: integer);
{$ENDIF}
var
  st: string;
begin
  if not libanomalie then
  begin
    if Noerreur <> 98 then
      Writeln(FR, 'Erreur non bloquante sur l''enregistrement ||   ' + chaine);
  end;
  Libanomalie := true;
  case NoErreur of
    26: st := 'Salarie en suspension de paye';
    46: st := S; { PT45  Cas d'un salarié sorti }
    98: st := S;
  else st := 'une erreur a été détectée dans votre fichier,veuillez le vérifier'
  end;
  Writeln(FR, st);
  taille := taille + 1;
end;

{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.ControleSalarie(TypeEnr, chaine, S: string; var CodeSalarie: string; Ecrit: Boolean);
{$ELSE}

procedure ControleSalarie(TypeEnr, chaine, S: string; var CodeSalarie: string; Ecrit: Boolean);
{$ENDIF}
var
  T, TAtt: tob;
begin
  DateEntree := 0;
  DateSortie := 0;
// DEB PT34
  if (tob_sal = nil) and (POS('S1', CodeSalarie) < 0) then // PT47
    Erreur(chaine, Codesalarie, 25) else
  begin
    if Tob_Sal <> nil then T := Tob_Sal.findfirst(['PSA_SALARIE'], [Codesalarie], true)
    else T := nil;
// FIN PT34
    // DEB PT30
    if T = nil then
    begin
      if Tob_SalATT <> nil then TAtt := Tob_SalATT.findfirst(['PIM_SALARIE'], [Codesalarie], true) // PT27
      else TAtt := nil;
    end
    else TAtt := nil;
    // FIN PT30
    if (T = nil) and (TAtt = nil) then Erreur(chaine, Codesalarie, 25) // PT27
    else
      if (T <> nil) and (T.GetValue('PSA_SUSPENSIONPAIE') = 'X') then Anomalie(chaine, Codesalarie, 26);
    // DEB PT27
    if (Ecrit) and (TAtt <> nil) then
    begin
      if (TAtt.GetValue('TRAITSAL') = 'C') then CodeSalarie := TAtt.GetValue('CodeS5');
    end;
    // FIN PT27
  end;
end;
// La periode contenent ces dates est elle cloturée ?
// result = 0 : non
//          1 : pas d'exercice social trouvé
//          2 : Mois date début cloturé
//          3 : Mois date de fin cloturé
//
//  PT8 15/02/02 PH test pour les mvts de cp sur mois clos N et N+1
{$IFNDEF COMSX}

function TOF_PGIMPORTFIC.IsPeriodecloture(Dated, DateF: TDateTime; CasCp: string = '-'): integer;
{$ELSE}

function IsPeriodecloture(Dated, DateF: TDateTime; CasCp: string = '-'): integer;
{$ENDIF}
var
  Q: TQuery;
  aa, aad, aaf, mmd, mmf, jj: WORD;
  trouve: boolean;
  d, f: integer;
begin
    // PT19
  if (Cloture = '') or (Dated < DebExer) or (DateF > FinExer) then
  begin
    trouve := false;
    Q := OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER', TRUE);
    while not Q.EOF do
    begin
      DebExer := Q.FindField('PEX_DATEDEBUT').AsDateTime;
      FinExer := Q.FindField('PEX_DATEFIN').AsDateTime;
      if ((CasCp = '-') and (debExer <= DateD) and (FinExer >= Datef))
        or ((CasCp = 'X') and (FinExer >= Datef)) then { PT41 }
      begin
        trouve := true;
        break;
      end;
      Q.Next;
    end;
    Cloture := Q.FindField('PEX_CLOTURE').asstring;
    Ferme(Q);
    if not trouve then
    begin
      result := 1; // pas d'exercice social actif correspondant à ces dates
      exit;
    end;
  end;
    // FIN PT19

  //  PT8 15/02/02 PH test pour les mvts de cp sur mois clos N et N+1
  if CasCp = 'X' then
  begin
    decodedate(Dated, aa, mmd, jj);
    decodedate(dateF, aa, mmf, jj);
    if GetParamSoc('SO_PGDECALAGE') = TRUE then // PT24
    begin
      if mmd = 12 then d := 1 else d := mmd + 1;
      if mmf = 12 then f := 1 else f := mmf + 1;
    end
    else
    begin
      d := mmd;
      f := mmf;
    end;
    result := 0;
    if (Cloture[d] = 'X') then
    begin
      if d < 12 then
      begin
        if Cloture[d + 1] = 'X' then result := 2;
      end else result := 2;
    end;
    if (result = 0) and (Cloture[f] = 'X') then
    begin
      if f < 12 then
      begin
        if Cloture[f + 1] = 'X' then result := 3;
      end else result := 3;
    end;
  end
  else
  begin
    if Cloture[d] = 'X' then result := 2
    else
      if Cloture[f] = 'X' then result := 3
      else
        result := 0;
  end;
  //  FIN PT8
// PT19
  if (result = 2) or (result = 3) then
  begin
    // Rajouter le test d'un paramsoc pour outrepasser les controles SPECIFCEGID
  end;
  // FIN PT19
{ DEB PT41 }
  if (CasCp = 'X') and (Result = 0) then
  begin
    if not GetParamSocSecur('SO_PGABSENCECHEVAL', False) then
    begin
      decodedate(Dated, aad, mmd, jj);
      decodedate(dateF, aaf, mmf, jj);
      if (MMD <> MMF) or (AAD <> AAF) then
      begin
        result := 4;
        exit;
      end;
    end;
  end;
{ FIN PT41 }
end;
{$IFNDEF COMSX}

function TOF_PGIMPORTFIC.RendNumOrdre(TypeMvt, CodeSalarie: string): integer;
{$ELSE}

function RendNumOrdre(TypeMvt, CodeSalarie: string): integer;
{$ENDIF}
var
  Ordre: Integer;
  T1: TOB;
begin
  T1 := Tob_OrdreAbs.findFirst(['PCN_SALARIE', 'PCN_TYPEMVT'], [CodeSalarie, TypeMvt], FALSE);
  if T1 <> nil then
  begin
    Ordre := T1.GetValue('NUMORDRE') + 1;
    T1.PutValue('NUMORDRE', Ordre);
  end
  else
  begin
    Ordre := 1;
    T1 := TOB.Create('LES NUMORDRES ABSENCESALARIE', Tob_OrdreAbs, -1);
    T1.AddChampSup('PCN_SALARIE', FALSE);
    T1.AddChampSup('PCN_TYPEMVT', FALSE);
    T1.AddChampSup('NUMORDRE', FALSE);
    T1.PutValue('PCN_SALARIE', CodeSalarie);
    T1.PutValue('PCN_TYPEMVT', TypeMvt);
    T1.PutValue('NUMORDRE', Ordre);
  end;
  result := Ordre;
end;
{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.ControleEnregVSA(TypeEnr, chaine, S: string; Ecrit: boolean);
{$ELSE}

procedure ControleEnregVSA(TypeEnr, chaine, S: string; Ecrit: boolean);
{$ENDIF}
var
  Codesalarie, DateDebutperiode, DateFinperiode, SNoOrdre: string;
  Axe, Section, CodeAxe, sMontant, LibSect: string;
  TDateDeb, TDateFin: tdatetime;
  T: tob;
  TopCloture: integer;
begin
  libErreur := false;
  libanomalie := false;
  if not GetParamSoc('SO_PGANALYTIQUE') then Erreur(chaine, '', 101); // PT24
  // contrôle sur le salarié
  Codesalarie := readtokenPipe(S, Separateur);
  if (CodeSalarie <> Salarieprec) or (POS('S1', CodeSalarie) > 0) then //PT32 //PT47
  begin
    ControleSalarie(TypeEnr, chaine, S, CodeSalarie, Ecrit);
  end;
  Salarieprec := Codesalarie;
  DateDebutperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DatedebutPeriode) then
    Erreur(chaine, DatedebutPeriode, 30);
  DateFinperiode := readtokenPipe(S, Separateur);
  if not Isvaliddate(DateFinPeriode) then
    Erreur(chaine, DateFinPeriode, 35);
  TDateDeb := 0;
  TDateFin := 0;
  if (Isvaliddate(DatedebutPeriode)) and (Isvaliddate(DatefinPeriode)) then
  begin
    TDateDeb := strtodate(DatedebutPeriode);
    TDateFin := strtodate(DatefinPeriode);
  end;
  if TDateFin < TDateDeb then
    Erreur(chaine, DateDebutPeriode + '//' + DateFinPeriode, 40);
  Topcloture := IsPeriodecloture(TDatedeb, TDateFin);
  if TopCloture = 1 then
    Erreur(Chaine, '', 44);
  if TopCloture = 2 then
    Erreur(Chaine, '', 47);
  if TopCloture = 3 then
    Erreur(Chaine, '', 48);
  if (DateSortie > 10) and (DateSortie < TDateDeb) then
    Erreur(Chaine, '', 46);
  Axe := readtokenPipe(S, Separateur);
  if (Axe < 'A1') or (Axe > 'A5') then Erreur(chaine, Axe, 103);
  Section := readtokenPipe(S, Separateur);
  Smontant := readtokenPipe(S, Separateur);
  if sMontant = '' then
    Erreur(chaine, sMontant, 90)
  else
    if not Isnumeric(sMontant) then Erreur(chaine, sMontant, 90);
  SNoOrdre := readtokenPipe(S, Separateur);
  if SNoOrdre = '' then
    Erreur(chaine, SNoOrdre, 102)
  else
    if not Isnumeric(SNoOrdre) then Erreur(chaine, SNoOrdre, 102);

  // PT16   28/08/2003 PH Création des sections en récupérant le libellé de la section
  if S <> '' then LibSect := readtokenPipe(S, Separateur)
  else LibSect := '';
  if not ExisteSection(Axe, Section, LibSect) then Erreur(Chaine, Section, 100);
  // Fin PT16

  if not Ecrit then exit;
  T := Tob.create('PAIEVENTIL', TOB_Paieventil, -1);
  if T = nil then exit;
  CodeAxe := Copy(Axe, 2, 1);
  T.PutValue('PAV_NATURE', 'VS' + CodeAxe);
  T.PutValue('PAV_COMPTE', CodeSalarie);
  T.PutValue('PAV_DATEDEBUT', strtodate(DateDebutperiode));
  T.PutValue('PAV_DATEFIN', strtodate(DateFinperiode));
  T.PutValue('PAV_SECTION', Section);
  T.PutValue('PAV_TAUXMONTANT', Valeur(sMontant));
  T.PutValue('PAV_TAUXQTE1', 0);
  T.PutValue('PAV_TAUXQTE2', 0);
  T.PutValue('PAV_NUMEROVENTIL', StrToInt(SNoOrdre));
  T.PutValue('PAV_SOCIETE', GetParamSoc('SO_SOCIETE'));
  T.PutValue('PAV_MONTANT', 0);
  T.PutValue('PAV_SOUSPLAN1', '');
  T.PutValue('PAV_SOUSPLAN2', '');
  T.PutValue('PAV_SOUSPLAN3', '');
  T.PutValue('PAV_SOUSPLAN4', '');
  T.PutValue('PAV_SOUSPLAN5', '');
  T.PutValue('PAV_SOUSPLAN6', '');
end;
{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 10/12/2004
Modifié le ... :   /  /
Description .. : Fonction de création automatique de salarié
Suite ........ : 10/12/2004 Ajout de fonctionnalité : prise en compte de
Suite ........ : la modification des infos salariés
Mots clefs ... : PAIE;
*****************************************************************}
// DEB PT27
{$IFNDEF COMSX}

function TOF_PGIMPORTFIC.ControleEnregSAL(TypeEnr, chaine, S: string; Ecrit: boolean): Boolean;
{$ELSE}

function ControleEnregSAL(TypeEnr, chaine, S: string; Ecrit: boolean): Boolean;
{$ENDIF}
var
  T, T1, T2, TSalACreer, TEtab: TOB;
  CodeSalarie, DatEnt, SNom, SPrenom, AD1, AD2, AD3, CP, Ville, Tel, DNai, Nat, Sexe, NOSS: string;
  Etab, DSortie, Civilite, Portable, st, Etablis: string;
  Anomalie: Boolean;
  i, LongLib, LongPre: integer;
  Q: TQuery;
  InfTiers: Info_Tiers;
  CodeAuxi, Libell, Prenom, LeRapport, ConvColl, TraitSal, Trait: string;
  QQuery: TQuery;
  IMax, MaxInterim: integer;
  // DEb PT35
  commentaire: string;
  CommNai, DepNai, PayNai, NumCarteSej, Delivpar, SitFam, MotifEntree, MotifSort, CatDucs, RegimeSS, StatutCat, StatutProf, CondEmploi, UniteTrav, SortieDef: string; //PT37
  DateExpir, HeureEmb: TDateTime; //PT37
  TauxTpsPartiel: Double; //PT37
  // FIN PT35
  function TraiteErreur(Chaine, CodeSal: string; Noerreur: Integer): Boolean;
  begin
    Erreur(Chaine, CodeSal, NoErreur);
    Result := TRUE;
  end;

begin
  if not Ecrit then
  begin
    libErreur := false;
    libanomalie := false;
    if Tob_SalATT = nil then Tob_SalATT := Tob.Create('Mes_Salaries_A_CREER', nil, -1);
    Codesalarie := readtokenPipe(S, Separateur);
//GGR PT35
    DatEnt := readtokenPipe(S, Separateur);
    SNom := readtokenPipe(S, Separateur);
    SPrenom := readtokenPipe(S, Separateur);
    AD1 := readtokenPipe(S, Separateur);
    AD2 := readtokenPipe(S, Separateur);
    AD3 := readtokenPipe(S, Separateur);
    CP := readtokenPipe(S, Separateur);
    Ville := readtokenPipe(S, Separateur);
    Tel := readtokenPipe(S, Separateur);
    DNai := readtokenPipe(S, Separateur);
    Nat := readtokenPipe(S, Separateur);
    Sexe := readtokenPipe(S, Separateur);
    NOSS := readtokenPipe(S, Separateur);
    Civilite := readtokenPipe(S, Separateur);
    Portable := readtokenPipe(S, Separateur);
    DSortie := readtokenPipe(S, Separateur);
    Etablis := readtokenPipe(S, Separateur);
    // DEB PT37
    if VersionFic >= '0700' then // PT47 Uniformisation versionfic
    begin
      Commentaire := readtokenPipe(S, Separateur);
      DepNai := readtokenPipe(S, Separateur);
      CommNai := readtokenPipe(S, Separateur);
      PayNai := readtokenPipe(S, Separateur);
      NumCarteSej := readtokenPipe(S, Separateur);
      Delivpar := readtokenPipe(S, Separateur);
      St := readtokenPipe(S, Separateur);
{      if (st <> '') and (isDateTimeText(St)) then DateExpir := StrTodate(st) PT47
      else DateExpir := IDate1900;}
      if (st <> '') then DateExpir := StrToDate(St)
      else DateExpir := IDate1900;
      SitFam := readtokenPipe(S, Separateur);
      MotifEntree := readtokenPipe(S, Separateur);
      st := readtokenPipe(S, Separateur);
      if (st <> '') then HeureEmb := StrToTime(st)
      else HeureEmb := IDate1900;
      MotifSort := readtokenPipe(S, Separateur);
      CatDucs := readtokenPipe(S, Separateur);
      RegimeSS := readtokenPipe(S, Separateur);
      StatutCat := readtokenPipe(S, Separateur);
      StatutProf := readtokenPipe(S, Separateur);
      UniteTrav := readtokenPipe(S, Separateur);
      CondEmploi := readtokenPipe(S, Separateur);
      st := readtokenPipe(S, Separateur);
      if (st <> '') and (IsNumeric(St)) then TauxTpsPartiel := StrToFloat(st) // PT47
      else TauxTpsPartiel := 00;
      SortieDef := readtokenPipe(S, Separateur);
    end;
    // FIN PT37
    if iduniq = '' then iduniq := AglGetGuid();
    if libanomalie then StatutEnreg := 'Anomalie';
    if liberreur then StatutEnreg := 'Erreur';
    NumOrdre := NumOrdre + 1;
    // PT44 Traitement des dates dans la requête
// DEB PT46
    st := 'insert into IMPORTPAIE(PIE_GUID,PIE_TYPEDONNEE,PIE_SALARIE,PIE_NUMORDRE,PIE_LIBELLE,PIE_PRENOM,PIE_NUMEROSS,' +
      'PIE_DATEENTREE,PIE_DATESORTIE,PIE_COMMENTAIRE,PIE_RUBOUMOTIF,PIE_LIBELLE2,PIE_BASEJOUR,PIE_TAUXHEURES,' +
      'PIE_COEFF,PIE_MONTANT,PIE_DATEDEBUT,PIE_DATEFIN,PIE_DEBUTDJ,PIE_FINDJ,PIE_STATUT) VALUES ("' +
      IDUNIQ + '","SAL","' + Codesalarie + '",' + IntToStr(NumOrdre) + ',"' + Snom + '","' + Sprenom + '","' + noss +
      '","' + UsDateTime(StrToDate(Datent)) + '","' + UsDateTime(StrToDate(DSORTIE)) + '","' + commentaire + '","","",0,0,"",0,0,0,"","","' + StatutEnreg + '")';
// FIN PT46
    ExecuteSQL(St);
//Fin GGS PT35
    if (Copy(Codesalarie, 1, 2) = 'S1') then TraitSal := 'C' else // PT47
      traitSal := 'M';
    if (TraitSal = 'C') then
    begin
      if (not GetParamSoc('SO_PGINCSALARIE')) or (GetParamSoc('SO_PGTYPENUMSAL') <> 'NUM') then
      begin
        Erreur(Chaine, Codesalarie, 258);
        result := FALSE;
        exit;
      end;
    end;
    // ATTENTION, si modification alors reporter les modifs dans UTOMSalaries
    if MatSouche = -999 then
    begin
      MaxInterim := 0;
      if (GetParamSoc('SO_PGCODEINTERIM') = True) and (GetParamSoc('SO_PGINTERIMAIRES') = True) then
      begin
        QQuery := OpenSQL('SELECT MAX(PSI_INTERIMAIRE) FROM INTERIMAIRES', True);
        if (not QQuery.EOF) and (QQuery.Fields[0].AsString <> '') then MaxInterim := StrToInt(QQuery.Fields[0].AsString)
        else MaxInterim := 0;
        Ferme(QQuery);
      end;

      QQuery := OpenSQL('SELECT MAX(PSA_SALARIE) FROM SALARIES', TRUE);
      if not QQuery.Eof then
      begin // Attention, respecter le type du numéro salarié soit alpha soit numérique avec increment auto
        if QQuery.Fields[0].AsString <> '' then Imax := StrToInt(QQuery.Fields[0].AsString) // On récupère le max du numéro de salarié
        else Imax := 0;
        if MaxInterim > IMax then IMax := MaxInterim;
        if IMax < 2147483647 then
          MatSouche := Imax;
      end
      else MatSouche := 0;
      Ferme(QQuery);
    end;
    // contrôle existance du salarié
// DEB PT34
    if Tob_sal <> nil then
    begin
      T := Tob_sal.FindFirst(['PSA_SALARIE'], [CodeSalarie], true);
      if (T <> nil) and (TraitSal = 'C') then // Mode creation salarie et déjà existant
      begin
        Erreur(Chaine, Codesalarie, 250);
        result := FALSE;
        exit;
      end;
    end;
// FIN PT34
    T := Tob_SalATT.FindFirst(['PIM_SALARIE'], [CodeSalarie], true); { PT45-2 PIM au lieu de PMI }
    if (T <> nil) and (TraitSal = 'C') then // Salarié déjà présent ==> doublon de code salarié à créer
    begin
      Erreur(Chaine, Codesalarie, 255);
      result := FALSE;
      exit;
    end;
    Anomalie := FALSE;
//    DatEnt := readtokenPipe(S, Separateur);
    if (DatEnt = '') or (not IsValidDate(DatEnt)) then Anomalie := TraiteErreur(Chaine, Codesalarie, 251);
//    SNom := readtokenPipe(S, Separateur);
    if Snom = '' then Anomalie := TraiteErreur(Chaine, Codesalarie, 252);
//    SPrenom := readtokenPipe(S, Separateur);
    if SPrenom = '' then Anomalie := TraiteErreur(Chaine, Codesalarie, 253);
    if Anomalie then
    begin
      result := FALSE;
      exit;
    end;
{    AD1 := readtokenPipe(S, Separateur);
    AD2 := readtokenPipe(S, Separateur);
    AD3 := readtokenPipe(S, Separateur);
    CP := readtokenPipe(S, Separateur);
    Ville := readtokenPipe(S, Separateur);
    Tel := readtokenPipe(S, Separateur);
    DNai := readtokenPipe(S, Separateur);
    Nat := readtokenPipe(S, Separateur);
    Sexe := readtokenPipe(S, Separateur);
    NOSS := readtokenPipe(S, Separateur);
    Civilite := readtokenPipe(S, Separateur);
    Portable := readtokenPipe(S, Separateur);
    DSortie := readtokenPipe(S, Separateur);
    Etablis := readtokenPipe(S, Separateur);}
    if Etablis <> '' then
    begin
      TEtab := Tob_Etab.FindFirst(['ETB_ETABLISSEMENT'], [Etablis], TRUE);
      if TEtab = nil then
      begin
        Erreur(Chaine, Etablis, 257);
//      Anomalie := TRUE;
        result := FALSE;
        exit;
      end;
    end;
{    Commentaire := readtokenPipe(S, Separateur);
    if iduniq = '' then iduniq := AglGetGuid();
    if libanomalie then StatutEnreg := 'Anomalie';
    if liberreur then StatutEnreg := 'Erreur';
    st:= 'insert into IMPORTPAIE(PIE_GUID,PIE_TYPEDONNEE,PIE_SALARIE,PIE_LIBELLE,PIE_PRENOM,PIE_NUMEROSS,' +
         'PIE_DATEENTREE,PIE_DATESORTIE,PIE_COMMENTAIRE,PIE_RUBOUMOTIF,PIE_LIBELLE2,PIE_BASEJOUR,PIE_TAUXHEURES,'+
         'PIE_COEFF,PIE_MONTANT,PIE_DATEDEBUT,PIE_DATEFIN,PIE_DEBUTDJ,PIE_FINDJ,PIE_STATUT) VALUES ("' +
         IDUNIQ + '","SAL","' + Codesalarie + '","' + Snom +'","' + Sprenom + '","'+ noss +
         '","' + Datent + '","'+ DSORTIE + '","' + commentaire + '","","",0,0,"",0,0,0,"","","'+ StatutEnreg +'")';
    ExecuteSQL(St);
    //Fin GGS PT35}
    T1 := TOB.Create('SALATTENTE', TOB_SalATT, -1);
    T1.PutValue('PIM_SALARIE', CodeSalarie);
    T1.PutValue('PIM_DATEENTREE', StrToDate(DatEnt));
    T1.PutValue('PIM_LIBELLE', Snom);
    T1.PutValue('PIM_PRENOM', SPrenom);
    T1.PutValue('PIM_ADRESSE1', AD1);
    T1.PutValue('PIM_ADRESSE2', AD2);
    T1.PutValue('PIM_ADRESSE3', AD3);
    T1.PutValue('PIM_CODEPOSTAL', CP);
    T1.PutValue('PIM_VILLE', Ville);
    T1.PutValue('PIM_TELEPHONE', Tel);
    if IsValidDate(DNai) then T1.PutValue('PIM_DATENAISSANCE', StrToDate(DNai))
    else T1.PutValue('PIM_DATENAISSANCE', IDate1900);
    T1.PutValue('PIM_NATIONALITE', Nat);
    T1.PutValue('PIM_SEXE', Sexe);
    T1.PutValue('PIM_NUMEROSS', NOSS);
    T1.AddChampSupValeur('PORTABLE', Portable, False);
    T1.AddChampSupValeur('CIVILITE', Civilite, False);
    if IsValidDate(DSortie) then T1.AddChampSupValeur('DATESORTIE', StrToDate(DSortie), FALSE)
    else T1.AddChampSupValeur('DATESORTIE', IDate1900, FALSE);
    if Etablis <> '' then T1.AddChampSupValeur('ETABLISSEMENT', Etablis, False)
    else T1.AddChampSupValeur('ETABLISSEMENT', '', False);
    T1.AddChampSupValeur('TRAITSAL', TraitSal, False); // Stockage du type de traitement à faire pour le salarié
    T1.AddChampSupValeur('MATRICULEBL', CodeSalarie, False); // Stockage du matricule origine BL

    // DEB PT37
    if VersionFic >= '0700' then // PT47 Uniformisation versionfic
    begin
      T1.AddChampSupValeur('DEPNAI', DepNai, False); // Département naissance
      T1.AddChampSupValeur('COMMNAI', CommNai, False); // Commune naissance
      T1.AddChampSupValeur('PAYNAI', PayNai, False); // Pays Naissance
      T1.AddChampSupValeur('NUMCARTESEJ', NumCarteSej, False); // Numéro carte de séjour
      T1.AddChampSupValeur('DELIVPAR', Delivpar, False); // Délivrée par
      T1.AddChampSupValeur('SITFAM', SitFam, False); //  Situation de fammille
      T1.AddChampSupValeur('MOTIFENTREE', MotifEntree, False); // Motif entrée
      T1.AddChampSupValeur('MOTIFSORT', MotifSort, False); // Motif sortie
      T1.AddChampSupValeur('CATDUCS', CatDucs, False); // Catégorie DUCS
      T1.AddChampSupValeur('REGIMESS', RegimeSS, False); // Régime SS
      T1.AddChampSupValeur('STATUTCAT', StatutCat, False); // Statut c&tégoriel
      T1.AddChampSupValeur('STATUFPROF', StatutProf, False); // Statut professionnel
      T1.AddChampSupValeur('CONDEMPLOI', CondEmploi, False); // Condition d'emploi
      T1.AddChampSupValeur('UNITETRAV', UniteTrav, False); // Unité de travail
      T1.AddChampSupValeur('SORTIEDEF', SortieDef, False); // Sortie définitive
      T1.AddChampSupValeur('DATEEXPIR', DateExpir, False); // Date expiration carte de séjour
      T1.AddChampSupValeur('HEUREMB', HeureEmb, False); // Heure d'embauche
      T1.AddChampSupValeur('TAUXTPSPARTIEL', TauxTpsPartiel, False); // Taux temps partiel
    end;
    // FIN PT37
    if TraitSal = 'C' then // Phase creation alors ajout 1 ligne dans la table de correspondance
    begin
      MatSouche := MatSouche + 1;
      T1.AddChampSupValeur('CodeS5', ColleZeroDevant(MatSouche, 10), False);
    end
    else
      T1.AddChampSupValeur('CodeS5', CodeSalarie, False);
  end // Fin du cas controle des enregistrements salaries
  else
  begin
    if (Tob_SalATT <> nil) and (not TopCreat) then
    begin
      TopCreat := TRUE;
      TSalACreer := TOB.Create('les_salaries_acreer', nil, -1);
      Etab := GetParamSoc('SO_ETABLISDEFAUT');
      if Etab <> '' then TEtab := Tob_Etab.FindFirst(['ETB_ETABLISSEMENT'], [Etab], TRUE)
      else TEtab := Tob_Etab.FindFirst([''], [''], TRUE);
      if TEtab = nil then
      begin
        Result := FALSE;
        Erreur(Chaine, Codesalarie, 256);
        exit;
      end;
      Etab := TEtab.GetValue('ETB_ETABLISSEMENT');
      ConvColl := TEtab.GetValue('ETB_CONVENTION');

      for i := 0 to Tob_SalATT.detail.count - 1 do
      begin
        T1 := Tob_SalATT.detail[i];
        Trait := T1.GetValue('TRAITSAL');
        if Trait = 'C' then
        begin
          T2 := TOB.Create('SALARIES', TSalACreer, -1);
          T2.PutValue('PSA_SALARIE', T1.GetValue('CODES5'));
          Etablis := T1.GetValue('ETABLISSEMENT');
          T2.PutValue('PSA_ETABLISSEMENT', Etab);
          T2.PutValue('PSA_CONVENTION', Convcoll);
          if Etablis <> '' then
          begin
            if Etablis <> Etab then
            begin
              TEtab := Tob_Etab.FindFirst(['ETB_ETABLISSEMENT'], [Etablis], TRUE);
              if TEtab <> nil then
              begin
                T2.PutValue('PSA_ETABLISSEMENT', Etablis);
                T2.PutValue('PSA_CONVENTION', TEtab.GetValue('ETB_CONVENTION'));
              end;
            end;
          end;
          T2.PutValue('PSA_DATEENTREE', T1.GetValue('PIM_DATEENTREE'));
          T2.PutValue('PSA_DATEANCIENNETE', T2.GetValue('PSA_DATEENTREE'));
          T2.PutValue('PSA_ANCIENPOSTE', T2.GetValue('PSA_DATEENTREE'));
          T2.PutValue('PSA_DATENAISSANCE', Idate1900);
          T2.PutValue('PSA_DATESORTIE', Idate1900);
          T2.PutValue('PSA_DATEENTREEPREC', Idate1900);
          T2.PutValue('PSA_DATESORTIEPREC', Idate1900);
          T2.PutValue('PSA_DATELIBRE1', Idate1900);
          T2.PutValue('PSA_DATELIBRE2', Idate1900);
          T2.PutValue('PSA_DATELIBRE3', Idate1900);
          T2.PutValue('PSA_DATELIBRE4', Idate1900);
          T2.PutValue('PSA_ORDREAT', '1');
          T2.PutValue('PSA_DADSFRACTION', '1');
          T2.PutValue('PSA_REGIMESS', '200');
          T2.PutValue('PSA_TAUXPARTSS', 0);
          T2.PutValue('PSA_TAUXPARTIEL', 0);
          T2.PutValue('PSA_DADSDATE', IDate1900);
          T2.PutValue('PSA_NATIONALITE', 'FRA');
          T2.PutValue('PSA_PAYSNAISSANCE', 'FRA');
          T2.PutValue('PSA_PRISEFFECTIF', 'X');
          T2.PutValue('PSA_UNITEPRISEFF', 1);
          T2.PutValue('PSA_LIBELLE', T1.GetValue('PIM_LIBELLE'));
          T2.PutValue('PSA_PRENOM', T1.GetValue('PIM_PRENOM'));
          T2.PutValue('PSA_ADRESSE1', T1.GetValue('PIM_ADRESSE1'));
          T2.PutValue('PSA_ADRESSE2', T1.GetValue('PIM_ADRESSE2'));
          T2.PutValue('PSA_ADRESSE2', T1.GetValue('PIM_ADRESSE3'));
          T2.PutValue('PSA_CODEPOSTAL', T1.GetValue('PIM_CODEPOSTAL'));
          T2.PutValue('PSA_VILLE', T1.GetValue('PIM_VILLE'));
          T2.PutValue('PSA_TELEPHONE', T1.GetValue('PIM_TELEPHONE'));
          T2.PutValue('PSA_DATENAISSANCE', T1.GetValue('PIM_DATENAISSANCE'));
          T2.PutValue('PSA_NATIONALITE', T1.GetValue('PIM_NATIONALITE'));
          T2.PutValue('PSA_SEXE', T1.GetValue('PIM_SEXE'));
          T2.PutValue('PSA_NUMEROBL', T1.GetValue('MATRICULEBL')); // Matricule Origine S1
          T2.PutValue('PSA_NUMEROSS', T1.GetValue('PIM_NUMEROSS'));
          T2.PutValue('PSA_CIVILITE', T1.GetValue('CIVILITE'));
          T2.PutValue('PSA_PORTABLE', T1.GetValue('PORTABLE'));
          T2.PutValue('PSA_DATESORTIE', T1.GetValue('DATESORTIE'));
          // DEB PT37
          if VersionFic >= '0700' then // PT47 Uniformisation versionfic
          begin
            T2.PutValue('PSA_DEPTNAISSANCE', T1.GetValue('DEPNAI')); // Département naissance
            T2.PutValue('PSA_COMMUNENAISS', T1.GetValue('COMMNAI')); // Département naissance
            T2.PutValue('PSA_PAYSNAISSANCE', T1.GetValue('PAYNAI')); // Pays Naissance
            T2.PutValue('PSA_CARTESEJOUR', T1.GetValue('NUMCARTESEJ')); // Numéro carte de séjour
            T2.PutValue('PSA_DELIVPAR', T1.GetValue('DELIVPAR')); // Délivrée par
            T2.PutValue('PSA_SITUATIONFAMIL', T1.GetValue('SITFAM')); //  Situation de fammille
            T2.PutValue('PSA_MOTIFENTREE', T1.GetValue('MOTIFENTREE')); // Motif entrée
            T2.PutValue('PSA_MOTIFSORTIE', T1.GetValue('MOTIFSORT')); // Motif sortie
            T2.PutValue('PSA_CATDADS', T1.GetValue('CATDUCS')); // Catégorie DUCS
            T2.PutValue('PSA_REGIMESS', T1.GetValue('REGIMESS')); // Régime SS
            T2.PutValue('PSA_DADSCAT', T1.GetValue('STATUTCAT')); // Statut c&tégoriel
            T2.PutValue('PSA_DADSPROF', T1.GetValue('STATUFPROF')); // Statut professionnel
            T2.PutValue('PSA_CONDEMPLOI', T1.GetValue('CONDEMPLOI')); // Condition d'emploi PT50
            T2.PutValue('PSA_UNITETRAVAIL', T1.GetValue('UNITETRAV')); // Unité de travail
            T2.PutValue('PSA_SORTIEDEFINIT', T1.GetValue('SORTIEDEF')); // Sortie définitive
            T2.PutValue('PSA_DATEXPIRSEJOUR', T1.GetValue('DATEEXPIR')); // Date expiration carte de séjour
            T2.PutValue('PSA_HEUREMBAUCHE', T1.GetValue('HEUREMB')); // Heure d'embauche
            T2.PutValue('PSA_TAUXPARTIEL', T1.GetValue('TAUXTPSPARTIEL')); // Taux temps partiel
          end;
          // FIN PT37
          T2.PutValue('PSA_TYPPROFIL', 'ETB');
          T2.PutValue('PSA_TYPPROFILREM', 'ETB');
          T2.PutValue('PSA_TYPPERIODEBUL', 'ETB');
          T2.PutValue('PSA_TYPPROFILRBS', 'ETB');
          T2.PutValue('PSA_TYPPROFILAFP', 'ETB');
          T2.PutValue('PSA_TYPPROFILAPP', 'ETB');
          T2.PutValue('PSA_TYPPROFILRET', 'ETB');
          T2.PutValue('PSA_TYPPROFILMUT', 'ETB');
          T2.PutValue('PSA_TYPPROFILPRE', 'ETB');
          T2.PutValue('PSA_TYPPROFILTSS', 'ETB');
          T2.PutValue('PSA_TYPPROFILCGE', 'ETB');
          T2.PutValue('PSA_TYPPROFILANC', 'ETB');
          T2.PutValue('PSA_TYPEDITBULCP', 'ETB');
          T2.PutValue('PSA_CPACQUISMOIS', 'ETB');
          T2.PutValue('PSA_CPACQUISSUPP', 'ETB');
          T2.PutValue('PSA_TYPNBACQUISCP', 'ETB');
          T2.PutValue('PSA_CPTYPEMETHOD', 'ETB');
          T2.PutValue('PSA_CPTYPERELIQ', 'ETB');
          T2.PutValue('PSA_CPACQUISANC', 'ETB');
          T2.PutValue('PSA_DATANC', 'ETB');
          T2.PutValue('PSA_TYPPROFILTRANS', 'ETB');
          T2.PutValue('PSA_TYPPROFILFNAL', 'DOS');
          T2.PutValue('PSA_STANDCALEND', 'ETB');
          T2.PutValue('PSA_TYPREDREPAS', 'ETB');
          T2.PutValue('PSA_TYPREDRTT1', 'ETB');
          T2.PutValue('PSA_TYPREDRTT2', 'ETB');
          T2.PutValue('PSA_TYPDADSFRAC', 'ETB');
          T2.PutValue('PSA_TYPPRUDH', 'ETB');
          T2.PutValue('PSA_CPTYPEVALO', 'ETB');
          T2.PutValue('PSA_TYPREGLT', 'ETB');
          T2.PutValue('PSA_TYPVIRSOC', 'ETB');
          T2.PutValue('PSA_TYPDATPAIEMENT', 'ETB');
          T2.PutValue('PSA_TYPPAIACOMPT', 'ETB');
          T2.PutValue('PSA_TYPACPSOC', 'ETB');
          T2.PutValue('PSA_TYPPAIFRAIS', 'ETB');
          T2.PutValue('PSA_TYPFRAISSOC', 'ETB');
          T2.PutValue('PSA_TYPJOURHEURE', 'DOS');
          T2.PutValue('PSA_TYPEDITORG', 'ETB');
          T2.PutValue('PSA_TYPACTIVITE', 'ETB');
          T2.PutValue('PSA_ACTIVITE', TEtab.GetValue('ETB_ACTIVITE'));
          T2.PutValue('PSA_EDITORG', TEtab.GetValue('ETB_EDITORG'));
          T2.PutValue('PSA_PROFILANCIEN', TEtab.GetValue('ETB_PROFILANCIEN'));
          T2.PutValue('PSA_PROFIL', TEtab.GetValue('ETB_PROFIL'));
          T2.PutValue('PSA_PROFILREM', TEtab.GetValue('ETB_PROFILREM'));
          T2.PutValue('PSA_PERIODBUL', TEtab.GetValue('ETB_PERIODBUL'));
          T2.PutValue('PSA_PROFILRBS', TEtab.GetValue('ETB_PROFILRBS'));
          T2.PutValue('PSA_REDREPAS', TEtab.GetValue('ETB_REDREPAS'));
          T2.PutValue('PSA_REDRTT1', TEtab.GetValue('ETB_REDRTT1'));
          T2.PutValue('PSA_REDRTT2', TEtab.GetValue('ETB_REDRTT2'));
          T2.PutValue('PSA_PROFILAFP', TEtab.GetValue('ETB_PROFILAFP'));
          T2.PutValue('PSA_PCTFRAISPROF', TEtab.GetValue('ETB_PCTFRAISPROF'));
          T2.PutValue('PSA_PROFILAPP', TEtab.GetValue('ETB_PROFILAPP'));
          T2.PutValue('PSA_PROFILRET', TEtab.GetValue('ETB_PROFILRET'));
          T2.PutValue('PSA_PROFILMUT', TEtab.GetValue('ETB_PROFILMUT'));
          T2.PutValue('PSA_PROFILPRE', TEtab.GetValue('ETB_PROFILPRE'));
          T2.PutValue('PSA_PROFILTSS', TEtab.GetValue('ETB_PROFILTSS'));
          T2.PutValue('PSA_PROFILTRANS', TEtab.GetValue('ETB_PROFILTRANS'));
          T2.PutValue('PSA_CALENDRIER', TEtab.GetValue('ETB_STANDCALEND'));
          T2.PutValue('PSA_PROFILFNAL', GetParamSoc('SO_PGPROFILFNAL'));
          T2.PutValue('PSA_PROFILCGE', TEtab.GetValue('ETB_PROFILCGE'));
          T2.PutValue('PSA_NBREACQUISCP', TEtab.GetValue('ETB_NBREACQUISCP'));
          T2.PutValue('PSA_NBACQUISCP', TEtab.GetValue('ETB_NBACQUISCP'));
          T2.PutValue('PSA_NBRECPSUPP', TEtab.GetValue('ETB_NBRECPSUPP'));
          T2.PutValue('PSA_VALORINDEMCP', TEtab.GetValue('ETB_VALORINDEMCP'));
          T2.PutValue('PSA_RELIQUAT', TEtab.GetValue('ETB_RELIQUAT'));
          T2.PutValue('PSA_BASANCCP', TEtab.GetValue('ETB_BASANCCP'));
          T2.PutValue('PSA_TYPDATANC', TEtab.GetValue('ETB_TYPDATANC'));
          T2.PutValue('PSA_MVALOMS', TEtab.GetValue('ETB_MVALOMS'));
          T2.PutValue('PSA_VALODXMN', TEtab.GetValue('ETB_VALODXMN'));
          T2.PutValue('PSA_EDITBULCP', TEtab.GetValue('ETB_EDITBULCP'));
          T2.PutValue('PSA_JOURHEURE', TEtab.GetValue('ETB_JOURHEURE'));
          T2.PutValue('PSA_JOURHEURE', GetParamSoc('SO_PGMETODHEURES'));
          T2.PutValue('PSA_PRUDHCOLL', TEtab.GetValue('ETB_PRUDHCOLL'));
          T2.PutValue('PSA_PRUDHSECT', TEtab.GetValue('ETB_PRUDHSECT'));
          T2.PutValue('PSA_PRUDHVOTE', TEtab.GetValue('ETB_PRUDHVOTE'));
          T2.PutValue('PSA_PGMODEREGLE', TEtab.GetValue('ETB_PGMODEREGLE'));
          T2.PutValue('PSA_RIBVIRSOC', TEtab.GetValue('ETB_RIBSALAIRE'));
          T2.PutValue('PSA_PAIACOMPTE', TEtab.GetValue('ETB_PAIACOMPTE'));
          T2.PutValue('PSA_RIBACPSOC', TEtab.GetValue('ETB_RIBACOMPTE'));
          T2.PutValue('PSA_PAIFRAIS', TEtab.GetValue('ETB_PAIFRAIS'));
          T2.PutValue('PSA_RIBFRAISSOC', TEtab.GetValue('ETB_RIBFRAIS'));
          T2.PutValue('PSA_MOISPAIEMENT', TEtab.GetValue('ETB_MOISPAIEMENT'));
          T2.PutValue('PSA_JOURPAIEMENT', TEtab.GetValue('ETB_JOURPAIEMENT'));
          T2.PutValue('PSA_CONFIDENTIEL', '0');
          if (GetParamSoc('SO_PGCONGES')) and (TEtab.GetValue('ETB_CONGESPAYES') = 'X') then T2.PutValue('PSA_CONGESPAYES', 'X')
          else T2.PutValue('PSA_CONGESPAYES', '-');
          // Traitement de création automatique des tiers salariés
          {if (GetParamSoc('SO_PGTYPENUMSAL') = 'NUM') and (GetParamSoc('SO_PGTIERSAUXIAUTO') = TRUE) then}//PT60
          if (GetParamSoc('SO_PGTIERSAUXIAUTO') = TRUE) then//PT60
          begin
            Libell := T2.GetValue('PSA_LIBELLE');
            LongLib := Length(Libell);
            Prenom := T2.GetValue('PSA_PRENOM');
            LongPre := Length(Prenom);
            Libell := Copy(Libell, 1, LongLib) + ' ' + Copy(Prenom, 1, LongPre);
            with InfTiers do
            begin
              Libelle := Copy(Libell, 1, 35);
              Adresse1 := T2.GetValue('PSA_ADRESSE1');
              Adresse2 := T2.GetValue('PSA_ADRESSE2');
              Adresse3 := T2.GetValue('PSA_ADRESSE3');
              Ville := T2.GetValue('PSA_VILLE');
              Telephone := T2.GetValue('PSA_TELEPHONE');
              CodePostal := T2.GetValue('PSA_CODEPOSTAL');
              Pays := '';
            end;
            CodeAuxi := CreationTiers(InfTiers, LeRapport, 'SAL', T2.GetValue('PSA_SALARIE'));
            if CodeAuxi <> '' then T2.PutValue('PSA_AUXILIAIRE', CodeAuxi);
          end;
          CreateYRS(T2.GetValue('PSA_SALARIE'), '', ''); //PT52
        end // Fin du cas création des salaries
        else
        begin // Cas de la modification des salariés
          st := 'UPDATE SALARIES SET PSA_DATEENTREE="' + UsDateTime(T1.GetValue('PIM_DATEENTREE')) + '",';
          st := st + 'PSA_LIBELLE="' + T1.GetValue('PIM_LIBELLE') + '",';
          st := st + 'PSA_PRENOM="' + T1.GetValue('PIM_PRENOM') + '",';
          st := st + 'PSA_ADRESSE1="' + T1.GetValue('PIM_ADRESSE1') + '",';
          st := st + 'PSA_ADRESSE2="' + T1.GetValue('PIM_ADRESSE2') + '",';
          st := st + 'PSA_ADRESSE3="' + T1.GetValue('PIM_ADRESSE3') + '",';
          st := st + 'PSA_CODEPOSTAL="' + T1.GetValue('PIM_CODEPOSTAL') + '",';
          st := st + 'PSA_VILLE="' + T1.GetValue('PIM_VILLE') + '",';
          st := st + 'PSA_TELEPHONE="' + T1.GetValue('PIM_TELEPHONE') + '",';
          st := st + 'PSA_DATENAISSANCE="' + UsDateTime(T1.GetValue('PIM_DATENAISSANCE')) + '",';
          st := st + 'PSA_NATIONALITE="' + T1.GetValue('PIM_NATIONALITE') + '",';
          st := st + 'PSA_SEXE="' + T1.GetValue('PIM_SEXE') + '",';
          st := st + 'PSA_NUMEROSS="' + T1.GetValue('PIM_NUMEROSS') + '",';
          st := st + 'PSA_CIVILITE="' + T1.GetValue('CIVILITE') + '",';
          st := st + 'PSA_PORTABLE="' + T1.GetValue('PORTABLE') + '",';
          st := st + 'PSA_DATESORTIE="' + UsDateTime(T1.GetValue('DATESORTIE')) + '",';
          // DEB PT37
          if VersionFic >= '0700' then // PT47 Uniformisation versionfic
          begin
            st := st + 'PSA_DEPTNAISSANCE="' + T1.GetValue('DEPNAI') + '",'; // Département naissance
            st := st + 'PSA_COMMUNENAISS="' + T1.GetValue('COMMNAI') + '",'; // Département naissance
            st := st + 'PSA_PAYSNAISSANCE="' + T1.GetValue('PAYNAI') + '",'; // Pays Naissance
            st := st + 'PSA_CARTESEJOUR="' + T1.GetValue('NUMCARTESEJ') + '",'; // Numéro carte de séjour
            st := st + 'PSA_DELIVPAR="' + T1.GetValue('DELIVPAR') + '",'; // Délivrée par
            st := st + 'PSA_SITUATIONFAMIL="' + T1.GetValue('SITFAM') + '",'; //  Situation de fammille
            st := st + 'PSA_MOTIFENTREE="' + T1.GetValue('MOTIFENTREE') + '",'; // Motif entrée
            st := st + 'PSA_MOTIFSORTIE="' + T1.GetValue('MOTIFSORT') + '",'; // Motif sortie
            st := st + 'PSA_CATDADS="' + T1.GetValue('CATDUCS') + '",'; // Catégorie DUCS
            st := st + 'PSA_REGIMESS="' + T1.GetValue('REGIMESS') + '",'; // Régime SS
            st := st + 'PSA_DADSCAT="' + T1.GetValue('STATUTCAT') + '",'; // Statut c&tégoriel
            st := st + 'PSA_DADSPROF="' + T1.GetValue('STATUFPROF') + '",'; // Statut professionnel
            st := st + 'PSA_CODEEMPLOI="' + T1.GetValue('CONDEMPLOI') + '",'; // Condition d'emploi
            st := st + 'PSA_UNITETRAVAIL="' + T1.GetValue('UNITETRAV') + '",'; // Unité de travail
            st := st + 'PSA_SORTIEDEFINIT="' + T1.GetValue('SORTIEDEF') + '",'; // Sortie définitive
            st := st + 'PSA_DATEXPIRSEJOUR="' + UsDateTime(T1.GetValue('DATEEXPIR')) + '",'; // Date expiration carte de séjour
            st := st + 'PSA_HEUREMBAUCHE="' + UsDateTime(T1.GetValue('HEUREMB')) + '",'; // Heure d'embauche
            st := st + 'PSA_TAUXPARTIEL=' + StrFPoint(T1.GetValue('TAUXTPSPARTIEL')) + ','; // Taux temps partiel PT47
          end;
          // FIN PT37
          st := st + 'PSA_DATEMODIF="' + UsDateTime(Date) + '"';
          st := st + ' WHERE PSA_SALARIE="' + T1.GetValue('PIM_SALARIE') + '"';
          ExecuteSQL(St);
        end;
      end; // Fin de la boucle de traitement des salariés à créer
      TSalACreer.InsertDB(nil, false);
      FreeAndNIL(TSalACreer);
      //    end;
      if Tob_sal <> nil then
      begin
        Tob_sal.Free;
        Tob_sal := nil;
      end;
      // permet le rechargement des salariés nouvellement crées
     { st := 'SELECT PSA_SALARIE,PSA_DATEENTREE,PSA_DATESORTIE,PSA_SUSPENSIONPAIE,PSA_ETABLISSEMENT, ' +
        'PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_CONFIDENTIEL,PSA_CONGESPAYES,' + { PT22 }
      {  'ETB_DATECLOTURECPN,ETB_CONGESPAYES FROM SALARIES LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT = ETB_ETABLISSEMENT';}
      st := 'SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,PSA_DATEENTREE,PSA_DATESORTIE,PSA_SUSPENSIONPAIE,PSA_ETABLISSEMENT, ' +
        'PSA_TRAVAILN1,PSA_TRAVAILN2,PSA_TRAVAILN3,PSA_TRAVAILN4,PSA_CODESTAT,PSA_CONFIDENTIEL,PSA_CONGESPAYES,' + { PT22 }
        'ETB_DATECLOTURECPN,ETB_CONGESPAYES FROM SALARIES LEFT JOIN ETABCOMPL ON PSA_ETABLISSEMENT = ETB_ETABLISSEMENT';
      Q := opensql(st, true);
      if (not Q.eof) then
      begin
        Tob_sal := TOB.Create('Table des Salariés', nil, -1);
        if tob_sal <> nil then
          Tob_sal.loaddetaildb('INFOS SALARIES', '', '', Q, false, false, -1, 0);
      end;
      Ferme(Q);
    end;
  end;
  result := true;
end;
{$IFNDEF COMSX}

function TOF_PGIMPORTFIC.ExisteRubrique(Rubrique: string): boolean;
{$ELSE}

function ExisteRubrique(Rubrique: string): boolean;
{$ENDIF}
var
  T: tob;
begin
  result := false;
  if not PgImpnotVide(tob_rubrique, true) then exit;

  T := Tob_rubrique.findfirst(['PRM_RUBRIQUE', 'PRM_PREDEFINI'], [rubrique, 'DOS'], true);
  if T = nil then
    T := Tob_rubrique.findfirst(['PRM_RUBRIQUE', 'PRM_PREDEFINI'], [rubrique, 'STD'], true);
  if T = nil then
    T := Tob_rubrique.findfirst(['PRM_RUBRIQUE', 'PRM_PREDEFINI'], [rubrique, 'CEG'], true);
  if T = nil then exit;
  TypeBase := T.getvalue('PRM_TYPEBASE');
  TypeTaux := T.getvalue('PRM_TYPETAUX');
  TypeCoeff := T.getvalue('PRM_TYPECOEFF');
  TypeMontant := T.getvalue('PRM_TYPEMONTANT');
  Libellerubrique := T.getvalue('PRM_LIBELLE');
  result := true;
end;

procedure TOF_PGIMPORTFIC.OnArgument(Arguments: string);
var
  NomFic, FicRapport: ThEdit;
  BValider, BImprime, ConsultErreur: TToolbarbutton97;
  st, LeNomFic, LeMode: string;
begin
  //  inherited;
  SetControlChecked('ANOABS', GetParamSocSecur('SO_PGEWSDEFAULTIGNORANO', False));    //PT58
  MatSouche := -999;
  TopCreat := FALSE;
  TFVierge(Ecran).Retour := 'NON'; // PT24 Initialisation avec un code retour
  st := Arguments;
  if st <> '' then
  begin
    LeNomFic := ReadTokenSt(St);
    LeMode := ReadTokenSt(St);
    Ecran.Caption := Ecran.Caption + ' NetExpert';
    PgImportDef(LeNomFic, LeMode);
  end
  else
  begin
    NomFic := THedit(GetControl('EDTFICIMPORT'));
    taille := 0;
    if Nomfic <> nil then
    begin
      NomFic.OnChange := FicChange;
      Nomfic.DataType := 'OPENFILE(*.*)'
    end;
    FicRapport := THedit(GetControl('FICRAPPORT'));
    if FicRapport <> nil then
    begin
      FicRapport.OnExit := FicChange;
      FicRapport.DataType := 'OPENFILE(*.*)';
    end;
  end; // PT24
  BValider := TToolbarbutton97(GetControl('BVALIDER'));
{$IFNDEF COMSX}
  if BValider <> nil then
    BValider.OnClick := ControleFichierImport;
{$ENDIF}
  BImprime := ttoolbarbutton97(getcontrol('BIMPRIMER'));
  if Bimprime <> nil then
    Bimprime.Onclick := ImprimeClick;

  Salarieprec := '';
  RapportErreur := THRichEditOLE(getcontrol('LISTEERREUR'));

  MarqueDebut := '***DEBUT***';
  MarqueFin := '***FIN***';
  Sep := THValComboBox(getcontrol('SEPARATEUR'));
  if (Sep <> nil) and (TypImport = '') then // PT24
  begin
    Sep.value := 'TAB';
    // PT18   10/09/2003 PH Controle du séparateur du fichier import et creation automatique fichier rapport
    Sep.OnExit := SepChange;
    Separateur := #9;
  end;
  ConsultErreur := TToolbarbutton97(getcontrol('CONSULTERREUR'));
  if ConsultErreur <> nil then
    ConsultErreur.Onclick := ConsultErreurClick;
  //  PT24  end;
  // DEB PT27
  Tob_Etab := TOB.Create('ETABCOMPL', nil, -1);
  Tob_Etab.LoadDetailDB('ETABCOMPL', '', '', nil, FALSE);
  //  FIN PT27

  Tob_MotifAbs := TOB.Create('MOTIFABSENCE', nil, -1); { PT39 }
  Tob_MotifAbs.LoadDetailDB('MOTIFABSENCE', '', '', nil, False); { PT39 }
end;

procedure TOF_PGIMPORTFIC.ConsultErreurClick(Sender: Tobject);
var
  FicRapport: THedit;
  F: TextFile;
  S: string;
  TPC: TPageControl;
  T: TTabSheet;
  i: integer;
  okok: Boolean;
begin
  if RapportErreur = nil then exit;
  FicRapport := THedit(GetControl('FICRAPPORT'));
  if FicRapport = nil then exit;
  RapportErreur.Clear;
  if FicRapport.Text = '' then exit;
  if not FileExists(FicRapport.text) then exit;

  valideEcran(false);
  AssignFile(F, FicRapport.Text);
  Reset(F);
  i := 0;
  InitMoveProgressForm(nil, 'Chargement du fichier d''erreurs à l''écran', 'Veuillez patienter SVP ...', taille, TRUE, TRUE);
  while not eof(F) do
  begin
    i := i + 1;
    Readln(F, S);
    okok := MoveCurProgressForm(IntToStr(i));
    if not okok then
    begin
      closeFile(F);
      FiniMoveProgressForm;
    end;
    RapportErreur.lines.Add(S);
  end;

  closeFile(F);
  ValideEcran(true);
  FiniMoveProgressForm;

  T := TTabSheet(GetControl('TT2'));
  if T <> nil then
  begin
    TPC := TPageControl(getcontrol('PAGECONTROL1'));
    if TPC <> nil then
      TPC.ActivePage := T;
  end;
end;

procedure TOF_PGIMPORTFIC.valideEcran(Top: boolean);
begin
  SetControlEnabled('EDTFICIMPORT', top);
  SetControlEnabled('FICRAPPORT', top);
  SetControlEnabled('SEPARATEUR', top);
  SetControlEnabled('BVALIDER', top);
  SetControlEnabled('BFERME', top);
  SetControlEnabled('CONSULTERREUR', top);
end;


// ligne absence
{function TOF_PGIMPORTFIC.EcritligneMAB(S:string,T:tob);
var
begin

end;}


procedure TOF_PGIMPORTFIC.SepChange(Sender: TObject);
begin
  // PT18   10/09/2003 PH Controle du séparateur du fichier import et creation automatique fichier rapport
  if Sep <> nil then
    if (Sep.value = 'TAB') or (Copy(Sep.Value, 1, 3) = 'Tab') then Separateur := #9
    else if (Sep.value = 'PIP') or (Copy(Sep.Value, 1, 3) = 'Pip') then Separateur := '|'
    else if (Sep.value = 'PVI') or (Copy(Sep.Value, 1, 3) = 'Poi') then Separateur := ';';
end;

procedure TOF_PGIMPORTFIC.FicChange(Sender: TObject);
var
  LeFic, St: string;
  F: TextFile;
begin
  // PT18   10/09/2003 PH Controle du séparateur du fichier import et creation automatique fichier rapport
  if not Assigned(Sender) then exit;
  LeFic := THEdit(Sender).Text;
  if LeFic = '' then exit;
  if THEdit(Sender).Name = 'FICRAPPORT' then
  begin
    if not FileExists(LeFic) then
    begin
      AssignFile(F, LeFic);
{$I-}ReWrite(F);
{$I+}if IoResult <> 0 then
      begin
        PGIBox('Fichier inaccessible : ' + LeFic, 'Abandon du traitement');
        Exit;
      end;
      closeFile(F);
    end;
  end
  else
  begin
    AssignFile(F, LeFic);
    reset(F);
    ReadLn(F, St);
    ReadLn(F, St);
    if Sep <> nil then
    begin
      if Pos(';', St) > 0 then
      begin
        Sep.Value := 'PVI';
        Separateur := ';';
      end
      else if Pos('|', St) > 0 then
      begin
        Sep.Value := 'PIP';
        Separateur := '|';
      end;
    end;
    closeFile(F);
  end;
  // FIN PT18
end;

procedure TOF_PGIMPORTFIC.ImprimeClick(Sender: TObject);
var
  MPages: tpagecontrol;
begin
{$IFNDEF EAGLCLIENT}
  MPages := TPageControl(getcontrol('PAGECONTROL1'));
  if MPages <> nil then
    PrintPageDeGarde(MPages, TRUE, TRUE, FALSE, Ecran.Caption, 0);
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Dev Paie
Créé le ...... : 11/09/2003
Modifié le ... :   /  /
Description .. : Le nom du fichier obligatoirement renseigné avec une
Suite ........ : extension tq .t (.txt,.trt,.ttt ....)
Suite ........ : Fichier existant obligatoirement
Suite ........ : Le Type d'import = 'AUTO' sinon considéré comme appli paie
Suite ........ : Le séparateur est un Char := |
Suite ........ :
Suite ........ : Test si creation du fichier de rapport peut être créer mais
Suite ........ : il aura une extension en .log au lieu de .trt ou .txt ou .nsv (NetExpert PCL)
Mots clefs ... : PAIE;GIGA
*****************************************************************}
{$IFNDEF COMSX}

function TOF_PGIMPORTFIC.PgImportDef(LeNomduFichier, LeTypImp: string; LeSeparateur: Char = '|'): Boolean;
{$ELSE}

function PgImportDef(LeNomduFichier, LeTypImp: string; LeSeparateur: Char = '|'): Boolean;
{$ENDIF}
var
  LeRapport: string;
  ll: Integer;
  Sep: THValComboBox;
begin
  result := FALSE;
  MatSouche := -999;
  Salarieprec := '';
  MarqueDebut := '***DEBUT***';
  MarqueFin := '***FIN***';
  TopCreat := FALSE;
  if LeNomDuFichier = '' then exit;
  //  if not FileExists(LeNomduFichier) then exit;
  ll := Pos('.NSV', LeNomduFichier);
  if ll <= 0 then exit; // PT24
  if LeTypImp <> 'AUTO' then exit;
{$IFNDEF COMSX}
  if (LeSeparateur <> '|') and (ll <= 0) then exit
  else
    if ll > 0 then LeSeparateur := ';';
  SetControlText('EDTFICIMPORT', LeNomDuFichier);
  SetControlEnabled('EDTFICIMPORT', FALSE);
  LeRapport := Copy(LeNomDuFichier, 1, ll) + '.log';
  SetControlText('FICRAPPORT', LeRapport);
  SetControlEnabled('FICRAPPORT', FALSE);
  SetControlEnabled('SEPARATEUR', FALSE);
  Sep := THValComboBox(getcontrol('SEPARATEUR'));
  if (Sep <> nil) then
  begin
    if (ll <= 0) then Sep.value := 'PIP'
    else Sep.value := 'PVI';
  end;
{$ELSE}
  FreeAndNil(Tob_Etab);
  Tob_Etab := TOB.Create('ETABCOMPL', nil, -1);
  Tob_Etab.LoadDetailDB('ETABCOMPL', '', '', nil, FALSE);
{$ENDIF}
  Separateur := LeSeparateur;
  NomDuFichier := LeNomDuFichier;
  TypImport := LeTypImp;
  result := TRUE;
end;

(* PT29 function TOF_PGIMPORTFIC.ImportRendPeriodeEnCours(var ExerPerEncours, DebPer, FinPer: string): Boolean;
var
  Q: TQuery;
begin
  result := FALSE;
  DebPer := DateToStr(idate1900);
  FinPer := DateToStr(idate1900);
  Q := OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC', TRUE);
  if not Q.EOF then
  begin
    DebPer := Q.FindField('PEX_DEBUTPERIODE').AsString;
    FinPer := Q.FindField('PEX_FINPERIODE').AsString;
    ExerPerEncours := Q.FindField('PEX_EXERCICE').AsString;
    result := TRUE;
    //  PT1
  end;
  Ferme(Q);
end;    *)

(* PT29 function TOF_PGIMPORTFIC.ImportRendExerSocialEnCours(var MoisE, AnneeE, ComboExer: string; var DebExer, FinExer: TDateTime): Boolean;
var
  Q: TQuery;
  DatF: TDateTime;
  Jour, Mois, Annee: WORD;
begin
  result := FALSE;
  DebExer := idate1900;
  FinExer := idate1900;
  Q := OpenSQL('SELECT * FROM EXERSOCIAL WHERE PEX_ACTIF="X" ORDER BY PEX_ANNEEREFER DESC', TRUE);
  if not Q.EOF then
  begin
    DatF := Q.FindField('PEX_FINPERIODE').AsFloat; //Q.Fields[7].AsFloat; // Recup date de fin periode en cours
    DecodeDate(DatF, Annee, Mois, Jour);
    MoisE := ImportColleZeroDevant(Mois, 2);
    ComboExer := Q.FindField('PEX_EXERCICE').AsString; //Q.Fields[0].AsString; // recup Combo identifiant exercice
    AnneeE := Q.FindField('PEX_ANNEEREFER').AsString; //Q.Fields[8].AsString; // recup Annee de exercice
    DebExer := Q.FindField('PEX_DATEDEBUT').AsDateTime;
    FinExer := Q.FindField('PEX_DATEFIN').AsDateTime;
    result := TRUE;
    // PT1
  end;
  Ferme(Q);
end;         *)

(* PT29 function TOF_PGIMPORTFIC.ImportColleZeroDevant(Nombre, LongChaine: integer): string;
var
  tabResult: string;
  TabInt: string;
  i, j: integer;
begin
  tabResult := '';
  for i := 1 to LongChaine do
  begin
    if Nombre < power(10, i) then
    begin
      TabInt := inttostr(Nombre);
      // colle (LongChaine-i zéro devant]
      for j := 0 to (LongChaine - i - 1)
        do insert('0', TabResult, j);
      result := concat(TabResult, Tabint);
      exit;
    end;
    if i > LongChaine then result := inttostr(Nombre);
  end;
end;    *)

// DEB PT27


procedure TOF_PGIMPORTFIC.OnClose;
begin
  inherited;
  if Tob_SalATT <> nil then
  begin
    Tob_SalATT.Free;
    Tob_SalATT := nil;
  end;
// d memcheck
  if Tob_Etab <> nil then
  begin
    Tob_Etab.Free;
    Tob_Etab := nil;
  end;
// f memcheck
  if Assigned(Tob_MotifAbs) then FreeAndNil(Tob_MotifAbs); { PT39 }
end;
// FIN PT27


{$IFNDEF COMSX}

procedure TOF_PGIMPORTFIC.ConversionJourHeure(chaine, NBj, Nbh: string; var VNbj, VNbh: string);
{$ELSE}

procedure ConversionJourHeure(chaine, NBj, Nbh: string; var VNbj, VNbh: string);
{$ENDIF}
var
  virgule, lgnb: integer;
begin
  if isNumeric(Nbj) then
  begin
    virgule := pos(',', Nbj);
    lgnb := length(Nbj);
    if (virgule > 0) and (lgnb > virgule) then
      vnbj := copy(nbj, 1, virgule - 1) + '.' + copy(nbj, virgule + 1, lgnb - virgule - 1);
  end;
  if not isNumeric(Nbj) then
    if Nbj <> '' then Erreur(chaine, Nbj, 150);

  if isNumeric(Nbh) then
  begin
    virgule := pos(',', Nbh);
    lgnb := length(Nbh);
    if (virgule > 0) and (lgnb > virgule) then
      vnbh := copy(nbh, 1, virgule - 1) + '.' + copy(nbh, virgule + 1, lgnb - virgule - 1);
  end;
  if not isNumeric(NbH) then
    if ((NbH <> '') and (isNumeric(NbJ))) then
      Erreur(chaine, NbH, 160);
end;



initialization
  registerclasses([TOF_PGIMPORTFIC]);
end.

