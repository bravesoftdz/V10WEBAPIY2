unit Recordcom;

interface
uses classes, SysUtils, Hctrls, DB, Ent1, Paramsoc,Hent1,
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  Dialogs, UtilTrans, MajTable, ControlParam, UTOB, utilPGI;  // AJOUT SUR 7XX

{$IFNDEF EAGLCLIENT}
procedure Ecritureentete( var F: TextFile; Typearchive, Mode : string;Origine : string=''; pp : string=''; Complement : string='');
procedure EcritureJournal(ListeJalB: TList; var F: TextFile; Libere : Boolean=TRUE);
procedure EcritureCpptegen(ListeCpteGen: TList; var F: TextFile);
procedure EcritureMouvement(ListeMouvt: TList; var F: TextFile; Formatm : string; var Corresp : string);
procedure EcritureCppteaux(ListeCpteaux: TList; var F: TextFile; stArg : string='');
procedure EcritureTierscomp(var F: TextFile; Where : string); // complement tiers
// ajout me  23/05/2003
procedure EcritureRib(ListeRib: TList; var F: TextFile);
procedure Ecrituregeneraux(var F: TextFile);
procedure EcritureExercice(var F: TextFile;  Where : string='');
procedure Ecrituretablelibre(ListeLibre: TList; var F: TextFile);
procedure EcritureChoixCod(var F: TextFile; Serie : string='S5');
procedure EcritureParamlib(var F: TextFile; TP : TOB);
procedure EcritureDesRessource(var F: TextFile; Where : string='');

procedure EcritureEtablissement(var F: TextFile; Where : string='');
procedure Ecrituretablesection(ListeSection: TList; var F: TextFile);
procedure EcritureModepaiement(var F: TextFile; var Corresp : string);
procedure Ecrituretableregle(Listeregle: TList; var F: TextFile; Origine : string ;stArg : string='');
procedure Ecrituredevise(var F: TextFile);
procedure Ecrituretva(var F: TextFile);
procedure EcritureSectionana(var F: TextFile; Where :string);
procedure EcritureSouche(var F: TextFile; Where : string);
procedure EcritureCorresp(var F: TextFile);
procedure EcritureCorrespimpot(var F: TextFile);
function Comparejournal (Item1,Item2:Pointer) : integer;
procedure EcritureBanquecp(var F: TextFile);
procedure EcritureBanques(var F: TextFile);
procedure EcritureReleveBanque(var F: TextFile);
procedure EcritureLigneReleveBanque(var F: TextFile);
procedure EcritureContact(var F: TextFile; WW : string='');
function CoherenceMDRS1 (var Code : string) : Boolean;
Function findcorrespMDR (mr : string; stArg : string) : string;
procedure LiberJournal(ListeJalB: TList);
procedure EcrireCpteBudgetaire(var F: TextFile; Where :string);
procedure EcrireJalBudgetaire(var F: TextFile; Where :string);
Procedure EcritureSectionBudgetaire(var F: TextFile; Where :string);
Procedure EcritureRelance(var F: TextFile; Where :string);
{$ENDIF}
procedure LibererGen (ListeCpteGen: TList);
function CompteEstalpha (ListeCpteGen: TList) : Boolean ;
Procedure EcritureBapTypeVisa(var F: TextFile);
Procedure EcritureBapCircuit(var F: TextFile);
procedure EcritureVentilType (var F: TextFile);
{$IFDEF CERTIFNF}
procedure EcritureSuiviValidation (var F: TextFile);
{$ENDIF}

const            //***S5EXPDOSSTD
  versionexport   = '008'; // version du fichier
  Sousversion     = '001';

  Formatentete = '***%2.2s%-3.3s%-3.3s%3.3s%-3.3s%-8.8s%-8.8s%-3.3s%-5.5s%-12.12s%-35.35s%-35.35s%1.1s%3.3s%6.6s%-3.3s%8.8s%-3.3s';
  Formatgeneraux1 = '***PS1%-35.35s%-35.35s%-35.35s%-35.35s%-9.9s%-35.35s%-3.3s%-25.25s%-25.25s%-25.25s%-35.35s%-35.35s%-35.35s%-17.17s%-35.35s%-17.17s'
    +
    '%-35.35s%-20.20s%-35.35s';
  Formatgeneraux2 = '***PS2%-2.2s%1.1s%-2.2s%1.1s%2.2s%1.1s%2.2s%1.1s%2.2s%1.1s%2.2s%1.1s%2.2s%1.1s'
    +
    '%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%17.17s%3.3s%1.1s%1.1s%1.1s';

  Formatgeneraux3 =
    '***PS3%-17.17s%-17.17s%-17.17s%-17.17s%-3.3s%-3.3s%-3.3s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s';
  Formatgeneraux4 = '***PS4%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%17.17s'
    +
    '%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%17.17s' +
    '%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%17.17s';
  Formatgeneraux5 =
    '***PS5%-3.3s%1.1s%1.1s%-3.3s%-17.17s%-17.17s%-10.10s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-2.2s%1.1s%1.1s%1.1s';
  Formatexercice = '***EXO%-3.3s%-8.8s%-8.8s%-3.3s%-3.3s%-35.35s%-3.3s';

  FormatJournal =
    '***JAL%-3.3s%-35.35s%-3.3s%-3.3s%-3.3s%-17.17s%-3.3s%-3.3s%-200.200s%-200.200s%-17.17s%1s';
  FormatComptegene =
    '***CGN%-17.17s%-35.35s%3.3s%1.1s%1.1s%1.1s%1.1s%1.1s%1.1s%1.1s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-3.3s%-17.17s%-17.17s%-3.3s%-3.3s%-3.3s'+
    //Sauvegarde de qualifiant qte  Reprendre dans 6.5
    '%1s%3.3s%1s%1s%-17.17s%-17.17s%1s%3.3s%3.3s';

  FormatCompteaux = '***CAE%-17.17s%-35.35s%3.3s%1.1s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s'+
    '%-17.17s%-17.17s%-17.17s%-17.17s%-35.35s%-35.35s%-35.35s%-9.9s%-35.35s%-24.24s%-5.5s%-5.5s%-11.11s%2.2s%3.3s' +
    '%-17.17s%3.3s%1.1s%3.3s%-25.25s%-25.25s%3.3s%3.3s%-35.35s%-17.17s%-17.17s%-5.5s%-35.35s%-35.35s%-35.35s%-25.25s%-25.25s%-25.25s%-50.50s%3.3s' +
    '%1.1s%3.3s%1.1s%-3.3s%-17.17s%1s%1s%3s%3s%1s%17.17s%17.17s%20.20s%20.20s%-17.17s%3.3s%-20.20s%-20.20s%-20.20s'; // fiche com 31027

  FormatTierscomp = '***YTC%-17.17s%-17.17s%3.3s%-6.6s%-6.6s%-6.6s%-6.6s%-6.6s%-6.6s%-6.6s%-6.6s%-6.6s%-6.6s'+
    '%-15.15s%-15.15s%-15.15s%-8.8s%-8.8s%-8.8s%-6.6s%-6.6s%-6.6s%-15.15s%-15.15s%-15.15s%-8.8s%-8.8s%-8.8s' +
    '%-35.35s%-35.35s%-35.35s%1s%1s%1s%-17.17s%-17.17s%-17.17s%3.3s%-8.8s%-8.8s%-35.35s%-35.35s%-3.3s%-17.17s%-17.17s'+
    '%-17.17s%-15.15s%-15.15s%-15.15s%-6.6s%-6.6s%-17.17s%-17.17s%-17.17s%-17.17s%1s%1s%35.35s%3.3s%3.3s%3.3s';

  Formatmvt =
    '%-3.3s%-8.8s%2.2s%-17.17s%1.1s%-17.17s%-35.35s%-35.35s%-3.3s%-8.8s%1.1s%20.20s%1.1s%8.8s%-3.3s%-10.10s%-3.3s%20.20s%20.20s%-3.3s%-2.2s%2.2s';
  Formatmvtetendu =
    '%-3.3s%-8.8s%2.2s%-17.17s%1.1s%-17.17s%-35.35s%-35.35s%-3.3s%-8.8s%1.1s%20.20s%1.1s%8.8s%-3.3s%-10.10s%-3.3s%20.20s%20.20s%-3.3s%-2.2s%2.2s' +
    '%-35.35s%8.8s%8.8s%3.3s%-17.17s%8.8s%3.3s%20.20s%20.20s%3.3s%3.3s%35.35s%1.1s%-3.3s%-3.3s%-3.3s%-17.17s%-17.17s%-17.17s' +
    '%8.8s%8.8s%8.8s%35.35s%10.10s%-17.17s%-30.30s%-30.30s%-30.30s%-30.30s%-30.30s%-30.30s%-30.30s%-30.30s' +
    '%-30.30s%-30.30s%-3.3s%-3.3s%-3.3s%-3.3s%-20.20s%-20.20s%-20.20s%-20.20s%-8.8s%-1.1s%-1.1s%-3.3s%-20.20s%-20.20s%-20.20s' +
    '%-8.8s%-8.8s%-5.5s%1.1s%1.1s%-3.3s%-17.17s%-17.17s%-17.17s%-17.17s%-35.35s%3.3s%10.10s%3.3s%-17.17s%-17.17s%1.1s' +
// ajout me pour ecrcompl
    '%-8.8s%-8.8s%-8.8s%-35.35s%1s%1s%3.3s%-8.8s'+ // fiche 10487 + BVE 28.08.07 Ajout E_DOCID (validation ecriture)
    '%17.17s%3.3s%8.8s%3.3s'; // ajout me des zones pour echange et e_qualiforigne

  Formatmvtetenduana =
    '%-3.3s%-8.8s%2.2s%-17.17s%1.1s%-17.17s%-35.35s%-35.35s%-3.3s%-8.8s%1.1s%20.20s%1.1s%8.8s%-3.3s%-10.10s%-3.3s%20.20s%20.20s%-3.3s%-2.2s%2.2s' +
    '%-35.35s%8.8s%8.8s%3.3s%17.17s%8.8s%3.3s%20.20s%20.20s%3.3s%3.3s%35.35s%1.1s%3.3s%3.3s%3.3s%17.17s%17.17s%1.1s%-17.17s%-17.17s' +
    '%17.17s%17.17s%17.17s%17.17s%-30.30s%-30.30s%-30.30s%-30.30s%-30.30s%-30.30s%-30.30s%-30.30s' +
    '%-30.30s%-30.30s%-3.3s%-3.3s%-3.3s%-3.3s%-20.20s%-20.20s%-20.20s%-20.20s%-8.8s%-1.1s%-1.1s%-3.3s%-20.20s%-20.20s%-20.20s' +
    '%-8.8s%-8.8s%-5.5s%1.1s%1.1s%-3.3s%-17.17s%-17.17s%-17.17s%-17.17s%-35.35s%3.3s%10.10s%3.3s%-17.17s%-17.17s%1.1s%1s';
  FormatGed =  // format Ged
    '%-3.3s%-8.8s%2.2s%-17.17s%1.1s%-17.17s%-35.35s%-255.255s';

  Formattablelibre =
    '***TL%1.1s%-17.17s%-35.35s%-3.3s%-35.35s%-35.35s%-35.35s%-35.35s%-35.35s%-35.35s%-35.35s%-35.35s%-35.35s%-35.35s';

// table libre personnalisé
  FormatChoixCode =
    '***%-3.3s%-3.3s%-17.17s%-70.70s%-17.17s%-70.70s';
  FormatParamlib =
    '***PTL%-3.3s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s'+
    '%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s'+
    '%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s%-35.35s%-1s';

// Ressource
  FormatRessource =
    '***ARS%-17.17s%-17.17s%-35.35s%-35.35s';

  Formatetablissement = '***ETB%-3.3s%-35.35s%-35.35s%-35.35s%-35.35s%-9.9s%-35.35s%-3.3s%-25.25s%-25.25s%-17.17s%5.5s';
  Formatsection = '***SSA%-17.17s%-35.35s%3.3s%3.3s%2.2s%2.2s%35.35s';
  Formatpaiement = '***MDP%-3.3s%-35.35s%3.3s%3.3s%1.1s%1.1s%1.1s%20.20s%3.3s%-17.17s'; // fiche 10604
  Formatregle = '***MDR%-3.3s%-35.35s%-3.3s%-3.3s%-3.3s%-2.2s%-3.3s%-20.20s%-3.3s%-3.3s'
    +
    '%-6.6s%-3.3s%-6.6s%-3.3s%-6.6s%-3.3s%-6.6s%-3.3s%-6.6s%-3.3s%-6.6s%-3.3s%-6.6s%-3.3s%-6.6s%-3.3s%-6.6s' +
    '%-3.3s%-6.6s%-3.3s%-6.6s%-3.3s%-6.6s';
  Formatdevise = '***DEV%-3.3s%-35.35s%3.3s%1.1s%1.1s%6.6s%1.1s%20.20s%3.3s%1.1s%-17.17s'
    +
    '%-17.17s%-17.17s%-17.17s%20.20s%20.20s';
  FormatTVA = '***REG%-17.17s%-35.35s';
  FormatSectionana =
    '***SAT%-17.17s%-35.35s%-3.3s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-3.3s';
  FormatSouche = '***SOU%-3.3s%-35.35s%-015d%1.1s%1.1s%1.1s';
  Formatcorresp = '***%3.3s%-17.17s%-35.35s%-3.3s';
  FormatcorrespImp = '***%3.3s%-17.17s%-17.17s%-3.3s';

  Formatrupture = '***RUP%-17.17s%-17.17s%-3.3s';
  // ajout me nature éco et nouveau format des comptes
  FormatRib = '***%-3.3s%-17.17s%6.6s%1.1s%-5.5s%-5.5s%-11.11s%2.2s%-24.24s%-35.35s%-3.3s%-3.3s%-35.35s%-3.3s%1.1s%1.1s%1.1s%70.70s%3.3s%1s%8.8s%20.20s';
  // AJOUT SUR 7XX
  FormatBqcp = '***BQC%-17.17s%-17.17s%35.35s%24.24s%-35.35s%-35.35s%-35.35s%9.9s%-35.35s%-9.9s%-3.3s%-3.3s%-3.3s%-25.25s%25.25s%35.35s%5.5s%11.11s%2.2s%-5.5s%-35.35s%-6.6s%6.6s'+
  '%6.6s%7.7s%40.40s%40.40s%40.40s%40.40s%40.40s%40.40s%40.40s%5d'+
  '%5d%5d%5d%5d%5d%-17.17s%20d%20d%10.10s%1s%3.3s%1s%1s%-3.3s%-3.3s%-3.3s%20d%20d%40.40s%1s%1s%3.3s'+
  '%3.3s%5.5s%5.5s%17.17s%3.3s%3.3s%35.35s%3.3s%6.6s%5.5s%1s%-70.70s%-8.8s%3.3s%3.3s';

  FormatBanque = '***BQE%3.3s%35.35s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s'+
  '%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s'+
  '%20.20s%20.20s%-05d%-05d%10.10s%10.10s%20.20s%-3.3s%-3.3s'+
  '%20.20s%20.20s%-05d%-05d%10.10s%10.10s%20.20s%-3.3s%-3.3s'+
  '%20.20s%20.20s%-05d%-05d%10.10s%10.10s%20.20s%-3.3s%-3.3s%-3.3s'+
  '%20.20s%20.20s%-05d%-05d%10.10s%10.10s%20.20s%20.20s%20.20s%10.10s%5.5s%17.17s';

  FormatRlvBqe = '***EXB%17.17s%17.17s%8.8s%8.8s%20.20s%20.20s%20.20s%20.20s%20.20s%20.20s%20.20s%20.20s'+
  '%-3.3s%10d%-35.35s%-3.3s%-8.8s%20.20s%-20.20s%-20.20s'+
  '%-20.20s%-10d%-3.3s%1s%-8.8s%-3.3s%10d%1s';

  FormatRlvLig = '***EXL%17.17s%10d%10d%-35.35s%-35.35s%1s%3.3s%8.8s%8.8s%-20.20s%-20.20s%-20.20s%-20.20s'+
  '%-35.35s%-17.17s%-17.17s%8.8s%2.2s%10d%10d%10d'+
  '%-50.50s%-50.50s%-50.50s%3.3s';

   // contact
   FormatContact = '***CON%-3.3s%-35.35s%-06d%3.3s%1s%-35.35s%-35.35s%-35.35s%-35.35s%-25.25s%-25.25s%-25.25s%-250.250s'+
   '%3.3s%3.3s%3.3s%3.3s%-06d%-06d%-06d%3.3s%1s%-35.35s%-35.35s%-35.35s%1s%1s%1s%3.3s%3.3s%3.3s%8.8s%8.8s%8.8s%-015d%-015d%-015d%-17.17s%-25.25s';

   // Les budgets
   FormatmvtBudg =
    '%-3.3s%-8.8s%2.2s%-17.17s%1.1s%-17.17s%-35.35s%-35.35s%-3.3s%-8.8s%1.1s%20.20s%1.1s%8.8s%-3.3s%-10.10s%-3.3s%20.20s%20.20s%-3.3s%-2.2s%2.2s' +
    '%-35.35s%8.8s%8.8s%3.3s%17.17s%8.8s%3.3s%20.20s%20.20s%3.3s%3.3s%35.35s%-35.35s%-35.35s%-35.35s%-35.35s'+
    '%-17.17s%-17.17s%-17.17s%-17.17s';

   FormatcpteBudg =
    '***FBG%-17.17s%-35.35s%-3.3s%-3.3s%1.1s%-3.3s%1.1s%-10.10s%-10.10s%1.1s%-10.10s%-3.3s%-17.17s%-17.17s%-17.17s%-17.17s' +
    '%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-250.250s%-250.250s';

   FormatJalBudg =
    '***FBJ%-3.3s%-35.35s%-3.3s%-10.10s%-10.10s%-10.10s%3.3s%-3.3s%-3.3s%3.3s%1.1s%-3.3s%-10.10s%-10.10s%-3.3s%-3.3s%-35.35s' +
    '%-1024.1024s%-1024.1024s%-1024.1024s%-1024.1024s';

   FormatSectionBudg =
    '***FBS%-17.17s%-3.3s%-35.35s%-17.17s%-6.6s%-10.10s%-10.10s%1s%-10.10s%3.3s%-3.3s%1s%-3.3s%-3.3s' +
    '%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-17.17s%-250.250s%-250.250s';

  // ajout me 26/08/2005
  FormatRelance =
      '***RRL%-3.3s%-3.3s%-35.35s%1s%1s%1s%-5.5s%-3.3s%-5.5s%3.3s%-5.5s%3.3s%-5.5s%3.3s%-5.5s%3.3s' +
      '%-5.5s%3.3s%-5.5s%3.3s%1s';

  // ajout me 27/07/2006
  FormatBap = '%-3.3s%-8.8s%2.2s%-17.17s%1.1s%-17.17s%8.8s%-3.3s%-3.3s%-3.3s%-3.3s%8.8s%-5.5s%-8.8s'+
  '%1s%-35.35s%-3.3s%3.3s%8.8s%8.8s%8.8s%-8.8s%-3.3s%-3.3s'+
  '%1s%1s%1s%1s';

  FormatBapVisa = '***CTI%-3.3s%-35.35s%-17.17s%5.5s%-3.3s%-20.20s%-20.20s%-3.3s%3.3s'+
  '%-3.3s%-3.3s%-3.3s%-3.3s%-3.3s%-35.35s%-35.35s%35.35s%-3.3s%-3.3s%-3.3s%-3.3s%-250.250s%-250.250s';

  FormatBapCircuit = '***CCI%-3.3s%5.5s%-3.3s%-3.3s%-5.5s%-35.35s';

  FormatVentilType = '***VEN%-3.3s%17.17s%17.17s%-20.20s%-20.20s%-20.20s%-8.8s%-3.3s%-20.20s%-17.17s%-17.17s%-17.17s%-17.17s%17.17s%17.17s';

  // BVE 28.08.07 : Suivi des validations
{$IFDEF CERTIFNF}
  FormatSuiviValidation = '***CPJ%-8.8s%8.8s%-3.3s%-8.8s%-8.8s%-8.8s';
{$ENDIF}

type
  TFTransfertcom = record
    FichierSortie, Exo   : string;
    Jr                   : string;
    Dateecr1, Dateecr2   : TDateTime;
    Etabi, ptiersautre   : string;
    Exporte, balance     : Boolean;
    Pgene, ptiers, pana  : string;
    TypeFormat           : string;
    naturejal, Serie     : string;
    Typearchive          : string;
    psection,pjournaux   : string;
    complement,Synch     : string;
    SuppCarAux,TLibre    : Boolean;
  end;

type
  TFJournal = record
    code, // code journal(3)
    libelle, // libelle (35)
    nature, // nature (3)
    souchen, // souche écriture (3)
    souches, // souche simulation (3)
    compte, // contrepartie (17)
    axe,    // axe (3)
    modesaisie, // -,BOR,LIB
    cptauto, // compte automatique (200)
    cptint, // comptes interdits (200)
    abrege: // libelle abrege (17)
    string;

  end;
  PListeJournal = ^TFJournal;

  TFTablelibre = record
    identifiant, // Tlx
    code, // code journal(3)
    libelle, // libelle (35)
    typetable, // nature (3)
    tx0, tx1, tx2, tx3, tx4, tx5, tx6, tx7, tx8, tx9:
    string;
  end;
  PLibre = ^TFTablelibre;

  TFTablesection = record
    code, // code (3)
    libelle, // libelle (35)
    axe, // axe (3)
    // table STRUCRSE : contient la définition des sous plans
    // SSSTRUCR : Contient le contenu des sous plan
    //Le PS_SOUSSECTION de  SSSTRUCR correspond au SS_SOUSSECTION de STRUCRSE
    //CODESP=PS_SOUSSECTION
    //LIBELLESP=PS_LIBELLE
    //CODE=SS_SOUSSECTION
    //LIBELLE=SS_LIBELLE
    codesp, debsp, lgsp, libellesp:
    string;
  end;
  PSection = ^TFTablesection;

  // mode de reglement
  TFTablereglement = record
    code, // code (3)
    libelle, // libelle (35)
    apartirde,
      plusjour, arrondijour, nbeche, separepar, montantmin, remplacemin,
      mp1, taux1, mp2, taux2, mp3, taux3, mp4, taux4, mp5, taux5, mp6, taux6,
      mp7, taux7, mp8, taux8, mp9, taux9, mp10, taux10, mp11, taux11, mp12,
        taux12:
    string;
  end;
  Preglement = ^TFTablereglement;

  TFCompteGencom = record
    Code: string; // code comptable (17)
    Libelle: string; // libelle (35)
    Nature: string; // char (3)
    Lettrable: string; // (1)
    Pointage: string; // (1)
    Ventilaxe1: string; // (1)
    Ventilaxe2: string; // (1)
    Ventilaxe3: string; // (1)
    Ventilaxe4: string; // (1)
    Ventilaxe5: string; // (1)
    Table1, Table2, Table3, Table4,
    Table5, Table6, Table7, Table8,
    Table9, Table10,
    abrege, // libelle abrege (17)
    sens, Corresp1, Corresp2,
    Tva, Encaissement, TPF, condifentiel,
    ctoff,ctoffper,cteche,visarev,cyclerev,ctcpte,
    //Sauvegarde de qualifiant qte  Reprendre dans 6.5
    qualqte1, qualqte2 : string;

  end;
  PListeCpteGen = ^TFCompteGencom;

type
  TFAuxiliaire = record
      code, libelle, nature, lettrable, collectif, ean,
      tb1, tb2, tb3, tb4, tb5, tb6, tb7, tb8, tb9, tb10,
      adresse1, adresse2, adresse3, codepostal, ville, domicile, etablissement,
      guichet, compte, cle, pays, abrege, langue, multidevise,
      devisetiers, tel, fax, regimetva, modereg, commentaire,
      nif, siret, ape, ctnom, ctservice, ctfonction, cttel, ctfax,
      cttelex, ctrva, ctcivilite, ctprincipal, formjur, rib,tvaenc,
      tierspayeur, ispayeur, remboursement,
      relanceregl, relancetrait, confidentiel, corresp1, corresp2, escompte, remise,
      facture, fromjur, ceditdemande, ceditaccorde, ceditplafond  : string;  // fiche com 31027
  end;
  PListeAux = ^TFAuxiliaire;
  TFRib = record
        ident,code,numero,prinpipal,codebanque,codeguichet,nocompte,cle,domicile,ville,pays,devise,bic,
        soc,ribsal,ribcompte,ribprof, iban, nateco,typpays : string;   // ajout me nature éco et nouveau format des comptes
  end;
  PListeRib = ^TFRib;

type
  TMouvement = record
    Journal, // code journal
    Datecomptable, // date pièce JJMMAAAA
    Typepiece, // type piece FC, AC, RC, FF, AF, RF, OD, OC, OF
    General, // compte général
    Typecpte, // type de compte X, A, H, O
    AuxSect, // compte auxiliaire  ou section analytique
    Refinterne, // ref du mouvement
    Libelle, // libelle du mouvement
    Modepaie, // mode de paiement
    Echeance, // date d'échéance
    Sens, // sens
    Montant1, // Montant de mouvement
    Typeecriture, // Type d'écriture
    NumPiece, // numéro de pièce
    Devise, // devise
    TauxDev, // taux devise
    CodeMontant, // code montant F, E, D,-
    Montant2, // montant devise
    Montant3, // montant contrevaleur
    Etablissement, //Etablissement
    Axe, // axe
    Numeche, // multi échéance
    // fin standard
    RefExterne, // référence externe
    Daterefexterne, // date ref externe
    Datecreation, // date systeme
    Societe, // code societe
    Affaire, // code affaire
    Datetauxdev, // date du taux de la devise
    Ecranouveau, // 'N',OAN:A,H:A
    Qte1, // dont 4 dec
    Qte2, // dont 4 dec
    QualifQte1, // qualifiant qte1
    QualifQte2, // qualifiant qte1
    Reflibre, // Ref libre mouvement
    Tvaencaiss, // tva encaissement
    Regimetva,
    TVA,
    TPF, // code TPF
    Contrepartigen, contrepartiaux, RefPointage, datepointage, daterelance,
    datevaleur,
    Rib, Refreleve, Numeroimmo,
    LT0, LT1, LT2, LT3, LT4, LT5, LT6, LT7, LT8, LT9,
    TA0, TA1, TA2, TA3,
    LM0, LM1, LM2, LM3,
    LD, LB0, LB1,
    Conso, // consolidation
    Couverture, // couverture
    Couverturedev, couverturelibre,
    Datepaquetmax, Datepaquetmin,
    lettrage, lettragedev, lettragelibre,
    etatlettrage, refgescom, typemvt, docid, treso, numtrait,numenca,
    souspl1, souspl2, souspl3, souspl4, souspl5, valide,
    ctdeb, ctfin, ctdate, cleecr, confidentiel, cfonbok,
    // BAP
    viseur, statutbap, viseur1, viseur2, numeroordre, Nbjour, tierspayeur, codevisa,
    circuitbap, daterelance2, DateEcheance, alerte1, alerte2,relance1, relance2, codeaccept : string;
    // BVE 28.08.07 Ajout E_DOCID : Suivi des validations
    idvalidation,
    qualiforigine : string;
    // fin etendu
  end;
  PListeMouvt = ^TMouvement;

var
   DateFichierClo : string;


implementation

{$IFNDEF EAGLCLIENT}

Function findcorrespMDP (mp : string; var Corresp : string) : string;
var
MDP     : string;
Q1      : TQuery;

begin
  MDP := mp;
  Result := MDP;
  if Corresp = 'FALSE' then exit;
  Q1 := OpenSql('Select CR_CORRESP,CR_LIBELLE,CR_ABREGE from CORRESP where CR_TYPE="MDP" and CR_CORRESP="'
  + mp +'"', TRUE);
  if not Q1.EOF then
  MDP := Q1.Findfield('CR_LIBELLE').asstring;
  Ferme (Q1);
  Result := MDP;
end;

Function findcorrespMDR (mr : string; stArg : string) : string;
var
  MDR     : string;
  Q1      : Tquery;
begin
      MDR := mr;
      Result := MDR;
     //pour agnes if stArg = '' then exit;
      Q1 := OpenSql('Select CR_CORRESP,CR_LIBELLE,CR_ABREGE from CORRESP where CR_TYPE="MDR" and CR_CORRESP="'
      + mr +'"', TRUE);
      if not Q1.EOF then
      MDR := Q1.Findfield('CR_LIBELLE').asstring;
      Ferme (Q1);
      Result := MDR;
end;


procedure EcritureJournal(ListeJalB: TList; var F: TextFile; Libere : Boolean=TRUE);
var
  i             : integer;
  TJournal      : PListeJournal;
  Ligne         : string;
  ReserveExpert : string;
  JExp          : string;
begin
  JExp := GetParamSocSecur ('SO_CPJOURNALEXPERT', '', TRUE);
  if (JExp <> '') and (JExp[length (JExp)] <> ';') then // fiche 10513
  JExp := JExp +';';

  for i := 0 to ListeJalB.Count - 1 do
  begin
    TJournal := ListeJalB.Items[i];
      // Ajout me pour 7.10 journal réservé à l'expert comptable
    if (pos(TJournal^.code+';', JExp) <> 0) then
       ReserveExpert := 'X'
    else
        ReserveExpert := '-';

    Ligne := Format(FormatJournal, [TJournal^.code, TJournal^.libelle,
      TJournal^.nature,
      TJournal^.souchen, TJournal^.souches, TJournal^.compte,
        TJournal^.axe, TJournal^.modesaisie, TJournal^.cptauto,
        TJournal^.cptint, TJournal^.abrege, ReserveExpert]);
    writeln(F, Ligne);
  end;
  if Libere then
  begin
      for i:=0 To  ListeJalB.Count-1 do
      begin
           TJournal := ListeJalB.Items[i];
           if (TJournal <> nil ) then Dispose (TJournal);
      end;
  end;
end;

procedure LiberJournal(ListeJalB: TList);
var
   i : integer;
   TJournal: PListeJournal;
begin
  for i:=0 To  ListeJalB.Count-1 do
  begin
       TJournal := ListeJalB.Items[i];
       if (TJournal <> nil ) then Dispose (TJournal);
  end;
end;


procedure EcritureCpptegen(ListeCpteGen: TList; var F: TextFile);
var
  i: integer;
  Ligne: string;
  TCGene: PListeCpteGen;
begin
  for i := 0 to ListeCpteGen.Count - 1 do
  begin
    TCGene := ListeCpteGen.Items[i];
    Ligne := Format(FormatComptegene, [TCGene^.Code, TCGene^.Libelle,
      TCGene^.Nature, TCGene^.Lettrable,
      TCGene^.Pointage, TCGene^.Ventilaxe1, TCGene^.Ventilaxe2,
      TCGene^.Ventilaxe3, TCGene^.Ventilaxe4, TCGene^.Ventilaxe5,
      TCGene^.Table1, TCGene^.Table2, TCGene^.Table3, TCGene^.Table4,
      TCGene^.Table5, TCGene^.Table6, TCGene^.Table7, TCGene^.Table8,
      TCGene^.Table9, TCGene^.Table10, TCGene^.abrege, TCGene^.sens,
      TCGene^.Corresp1,TCGene^.Corresp2, TCGene^.Tva, TCGene^.Encaissement, TCGene^.TPF,
      TCGene^.ctoff, TCGene^.ctoffper, TCGene^.cteche, TCGene^.visarev, TCGene^.cyclerev,
          //Sauvegarde de qualifiant qte  Reprendre dans 6.5
      TCGene^.ctcpte, TCGene^.condifentiel, TCGene^.qualqte1, TCGene^.qualqte2]);

    writeln(F, Ligne);
  end;
  for i:=0 To  ListeCpteGen.Count-1 do
  begin
       TCGene := ListeCpteGen.Items[i];
       if (TCGene <> nil) then Dispose (TCGene);
  end;
end;


procedure EcritureTierscomp(var F: TextFile; Where : string);
var
QT        : Tquery;
Ligne     : string;
begin
  QT := OpenSQl ('SELECT * from TIERSCOMPL  '+ Where , TRUE);
  while not QT.EOF do
  begin
        Ligne := Format(FormatTierscomp,
        [QT.FindField('YTC_AUXILIAIRE').asstring, QT.FindField('YTC_TIERS').asstring,  '',
        QT.FindField('YTC_TABLELIBRETIERS1').asstring, QT.FindField('YTC_TABLELIBRETIERS2').asstring,
        QT.FindField('YTC_TABLELIBRETIERS3').asstring, QT.FindField('YTC_TABLELIBRETIERS4').asstring,
        QT.FindField('YTC_TABLELIBRETIERS5').asstring, QT.FindField('YTC_TABLELIBRETIERS6').asstring,
        QT.FindField('YTC_TABLELIBRETIERS7').asstring, QT.FindField('YTC_TABLELIBRETIERS8').asstring,
        QT.FindField('YTC_TABLELIBRETIERS9').asstring, QT.FindField('YTC_TABLELIBRETIERSA').asstring,
        FloatToStr(QT.FindField('YTC_VALLIBRE1').AsFloat), FloatToStr(QT.FindField('YTC_VALLIBRE2').AsFloat),
        FloatToStr(QT.FindField('YTC_VALLIBRE3').AsFloat),
        FormatDateTime(Traduitdateformat('ddmmyyyy'),QT.FindField ('YTC_DATELIBRE1').AsDateTime), FormatDateTime(Traduitdateformat('ddmmyyyy'),QT.FindField ('YTC_DATELIBRE2').AsDateTime),
        FormatDateTime(Traduitdateformat('ddmmyyyy'),QT.FindField ('YTC_DATELIBRE3').AsDateTime),
        QT.FindField('YTC_TABLELIBREFOU1').asstring, QT.FindField('YTC_TABLELIBREFOU2').asstring,
        QT.FindField('YTC_TABLELIBREFOU3').asstring,
        FloatToStr(QT.FindField('YTC_VALLIBREFOU1').AsFloat), FloatToStr(QT.FindField('YTC_VALLIBREFOU2').AsFloat),
        FloatToStr(QT.FindField('YTC_VALLIBREFOU3').AsFloat),
        FormatDateTime(Traduitdateformat('ddmmyyyy'),QT.FindField ('YTC_DATELIBREFOU1').AsDateTime), FormatDateTime(Traduitdateformat('ddmmyyyy'),QT.FindField ('YTC_DATELIBREFOU2').AsDateTime),
        FormatDateTime(Traduitdateformat('ddmmyyyy'),QT.FindField ('YTC_DATELIBREFOU3').AsDateTime),
        QT.FindField('YTC_TEXTELIBRE1').asstring, QT.FindField('YTC_TEXTELIBRE2').asstring,
        QT.FindField('YTC_TEXTELIBRE3').asstring,
        QT.FindField('YTC_BOOLLIBRE1').asstring, QT.FindField('YTC_BOOLLIBRE2').asstring,
        QT.FindField('YTC_BOOLLIBRE3').asstring,
        QT.FindField('YTC_RESSOURCE1').asstring, QT.FindField('YTC_RESSOURCE2').asstring,
        QT.FindField('YTC_RESSOURCE3').asstring, QT.FindField('YTC_DOCIDENTITE').asstring,
        FormatDateTime(Traduitdateformat('ddmmyyyy'),QT.FindField ('YTC_DOCDATEDELIV').AsDateTime), FormatDateTime(Traduitdateformat('ddmmyyyy'),QT.FindField ('YTC_DOCDATEEXPIR').AsDateTime),
        QT.FindField('YTC_DOCOBSERV').asstring, QT.FindField('YTC_DOCORIGINE').asstring, QT.FindField('YTC_FAMREG').asstring,
        QT.FindField('YTC_COMMSPECIAL').asstring, QT.FindField('YTC_REPRESENTANT2').asstring, QT.FindField('YTC_REPRESENTANT3').asstring,
        FloatToStr(QT.FindField('YTC_TAUXREPR1').AsFloat), FloatToStr(QT.FindField('YTC_TAUXREPR2').AsFloat),
        FloatToStr(QT.FindField('YTC_TAUXREPR3').AsFloat),
        IntToStr(QT.FindField('YTC_NADRESSELIV').asinteger), IntToStr(QT.FindField('YTC_NADRESSEFAC').Asinteger),
        QT.FindField('YTC_STATIONEDI').asstring, QT.FindField('YTC_NOTRECODETIERS').asstring,
        QT.FindField('YTC_NOTRECODCOMPTA').asstring, QT.FindField('YTC_SCHEMAGEN').asstring,
        QT.FindField('YTC_ACCELERATEUR').asstring, QT.FindField('YTC_DAS2').asstring,
        QT.FindField('YTC_PROFESSION').asstring, QT.FindField('YTC_REMUNERATION').asstring,
        QT.FindField('YTC_INDEMNITE').asstring, QT.FindField('YTC_AVANTAGE').asstring]);
        writeln(F, Ligne);
        QT.next;
   end;
  Ferme(QT);
end;

procedure EcritureCppteaux(ListeCpteaux: TList; var F: TextFile; stArg : string='');
var
  i: integer;
  Ligne: string;
  TCaux: PListeAux;
  MDR  : string;
begin
  for i := 0 to ListeCpteaux.Count - 1 do
  begin
    TCaux := ListeCpteaux.Items[i];
    // ajout me si code à blanc Cela provient du fait que les comptes aient plusieurs rib associés
    if TCaux^.Code <> '' then
    begin
     MDR := findcorrespMDR (TCaux^.modereg, stArg);
        Ligne := Format(FormatCompteaux, [TCaux^.Code, TCaux^.libelle,
          TCaux^.nature, TCaux^.lettrable, TCaux^.collectif, TCaux^.ean, TCaux^.tb1,
          TCaux^.tb2, TCaux^.tb3, TCaux^.tb4, TCaux^.tb5, TCaux^.tb6, TCaux^.tb7, TCaux^.tb8,
          TCaux^.tb9, TCaux^.tb10, TCaux^.adresse1, TCaux^.adresse2, TCaux^.adresse3, TCaux^.codepostal,
          TCaux^.ville, TCaux^.domicile, TCaux^.etablissement, TCaux^.guichet, TCaux^.compte,
          TCaux^.cle, TCaux^.pays, TCaux^.abrege, TCaux^.langue, TCaux^.multidevise,
          TCaux^.devisetiers, TCaux^.tel, TCaux^.fax, TCaux^.regimetva, MDR, TCaux^.commentaire,
          TCaux^.nif, TCaux^.siret, TCaux^.ape, TCaux^.ctnom, TCaux^.ctservice,
          TCaux^.ctfonction, TCaux^.cttel, TCaux^.ctfax, TCaux^.cttelex, TCaux^.ctrva,
          TCaux^.ctcivilite, TCaux^.ctprincipal, TCaux^.formjur, TCaux^.rib, TCaux^.tvaenc,
                        // ajout me 18-08-2005
          TCaux^.tierspayeur, TCaux^.ispayeur, TCaux^.remboursement,
                       // ajout me 26/08/2005
          TCaux^.relanceregl, TCaux^.relancetrait, TCaux^.confidentiel,
          // fiche 10505
          TCaux^.corresp1, TCaux^.corresp2, TCaux^.escompte, TCaux^.remise, TCaux^.facture,
          TCaux^.fromjur, TCaux^.ceditdemande, TCaux^.ceditaccorde, TCaux^.ceditplafond ]);  // fiche com 31027
        writeln(F, Ligne);
    end;
  end;
  for i:=0 To  ListeCpteaux.Count-1 do
  begin
       TCaux := ListeCpteaux.Items[i];
       if TCaux <> nil then Dispose (TCaux);
  end;

end;

procedure EcritureRib(ListeRib: TList; var F: TextFile);
var
  i       : integer;
  Ligne   : string;
  TRib    : PListeRib;
begin
  for i := 0 to ListeRib.Count - 1 do
  begin
    TRib := ListeRib.Items[i];
    // ajout me si code à blanc Cela provient du fait que les comptes aient plusieurs rib associés
    if TRib^.Code <> '' then
    begin
        Ligne := Format(FormatRib, [TRib^.ident, TRib^.code,
        TRib^.numero,TRib^.prinpipal,TRib^.codebanque,TRib^.codeguichet,TRib^.nocompte,
        TRib^.cle,TRib^.domicile,TRib^.ville,TRib^.pays,TRib^.devise,TRib^.bic,
          // ajout me nature éco et nouveau format des comptes
        TRib^.soc,TRib^.ribsal,TRib^.ribcompte,TRib^.ribprof, TRib^.iban, TRib^.nateco, TRib^.typpays,
        TRib^.codebanque, TRib^.nocompte ]);
        writeln(F, Ligne);
    end;
  end;
  for i:=0 To  ListeRib.Count-1 do
  begin
           TRib := ListeRib.Items[i];
           if TRib <> nil then Dispose (TRib);
  end;

end;


procedure EcritureBanquecp(var F: TextFile);
var Qr        : Tquery;
    Ligne     : string;
    Where     : string;
begin
 // AJOUT SUR 7XX   AVOIR
if EstMultiSoc then
 Where := ' Where BQ_NODOSSIER="' + V_PGI.NoDossier +'" OR BQ_NODOSSIER="000000" ' +
 'or (BQ_NODOSSIER is null)';

Qr := OpenSql ('SELECT * FROM BANQUECP' + Where, TRUE);
    while not Qr.EOF do
    begin
        Ligne := Format(FormatBqcp,
        [Qr.FindField('BQ_CODE').asstring, Qr.FindField('BQ_GENERAL').asstring,
        Qr.FindField('BQ_LIBELLE').asstring,
        Qr.FindField('BQ_DOMICILIATION').asstring,
        Qr.FindField('BQ_ADRESSE1').asstring,Qr.FindField('BQ_ADRESSE2').asstring,
        Qr.FindField('BQ_ADRESSE3').asstring,Qr.FindField('BQ_CODEPOSTAL').asstring,
        Qr.FindField('BQ_VILLE').asstring,
        Qr.FindField('BQ_DIVTERRIT').asstring,Qr.FindField('BQ_PAYS').asstring,
        Qr.FindField('BQ_LANGUE').asstring,Qr.FindField('BQ_DEVISE').asstring,
        Qr.FindField('BQ_TELEPHONE').asstring,Qr.FindField('BQ_FAX').asstring,
        Qr.FindField('BQ_CONTACT').asstring,Qr.FindField('BQ_ETABBQ').asstring,
        Qr.FindField('BQ_NUMEROCOMPTE').asstring,Qr.FindField('BQ_CLERIB').asstring,
        Qr.FindField('BQ_GUICHET').asstring,Qr.FindField('BQ_CODEBIC').asstring,
        Qr.FindField('BQ_NUMEMETLCR').asstring,Qr.FindField('BQ_CONVENTIONLCR').asstring,
        Qr.FindField('BQ_NUMEMETVIR').asstring,Qr.FindField('BQ_JOURFERMETUE').asstring,
        Qr.FindField('BQ_REPRELEVE').asstring,Qr.FindField('BQ_REPLCR').asstring,
        Qr.FindField('BQ_REPLCRFOURN').asstring,Qr.FindField('BQ_REPVIR').asstring,
        Qr.FindField('BQ_REPPRELEV').asstring,Qr.FindField('BQ_REPBONAPAYER').asstring,
        Qr.FindField('BQ_REPIMPAYELCR').asstring,Qr.FindField('BQ_DELAIVIRORD').asinteger,
        Qr.FindField('BQ_DELAIVIRCHAUD').asinteger,Qr.FindField('BQ_DELAIVIRBRULANT').asinteger,
        Qr.FindField('BQ_DELAIPRELVORD').asinteger,Qr.FindField('BQ_DELAIPRELVACC').asinteger,
        Qr.FindField('BQ_DELAILCR').asinteger,Qr.FindField('BQ_GUIDECOMPATBLE').asstring,
        Qr.FindField('BQ_DERNSOLDEFRS').asinteger,Qr.FindField('BQ_DERNSOLDEDEV').asinteger,
        Qr.FindField('BQ_DATEDERNSOLDE').asstring,Qr.FindField('BQ_RELEVEETRANGER').asstring,
        Qr.FindField('BQ_CALENDRIER').asstring,Qr.FindField('BQ_RAPPAUTOREL').asstring,
        Qr.FindField('BQ_RAPPROAUTOLCR').asstring,Qr.FindField('BQ_LETTREVIR').asstring,
        Qr.FindField('BQ_LETTREPRELV').asstring,Qr.FindField('BQ_LETTRELCR').asstring,
        Qr.FindField('BQ_ENCOURSLCR').asinteger,Qr.FindField('BQ_PLAFONDLCR').asinteger,
        Qr.FindField('BQ_REPIMPAYEPRELV').asstring,Qr.FindField('BQ_ECHEREPPRELEV').asstring,
        Qr.FindField('BQ_ECHEREPLCR').asstring,Qr.FindField('BQ_SOCIETE').asstring,
        Qr.FindField('BQ_BANQUE').asstring,Qr.FindField('BQ_DELAIBAPLCR').asstring,
        Qr.FindField('BQ_DELAITRANSINT').asstring,Qr.FindField('BQ_COMPTEFRAIS').asstring,
        Qr.FindField('BQ_TYPEREMTRANS').asstring,Qr.FindField('BQ_INDREMTRANS').asstring,
        Qr.FindField('BQ_COMMENTAIRE').asstring,Qr.FindField('BQ_LETTRECHQ').asstring,
        Qr.FindField('BQ_NUMEMETPRE').asstring,Qr.FindField('BQ_DESTINATAIRE').asstring,
        Qr.FindField('BQ_MULTIDEVISE').asstring,Qr.FindField('BQ_CODEIBAN').asstring,
        Qr.FindField('BQ_NODOSSIER').asstring, Qr.FindField('BQ_CODECIB').asstring,
        Qr.FindField('BQ_NATURECPTE').asstring]);
        writeln(F, Ligne);
       Qr.next;
    end;
ferme (Qr);
end;

procedure EcritureBanques(var F: TextFile);
var Qr              : Tquery;
    Ligne           : string;
    oketab          : Boolean;
    etab, abreg     : string;
begin

Qr := OpenSql ('SELECT * FROM BANQUES', TRUE);
    oketab := FALSE;
    if ChampphysiqueExiste ( 'BANQUES', 'PQ_ETABBQ') then  oketab := TRUE;

    while not Qr.EOF do
    begin
        if oketab then
        begin
          etab := Qr.FindField('PQ_ETABBQ').asstring;
         abreg := Qr.FindField('PQ_ABREGE').asstring;
        end;
        Ligne := Format(FormatBanque,
        [Qr.FindField('PQ_BANQUE').asstring, Qr.FindField('PQ_LIBELLE').asstring,
        Qr.FindField('PQ_RESTCB').asstring,  Qr.FindField('PQ_REMCB').asstring,
        Qr.FindField('PQ_RESTCHQ').asstring, Qr.FindField('PQ_REMCHQ').asstring,
        Qr.FindField('PQ_RESTESP').asstring, Qr.FindField('PQ_REMESP').asstring,
        Qr.FindField('PQ_RESTLCR').asstring, Qr.FindField('PQ_REMLCR').asstring,
        Qr.FindField('PQ_RESTPRE').asstring, Qr.FindField('PQ_REMPRE').asstring,
        Qr.FindField('PQ_RESTTRI').asstring, Qr.FindField('PQ_REMTRI').asstring,
        Qr.FindField('PQ_RESTVIR').asstring, Qr.FindField('PQ_REMVIR').asstring,
        Qr.FindField('PQ_RESTVIT').asstring, Qr.FindField('PQ_REMVIT').asstring,
        Qr.FindField('PQ_RESTTIP').asstring, Qr.FindField('PQ_REMTIP').asstring,
        Qr.FindField('PQ_RESTTEP').asstring, Qr.FindField('PQ_REMTEP').asstring,
        Qr.FindField('PQ_RESTVIC').asstring, Qr.FindField('PQ_REMVIC').asstring,
        Qr.FindField('PQ_DE_MODE').asstring, Qr.FindField('PQ_DE_TXREF').asstring,

        FloatToStr(Qr.FindField('PQ_DE_CORRECTION').asfloat), FloatToStr(Qr.FindField('PQ_DE_SAISIE').asfloat),
        Qr.FindField('PQ_DE_NUMERATEUR').asinteger, Qr.FindField('PQ_DE_DENOMINATEUR').asinteger,
        FloatToStr(Qr.FindField('PQ_DE_TAUX').asfloat), Qr.FindField('PQ_DE_DATETAUX').asstring,
        FloatToStr(Qr.FindField('PQ_DE_PLAFOND').asfloat), Qr.FindField('PQ_PD_MODE').asstring,
        Qr.FindField('PQ_PD_TXREF').asstring,

        FloatToStr(Qr.FindField('PQ_PD_CORRECTION').asfloat),FloatToStr(Qr.FindField('PQ_PD_SAISIE').asfloat),
        Qr.FindField('PQ_PD_NUMERATEUR').asinteger,Qr.FindField('PQ_PD_DENOMINATEUR').asinteger,
        FloatToStr(Qr.FindField('PQ_PD_TAUX').asfloat),Qr.FindField('PQ_PD_DATETAUX').asstring,
        FloatToStr(Qr.FindField('PQ_PD_PLAFOND').asfloat),Qr.FindField('PQ_CR_MODE').asstring,
        Qr.FindField('PQ_CR_TXREF').asstring,

        FloatToStr(Qr.FindField('PQ_CR_CORRECTION').asfloat),FloatToStr(Qr.FindField('PQ_CR_SAISIE').asfloat),
        Qr.FindField('PQ_CR_NUMERATEUR').asinteger,Qr.FindField('PQ_CR_DENOMINATEUR').asinteger,
        FloatToStr(Qr.FindField('PQ_CR_TAUX').asfloat),Qr.FindField('PQ_CR_DATETAUX').asstring,
        FloatToStr(Qr.FindField('PQ_CR_PLAFOND').asfloat),Qr.FindField('PQ_CR_TYPEPLAFOND').asstring,
        Qr.FindField('PQ_CO_MODE').asstring, Qr.FindField('PQ_CO_TXREF').asstring,

        FloatToStr(Qr.FindField('PQ_CO_CORRECTION').asfloat),FloatToStr(Qr.FindField('PQ_CO_SAISIE').asfloat),
        Qr.FindField('PQ_CO_NUMERATEUR').asinteger,Qr.FindField('PQ_CO_DENOMINATEUR').asinteger,
        FloatToStr(Qr.FindField('PQ_CO_TAUX').asfloat),Qr.FindField('PQ_CO_DATETAUX').asstring,
        FloatToStr(Qr.FindField('PQ_CO_FRAIS').asfloat),FloatToStr(Qr.FindField('PQ_CO_TVA').asfloat),
        FloatToStr(Qr.FindField('PQ_BB_TAUX').asfloat), Qr.FindField('PQ_BB_DATETAUX').asstring,
        etab, abreg]);
        writeln(F, Ligne);
       Qr.next;
    end;
ferme (Qr);
end;

procedure EcritureReleveBanque(var F: TextFile);
var Qr        : Tquery;
    Ligne     : string;
begin
Qr := OpenSql ('SELECT * FROM EEXBQ', TRUE);
    while not Qr.EOF do
    begin
        Ligne := Format(FormatRlvBqe,
        [QR.FindField ('EE_GENERAL').asstring,
        Qr.FindField('EE_REFPOINTAGE').asstring,
        FormatDateTime(Traduitdateformat('ddmmyyyy'),Qr.FindField('EE_DATEOLDSOLDE').AsDateTime),
        FormatDateTime(Traduitdateformat('ddmmyyyy'),Qr.FindField('EE_DATEPOINTAGE').AsDateTime),
        FloatToStr(Qr.FindField('EE_OLDSOLDECRE').asfloat),
        FloatToStr(Qr.FindField('EE_NEWSOLDECRE').asfloat),
        FloatToStr(Qr.FindField('EE_OLDSOLDEDEB').asfloat),
        FloatToStr(Qr.FindField('EE_NEWSOLDEDEB').asfloat),
        FloatToStr(Qr.FindField('EE_OLDSOLDEDEBEURO').asfloat),
        FloatToStr(Qr.FindField('EE_OLDSOLDECREEURO').asfloat),
        FloatToStr(Qr.FindField('EE_NEWSOLDEDEBEURO').asfloat),
        FloatToStr(Qr.FindField('EE_NEWSOLDECREEURO').asfloat),
        Qr.FindField('EE_SOCIETE').asstring,
        Qr.FindField('EE_NUMRELEVE').asinteger,
        Qr.FindField('EE_RIB').asstring,
        Qr.FindField('EE_DEVISE').asstring,
        FormatDateTime(Traduitdateformat('ddmmyyyy'),Qr.FindField('EE_DATESOLDE').AsDateTime),
        FloatToStr(Qr.FindField('EE_OLDSOLDEDEBDEV').asfloat),
        FloatToStr(Qr.FindField('EE_OLDSOLDECREDEV').asfloat),
        FloatToStr(Qr.FindField('EE_NEWSOLDEDEBDEV').asfloat),
        FloatToStr(Qr.FindField('EE_NEWSOLDECREDEV').asfloat),
        Qr.FindField('EE_NBMVT').asinteger,
        Qr.FindField('EE_ORIGINERELEVE').asstring,
        Qr.FindField('EE_VALIDE').asstring,
        FormatDateTime(Traduitdateformat('ddmmyyyy'),Qr.FindField('EE_DATEINTEGRE').AsDateTime),
        Qr.FindField('EE_STATUTRELEVE').asstring,
        Qr.FindField('EE_NUMERO').asinteger,
        Qr.FindField('EE_AVANCEMENT').asstring]);
        writeln(F, Ligne);
       Qr.next;
    end;
ferme (Qr);
end;

procedure EcritureLigneReleveBanque(var F: TextFile);
var Qr        : Tquery;
    Ligne     : string;
begin
Qr := OpenSql ('SELECT * FROM EEXBQLIG', TRUE);
    while not Qr.EOF do
    begin
        Ligne := Format(FormatRlvLig,
        [QR.FindField ('CEL_GENERAL').asstring,
        Qr.FindField('CEL_NUMRELEVE').asinteger,
        Qr.FindField('CEL_NUMLIGNE').asinteger,
        QR.FindField ('CEL_LIBELLE').asstring,
        QR.FindField ('CEL_RIB').asstring,
        QR.FindField ('CEL_VALIDE').asstring,
        QR.FindField ('CEL_CODEAFB').asstring,
        FormatDateTime(Traduitdateformat('ddmmyyyy'),QR.FindField ('CEL_DATEOPERATION').AsDateTime),
        FormatDateTime(Traduitdateformat('ddmmyyyy'),QR.FindField ('CEL_DATEVALEUR').AsDateTime),
        FloatToStr(Qr.FindField('CEL_DEBITDEV').asfloat),
        FloatToStr(Qr.FindField('CEL_CREDITDEV').asfloat),
        FloatToStr(Qr.FindField('CEL_DEBITEURO').asfloat),
        FloatToStr(Qr.FindField('CEL_CREDITEURO').asfloat),
        QR.FindField ('CEL_REFPIECE').asstring,
        QR.FindField ('CEL_REFORIGINE').asstring,
        QR.FindField ('CEL_REFPOINTAGE').asstring,
        FormatDateTime(Traduitdateformat('ddmmyyyy'),QR.FindField ('CEL_DATEPOINTAGE').AsDateTime),
        QR.FindField ('CEL_IMO').asstring,
        QR.FindField ('CEL_CODERAPPRO').asinteger,
        QR.FindField ('CEL_EXONERE').asinteger,
        QR.FindField ('CEL_DISPONIBLE').asinteger,
        QR.FindField ('CEL_LIBELLE1').asstring,
        QR.FindField ('CEL_LIBELLE2').asstring,
        QR.FindField ('CEL_LIBELLE3').asstring,
        QR.FindField ('CEL_NATUREINTERNE').asstring]);
        writeln(F, Ligne);
       Qr.next;
    end;
    ferme (Qr);
end;

procedure EcritureContact(var F: TextFile; WW : string='');
var Qr        : Tquery;
    Ligne     : string;
begin

Qr := OpenSql ('SELECT * FROM CONTACT' + WW, TRUE);

    while not Qr.EOF do
    begin
        Ligne := Format(FormatContact,
        [Qr.FindField('C_TYPECONTACT').asstring, Qr.FindField('C_AUXILIAIRE').asstring,
        Qr.FindField('C_NUMEROCONTACT').asinteger,
        Qr.FindField('C_NATUREAUXI').asstring,
        Qr.FindField('C_PRINCIPAL').asstring,Qr.FindField('C_NOM').asstring,
        Qr.FindField('C_PRENOM').asstring,Qr.FindField('C_SERVICE').asstring,
        Qr.FindField('C_FONCTION').asstring,
        Qr.FindField('C_TELEPHONE').asstring,Qr.FindField('C_FAX').asstring,
        Qr.FindField('C_TELEX').asstring,Qr.FindField('C_RVA').asstring,
        Qr.FindField('C_SOCIETE').asstring,Qr.FindField('C_CIVILITE').asstring,
        Qr.FindField('C_FONCTIONCODEE').asstring,Qr.FindField('C_LIPARENT').asstring,
        Qr.FindField('C_JOURNAIS').asinteger,Qr.FindField('C_MOISNAIS').asinteger,
        Qr.FindField('C_ANNEENAIS').asinteger,Qr.FindField('C_SEXE').asstring,
        Qr.FindField('C_PUBLIPOSTAGE').asstring,Qr.FindField('C_TEXTELIBRE1').asstring,
        Qr.FindField('C_TEXTELIBRE2').asstring,Qr.FindField('C_TEXTELIBRE3').asstring,
        Qr.FindField('C_BOOLLIBRE1').asstring,Qr.FindField('C_BOOLLIBRE2').asstring,
        Qr.FindField('C_BOOLLIBRE3').asstring,
        Qr.FindField('C_LIBRECONTACT1').asstring,Qr.FindField('C_LIBRECONTACT2').asstring,
        Qr.FindField('C_LIBRECONTACT3').asstring,
        FormatDateTime(Traduitdateformat('ddmmyyyy'), Qr.Findfield('C_DATELIBRE1').AsDateTime),
        FormatDateTime(Traduitdateformat('ddmmyyyy'), Qr.Findfield('C_DATELIBRE2').AsDateTime),
        FormatDateTime(Traduitdateformat('ddmmyyyy'), Qr.Findfield('C_DATELIBRE3').AsDateTime),
        Qr.FindField('C_VALLIBRE1').asinteger,Qr.FindField('C_VALLIBRE2').asinteger,
        Qr.FindField('C_VALLIBRE3').asinteger,
        Qr.FindField('C_TIERS').asstring, Qr.FindField('C_CLETELEPHONE').asstring]);
        writeln(F, Ligne);
       Qr.next;
    end;
ferme (Qr);
end;


procedure EcritureMouvement(ListeMouvt: TList; var F: TextFile; Formatm : string; var Corresp : string);
var
  i       : integer;
  Ligne   : string;
  Pmvt    : PListeMouvt;
  MDP     : string;
begin
  for i := 0 to ListeMouvt.Count - 1 do
  begin
    Pmvt := ListeMouvt.Items[i];
    MDP := findcorrespMDP (Pmvt^.Modepaie,  Corresp);
    if (Formatm = 'STD') or (Pmvt^.Typecpte ='G') or (Pmvt^.Typecpte ='P')then  // G=GED P=BAP
    begin
     case Pmvt^.Typecpte[1] of
           'G' :  // Ged
            Ligne := Format(FormatGed, [Pmvt^.Journal, Pmvt^.Datecomptable,
              Pmvt^.Typepiece, Pmvt^.General, Pmvt^.Typecpte, Pmvt^.AuxSect, Pmvt^.Refinterne, Pmvt^.docid]);
            'P' : // BAP
              Ligne := Format(FormatBap,
              [Pmvt^.Journal,  Pmvt^.Datecomptable,
              Pmvt^.Typepiece, Pmvt^.General, Pmvt^.Typecpte,Pmvt^.AuxSect,
              Pmvt^.NumPiece , Pmvt^.viseur, Pmvt^.statutbap,Pmvt^.viseur1,
              Pmvt^.viseur2, Pmvt^.numeroordre, Pmvt^.Nbjour, Pmvt^.Echeance, Pmvt^.tierspayeur, Pmvt^.cleecr, Pmvt^.codevisa,
              Pmvt^.circuitbap,  Pmvt^.ctdate, Pmvt^.daterelance, Pmvt^.daterelance2, Pmvt^.DateEcheance,Pmvt^.societe,
              Pmvt^.Etablissement,Pmvt^.alerte1, Pmvt^.alerte2,Pmvt^.relance1, Pmvt^.relance2]);
      else begin
            Ligne := Format(Formatmvt, [Pmvt^.Journal, Pmvt^.Datecomptable,
              Pmvt^.Typepiece, Pmvt^.General, Pmvt^.Typecpte, Pmvt^.AuxSect, Pmvt^.Refinterne, Pmvt^.Libelle,
              MDP, Pmvt^.Echeance, Pmvt^.Sens, Pmvt^.Montant1, Pmvt^.Typeecriture,
              Pmvt^.NumPiece, Pmvt^.Devise, Pmvt^.TauxDev, Pmvt^.CodeMontant,
              Pmvt^.Montant2, Pmvt^.Montant3, Pmvt^.Etablissement, Pmvt^.Axe,
              Pmvt^.Numeche]);
            end;
      end;
    end
    else
    begin
         if (Formatm = 'SAGE') then
           Ligne :=Format_String(Pmvt^.Journal,3)+Format_String(Pmvt^.Datecomptable,6)+Format_String(Pmvt^.Typepiece,2)
           +Format_String(Pmvt^.General,13)+Format_String(Pmvt^.Typecpte,1)+Format_String(Pmvt^.AuxSect,13)
           +Format_String(Pmvt^.Refinterne,13)+Format_String(Pmvt^.Libelle,25)+Format_String(MDP,1)
           +Format_String(Pmvt^.Echeance,6)+Format_String(Pmvt^.Sens,1)+AlignDroite(Pmvt^.Montant1,20)
           +Format_String(Pmvt^.Typeecriture,1)+AlignDroite(Pmvt^.NumPiece,7)
         else
         begin

            if Pmvt^.Typecpte ='A' then  // analytique

             Ligne := Format(Formatmvtetenduana, [Pmvt^.Journal, Pmvt^.Datecomptable,
                Pmvt^.Typepiece, Pmvt^.General, Pmvt^.Typecpte, Pmvt^.AuxSect,
                Pmvt^.Refinterne, Pmvt^.Libelle, MDP, Pmvt^.Echeance,
                Pmvt^.Sens, Pmvt^.Montant1, Pmvt^.Typeecriture,
                Pmvt^.NumPiece, Pmvt^.Devise, Pmvt^.TauxDev, Pmvt^.CodeMontant,
                Pmvt^.Montant2, Pmvt^.Montant3, Pmvt^.Etablissement, Pmvt^.Axe,
                Pmvt^.Numeche, Pmvt^.RefExterne, Pmvt^.Daterefexterne,
                Pmvt^.Datecreation, Pmvt^.Societe, Pmvt^.Affaire, Pmvt^.Datetauxdev,
                Pmvt^.Ecranouveau, Pmvt^.Qte1, Pmvt^.Qte2, Pmvt^.QualifQte1,
                Pmvt^.QualifQte2, Pmvt^.Reflibre, Pmvt^.Tvaencaiss, Pmvt^.Regimetva,
                Pmvt^.TVA, Pmvt^.TPF, Pmvt^.Contrepartigen, Pmvt^.contrepartiaux,
                Pmvt^.RefPointage, Pmvt^.souspl1, Pmvt^.souspl2, Pmvt^.souspl3,
                Pmvt^.souspl4, Pmvt^.souspl5,
                (*Pmvt^.Rib, Pmvt^.Refreleve,*) Pmvt^.Numeroimmo,
                Pmvt^.LT0, Pmvt^.LT1, Pmvt^.LT2, Pmvt^.LT3, Pmvt^.LT4, Pmvt^.LT5,
                Pmvt^.LT6, Pmvt^.LT7, Pmvt^.LT8, Pmvt^.LT9,
                Pmvt^.TA0, Pmvt^.TA1, Pmvt^.TA2, Pmvt^.TA3, Pmvt^.LM0, Pmvt^.LM1,
                Pmvt^.LM2, Pmvt^.LM3, Pmvt^.LD, Pmvt^.LB0, Pmvt^.LB1, Pmvt^.Conso, Pmvt^.Couverture,
                Pmvt^.Couverturedev, Pmvt^.couverturelibre,
                Pmvt^.Datepaquetmax, Pmvt^.Datepaquetmin, Pmvt^.lettrage,
                // ajout me les tables libres
                Pmvt^.lettragedev, Pmvt^.lettragelibre, Pmvt^.etatlettrage,Pmvt^.TA0, Pmvt^.TA1, Pmvt^.TA2,
                Pmvt^.TA3, Pmvt^.refgescom, Pmvt^.typemvt, Pmvt^.docid, Pmvt^.treso, Pmvt^.numtrait, Pmvt^.numenca, Pmvt^.valide, Pmvt^.confidentiel])
             else
             begin
                   Ligne := Format(Formatmvtetendu, [Pmvt^.Journal, Pmvt^.Datecomptable,
                      Pmvt^.Typepiece, Pmvt^.General, Pmvt^.Typecpte, Pmvt^.AuxSect,
                      Pmvt^.Refinterne, Pmvt^.Libelle, MDP, Pmvt^.Echeance,
                      Pmvt^.Sens, Pmvt^.Montant1, Pmvt^.Typeecriture,
                      Pmvt^.NumPiece, Pmvt^.Devise, Pmvt^.TauxDev, Pmvt^.CodeMontant,
                      Pmvt^.Montant2, Pmvt^.Montant3, Pmvt^.Etablissement, Pmvt^.Axe,
                      Pmvt^.Numeche, Pmvt^.RefExterne, Pmvt^.Daterefexterne,
                      Pmvt^.Datecreation, Pmvt^.Societe, Pmvt^.Affaire, Pmvt^.Datetauxdev,
                      Pmvt^.Ecranouveau, Pmvt^.Qte1, Pmvt^.Qte2, Pmvt^.QualifQte1,
                      Pmvt^.QualifQte2, Pmvt^.Reflibre, Pmvt^.Tvaencaiss, Pmvt^.Regimetva,
                      Pmvt^.TVA, Pmvt^.TPF, Pmvt^.Contrepartigen, Pmvt^.contrepartiaux,
                      Pmvt^.RefPointage, Pmvt^.datepointage, Pmvt^.daterelance,
                      Pmvt^.datevaleur, Pmvt^.Rib, Pmvt^.Refreleve, Pmvt^.Numeroimmo,
                      Pmvt^.LT0, Pmvt^.LT1, Pmvt^.LT2, Pmvt^.LT3, Pmvt^.LT4, Pmvt^.LT5,
                      Pmvt^.LT6, Pmvt^.LT7, Pmvt^.LT8, Pmvt^.LT9,
                      Pmvt^.TA0, Pmvt^.TA1, Pmvt^.TA2, Pmvt^.TA3, Pmvt^.LM0, Pmvt^.LM1,
                      Pmvt^.LM2, Pmvt^.LM3, Pmvt^.LD, Pmvt^.LB0, Pmvt^.LB1, Pmvt^.Conso, Pmvt^.Couverture,
                      Pmvt^.Couverturedev, Pmvt^.couverturelibre,
                      Pmvt^.Datepaquetmax, Pmvt^.Datepaquetmin, Pmvt^.lettrage,
                      // ajout me les tables libres
                      Pmvt^.lettragedev, Pmvt^.lettragelibre, Pmvt^.etatlettrage,Pmvt^.TA0, Pmvt^.TA1, Pmvt^.TA2,
                      Pmvt^.TA3, Pmvt^.refgescom, Pmvt^.typemvt, Pmvt^.docid, Pmvt^.treso, Pmvt^.numtrait, Pmvt^.numenca, Pmvt^.valide,
                      // ajout me 26-05-2005 ecrcompl                                                         // Fiche 10487
                      Pmvt^.ctdeb, Pmvt^.ctfin, Pmvt^.ctdate, Pmvt^.cleecr, Pmvt^.confidentiel, Pmvt^.cfonbok, PMvt^.codeaccept
                      // BVE 28.08.07 Ajout E_DOCID : Suivi des validations
                      ,Pmvt^.idvalidation, '', '', '', Pmvt^.qualiforigine]);
             end;
             end;
    end;
    writeln(F, Ligne);
  end;
  for i:=0 To  ListeMouvt.Count-1 do
  begin
       Pmvt := ListeMouvt.Items[i];
       if Pmvt <> nil then Dispose (Pmvt);
  end;

end;


procedure Ecrituregeneraux2(var F: TextFile);
var
  Ligne            : string;
  Qa               : Tquery;
  lgsect           : array[0..5] of Integer;
  Bourrana         : array[0..5] of string;
  Sectattend       : array[0..5] of string;
  i                : integer;
  sj,ss,sq         : string;
  Pointage         : string;
begin

  Qa := Opensql('SELECT X_LONGSECTION,X_BOURREANA,X_SECTIONATTENTE from AXE',
    TRUE);
  i := 0;
  while not Qa.EOF do
  begin
    lgsect[i] := Qa.findField('X_LONGSECTION').asinteger;
    Bourrana[i] := Qa.findField('X_BOURREANA').asstring;
    Sectattend[i] := Qa.findField('X_SECTIONATTENTE').asstring;
    inc(i);
    Qa.next;
  end;
  ferme(Qa);
  if GetParamSocSecur('SO_POINTAGEJAL', TRUE) then sj :='X' else sj := '-';
  if GetParamSocSecur('SO_CPPOINTAGESECU', TRUE) then ss :='X' else ss := '-';
  if GetParamSocSecur('SO_CPPCLSAISIEQTE', TRUE) then sq :='X' else sq := '-';

  // fiche 18836
  if GetParamSocSecur ('SO_CPMODESYNCHRO', '', TRUE) then
    Pointage :=  GetParamSocSecur ('SO_CPPOINTAGESX', 'EXP')
  else
    Pointage := '';

  Ligne := Format(Formatgeneraux2, [GetParamSocSecur('SO_LGCPTEGEN', '6'),
    GetParamSocSecur('SO_BOURREGEN', '0'), GetParamSocSecur('SO_LGCPTEAUX', '6'),
    GetParamSocSecur('SO_BOURREAUX', '0'), IntToStr(lgsect[0]), Bourrana[0],
    IntToStr(lgsect[1]), Bourrana[1], IntToStr(lgsect[2]), Bourrana[2],
    IntToStr(lgsect[3]), Bourrana[3],
    IntToStr(lgsect[4]), Bourrana[4], GetParamSocSecur('SO_GENATTEND', ''),
    GetParamSocSecur('SO_CLIATTEND', ''), GetParamSocSecur('SO_FOUATTEND', ''),
    GetParamSocSecur('SO_SALATTEND', ''),
    GetParamSocSecur('SO_DIVATTEND', ''), Sectattend[0], Sectattend[1], Sectattend[2],
    Sectattend[3], Sectattend[4], Pointage, sj, ss, sq]);
  writeln(F, Ligne);
end;

Function RendTypeAno( Exo : string) : string;
var Qa  : Tquery;
Ecranv      : string;
begin
  Ecranv := ''; Result := '';
  Qa := OpenSQl ('select E_ECRANOUVEAU FROM ECRITURE Where E_EXERCICE="'+Exo+'" '+
  ' and (E_ECRANOUVEAU = "H" or E_ECRANOUVEAU ="OAN") ORDER BY E_DATECOMPTABLE DESC', TRUE);
  if not Qa.EOF then
  Result := Qa.findfield ('E_ECRANOUVEAU').asstring;
  ferme(Qa);
end;

procedure EcritureExercice(var F: TextFile; Where : string='');
var
  Qe      : Tquery;
  Ligne   : string;
  Typeano : string;
  EtatCpta: string;
begin
  Qe :=
  Opensql('SELECT EX_EXERCICE,EX_LIBELLE, EX_DATEDEBUT,EX_DATEFIN,EX_ETATCPTA,EX_ETATBUDGET FROM EXERCICE '+Where, TRUE);
  while not Qe.EOF do
  begin
    Typeano := RendTypeAno(Qe.Findfield('EX_EXERCICE').asstring);
    EtatCpta := Qe.Findfield('EX_ETATCPTA').asstring;
    if EtatCpta = 'NON' then EtatCpta := 'OUV';

    Ligne := Format(Formatexercice, [Qe.Findfield('EX_EXERCICE').asstring,
      FormatDateTime(Traduitdateformat('ddmmyyyy'), Qe.Findfield('EX_DATEDEBUT').AsDateTime),
        FormatDateTime(Traduitdateformat('ddmmyyyy'), Qe.Findfield('EX_DATEFIN').AsDateTime),
          EtatCpta, Qe.Findfield('EX_ETATBUDGET').asstring,
          Qe.Findfield('EX_LIBELLE').asstring,Typeano]);
    writeln(F, Ligne);
    Qe.Next;
  end;
  Ferme(Qe);
end;


procedure Ecritureentete( var F: TextFile; Typearchive, Mode : string;Origine : string=''; pp : string=''; Complement : string='');
var
   Ligne,DatePurge, context,
   serie, Exoclo, Dateclo,
   Datebascul, DateFichierCom,
   NoDossier, ModeSynchro : String ;
begin
    if ctxPCL in V_PGI.PGIContexte then context := 'EXP'
    else context := 'CLI';
    if V_PGI.LaSerie = S7 then serie := 'S5';
    if V_PGI.LaSerie = S5 then serie := 'S5';
    if V_PGI.LaSerie = S3 then serie := 'S3';
    if V_PGI.LaSerie = S1 then serie := 'S1';
    if Origine='SIS' then Serie := 'SI';
// ajout me pour la version V590
    Exoclo := GetParamSocSecur ('SO_CPDERNEXOCLO', TRUE);

    if Exoclo = '' then Exoclo := '   ';
    Datebascul := FormatDateTime(Traduitdateformat('ddmmyyyy'),GetParamSocSecur ('SO_DATEBASCULE', idate1900));
    if (GetParamSocSecur ('SO_DATEBASCULE', iDate1900) <= iDate1900) or (Datebascul = '') then
       Datebascul := '01011900';


    DateFichierCom := FormatDateTime(Traduitdateformat('ddmmyyyyhhnn'), now);

    if DateFichierClo  = '' then
    begin
         Dateclo := FormatDateTime(Traduitdateformat('ddmmyyyy'), Getprecedent.fin);
         if (Getprecedent.fin <= iDate1900) or (Dateclo = '') then
            Dateclo := '01011900';
    end
    else
       Dateclo := DateFichierClo;

    NoDossier:=V_PGI.NODOSSIER;
    DatePurge := FormatDateTime(Traduitdateformat('ddmmyyyy'),GetParamSocSecur('SO_CPDATEREFPURGE', iDate1900));
    if GetparamsocSecur ('SO_CPMODESYNCHRO', TRUE)  then
       ModeSynchro := 'SYN'
    else
       ModeSynchro := '';

// la version de fichier venant de sisco2 reste à 001 car
// les generaux restent à  CGE
       Ligne := Format(Formatentete,[serie, context, Typearchive, Mode,Exoclo,Datebascul,Dateclo,
       versionexport,NoDossier,DateFichierCom,V_PGI.UserLogin,'',pp, Complement, NoDossier, ModeSynchro, DatePurge, Sousversion]);
    writeln(F, Ligne);
end;


procedure Ecrituregeneraux(var F: TextFile);
var
  Ligne        : string;
  Tenueuro     : string;
  Geneana      : string;
  AnvDynamique : string;
  croisaxe     : string;
begin
  if GetParamSocSecur('SO_TENUEEURO', 'X') = FALSE then
     Tenueuro := '-'
  else
     Tenueuro := 'X';
  if  GetParamSocSecur('SO_ZGEREANAL', TRUE) then Geneana := 'X' else  Geneana := '-';
  // fiche 10393
  if  GetParamSocSecur('SO_CPANODYNA', TRUE) then AnvDynamique := 'X' else  AnvDynamique := '-';
  if GetParamSocSecur('SO_CROISAXE', False) then croisaxe := 'X' else croisaxe := '-';

  Ligne := Format(Formatgeneraux1, [GetParamSocSecur('SO_LIBELLE', ''),
    GetParamSocSecur('SO_ADRESSE1', ''), GetParamSocSecur('SO_ADRESSE2', ''),
      GetParamSocSecur('SO_ADRESSE3', ''),
      GetParamSocSecur('SO_CODEPOSTAL', ''), GetParamSocSecur('SO_VILLE', ''),
        GetParamSocSecur('SO_PAYS', 'FRF'), GetParamSocSecur('SO_TELEPHONE', ''),
      GetParamSocSecur('SO_FAX', ''), GetParamSocSecur('SO_TELEX', ''), GetParamSocSecur('SO_MAIL', ''),
        GetParamSocSecur('SO_RVA', ''),
      GetParamSocSecur('SO_CONTACT', ''), GetParamSocSecur('SO_NIF', ''), GetParamSocSecur('SO_SIRET', ''),
        GetParamSocSecur('SO_RC', ''),
      GetParamSocSecur('SO_APE', ''), GetParamSocSecur('SO_CAPITAL', ''),
        GetParamSocSecur('SO_TXTJURIDIQUE', '')]);
  writeln(F, Ligne);
  // ecriture paramètres généraux 2 et 3
  Ecrituregeneraux2(F);
  Ligne := Format(Formatgeneraux3, [GetParamSocSecur('SO_OUVREBIL', ''),
    GetParamSocSecur('SO_RESULTAT', ''),
    GetParamSocSecur('SO_OUVREBEN', ''), GetParamSocSecur('SO_OUVREPERTE', ''),
      GetParamSocSecur('SO_JALOUVRE', ''),
      GetParamSocSecur('SO_JALFERME', ''), GetParamSocSecur('SO_JALREPBALAN', ''),
      GetParamSocSecur('SO_DEFCOLCLI', ''), GetParamSocSecur('SO_DEFCOLFOU', ''),
      GetParamSocSecur('SO_DEFCOLSAL', ''), GetParamSocSecur('SO_DEFCOLDDIV', ''),
      GetParamSocSecur('SO_DEFCOLCDIV', ''), GetParamSocSecur('SO_DEFCOLDIV', '')]);
  writeln(F, Ligne);

  Ligne := Format(Formatgeneraux4,
    [GetParamSocSecur('SO_BILDEB1', ''), GetParamSocSecur('SO_BILFIN1', ''),
    GetParamSocSecur('SO_BILDEB2', ''), GetParamSocSecur('SO_BILFIN2', ''),
      GetParamSocSecur('SO_BILDEB3', ''), GetParamSocSecur('SO_BILFIN3', ''),
      GetParamSocSecur('SO_BILDEB4', ''), GetParamSocSecur('SO_BILFIN4', ''),
      GetParamSocSecur('SO_BILDEB5', ''), GetParamSocSecur('SO_BILFIN5', ''),
      GetParamSocSecur('SO_CHADEB1', ''), GetParamSocSecur('SO_CHAFIN1', ''),
      GetParamSocSecur('SO_CHADEB2', ''), GetParamSocSecur('SO_CHAFIN2', ''),
      GetParamSocSecur('SO_CHADEB3', ''), GetParamSocSecur('SO_CHAFIN3', ''),
      GetParamSocSecur('SO_CHADEB4', ''), GetParamSocSecur('SO_CHAFIN4', ''),
      GetParamSocSecur('SO_CHADEB5', ''), GetParamSocSecur('SO_CHAFIN5', ''),
      GetParamSocSecur('SO_PRODEB1', ''), GetParamSocSecur('SO_PROFIN1', ''),
      GetParamSocSecur('SO_PRODEB2', ''), GetParamSocSecur('SO_PROFIN2', ''),
      GetParamSocSecur('SO_PRODEB3', ''), GetParamSocSecur('SO_PROFIN3', ''),
      GetParamSocSecur('SO_PRODEB4', ''), GetParamSocSecur('SO_PROFIN4', ''),
      GetParamSocSecur('SO_PRODEB5', ''), GetParamSocSecur('SO_PROFIN5', '')]);
  writeln(F, Ligne);
  Ligne := Format(Formatgeneraux5, [GetParamSocSecur('SO_DEVISEPRINC', 'EUR'),
    GetParamSocSecur('SO_DECVALEUR', '2'),
      Tenueuro, '','','',
      GetParamSocSecur('SO_TAUXEURO', '6,55957'),
        GetParamSocSecur('SO_REGIMEDEFAUT', ''),
      GetParamSocSecur('SO_GCMODEREGLEDEFAUT', ''), GetParamSocSecur('SO_CODETVADEFAUT', ''),
        GetParamSocSecur('SO_CODETVAGENEDEFAULT', ''),
      GetParamSocSecur('SO_ETABLISDEFAUT', ''), GetParamSocSecur('SO_NUMPLANREF', '7'), Geneana, croisaxe, AnvDynamique]);
  writeln(F, Ligne);
end;

// Table libre personnalisé
procedure EcritureChoixCod(var F: TextFile; Serie : string='S5');
var
  Qe: Tquery;
  Ligne: string;
  procedure EcritureFich(Code, Ext, Requete : string);
  begin
        Qe := Opensql(Requete, TRUE);
        while not Qe.EOF do
        begin
          Ligne := Format(FormatChoixCode,
          [Code, Qe.Findfield(Ext+'TYPE').asstring, Qe.Findfield(Ext+'CODE').asstring, Qe.Findfield(Ext+'LIBELLE').asstring,
             Qe.Findfield(Ext+'ABREGE').asstring, Qe.Findfield(Ext+'LIBRE').asstring]);
          writeln(F, Ligne);
          Qe.Next;
        end;
        Ferme(Qe);
  end;

begin
  if Serie = 'S1' then
        EcritureFich('CTL', 'CC_', 'SELECT * FROM  CHOIXCOD where CC_TYPE="QME"')
  else
  begin
        EcritureFich('CTL', 'CC_', 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="NAT" and CC_ABREGE="X" ');
        EcritureFich('CTL', 'CC_', 'SELECT * FROM  CHOIXCOD where CC_TYPE="ZLI" and (CC_CODE like "CT%" or CC_CODE like "CR%") and CC_LIBELLE<>".-"');
      (*
        suite à UPDATE CHOIXCOD SET CC_TYPE="ZLA" WHERE CC_TYPE="ZLI" AND CC_CODE LIKE "A%"
        UPDATE CHOIXCOD SET CC_TYPE="ZLT" WHERE CC_TYPE="ZLI" AND (CC_CODE LIKE "C%" OR CC_CODE LIKE "F%")
      *)
        EcritureFich('CTL', 'CC_', 'SELECT * FROM  CHOIXCOD where CC_TYPE="ZLT" and (CC_CODE like "CT%" or CC_CODE like "CR%") and CC_LIBELLE<>".-"');

        //Sauvegarde de qualifiant qte  Reprendre dans 6.5
        EcritureFich('CTL', 'CC_', 'SELECT * FROM  CHOIXCOD where CC_TYPE="QME"');

        EcritureFich('CTL', 'CC_', 'SELECT * FROM  CHOIXCOD where CC_TYPE="JUR"');

        EcritureFich('XTL', 'YX_', 'SELECT * FROM  CHOIXEXT where YX_TYPE like  "LT%" ');
  end;
end;

procedure EcritureDesRessource(var F: TextFile; Where : string='');
var
  Qe   : Tquery;
  Ligne: string;
begin
     Qe := Opensql(' select ARS_TYPERESSOURCE,ARS_RESSOURCE,ARS_LIBELLE,ARS_LIBELLE2 '+
                ' from RESSOURCE where ars_ressource '+
                ' in (select distinct ytc_ressource1 from tierscompl) or '+
                ' ars_ressource in (select distinct ytc_ressource2 from tierscompl) '+
                ' or ars_ressource in (select distinct ytc_ressource3 from tierscompl )', TRUE);
      while not Qe.EOF do
      begin
        Ligne := Format(FormatRessource,
        [Qe.Findfield('ARS_TYPERESSOURCE').asstring, Qe.Findfield('ARS_RESSOURCE').asstring, Qe.Findfield('ARS_LIBELLE').asstring,
           Qe.Findfield('ARS_LIBELLE2').asstring]);
        writeln(F, Ligne);
        Qe.Next;
      end;
end;


// Paramétrage Table libre personnalisé
procedure EcritureParamlib(var F: TextFile; TP : TOB);
var
  Ligne: string;
  i    : integer;
  TA   : TOB;
begin

For i := 0 to TP.detail.count-1 do
begin
     TA := TP.detail[i];
    if TA.FieldExists('NT_TEXTE0') then
    Ligne := Format(FormatParamlib,
    [TA.GetValue('PL_TABLE'), TA.GetValue('LIBELLE'), TA.GetValue('VISIBLE'),
     TA.GetValue('NT_TEXTE0'), TA.GetValue('VNT_TEXTE0'),
     TA.GetValue('NT_TEXTE1'), TA.GetValue('VNT_TEXTE1'),
     TA.GetValue('NT_TEXTE2'), TA.GetValue('VNT_TEXTE2'),
     TA.GetValue('NT_TEXTE3'), TA.GetValue('VNT_TEXTE3'),
     TA.GetValue('NT_TEXTE4'), TA.GetValue('VNT_TEXTE4'),
     TA.GetValue('NT_TEXTE5'), TA.GetValue('VNT_TEXTE5'),
     TA.GetValue('NT_TEXTE6'), TA.GetValue('VNT_TEXTE6'),
     TA.GetValue('NT_TEXTE7'), TA.GetValue('VNT_TEXTE7'),
     TA.GetValue('NT_TEXTE8'), TA.GetValue('VNT_TEXTE8'),
     TA.GetValue('NT_TEXTE9'), TA.GetValue('VNT_TEXTE9'),
     TA.GetValue('NT_MONTANT1'), TA.GetValue('VNT_MONTANT1'),
     TA.GetValue('NT_MONTANT2'), TA.GetValue('VNT_MONTANT2'),
     TA.GetValue('NT_MONTANT3'), TA.GetValue('VNT_MONTANT3'),
     TA.GetValue('NT_MONTANT4'), TA.GetValue('VNT_MONTANT4'),
     TA.GetValue('NT_BOOLEAN1'), TA.GetValue('VNT_BOOLEAN1'),
     TA.GetValue('NT_BOOLEAN2'), TA.GetValue('VNT_BOOLEAN2'),
     TA.GetValue('NT_BOOLEAN3'), TA.GetValue('VNT_BOOLEAN3'),
     TA.GetValue('NT_BOOLEAN4'), TA.GetValue('VNT_BOOLEAN4'),
     TA.GetValue('NT_DATE1'), TA.GetValue('VNT_DATE1'),
     TA.GetValue('NT_DATE2'), TA.GetValue('VNT_DATE2'),
     TA.GetValue('NT_DATE3'), TA.GetValue('VNT_DATE3'),
     TA.GetValue('NT_DATE4'), TA.GetValue('VNT_DATE4')])
     else
     if TA.FieldExists('Y_LIBREBOOL0') then
    Ligne := Format(FormatParamlib,
    [TA.GetValue('PL_TABLE'), TA.GetValue('LIBELLE'), TA.GetValue('VISIBLE'),
     TA.GetValue('Y_LIBRETEXTE0'), TA.GetValue('VY_LIBRETEXTE0'),
     TA.GetValue('Y_LIBRETEXTE1'), TA.GetValue('VY_LIBRETEXTE1'),
     TA.GetValue('Y_LIBRETEXTE2'), TA.GetValue('VY_LIBRETEXTE2'),
     TA.GetValue('Y_LIBRETEXTE3'), TA.GetValue('VY_LIBRETEXTE3'),
     TA.GetValue('Y_LIBRETEXTE4'), TA.GetValue('VY_LIBRETEXTE4'),
     TA.GetValue('Y_LIBRETEXTE5'), TA.GetValue('VY_LIBRETEXTE5'),
     TA.GetValue('Y_LIBRETEXTE6'), TA.GetValue('VY_LIBRETEXTE6'),
     TA.GetValue('Y_LIBRETEXTE7'), TA.GetValue('VY_LIBRETEXTE7'),
     TA.GetValue('Y_LIBRETEXTE8'), TA.GetValue('VY_LIBRETEXTE8'),
     TA.GetValue('Y_LIBRETEXTE9'), TA.GetValue('VY_LIBRETEXTE9'),
     TA.GetValue('Y_LIBREMONTANT0'), TA.GetValue('VY_LIBREMONTANT0'),
     TA.GetValue('Y_LIBREMONTANT1'), TA.GetValue('VY_LIBREMONTANT1'),
     TA.GetValue('Y_LIBREMONTANT2'), TA.GetValue('VY_LIBREMONTANT2'),
     TA.GetValue('Y_LIBREMONTANT3'), TA.GetValue('VY_LIBREMONTANT3'),
     TA.GetValue('Y_LIBREBOOL0'), TA.GetValue('VY_LIBREBOOL0'),
     TA.GetValue('Y_LIBREBOOL1'), TA.GetValue('VY_LIBREBOOL1'),
     '', '',
     TA.GetValue('Y_LIBREDATE'), TA.GetValue('VY_LIBREDATE'),
     TA.GetValue('Y_TABLE0'), TA.GetValue('VY_TABLE0'),
     TA.GetValue('Y_TABLE1'), TA.GetValue('VY_TABLE1'),
     TA.GetValue('Y_TABLE2'), TA.GetValue('VY_TABLE2'),
     TA.GetValue('Y_TABLE3'), TA.GetValue('VY_TABLE3')])
    else
     if TA.FieldExists('E_LIBREBOOL0') then
    Ligne := Format(FormatParamlib,
    [TA.GetValue('PL_TABLE'), TA.GetValue('LIBELLE'), TA.GetValue('VISIBLE'),
     TA.GetValue('E_LIBRETEXTE0'), TA.GetValue('VE_LIBRETEXTE0'),
     TA.GetValue('E_LIBRETEXTE1'), TA.GetValue('VE_LIBRETEXTE1'),
     TA.GetValue('E_LIBRETEXTE2'), TA.GetValue('VE_LIBRETEXTE2'),
     TA.GetValue('E_LIBRETEXTE3'), TA.GetValue('VE_LIBRETEXTE3'),
     TA.GetValue('E_LIBRETEXTE4'), TA.GetValue('VE_LIBRETEXTE4'),
     TA.GetValue('E_LIBRETEXTE5'), TA.GetValue('VE_LIBRETEXTE5'),
     TA.GetValue('E_LIBRETEXTE6'), TA.GetValue('VE_LIBRETEXTE6'),
     TA.GetValue('E_LIBRETEXTE7'), TA.GetValue('VE_LIBRETEXTE7'),
     TA.GetValue('E_LIBRETEXTE8'), TA.GetValue('VE_LIBRETEXTE8'),
     TA.GetValue('E_LIBRETEXTE9'), TA.GetValue('VE_LIBRETEXTE9'),
     TA.GetValue('E_LIBREMONTANT0'), TA.GetValue('VE_LIBREMONTANT0'),
     TA.GetValue('E_LIBREMONTANT1'), TA.GetValue('VE_LIBREMONTANT1'),
     TA.GetValue('E_LIBREMONTANT2'), TA.GetValue('VE_LIBREMONTANT2'),
     TA.GetValue('E_LIBREMONTANT3'), TA.GetValue('VE_LIBREMONTANT3'),
     TA.GetValue('E_LIBREBOOL0'), TA.GetValue('VE_LIBREBOOL0'),
     TA.GetValue('E_LIBREBOOL1'), TA.GetValue('VE_LIBREBOOL1'),
     '', '',
     TA.GetValue('E_LIBREDATE'), TA.GetValue('VE_LIBREDATE'),
     TA.GetValue('E_TABLE0'), TA.GetValue('VE_TABLE0'),
     TA.GetValue('E_TABLE1'), TA.GetValue('VE_TABLE1'),
     TA.GetValue('E_TABLE2'), TA.GetValue('VE_TABLE2'),
     TA.GetValue('E_TABLE3'), TA.GetValue('VE_TABLE3')]);

     if TA.FieldExists('BE_LIBRETEXTE1') then
    Ligne := Format(FormatParamlib,
    [TA.GetValue('PL_TABLE'), TA.GetValue('LIBELLE'), TA.GetValue('VISIBLE'),
     TA.GetValue('BE_LIBRETEXTE1'), TA.GetValue('VBE_LIBRETEXTE1'),
     TA.GetValue('BE_LIBRETEXTE2'), TA.GetValue('VBE_LIBRETEXTE2'),
     TA.GetValue('BE_LIBRETEXTE3'), TA.GetValue('VBE_LIBRETEXTE3'),
     TA.GetValue('BE_LIBRETEXTE4'), TA.GetValue('VBE_LIBRETEXTE4'),
     TA.GetValue('BE_LIBRETEXTE5'), TA.GetValue('VBE_LIBRETEXTE5'),
     '','', '','', '','', '','', '','', '','', '','', '','',
     '','', '','', '', '', '','', '','',
     TA.GetValue('BE_TABLE0'), TA.GetValue('VBE_TABLE0'),
     TA.GetValue('BE_TABLE1'), TA.GetValue('VBE_TABLE1'),
     TA.GetValue('BE_TABLE2'), TA.GetValue('VBE_TABLE2'),
     TA.GetValue('BE_TABLE3'), TA.GetValue('VBE_TABLE3')]);

     writeln(F, Ligne);
end;


end;

procedure Ecrituretablelibre(ListeLibre: TList; var F: TextFile);
var
  Ligne: string;
  i: integer;
  Tlibre: PLibre;
begin
  for i := 0 to ListeLibre.Count - 1 do
  begin
    Tlibre := ListeLibre.Items[i];
    Ligne := Format(Formattablelibre, [Tlibre^.identifiant,
      Tlibre^.code, Tlibre^.libelle, Tlibre^.typetable,
        Tlibre^.tx0, Tlibre^.tx1, Tlibre^.tx2, Tlibre^.tx3, Tlibre^.tx4,
          Tlibre^.tx5, Tlibre^.tx6, Tlibre^.tx7, Tlibre^.tx8, Tlibre^.tx9]);
    writeln(F, Ligne);
  end;

  for i:=0 To  ListeLibre.Count-1 do
  begin
       Tlibre := ListeLibre.Items[i];
       if Tlibre <> nil then Dispose (Tlibre);
  end;

end;

procedure EcritureEtablissement(var F: TextFile; Where : string='');
var
  Qe: Tquery;
  Ligne: string;
begin
  Qe := Opensql('SELECT * FROM ETABLISS '+ Where, TRUE);
  while not Qe.EOF do
  begin
    Ligne := Format(Formatetablissement,
      [Qe.Findfield('ET_ETABLISSEMENT').asstring,
      Qe.Findfield('ET_LIBELLE').asstring,
      Qe.Findfield('ET_ADRESSE1').asstring,
      Qe.Findfield('ET_ADRESSE2').asstring,
      Qe.Findfield('ET_ADRESSE3').asstring,
      Qe.Findfield('ET_CODEPOSTAL').asstring,
      Qe.Findfield('ET_VILLE').asstring,
      Qe.Findfield('ET_PAYS').asstring,
      Qe.Findfield('ET_TELEPHONE').asstring,
      Qe.Findfield('ET_TELEX').asstring,
      Qe.Findfield('ET_SIRET').asstring,
      Qe.Findfield('ET_APE').asstring
      ]);
    writeln(F, Ligne);
    Qe.Next;
  end;
  Ferme(Qe);
end;

procedure Ecrituretablesection(ListeSection: TList; var F: TextFile);
var
  Ligne: string;
  i: integer;
  TSection: PSection;
begin
  for i := 0 to ListeSection.Count - 1 do
  begin
    TSection := ListeSection.Items[i];
    Ligne := Format(Formatsection, [TSection^.code, TSection^.libelle,
      TSection^.axe, TSection^.codesp, TSection^.debsp, TSection^.lgsp, TSection^.libellesp]);
    writeln(F, Ligne);
  end;
  for i:=0 To  ListeSection.Count-1 do
  begin
       TSection := ListeSection.Items[i];
       if TSection <> nil then Dispose (TSection);
  end;


end;

//EDTLC => MP_LETTRECHEQUE Boolean : Si oui, les flux sur ce mode de paiement peuvent faire l'objet d'une édition de lettre chèque
//EDTBOR => MP_LETTRETRAITE Boolean : Si oui, les flux sur ce mode de paiement peuvent faire l'objet d'une édition de lettre traite
//CONDITIONM => MP_CONDITION Boolean : Condition sur l'affectation du Mode de paiement en saisie. Voir les deux zones aprés :
//MONTANT => MP_MONTANTMAX
//Si MP CONDITION EST à X Tout flux créé avec un montant supérieur à MP_MONTANTMAX prent automatiquement MP_REMPLACEMAX comme mode de paiement
//MODEREMPLACE => MP_REMPLACEMAX

procedure EcritureModepaiement(var F: TextFile; var Corresp : string);
var
  Qe      : Tquery;
  Ligne   : string;
  MDP     : string;
begin
  Qe := Opensql('SELECT * FROM MODEPAIE', TRUE);
  while not Qe.EOF do
  begin
    MDP := findcorrespMDP(Qe.Findfield('MP_MODEPAIE').asstring, Corresp);
    Ligne := Format(Formatpaiement, [MDP,
      Qe.Findfield('MP_LIBELLE').asstring, Qe.Findfield('MP_CATEGORIE').asstring,
        Qe.Findfield('MP_CODEACCEPT').asstring, Qe.Findfield('MP_LETTRECHEQUE').asstring,
        Qe.Findfield('MP_LETTRETRAITE').asstring, Qe.Findfield('MP_CONDITION').asstring,
        FloatToStr(Qe.Findfield('MP_MONTANTMAX').asfloat), Qe.Findfield('MP_REMPLACEMAX').asstring,
        Qe.Findfield('MP_GENERAL').asstring]); // fiche 10604
    writeln(F, Ligne);
    Qe.Next;
  end;
  Ferme(Qe);
end;

function CoherenceMDRS1 (var Code : string) : Boolean;
var
ii : integer;
begin
     Result := FALSE;
     for ii := 0 to 18 do
             if (ControlMDR[ii].MDR  = Code) then begin Result := TRUE; exit; end;
     Code := 'DIV';
end;

procedure Ecrituretableregle(Listeregle: TList; var F: TextFile; Origine : string ; stArg : string='');
var
  Ligne   : string;
  i       : integer;
  TRegle  : Preglement;
  MDR     : string;
begin
  for i := 0 to Listeregle.Count - 1 do
  begin
    TRegle := Listeregle.Items[i];
    MDR := findcorrespMDR (TRegle^.code, stArg);

    Ligne := Format(Formatregle, [MDR, TRegle^.libelle, TRegle^.apartirde,
      TRegle^.plusjour, TRegle^.arrondijour, TRegle^.nbeche, TRegle^.separepar,
      TRegle^.montantmin, TRegle^.remplacemin, TRegle^.mp1, TRegle^.taux1, TRegle^.mp2,
      TRegle^.taux2, TRegle^.mp3, TRegle^.taux3, TRegle^.mp4, TRegle^.taux4, TRegle^.mp5, TRegle^.taux5, TRegle^.mp6,
      TRegle^.taux6, TRegle^.mp7, TRegle^.taux7, TRegle^.mp8, TRegle^.taux8, TRegle^.mp9,
      TRegle^.taux9, TRegle^.mp10, TRegle^.taux10, TRegle^.mp11, TRegle^.taux11,
      TRegle^.mp12, TRegle^.taux12]);
    writeln(F, Ligne);
  end;
  for i:=0 To  Listeregle.Count-1 do
  begin
       TRegle := Listeregle.Items[i];
       if TRegle <> nil then Dispose (TRegle);
  end;

end;

procedure Ecrituredevise(var F: TextFile);
var
  Qd: Tquery;
  Ligne: string;
begin
  Qd := Opensql('SELECT * FROM DEVISE', TRUE);
  while not Qd.EOF do
  begin
    Ligne := Format(Formatdevise, [Qd.Findfield('D_DEVISE').asstring,
      Qd.Findfield('D_LIBELLE').asstring, Qd.Findfield('D_SYMBOLE').asstring,
        Qd.Findfield('D_FERME').asstring,
          IntTostr(Qd.Findfield('D_DECIMALE').asinteger),
        FloatToStr(Qd.Findfield('D_QUOTITE').asfloat),
          Qd.Findfield('D_MONNAIEIN').asstring,
        FloatToStr(Qd.Findfield('D_PARITEEURO').asfloat),
        Qd.Findfield('D_CODEISO').asstring, Qd.Findfield('D_FONGIBLE').asstring,
        Qd.Findfield('D_CPTLETTRDEBIT').asstring,
          Qd.Findfield('D_CPTLETTRCREDIT').asstring,
        Qd.Findfield('D_CPTPROVDEBIT').asstring,
          Qd.Findfield('D_CPTPROVCREDIT').asstring,
        FloatToStr(Qd.Findfield('D_MAXDEBIT').asfloat),
          FloatToStr(Qd.Findfield('D_MAXCREDIT').asfloat)]);
    writeln(F, Ligne);
    Qd.Next;
  end;
  Ferme(Qd);
end;

procedure Ecrituretva(var F: TextFile);
var
  Qt: Tquery;
  Ligne: string;
begin
  Qt :=
    Opensql('SELECT CC_TYPE,CC_CODE,CC_LIBELLE FROM CHOIXCOD where CC_TYPE="RTV"',
    TRUE);
  while not Qt.EOF do
  begin
    Ligne := Format(Formattva, [Qt.Findfield('CC_CODE').asstring,
      Qt.Findfield('CC_LIBELLE').asstring]);
    writeln(F, Ligne);
    Qt.Next;
  end;
  Ferme(Qt);
end;

procedure EcritureSectionana(var F: TextFile; Where :string);
var
  Ligne: string;
  Qa: Tquery;
begin

  Qa := Opensql('SELECT * from SECTION '+ Where, TRUE);
  while not Qa.EOF do
  begin
    Ligne := Format(FormatSectionana,
      [Qa.Findfield('S_SECTION').asstring,
      Qa.Findfield('S_LIBELLE').asstring,
        Qa.Findfield('S_AXE').asstring,
        Qa.Findfield('S_TABLE0').asstring, Qa.Findfield('S_TABLE1').asstring,
        Qa.Findfield('S_TABLE2').asstring, Qa.Findfield('S_TABLE3').asstring,
        Qa.Findfield('S_TABLE4').asstring, Qa.Findfield('S_TABLE5').asstring,
        Qa.Findfield('S_TABLE6').asstring, Qa.Findfield('S_TABLE7').asstring,
        Qa.Findfield('S_TABLE8').asstring, Qa.Findfield('S_TABLE9').asstring,
        Qa.Findfield('S_ABREGE').asstring, Qa.Findfield('S_SENS').asstring]);
    writeln(F, Ligne);
    Qa.next;
  end;
  ferme(Qa);
end;

procedure EcritureSouche(var F: TextFile; Where : string);
var
  Ligne   : string;
  Qa      : Tquery;
  Rel     : string;
begin
  Rel := 'N';
  Qa := Opensql('SELECT SH_SOUCHE,SH_LIBELLE,SH_NUMDEPART,SH_SIMULATION,SH_ANALYTIQUE,SH_SOUCHEEXO from SOUCHE ' + Where, TRUE);
  while not Qa.EOF do
  begin
  if  Qa.Findfield('SH_SOUCHE').asstring = 'REL' then Rel := 'O';
    Ligne := Format(FormatSouche,
      [Qa.Findfield('SH_SOUCHE').asstring, Qa.Findfield('SH_LIBELLE').asstring,
        Qa.Findfield('SH_NUMDEPART').asinteger, Qa.Findfield('SH_SIMULATION').asstring,
          Qa.Findfield('SH_ANALYTIQUE').asstring, Rel]);
    writeln(F, Ligne);
    Qa.next;
  end;
  ferme(Qa);
end;


procedure EcritureCorresp(var F: TextFile);
var
  Ligne: string;
  Qa: Tquery;
  Code: string;
begin
  Code := 'CRR';
  // IGE : Cpt généraux ; IAU : Tiers ; IA1 : Axe 1 .. IA5 : Axe 5 ; IPM : Mode paiement ; IJA : journal

  Qa := Opensql('SELECT CR_CORRESP,CR_LIBELLE,CR_TYPE from CORRESP Where' +
    ' CR_TYPE like "GE%" OR CR_TYPE like "AU%" OR CR_TYPE="A%"', TRUE);
  while not Qa.EOF do
  begin
    Ligne := Format(Formatcorresp,
      [Code, Qa.Findfield('CR_CORRESP').asstring,
        Qa.Findfield('CR_LIBELLE').asstring,
      Qa.Findfield('CR_TYPE').asstring]);
    writeln(F, Ligne);
    Qa.next;
  end;
  ferme(Qa);
end;

procedure EcritureCorrespimpot(var F: TextFile);
var
  Ligne: string;
  Qa: Tquery;
  Code: string;
begin
  Code := 'CRI';

  // IGE : Cpt généraux ; IAU : Tiers ; IA1 : Axe 1 .. IA5 : Axe 5 ; IPM : Mode paiement ; IJA : journal

  Qa := Opensql('SELECT CR_CORRESP,CR_LIBELLE,CR_TYPE from CORRESP Where' +
    ' CR_TYPE like "I%" ', TRUE);
  while not Qa.EOF do
  begin
    Ligne := Format(FormatcorrespImp,
      [Code, Qa.Findfield('CR_CORRESP').asstring,
        Qa.Findfield('CR_LIBELLE').asstring,
      Qa.Findfield('CR_TYPE').asstring]);
    writeln(F, Ligne);
    Qa.next;
  end;
  ferme(Qa);
end;


function Comparejournal (Item1,Item2:Pointer) : integer;
var
TMvt1,TMvt2             : PListeMouvt;
begin
  Result := -1;
  if Item1 = nil then Result := -1;
  if Item2 = nil then Result := 1;
  if (Item1 = nil) or (Item2 = nil) then exit;
  Result := 0;

  TMvt1 := Item1; TMvt2 := Item2;
  if TMvt1^.Journal > TMvt2^.Journal then Result := 1
  else if TMvt1^.Journal < TMvt2^.Journal then Result := -1
  else
  if TMvt1^.Journal = TMvt2^.Journal then
  begin
       if TMvt1^.General > TMvt2^.General then Result := 1
       else if TMvt1^.General < TMvt2^.General then Result := -1
       else
       begin
            if TMvt1^.General = TMvt2^.General then
            begin
                      if TMvt1^.AuxSect > TMvt2^.AuxSect then Result := 1
                      else
                      if TMvt1^.AuxSect < TMvt2^.AuxSect then Result := -1;
            end;
       end;
  end;
end;
procedure EcrireCpteBudgetaire(var F: TextFile; Where :string);
var
Q1           :TQuery;
WW, Ligne    : string;
begin

WW := 'SELECT BG_ATTENTE,BG_BLOQUANT,BG_BUDGENE,BG_DATECREATION,BG_DATEFERMETURE,'+
'BG_DATEOUVERTURE,BG_FERME,BG_LIBELLE,BG_REPORTDISPO,BG_SENS,BG_SIGNE,'+
'BG_TABLE0,BG_TABLE1,BG_TABLE2,BG_TABLE3,BG_TABLE4,BG_TABLE5,BG_TABLE6,'+
'BG_TABLE7,BG_TABLE8,BG_TABLE9,BG_COMPTERUB,BG_EXCLURUB FROM BUDGENE'+
Where + ' ORDER BY BG_BUDGENE ';

 Q1:=OpenSQL(WW,TRUE) ;
 While Not Q1.Eof Do BEGIN
    Ligne := Format(FormatcpteBudg, [Q1.FindField ('BG_BUDGENE').asstring, Q1.FindField ('BG_LIBELLE').asstring,
      Q1.FindField ('BG_SENS').asstring, Q1.FindField ('BG_SIGNE').asstring,
      Q1.FindField ('BG_ATTENTE').asstring, Q1.FindField ('BG_BLOQUANT').asstring, Q1.FindField ('BG_ATTENTE').asstring,
      FormatDateTime(Traduitdateformat('dd/mm/yyyy'),Q1.FindField ('BG_DATECREATION').AsDateTime),
      FormatDateTime(Traduitdateformat('dd/mm/yyyy'),Q1.FindField ('BG_DATEOUVERTURE').AsDateTime),
      Q1.FindField ('BG_FERME').asstring, FormatDateTime(Traduitdateformat('dd/mm/yyyy'),Q1.FindField ('BG_DATEFERMETURE').AsDateTime),
      Q1.FindField ('BG_REPORTDISPO').asstring, Q1.FindField ('BG_TABLE0').asstring,
      Q1.FindField ('BG_TABLE1').asstring, Q1.FindField ('BG_TABLE2').asstring,
      Q1.FindField ('BG_TABLE3').asstring, Q1.FindField ('BG_TABLE4').asstring,
      Q1.FindField ('BG_TABLE5').asstring, Q1.FindField ('BG_TABLE6').asstring,
      Q1.FindField ('BG_TABLE7').asstring, Q1.FindField ('BG_TABLE8').asstring,
      Q1.FindField ('BG_TABLE9').asstring, Q1.FindField ('BG_COMPTERUB').asstring,
      Q1.FindField ('BG_EXCLURUB').asstring]);
    writeln(F, Ligne);
    Q1.NExt ;
END ;
Ferme(Q1) ;
end;


procedure EcrireJalBudgetaire(var F: TextFile; Where :string);
var
Q1           :TQuery;
WW, Ligne    : string;
begin

WW := 'SELECT BJ_BUDJAL, BJ_LIBELLE, BJ_AXE, BJ_DATECREATION, BJ_DATEOUVERTURE,' +
'BJ_DATEFERMETURE, BJ_EXODEB, BJ_EXOFIN, BJ_COMPTEURNORMAL, BJ_COMPTEURSIMUL,' +
'BJ_FERME, BJ_NATJAL, BJ_PERDEB, BJ_PERFIN, BJ_SENS, BJ_CATEGORIE,' +
'BJ_SOUSPLAN, BJ_BUDGENES, BJ_BUDGENES2, BJ_BUDSECTS, BJ_BUDSECTS2 FROM BUDJAL'+
Where + ' ORDER BY BJ_BUDJAL ';

 Q1:=OpenSQL(WW,TRUE) ;
 While Not Q1.Eof Do BEGIN
    Ligne := Format(FormatJalBudg, [Q1.FindField ('BJ_BUDJAL').asstring, Q1.FindField ('BJ_LIBELLE').asstring,
      Q1.FindField ('BJ_AXE').asstring,
      FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('BJ_DATECREATION').AsDateTime),
      FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('BJ_DATEOUVERTURE').AsDateTime),
      FormatDateTime(Traduitdateformat('ddmmyyyy'),Q1.FindField ('BJ_DATEFERMETURE').AsDateTime),
      Q1.FindField ('BJ_EXODEB').asstring, Q1.FindField ('BJ_EXOFIN').asstring,
      Q1.FindField ('BJ_COMPTEURNORMAL').asstring, Q1.FindField ('BJ_COMPTEURSIMUL').asstring,
      Q1.FindField ('BJ_FERME').asstring, Q1.FindField ('BJ_NATJAL').asstring,
      FormatDateTime(Traduitdateformat('ddmmyyyy'), Q1.FindField ('BJ_PERDEB').AsDateTime),
      FormatDateTime(Traduitdateformat('ddmmyyyy'), Q1.FindField ('BJ_PERFIN').AsDateTime),
      Q1.FindField ('BJ_SENS').asstring, Q1.FindField ('BJ_CATEGORIE').asstring,
      Q1.FindField ('BJ_SOUSPLAN').asstring,
      Q1.FindField ('BJ_BUDGENES').asstring,
      Q1.FindField ('BJ_BUDGENES2').asstring,
      Q1.FindField ('BJ_BUDSECTS').asstring,
      Q1.FindField ('BJ_BUDSECTS2').asstring]);

    writeln(F, Ligne);
    Q1.NExt ;
END ;
Ferme(Q1) ;
end;

Procedure EcritureSectionBudgetaire(var F: TextFile; Where :string);
var
WW,Ligne : string;
Q1       : TQuery;
begin

//WHERE BS_DATEMODIF>="01/01/1900" AND BS_DATEMODIF<="12/31/2099" AND BS_BUDSECT>="100000" AND BS_BUDSECT<="SECATT" AND BS_AXE="A1" AND BS_EXPORTE<>"X" ORDER BY BS_BUDSECT
WW := ' SELECT * FROM BUDSECT ' + Where + ' ORDER BY BS_AXE, BS_BUDSECT ';
 Q1:=OpenSQL(WW,TRUE) ;
 While Not Q1.Eof Do BEGIN
    Ligne := Format(FormatSectionBudg, [Q1.FindField ('BS_BUDSECT').asstring,
      Q1.FindField ('BS_AXE').asstring, Q1.FindField ('BS_LIBELLE').asstring,Q1.FindField ('BS_ABREGE').asstring,
      Q1.FindField ('BS_RUB').asstring,
      FormatDateTime(Traduitdateformat('dd/mm/yyyy'),Q1.FindField ('BS_DATECREATION').AsDateTime),
      FormatDateTime(Traduitdateformat('dd/mm/yyyy'),Q1.FindField ('BS_DATEOUVERTURE').AsDateTime),
      Q1.FindField ('BS_FERME').asstring,
      FormatDateTime(Traduitdateformat('dd/mm/yyyy'),Q1.FindField ('BS_DATEFERMETURE').AsDateTime),
      Q1.FindField ('BS_CREERPAR').asstring, Q1.FindField ('BS_REPORTDISPO').asstring,
      Q1.FindField ('BS_ATTENTE').asstring, Q1.FindField ('BS_SENS').asstring,
      Q1.FindField ('BS_SIGNE').asstring,
      Q1.FindField ('BS_TABLE0').asstring, Q1.FindField ('BS_TABLE1').asstring,
      Q1.FindField ('BS_TABLE2').asstring, Q1.FindField ('BS_TABLE3').asstring,
      Q1.FindField ('BS_TABLE4').asstring, Q1.FindField ('BS_TABLE5').asstring,
      Q1.FindField ('BS_TABLE6').asstring, Q1.FindField ('BS_TABLE7').asstring,
      Q1.FindField ('BS_TABLE8').asstring, Q1.FindField ('BS_TABLE9').asstring,
      Q1.FindField ('BS_SECTIONRUB').asstring, Q1.FindField ('BS_EXCLURUB').asstring]);
    writeln(F, Ligne);
    Q1.NExt ;
 END;

end;


Procedure EcritureRelance(var F: TextFile; Where :string);
var
Ligne     : string;
Q1        : TQuery;
begin
 Q1:=OpenSQL(' SELECT * FROM RELANCE ' + Where,TRUE) ;
 While Not Q1.Eof Do BEGIN
    Ligne := Format(FormatRelance, [Q1.FindField ('RR_TYPERELANCE').asstring,
      Q1.FindField ('RR_FAMILLERELANCE').asstring, Q1.FindField ('RR_LIBELLE').asstring,
      Q1.FindField ('RR_GROUPELETTRE').asstring, Q1.FindField ('RR_NONECHU').asstring,
      Q1.FindField ('RR_SCOORING').asstring,
      IntTostr(Q1.FindField ('RR_DELAI1').asinteger), Q1.FindField ('RR_MODELE1').asstring,
      IntTostr(Q1.FindField ('RR_DELAI2').asinteger), Q1.FindField ('RR_MODELE2').asstring,
      IntTostr(Q1.FindField ('RR_DELAI3').asinteger), Q1.FindField ('RR_MODELE3').asstring,
      IntTostr(Q1.FindField ('RR_DELAI4').asinteger), Q1.FindField ('RR_MODELE4').asstring,
      IntTostr(Q1.FindField ('RR_DELAI5').asinteger), Q1.FindField ('RR_MODELE5').asstring,
      IntTostr(Q1.FindField ('RR_DELAI6').asinteger), Q1.FindField ('RR_MODELE6').asstring,
      IntTostr(Q1.FindField ('RR_DELAI7').asinteger), Q1.FindField ('RR_MODELE7').asstring,
      Q1.FindField ('RR_ENJOURS').asstring]);
    writeln(F, Ligne);
    Q1.NExt ;
 END;

end;

{$ENDIF}

function CompteEstalpha (ListeCpteGen: TList) : Boolean ;
var
  i,j: integer;
  TCGene: PListeCpteGen;
begin
    Result := TRUE;
    for i:=0 To  ListeCpteGen.Count-1 do
    begin
         TCGene := ListeCpteGen.Items[i];
         for j:=1 To  length(TCGene^.Code) do
             if TCGene^.Code[j] in Alpha then exit;
    end;
    Result := FALSE;
end;

procedure LibererGen (ListeCpteGen: TList);
var
  i: integer;
  TCGene: PListeCpteGen;
begin
    for i:=0 To  ListeCpteGen.Count-1 do
    begin
         TCGene := ListeCpteGen.Items[i];
         if (TCGene <> nil) then Dispose (TCGene);
    end;
end;

// paramétrage des BAP
Procedure EcritureBapTypeVisa(var F: TextFile);
var
Ligne     : string;
Q1        : TQuery;
begin
 Q1:=OpenSQL(' SELECT * FROM CPTYPEVISA ',TRUE) ;
 While Not Q1.Eof Do BEGIN
    Ligne := Format(FormatBapVisa, [Q1.FindField ('CTI_CODEVISA').asstring,
      Q1.FindField ('CTI_LIBELLE').asstring, Q1.FindField ('CTI_ABREGE').asstring,
      IntTostr(Q1.FindField ('CTI_NBVISA').asinteger), Q1.FindField ('CTI_NATUREPIECE').asstring,
      AlignDroite(StrfMontant(Q1.FindField('CTI_MONTANTMIN').AsFloat,20,4,'',False), 20),
      AlignDroite(StrfMontant(Q1.FindField('CTI_MONTANTMAX').AsFloat,20,4,'',False), 20),
      Q1.FindField ('CTI_ETABLISSEMENT').asstring, Q1.FindField ('CTI_TYPELIBRE1').asstring,
      Q1.FindField ('CTI_TYPELIBRE2').asstring, Q1.FindField ('CTI_TYPELIBRE3').asstring,
      Q1.FindField ('CTI_CODELIBRE1').asstring, Q1.FindField ('CTI_CODELIBRE2').asstring,
      Q1.FindField ('CTI_CODELIBRE3').asstring, Q1.FindField ('CTI_TEXTELIBRE1').asstring,
      Q1.FindField ('CTI_TEXTELIBRE2').asstring, Q1.FindField ('CTI_TEXTELIBRE3').asstring,
      Q1.FindField ('CTI_AXE1').asstring, Q1.FindField ('CTI_AXE2').asstring,
      Q1.FindField ('CTI_AXE3').asstring, Q1.FindField ('CTI_CIRCUITBAP').asstring,
      Q1.FindField ('CTI_COMPTE').asstring, Q1.FindField ('CTI_EXCLUSION').asstring]);
    writeln(F, Ligne);
    Q1.NExt ;
 END;

end;


Procedure EcritureBapCircuit(var F: TextFile);
var
Ligne     : string;
Q1        : TQuery;
begin
 Q1:=OpenSQL(' SELECT * FROM CPCIRCUIT ',TRUE) ;
 While Not Q1.Eof Do BEGIN
    Ligne := Format(FormatBapCircuit, [Q1.FindField ('CCI_CIRCUITBAP').asstring,
      IntTostr(Q1.FindField ('CCI_NUMEROORDRE').asinteger), Q1.FindField ('CCI_VISEUR1').asstring,
      Q1.FindField ('CCI_VISEUR2').asstring, IntToStr(Q1.FindField ('CCI_NBJOUR').asinteger),
      Q1.FindField ('CCI_LIBELLE').asstring]);
    writeln(F, Ligne);
    Q1.NExt ;
 END;

end;


procedure EcritureVentilType (var F: TextFile);
var
Qe        : TQuery;
Ligne     : string;
begin
          // ventilation type
        Qe := Opensql('SELECT * FROM  CHOIXCOD WHERE CC_TYPE="VTY"', TRUE);
        while not Qe.EOF do
        begin
          Ligne := Format(FormatChoixCode,
          ['CTL', Qe.Findfield('CC_TYPE').asstring, Qe.Findfield('CC_CODE').asstring, Qe.Findfield('CC_LIBELLE').asstring,
             Qe.Findfield('CC_ABREGE').asstring, Qe.Findfield('CC_LIBRE').asstring]);
          writeln(F, Ligne);
          Qe.Next;
        end;
        Ferme(Qe);
        Qe := Opensql('SELECT * FROM  VENTIL,CHOIXCOD Where V_NATURE like "TY%" and CC_TYPE="VTY" and CC_CODE=V_COMPTE', TRUE);
        while not Qe.EOF do
        begin
          Ligne := Format(FormatVentilType,
          [Qe.Findfield('V_NATURE').asstring, Qe.Findfield('V_COMPTE').asstring, Qe.Findfield('V_SECTION').asstring,
             AlignDroite(StrfMontant(Qe.FindField('V_TAUXMONTANT').AsFloat,20,4,'',False),20), AlignDroite(StrfMontant(Qe.FindField('V_TAUXQTE1').AsFloat,20,4,'',False),20),
             AlignDroite(StrfMontant(Qe.FindField('V_TAUXQTE2').AsFloat,20,4,'',False),20),
             IntTostr(Qe.Findfield('V_NUMEROVENTIL').asinteger), Qe.Findfield('V_SOCIETE').asstring,
             AlignDroite(StrfMontant(Qe.FindField('V_MONTANT').AsFloat,20,4,'',False),20),
             Qe.Findfield('V_SOUSPLAN1').asstring,Qe.Findfield('V_SOUSPLAN2').asstring,
             Qe.Findfield('V_SOUSPLAN3').asstring,Qe.Findfield('V_SOUSPLAN4').asstring,
             Qe.Findfield('V_SOUSPLAN5').asstring,Qe.Findfield('V_SOUSPLAN6').asstring]);
          writeln(F, Ligne);
          Qe.Next;
        end;
        Ferme(Qe);
        // fiche 10629
        Qe := Opensql('SELECT * FROM  VENTIL Where V_NATURE like "GE%" ORDER BY V_COMPTE,V_NATURE ', TRUE);
        while not Qe.EOF do
        begin
          Ligne := Format(FormatVentilType,
          [Qe.Findfield('V_NATURE').asstring, Qe.Findfield('V_COMPTE').asstring, Qe.Findfield('V_SECTION').asstring,
             AlignDroite(StrfMontant(Qe.FindField('V_TAUXMONTANT').AsFloat,20,4,'',False),20), AlignDroite(StrfMontant(Qe.FindField('V_TAUXQTE1').AsFloat,20,4,'',False),20),
             AlignDroite(StrfMontant(Qe.FindField('V_TAUXQTE2').AsFloat,20,4,'',False),20),
             IntTostr(Qe.Findfield('V_NUMEROVENTIL').asinteger), Qe.Findfield('V_SOCIETE').asstring,
             AlignDroite(StrfMontant(Qe.FindField('V_MONTANT').AsFloat,20,4,'',False),20),
             Qe.Findfield('V_SOUSPLAN1').asstring,Qe.Findfield('V_SOUSPLAN2').asstring,
             Qe.Findfield('V_SOUSPLAN3').asstring,Qe.Findfield('V_SOUSPLAN4').asstring,
             Qe.Findfield('V_SOUSPLAN5').asstring,Qe.Findfield('V_SOUSPLAN6').asstring]);
          writeln(F, Ligne);
          Qe.Next;
        end;
        Ferme(Qe);

end;

{$IFDEF CERTIFNF}
procedure EcritureSuiviValidation (var F: TextFile);
var
Qe        : TQuery;
Ligne     : string;
begin
  // Suivi des Validations
  Qe := Opensql('SELECT * FROM CPJALVALIDATION', TRUE);
  while not Qe.EOF do
  begin
     Ligne := Format(FormatSuiviValidation,
             [IntToStr(Qe.Findfield('CPV_SESSION').AsInteger),
              FormatDateTime(Traduitdateformat('ddmmyyyy'),Qe.FindField ('CPV_DATE').AsDateTime),
              Qe.Findfield('CPV_UTILISATEUR').Asstring,
              IntToStr(Qe.Findfield('CPV_IDDEBUTVAL').AsInteger),
              IntToStr(Qe.Findfield('CPV_IDFINVAL').AsInteger),
              IntToStr(Qe.Findfield('CPV_NUMEVENT').AsInteger)]);
     WriteLN(F, Ligne);
     Qe.Next;
  end;
  Ferme(Qe);
end;
{$ENDIF}

end.

