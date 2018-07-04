{***********UNITE*************************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... : 01/03/2002
Description .. :
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
{
PT1-1  : 04/09/2001 VG V547 On alimente aussi la base 1420 comme la base 1060
   -2  : 04/09/2001 VG V547 Ajout de cumuls par défaut
PT2-1  : 07/09/2001 VG V547 initialisation de la coche sélectionné au niveau du
                            salarié sur l'ouverture de la fiche
   -2  : 07/09/2001 VG V547 La fiche FCUMULS s'appelle CUMULS
   -3  : 07/09/2001 VG V547 Réponse NON à la question "Certains cumuls vont être
                            perdus" : il faut garder les cumuls sélectionnés
   -4  : 10/09/2001 VG V547 Quand le code salarié n'est pas numérique, on
                            interdit d'inclure le code établissement dans le
                            numéro du salarié
   -5  : 12/09/2001 VG V547 Même si cela ne fait pas partie du cahier des
                            charges TDS, on considère que les civilités autre
                            que MME et MLE équivalent à Monsieur
PT3-1  : 12/10/2001 VG V562 Quand le code salarié n'est pas numérique dans le
                            fichier TDS, il ne faut pas faire de ColleZerodevant
PT4-1  : 08/11/2001 VG V563 Quand on choisit "Sans changement du code" et que le
                            code salarié est à blanc dans le fichier TDS, il ne
                            faut pas faire de ColleZerodevant
   -2  : 09/11/2001 VG V563 Quand on choisit "Renumérotation" pour un fichier
                            TDS, on prend un index composé de code
                            établissement, code salarié, code SS (cas des codes
                            salariés à blanc)
PT5-1  : 26/11/2001 VG V563 La tablette de la catégorie DADS (utilisée désormais
                            uniquement pour la DUCS) a changé . Le champ
                            retourné est désormais CO_CODE au lieu de CO_ABREGE
PT6    : 28/11/2001 SB V563 Pour ORACLE, Le champ PCN_CODETAPE doit être
                            renseigné
PT7    : 04/01/2002 VG V571 Gestion de l'aide
PT8    : 28/01/2002 VG V571 Sous ORACLE, la différence entre 2 doubles
                            identiques donnait une différence
PT9    : 30/01/2002 VG V571 La lecture des enregistrements salariés doit se
                            faire 3 ligne à la fois
PT10   : 31/01/2002 VG V571 Changement du nom du fichier d'import :
                            IS2PGIXXXXX.SAV => PGIXXXXX.SAV
PT11   : 01/03/2002 VG V571 Suppression de champs dans la fiche salarié -
                            Fiche qualité N° 10036
PT12   : 21/05/2002 VG V582 Evolution de l'AGL : Sur les modifications
                            d'enregistrements par un tob.insertorupdatedb, les
                            champs _datemodif et _utilisateur ne sont plus gérés
                            automatiquement
PT13-1 : 23/09/2002 VG V585 Alimentation différente des bases de cotisation de
                            reprise selon la catégorie DADS . Cadre + Dirigeant:
                            on alimente les rubriques 1060, 1062 et 1420 sinon
                            rubriques 1040 et 1042
PT13-2 : 25/09/2002 VG V585 Inversion de l'affichage des cumuls : A 1 cumul PGI,
                            on associe maintenant 1 cumul ISIS2
PT14   : 04/10/2002 VG V585 Plantage lors de l'import avec un message
                            "... truncated" : La domiciliation du RIB du salarié
                            était repris sur 25 caractères au lieu de 24 dans
                            PGI
PT15   : 14/11/2002 VG V585 Les cumuls 07 et 08 doivent être repris avec le
                            signe opposé mais pas les autres cumuls
PT16   : 27/01/2003 VG V591 Modification pour ne pas reprendre de données pour
                            des salariés dont l'enregistrement T n'existe pas
PT17   : 26/02/2003 VG V_42 Dans le cas ou le fichier d'import ne contient pas
                            de salariés (déjà importés) mais des cumuls ou des
                            congés payés, ces données n'étaient pas reprises
                            avec le bon code salarié - FQ N° 10516
PT18   : 27/02/2003 VG V_42 On interdit les caractères '"_%*? dans le libellé de
                            la domiciliation bancaire - FQ N° 10519
PT19   : 28/02/2003 VG V_42 L'affichage des cumuls n'était pas classé sous
                            Oracle - FQ N° 10520
PT20   : 04/03/2003 VG V_42 On ne doit créer que les conventions collectives
                            impaires - FQ N°10473
PT21   : 05/03/2003 VG V_42 Dans le cas du choix "Sans changement du numéro de
                            salarié", on n'autorise pas l'import s'il y a un
                            risque de doublons ou d'écrasement de données
PT22   : 20/03/2003 VG V_42 Choix de la Réaffectation par ordre chronologique :
                            On ne crée qu'un seul salarié avec comme matricule
                            0000000000 - FQ N°10599
PT23   : 26/03/2003 SB V_42 FQ 10616 Mvt Rep absencesalarié CodeTape topé "..."
PT24-1 : 15/04/2003 VG V_42 Le montant de la tranche 1 de la base CET doit être
                            limitée à 8 fois le montant du plafond - FQ N°10625
PT24-2 : 15/04/2003 VG V_42 Contrôle supplémentaire dans le cas ou les
                            matricules des salariés sont gérés en numérique, le
                            code de l'établissement est alphanumérique et que
                            l'on choisit code établissement + code salarié ou
                            inversemment - FQ N°10625
PT25   : 16/05/2003 VG V_42 Gestion du fichier TRA
PT26   : 13/06/2003 VG V_42 Pour l'import du fichier TDS, prise en compte du cas
                            ou le code du salarié est sur moins de 10 caractères
PT27   : 16/06/2003 PH V_42 Refonte accès V_PGI_env : V_PGI_env.nodossier
                            remplacé par PgRendNoDossier()
PT28   : 27/06/2003 VG V_42 Modification de la date de validité des mouvements
                            CPA de l'en-cours : date de début d'exercice au lieu
                            de date du jour de l'import - FQ N°10739
PT29-1 : 04/08/2003 VG V_42 Pour le fichier TRA, des informations PSB et PSC
                            n'étaient pas reprises
PT29-2 : 04/08/2003 VG V_42 PSA_MODEREGLE remplacé par PSA_PGMODEREGLE
PT30-1 : 19/09/2003 VG V_42 Contrôles sur les congés payés avant la reprise des
                            données
PT30-2 : 19/09/2003 VG V_42 Reprise des informations prud'hommales en
                            personnalisé si elles sont renseignées, sinon idem
                            établissement - FQ N° 10766
PT31   : 25/09/2003 SB V_42 Affectation ETB new champ PSA_TYPACTIVITE
PT32   : 25/09/2003 VG V_42 Les montants des cumuls de reprise ne sont pas pris
                            en compte dans les cumuls PGI
PT33   : 24/11/2003 VG V_50 Correction du PT30-1 pour ne poser qu'une seule fois
                            la question
PT34   : 27/11/2003 VG V_50 Correction pour reprise fichier TDS
PT35   : 15/12/2003 VG V_50 Correction pour les plafond des bases de cotisation
                            qui n'avaient pas les bonnes valeurs
PT36   : 27/01/2004 VG V_50 Correction pour Oracle : "conversion de type variant
                            incorrect" si pas de convention paramétrée au niveau
                            de l'établissement - FQ N°11073
PT37   : 10/01/2004 VG V_50 Ne pas permettre le traitement si le type de
                            matricule (numérique ou alphanumérique) n'est pas
                            renseigné dans les Paramètres société - FQ N°11073
PT38-1 : 20/02/2004 VG V_50 Message si le matricule est supérieur à 2147483647
                            FQ N°11013
PT38-2 : 20/02/2004 VG V_50 Message pour inciter l'utilisateur à vérifier les
                            cumuls - FQ N°11047
PT39   : 23/02/2004 VG V_50 Message interdisant l'import si certains paramètres
                            sociétés ne sont pas renseignés - FQ N°11092
PT40   : 02/04/2004 VG V_50 Modification du sens de l'import du nombre de mois
                            d'ancienneté - FQ N°11184
PT41   : 18/05/2004 VG V_50 Complément du PT38-1 - FQ N°11312
PT42-1 : 24/05/2004 SB V_50 FQ 11237 Affectation ETB champ PSA_CPACQUISSUPP
PT42-2 : 24/05/2004 SB V_50 FQ 11136 Affectation ETB_CONGESPAYES =>
                            PSA_CONGESPAYES
PT42-3 : 24/05/2004 SB V_50 FQ 10370 Affectation ETB champ PSA_TYPEDITBULCP
PT43   : 11/08/2004 VG V_50 Q.EOF renvoie False pour une requête du type
                            SELECT MAX()
PT45   : 19/08/2004 VG V_50 Récupération de la base prévoyance idem base SS
                            FQ N°11440
PT46-1 : 06/10/2004 VG V_50 Pour récupération fichier TDS, la numérotation était
                            erronnée dans le cas de numérotation étab+matricule
                            ou matricule+étab - FQ N°11540
PT46-2 : 06/10/2004 VG V_50 Dans le fichier TDS, l'adresse du salarié n'est
                            peut-être pas bien formatée - FQ N°11540
PT47   : 04/11/2004 VG V_60 Contrôle supplémentaire afin de reprendre les bons
                            cumuls - FQ N°11761
PT48   : 22/11/2004 VG V_60 Lorsqu'on gère le dossier en numérique, qu'il existe
                            des matricules alphanumériques dans le fichier TDS
                            et que l'on choisit réaffectation des matricules,
                            l'import n'était pas autorisé - FQ N°11792
PT49   : 02/12/2004 VG V_60 Modification du message d'erreur pour un code
                            alphanumérique importé dans un dossier numérique
                            FQ N°11792
PT50   : 06/12/2004 VG V_60 Modification du traitement de la régul ancienneté
                            (cas ou non renseigné) pour import fichier TRA
PT51-1 : 06/01/2005 VG V_60 Import des nouvelles zones de la fiche du salarié
                            (zones DADS-U) - FQ N°11768
PT51-2 : 06/01/2005 VG V_60 Alimentation de la base de cotisation PGI
                            "1500 Base taxe sur les salaires" avec le montant du
                            cumul ISIS "29 Base taxe sur salaires" - FQ N°11729
PT52   : 03/02/2005 SB V_60 FQ 11692 Initialisation des dates à Idate1900
PT53   : 07/02/2005 VG V_60 correction PT51-2 - FQ N°11729
PT54   : 10/03/2005 PH V_602 rajout email dans la structure du fichier TRA
PT55   : 31/03/2005 PH V_603 Rajout conservation du numéro de matricule origine
                             dans le champ PSA_LIBRE8
PT56   : 13/10/2005 SB V_65 procedure EnregSalDefaut pour valeur par défaut
                            tables salaries
PT57   : 17/01/2006 VG V_65 Modification de l'initialisation des dates : null
                            remplacé par IDate1900 - FQ N°12558
PT58   : 26/01/2006 VG V_65 On alimente PHB_PLAFOND1 et PHB_TRANCHE1 avec la
                            valeur 0 pour la base de cotisation PGI
                            "1500 Base taxe sur les salaires" - FQ N°11729
PT59   : 05/04/2006 VG V_65 Récupération d'un fichier au format DADS-U
                            FQ N°12938
PT60   : 06/04/2006 VG V_65 Correction numérotation pour import TDS - FQ N°12085
PT61   : 02/05/2006 VG V_65 Correction import DADS-U : S41.G01.00.034 dans
                            PSA_TAUXPARTIEL au lieu de PSA_TRAVETRANG
                            FQ N°12938
PT62-2 : 01/09/2006 VG V_70 Import DADS-U : ajout de la gestion du champ
                            PSA_CONGESPAYES - FQ N°13282
PT62-3 : 01/09/2006 VG V_70 Depuis mise en place reprise DADS-U, Reprise
                            ISIS II/TRA ne reprenait plus les CP - FQ N°13417
PT62-4 : 01/09/2006 VG V_70 Gestion de la sortie pour erreur                            
PT63   : 30/11/2006 VG V_70 Correction plantage codes statistiques
PT64   : 01/12/2006 VG V_70 Modification du titre des messages en fonction du
                            type de fichier - FQ N°13311
PT65   : 06/12/2006 SB V_70 FQ 13584 Correction anomalie reprise cp
PT66   : 16/05/2007 VG V_72 Intégration du planning unifié
PT67-1 : 26/06/2007 VG V_72 Ergonomie - FQ N°13117
PT67-2 : 26/06/2007 VG V_72 Optimisation
PT68   : 19/07/2007 VG V_72 Alimentation automatique de la rubrique de base GMP
                            à partir du cumul 70
PT69   : 25/10/2007 VG V_80 Alimentation par défaut du champ psa_typpaievaloms
                            FQ N°14901
PT71  : 10/09/2008 JS      Alimentation par défaut du champ PSA_TYPCONVENTION
                            FQ N°14461
PT72  : 15/09/2008 JS      Message significatif si incohérence des matricules : numérique/alphanumérique
                           FQ N°13312
PT73  : 17/09/2008 JS      Ajouter rapport erreur + modifier création tiers
                           FQ N° 12986
PT74 :  17/09/2008 JS      Alimentation du champ R_PAYS dans la table RIB
                           FQ N° 15759
PT75  : 19/09/2008 JS      Mise a vide du champs nom de jeune fille PSA_NOMJF
                           FQ N° 14902
PT76  : 24/09/2008 JS      calcul automatique clé SS par le programme lors reprise fichier dadsu
                           Rajout du 0 pour obtenir 2 possition  FQ N° 14747
PT77  : 03/10/2008 JS      FQ n°15447 régime SS : si code 2 ou 3, alimenter le code 900 "Autres"
                           Pour l'import de fichier TRA, TDS, ISIS
PT78  : 13/10/2008 JS      FQ n°15288 Création automatique du compte tiers lorsque le matricule est alpha-numérique
}
unit ImportISIS;

interface

uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  SysUtils,
  Classes,
  Controls,
  Forms,
  assist,
  StdCtrls,
  ComCtrls,
  HMsgBox,
  HTB97,
  Hctrls,
  ExtCtrls,
  UTob,
  HPanel,
  UiUtil,
  HEnt1,
  HStatus,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_Main,
{$ELSE}
  MaineAgl,
{$ENDIF}
  Mask,
  EntPaie,
  PgOutils,
  Grids,
  utilPGI,
  ParamSoc,
  Ent1,
  PgOutils2,
  HSysMenu,
  UTobDebug,
  YRessource;

procedure LanceImportISIS();
type
  TFImpIsis = class(TFAssist)
    PFileName: TTabSheet;
    PCodageSal: TTabSheet;
    tedtFichier: TLabel;
    edtFichier: THCritMaskEdit;
    Label1: TLabel;
    FCodeSal: TRadioGroup;
    TLCodage: TLabel;
    PReport: TTabSheet;
    FReport: TListBox;
    PLancer: TTabSheet;
    BLancer: TToolbarButton97;
    Label3: TLabel;
    PCumuls: TTabSheet;
    GCumuls: THGrid;
    FCumulDest: THValComboBox;
    Panel1: TPanel;
    BCumuls: TToolbarButton97;
    FCumul: TCheckBox;
    FTypeFichier: TRadioGroup;
    ChckBxRem: TCheckBox;
    Label_cumul: TLabel;
    RecupRib: TCheckBox;
    HelpBtn: TToolbarButton97;
    function NextPage: TTabSheet; override;
    procedure edtFichierChange(Sender: TObject);
    procedure BLancerClick(Sender: TObject);
    procedure PChange(Sender: TObject); override;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure BCumulsClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FTypeFichierClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
  private { Déclarations privées }
    ListeCodeStat : TStringList ;
    TCP, TEtab, TOB_Cumuls, TRub : TOB ;
    DebutExo, FinExo : TDateTime ;
    DateTest, DateTestPrec, Ouvert, Siren, StMessage, StTitre : string;
    StMessage2 : string; //PT72
    NbMois : integer;
    F, FRapport : TextFile;
    Histo, Premier, Suivant : boolean;
    QRechStat:TQuery;
    function Traitement_ISISII_TRA(TSalaries : Tob) : boolean;
    function Traitement_TDS(TSalaries : TOB) : boolean;
    function Traitement_DADSU(TSalaries : TOB) : boolean;
    procedure BoucleDADSU (var TSalaries : TOB; var S, SSiret : string;
                           CodeInit : string);
    procedure LireFichierDADS(var j:integer;Longueur:integer;var S:string);
    procedure EnregSalDefaut(T:Tob);   { PT56 }
    procedure EnregSalarie(T:Tob;S, Etab:string);
    procedure EnregSalarie2(T:TOB;S,Etab:string);
    procedure EnregSalarie3(T:TOB;S,Etab:string);
    procedure EnregSalTDS(T,TSalPGI:TOB;S,Etab:string);
    procedure SauveRubriqueDADSU (var TSalaries : TOB; var SSiret : string;
                                  Code, Rubrique, Valeur : string);
    procedure EnregSalDADSU(T, TSalPGI : TOB; SIRET : string);
    procedure EnregHistoCum(H,HB:TOB;S,S2,S3,S4,Etab:string);
    procedure CreerTOBFille(HB: TOB; Etab, S, S2, CatDADS, Rub : string);
    procedure EnregStat(St : string);
    procedure EnregConvention(St : string);
    procedure EnregDossier(St : string);
    procedure EnregCP(St : string; TCP : TOB; Etab : string);
    function ControleCP(St, Etab : string) : string;
    procedure EnregHisto(H: TOB; S,Etab: string);
    procedure EnregMiniConv(TMC : TOB; St : string);
    procedure EnregRIB(TM:Tob; S:string; QRechSal:TQuery; Etab : string);
    function ControlSIRETEtab(S : string) : string;
    function EnregEtab(S : string) : string;
    function EnregEtab3(S, Etab:string) : string;
    function EnregEtab3bis(S:string) : string;
    procedure RecupEtab3(S:string);
    procedure ChargeCumuls ;
    procedure CreeFilleCumul(var Mere:TOB;NomFille,CumulPGI,CumulISIS:string);
    procedure VerifCumuls ;
    function ChangeCodeSalarie (T : TOB; Prefixe : string;
                                PeutCreer : boolean) : boolean;
    Function Verifokok( okok : boolean) : String ;
    procedure InitCodeStat ;
    procedure FinCodeStat ;
    function ChargeSalariesDansTOB : integer;
    procedure ChargeTOBEtab;
    procedure RemplitTOBEtab(S : string; var Etab : string);
    function RecupCompl(SSiret : string; var Etab : string) : string;
    procedure InitCodeSal;
    procedure FinCodeSal;
    function SalarieAbsent (St : string) : boolean;
    function SalarieDansListe (St : string; PeutCreer : boolean;
                               var Code : Integer) : boolean;
    function VerifType ( Deftype : string) : boolean; //PT72

  public { Déclarations publiques }
    Fichier : file;
  end;

var
  frmLitSav: TFImpIsis;
  ListeCodeSal : TStringList;
  Index : integer;
  Num : Longint;

implementation
uses TiersUtil,
{$IFDEF EAGLCLIENT}
MenuOLX;
{$ELSE}
MenuOLG;
{$ENDIF EAGLCLIENT}
{$R *.DFM}


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Fonction principale de l'import
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure LanceImportISIS();
var
PP: THpanel;
begin
frmLitSav:=TFImpIsis.Create(Application);
PP:=FindInsidePanel ;
if PP=Nil then
   begin
        try
        frmLitSav.ShowModal;
        finally
        frmLitSav.Free;
        end;
   end
else
   begin
   InitInside(frmLitSav,PP) ;
   frmLitSav.Show ;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Changement de page
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpIsis.NextPage : TTabSheet;
var
St : string;
begin
result := inherited NextPage;
if P.ActivePage=PFileName then
   begin
   if (edtFichier.text='') then
      result:=nil  ;
   if not FileExists(edtFichier.text) then
      result:=nil ;

   if ((VH_Paie.PgTypeNumSal <> 'NUM') and (VH_Paie.PgTypeNumSal <> 'ALP')) then
      begin
      St:='Le type de matricule (numérique ou alphanumérique) n''est pas#10#13'+
          'renseigné dans les Paramètres Société.#10#13'+
          'Vous devez renseigner le type de matricule dans les#10#13'+
          'Paramètres société. L''import des données est abandonné.';
      PGIBox(St, StTitre);     //PT64
      result:=nil ;
      end;

   if ((GetParamSoc('SO_PGRACINEAUXI')='') or (VH^.Cpta[fbAux].Lg=0) or
      (VH^.Cpta[fbAux].Cb='')) then
      begin
      St:= 'La racine tiers associée, la longueur des comptes auxilaires#10#13'+
           'ou le caractère de bourrage du compte auxiliaire ne sont pas#10#13'+
           'paramétrés sur ce dossier. Vous devez les renseigner dans#10#13'+
           'les Paramètres Société pour pouvoir faire l''import :#10#13'+
           'Menu "Paramètres - Dossier - Paramètres société"#10#13'+
           'Menu "Paramètres - Comptabilité - Paramètres comptables"#10#13';
      PGIBox(St, StTitre);   //PT64
      result:=nil ;
      end;
   END;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Changement du type de fichier (ISIS2/TDS)
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.edtFichierChange(Sender: TObject);
begin
inherited;
PChange(Nil) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Click sur le bouton "Lancer l'importation" : récupération des
Suite ........ : données
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.BLancerClick(Sender: TObject);
var
StNum, StSal, StStat : string;
QRechSal : TQuery;
TSalaries : TOB;
reponse : integer;
ContinueT : boolean;
begin
BLancer.Enabled:= FALSE;
BSuivant.Enabled:= FALSE;
BPrecedent.Enabled:= FALSE;
ContinueT:= True;    //PT62-4
Index := 0;
Histo := False;
{PT62-3
RepCP:='0';
}

// Recup du fichier de sauvegarde ...
InitCodeSal ;
InitCodeStat ;
AssignFile(F, edtFichier.Text);
Reset(F);

TSalaries := TOB.Create ('SALARIES', nil,-1);

// Si on gère les codes salariés numériquement, on récupère le numéro le + élevé
if (VH_Paie.PgTypeNumSal = 'NUM') then
   begin
   StSal:= 'SELECT MAX(PSA_SALARIE) AS MAXSAL'+
           ' FROM SALARIES';
   QRechSal:=OpenSQL(StSal,TRUE);
   if (not QRechSal.EOF) then
{PT67-2
      try
      Num := QRechSal.FindField('MAXSAL').AsInteger;
      except
            on E: EConvertError do
               Num:=0;
      end
}
      begin
      StNum := QRechSal.FindField('MAXSAL').AsString;
      if (StNum<>'') then
         Num:= StrToInt (StNum)
      else
         Num:= 0;
      end
//FIN PT67-2
   else
      Num:= 0;
   Ferme(QRechSal);
   end;

//Sélection des codes statistiques existant déjà
StStat:= 'SELECT CC_CODE, CC_ABREGE'+
         ' FROM CHOIXCOD WHERE'+
         ' CC_TYPE = "PSQ"';
QRechStat:=OpenSQL(StStat,TRUE) ;

//PT59
//Cas d'un fichier provenant d'Isis 2 ou format TRA
if ((FTypeFichier.ItemIndex = 0) or (FTypeFichier.ItemIndex = 2)) then
   ContinueT:= Traitement_ISISII_TRA (TSalaries);           //PT62-4

//Cas d'un fichier TDS
if (FTypeFichier.ItemIndex = 1) then
   ContinueT:= Traitement_TDS(TSalaries);                   //PT62-4

//Cas d'un fichier DADS-U
if (FTypeFichier.ItemIndex = 3) then
   ContinueT:= Traitement_DADSU(TSalaries);                 //PT62-4
//FIN PT59

Ferme(QRechStat) ;
FinCodeSal ;
FinCodeStat ;
bSuivantClick(Nil) ;

{PT59
if ((FTypeFichier.ItemIndex <> 1) and (FCumul.Checked)) then
}
if (((FTypeFichier.ItemIndex = 0) or (FTypeFichier.ItemIndex = 2)) and
   (FCumul.Checked) and (ContinueT = True)) then             //PT62-4
   begin
   StMessage:= 'Important : vous devez vérifier les cumuls des salariés#13#10'+
               'AVANT de commencer la saisie des bulletins de paie.#13#10'+
               'Voulez-vous les éditer immédiatement ?';
   reponse := PGIAsk(StMessage, StTitre);            //PT64
   if (reponse=mrYes) then
      AglLanceFiche ('PAY','FICHECUMSAL', '', '' , '' );

   StMessage:= 'Important : vous devez vérifier les bases de cotisation#13#10'+
               'AVANT de commencer la saisie des bulletins de paie.#13#10'+
               'Voulez-vous les éditer immédiatement ?';
   reponse := PGIAsk(StMessage, StTitre);            //PT64
   if (reponse=mrYes) then
      AglLanceFiche ('PAY','BASECOT_ETAT', '', '' , '' );
   StMessage:='Traitement terminé';
   end
else
   if (ContinueT = True) then      //PT62-4
      begin
      StMessage:='Traitement terminé';
      PGIBox(StMessage, StTitre);     //PT64
      end;
FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now));//PT73
Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
CloseFile(FRapport);
FreeAndNil(TRub);
FreeAndNil(TSalaries);
FreeAndNil(TEtab);
end;

//PT59
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/04/2006
Modifié le ... :   /  /
Description .. : Traitement concernant le fichier ISIS II ou TRA
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpIsis.Traitement_ISISII_TRA(TSalaries : Tob) : boolean;
var
compl, Drapeau, Etab, FileName, Id, RepCP, S, S1, S2, S3, S4, SSiret : string;
StCot, STRA, StSal, StSalBis, StTyp : string;
THistoBull, THistoCum, TMinConv, TRIB : TOB;
QRechCot, QRechSal : TQuery;
NbEnreg : integer;
ContinueT, ReadFait : boolean;
begin
RepCP:='0';        //PT62-3
ReadFait:= False;
ContinueT:= True;  //PT62-4
//Création ou modification du fichier .log
if (FTypeFichier.ItemIndex = 0) then
   FileName := 'is2pgi.log'
else
   FileName := 'trapgi.log';
AssignFile(FRapport, FileName);
if FileExists(FileName) then
   Append(FRapport)
else
   ReWrite(FRapport);
Writeln(FRapport, '');
Writeln(FRapport, 'Début de traitement : '+DateTimeToStr(Now));
FReport.Items.Add('Début de traitement : '+DateTimeToStr(Now));//PT73
//Création des TOB
THistoCum:= TOB.Create ('HISTOCUMSAL', nil,-1);
TMinConv := TOB.Create ('Mère MINIMUMCONVENT', nil,-1);
TCP := TOB.Create('Mère CP', nil, -1);
TRIB := TOB.Create ('Mère RIB', nil,-1);

//TOB pour la reprise des bases de cotisation
THistoBull:= TOB.Create ('Mère HISTOBULLETIN', nil,-1);
StCot:= 'SELECT PCT_RUBRIQUE, PCT_LIBELLE, PCT_THEMECOT, PCT_ORGANISME,'+
        ' PCT_BASEIMP, PCT_TXSALIMP, PCT_TXPATIMP, PCT_TRANCHE1, PCT_TRANCHE2,'+
        ' PCT_TRANCHE3, PCT_TYPETRANCHE1, PCT_TYPETRANCHE2, PCT_TYPETRANCHE3,'+
        ' PCT_ORDREETAT'+
        ' FROM COTISATION WHERE '+
        ' ##PCT_PREDEFINI## (PCT_RUBRIQUE = "1000" OR'+
        ' PCT_RUBRIQUE = "1020" OR'+
        ' PCT_RUBRIQUE = "1040" OR'+
        ' PCT_RUBRIQUE = "1042" OR'+
        ' PCT_RUBRIQUE = "1060" OR'+
        ' PCT_RUBRIQUE = "1070" OR'+
        ' PCT_RUBRIQUE = "1062" OR'+
        ' PCT_RUBRIQUE = "1080" OR'+
        ' PCT_RUBRIQUE = "1420" OR'+
        ' PCT_RUBRIQUE = "1500")';
QRechCot:=OpenSQL(StCot,TRUE);
TRub := TOB.Create ('Les RUBRIQUES', nil,-1);
TRub.LoadDetailDB('COTISATION','','',QRechCot,False);
Ferme(QRechCot);

//Lecture séquentielle du fichier d'import
Readln(F, S);

//Drapeau de début
if ((FTypeFichier.ItemIndex = 0) and (S[1]='*')) then
   begin
   Drapeau := Copy(S, 1, 11);
   if (Drapeau <> '***DEBUT***') then
      BEGIN
      Ferme(QRechStat) ;
      StMessage := 'Erreur drapeau de début';
      PGIBox(StMessage, StTitre);  //PT64
      FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now)); //PT73
      Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
      ContinueT:= False;            //PT62-4
      END;
   end;
Reset(F) ;
NbEnreg:=0;
while ((not Eof(F)) and (ContinueT=True)) do //premier passage            PT62-4
      begin
      NbEnreg:=NbEnreg+1;
      Readln(F, S);
      if Trim(S)='' then
         continue ;

      if (FTypeFichier.ItemIndex = 0) then
         begin
         case S[1] of
              'C': EnregDossier(S) ;
              'D': begin //Etablissement 1
                   Etab := EnregEtab (S);
                   if (Etab = '-1') then
                      BEGIN
                      Ferme(QRechStat) ;
                      StMessage:='Traitement annulé';
                      PGIBox(StMessage, StTitre);    //PT64
                      FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now)); //PT73
                      Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
                      ContinueT:= False;            //PT62-4
                      END;
                   end;
              'F': begin //Etablissement 3
                   compl := EnregEtab3(S, Etab) ;
                   if (Compl = '-1') then
                      BEGIN
                      Ferme(QRechStat) ;
                      StMessage:='Traitement annulé';
                      PGIBox(StMessage, StTitre);    //PT64
                      FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now)); //PT73
                      Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
                      ContinueT:= False;            //PT62-4
                      END;
                   end;
              'V': begin //Historique mois par mois
                   Histo := True;
                   end;
              'T': BEGIN //Salarié
                   S1:=S ;
                   Readln(F, S);
                   S1:=S1+S ;
                   Readln(F, S);
                   S1:=S1+S ;
                   END ;
              'x': BEGIN //Compléments Salarié
                   S1:=S ;
                   Readln(F, S);
                   S1:=S1+S ;
                   Readln(F, S);
                   S1:=S1+S ;
                   END ;
              'u': BEGIN //Congés payés
                   if (RepCP = '0') then
                      begin
                      RepCP := ControleCP(S, Etab);
                      if (RepCP = '-1') then
                         BEGIN
                         Ferme(QRechStat) ;
                         StMessage:='Traitement annulé';
                         PGIBox(StMessage, StTitre);    //PT64
                         FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now)); //PT73
                         Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
                         ContinueT:= False;            //PT62-4
                         END;
                      end;
                   END;
              end;
         end
      else
         begin
         Id := Copy(S, 1, 6);
         if (Id='***PAA') then
            begin //Dossier
            STRA:=Copy(S, 6, 31);
            EnregDossier(STRA) ;
            end;
         if (Id='***PEU') then
            begin //Etablissement 1
            STRA:=Copy(S, 6, 216);
            Etab := EnregEtab (STRA);
            if (Etab = '-1') then
               BEGIN
               Ferme(QRechStat) ;
               StMessage:='Traitement annulé';
               PGIBox(StMessage, StTitre);    //PT64
               FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now)); //PT73
               Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
               ContinueT:= False;            //PT62-4
               END;
            end;
         if (Id='***ETB') then
            begin //Etablissement 3
            STRA:=Copy(S, 20, 100);
            RemplitTOBEtab(S, Etab);
            compl := EnregEtab3(STRA, Etab) ;
            if (Compl = '-1') then
               BEGIN
               Ferme(QRechStat) ;
               StMessage:='Traitement annulé';
               PGIBox(StMessage, StTitre);    //PT64
               FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now)); //PT73
               Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
               ContinueT:= False;            //PT62-4
               END;
            STRA:=Copy(S, 7, 113);
            end;

         if (Id='***PCN') then
            BEGIN
            if (RepCP = '0') then
               begin
               STRA:=Copy(S, 20, 177);
               SSiret:=Copy(S, 7, 14);
               compl := RecupCompl(SSiret, Etab);
               if (Compl <> '-2') then
                  begin
                  RepCP := ControleCP(STRA, Etab);
                  if (RepCP = '-1') then
                     BEGIN
                     Ferme(QRechStat) ;
                     StMessage:='Traitement annulé';
                     PGIBox(StMessage, StTitre);    //PT64
                     FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now)); //PT73
                     Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
                     ContinueT:= False;            //PT62-4
                     END;
                  end;
               end;
            END;
         end;
      end;
Reset(F) ;
NbEnreg:=NbEnreg*2;
InitMove(NbEnreg,'') ;
while ((not Eof(F)) and (ContinueT=True)) do //Deuxieme passage           PT62-4
      begin
      MoveCur(False) ;
      Readln(F, S);
      if Trim(S)='' then
         continue ;
      if (FTypeFichier.ItemIndex = 0) then
         begin
         case S[1] of
              'D': begin //Etablissement 1
                   Etab := EnregEtab (S);
                   if (Etab = '-1') then
                      BEGIN
                      Ferme(QRechStat) ;
                      StMessage:='Traitement annulé';
                      PGIBox(StMessage, StTitre);    //PT64
                      FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now)); //PT73
                      Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
                      ContinueT:= False;            //PT62-4
                      END;
                   end;
              'F': begin //Etablissement 3
                   compl := EnregEtab3bis(S) ;
                   end;
              'N': begin //Codes Statistiques
{PT63
                   EnregStat(S, QRechStat);
}
                   EnregStat(S);
                   end;
              'P': begin //Conventions collectives
                   if (Compl <> '-2') then
                      EnregConvention(S);
                   end;
              'T': BEGIN //Salarié
                   S1:=S ;
                   Readln(F, S);
                   S1:=S1+S ;
                   Readln(F, S);
                   S1:=S1+S ;
                   if (Compl <> '-2') then
                      EnregSalarie(TSalaries,S1, Etab);
                   END ;
              'X': begin //Table de qualification (minimum conventionnel)
                   EnregMiniConv(TMinConv, S);
                   end;
              'x': BEGIN //Compléments Salarié
                   S1:=S ;
                   Readln(F, S);
                   S1:=S1+S ;
                   Readln(F, S);
                   S1:=S1+S ;
                   if (Compl <> '-2') then
                      EnregSalarie2(TSalaries,S1,Etab);
                   END ;
              '&': begin //Code catégorie
                   end;
              end;
         end
      else
         begin
         Id := Copy(S, 1, 6);
         if (Id='***PCC') then
            begin //Codes Statistiques
            STRA:=Copy(S, 6, 30);
{PT63
            EnregStat(STRA, QRechStat);
}
            EnregStat(STRA);
            end;
         if (Id='***PCV') then
            begin //Conventions collectives
            STRA:=Copy(S, 20, 204);
            SSiret:=Copy(S, 7, 14);
            compl := RecupCompl(SSiret, Etab);
            if (Compl <> '-2') then
               EnregConvention(STRA);
            end;
         if (Id='***PSA') then
            BEGIN //Salarié
            STRA:=Copy(S, 20, 529);
            SSiret:=Copy(S, 7, 14);
            compl := RecupCompl(SSiret, Etab);
            if (Compl <> '-2') then
               EnregSalarie(TSalaries,STRA, Etab);
            END ;
         if (Id='***PMI') then
            begin //Table de qualification (minimum conventionnel)
            STRA:=Copy(S, 6, 30);
            EnregMiniConv(TMinConv, STRA);
            end;
         if (Id='***PSB') then
            BEGIN //Compléments Salarié
            STRA:=Copy(S, 20, 60);
            SSiret:=Copy(S, 7, 14);
            compl := RecupCompl(SSiret, Etab);
            if (Compl <> '-2') then
               EnregSalarie2(TSalaries,STRA,Etab);
            END ;
         if (Id='***PSC') then
            BEGIN //Compléments Salarié
            STRA:=Copy(S, 20, 350);
            SSiret:=Copy(S, 7, 14);
            compl := RecupCompl(SSiret, Etab);
            if (Compl <> '-2') then
               EnregSalarie3(TSalaries,STRA,Etab);
            END ;
         end;
      end;
Reset(F) ;
StSal:= 'SELECT PSA_SALARIE, PSA_AUXILIAIRE'+
        ' FROM SALARIES';
QRechSal:=OpenSql(StSal,TRUE);
while ((not Eof(F)) and (ContinueT=True)) do //Troisième passage  PT62-4
      begin
      MoveCur(False) ;
      if (ReadFait=False) then
         Readln(F, S);
      ReadFait:= False;
      if Trim(S)='' then
         continue ;
      if (FTypeFichier.ItemIndex = 0) then
         begin
         case S[1] of
              'D': begin //Etablissement
                   Etab := EnregEtab (S);
                   if (Etab = '-1') then
                      BEGIN
                      Ferme(QRechStat) ;
                      Ferme(QRechSal) ;
                      StMessage:='Traitement annulé';
                      PGIBox(StMessage, StTitre);    //PT64
                      FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now)); //PT73
                      Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
                      ContinueT:= False;            //PT62-4
                      END;
                   end;
              'F': begin //Etablissement 3
                   compl := EnregEtab3bis(S) ;
                   end;
              'S': BEGIN //Cumuls individuels salariés
                   S1:= ''; S2:= ''; S3:= ''; S4:= '';
                   StSal:= Copy (S, 2, 6);
                   StTyp:= Copy (S, 8, 1);
                   if (StTyp='I') then
                      begin
                      S1:=S ; Readln(F, S); S1:=S1+S ;Readln(F, S); S1:=S1+S ;
                      ReadFait:= False;
                      end
                   else
                      ReadFait:= True;
                   if (ReadFait=False) then
                      Readln(F, S)
                   else
                      StSal:= Copy (S, 2, 6);
                   StSalBis:= Copy (S, 2, 6);
                   StTyp:= Copy (S, 8, 1);
                   if ((StSal=StSalBis) and (StTyp='J')) then
                      begin
                      S2:= S; Readln(F, S); S2:=S2+S ;Readln(F, S); S2:=S2+S ;
                      ReadFait:= False;
                      end
                   else
                      ReadFait:= True;
                   if (ReadFait=False) then
                      Readln(F, S)
                   else
                      StSal:= Copy (S, 2, 6);
                   StSalBis:= Copy (S, 2, 6);
                   StTyp:= Copy (S, 8, 1);
                   if ((StSal=StSalBis) and (StTyp='R')) then
                      begin
                      S3:= S; Readln(F, S); S3:=S3+S ;Readln(F, S); S3:=S3+S ;
                      ReadFait:= False;
                      end
                   else
                      ReadFait:= True;
                   if (ReadFait=False) then
                      Readln(F, S)
                   else
                      StSal:= Copy (S, 2, 6);
                   StSalBis:= Copy (S, 2, 6);
                   StTyp:= Copy (S, 8, 1);
                   if ((StSal=StSalBis) and (StTyp='S')) then
                      begin
                      S4:= S; Readln(F, S); S4:=S4+S ;Readln(F, S); S4:=S4+S ;
                      ReadFait:= False;
                      end
                   else
                      ReadFait:= True;
                   if ((Compl <> '-2') and (FCumul.Checked)) then
                      EnregHistoCum(THistoCum,THistoBull,S1,S2,S3,S4,Etab) ;
                   END ;
              'T': BEGIN //RIB
                   S1:=S ;
                   Readln(F, S);
                   S1:=S1+S ;
                   Readln(F, S);
                   S1:=S1+S ;
                   if (Compl <> '-2') then
                      EnregRIB(TRIB,S1,QRechSal, Etab);
                   END ;
              'V': begin //Historique de paie
                   if (Compl <> '-2') and (FCumul.Checked) then
                      EnregHisto(THistoCum,S,Etab);
                   end;
              'u': begin //Congés Payés
                   if ((Compl <> '-2') and (RepCP>='0')) then
                      EnregCP(S, TCP, Etab);
                   end;
              end;
         end
      else
         begin
         Id := Copy(S, 1, 6);
         if (Id='***PHC') then
            BEGIN //Cumuls individuels salariés
            STRA:=Copy(S, 6, 529);
            SSiret:=Copy(S, 7, 14);
            compl := RecupCompl(SSiret, Etab);
            S1:= ''; S2:= ''; S3:= ''; S4:= '';
            StSal:= Copy (S, 21, 6);
            StTyp:= Copy (S, 27, 1);
            if (StTyp='I') then
               S1:=Copy(S, 20, 658) ;
            StSalBis:= Copy (S, 678, 6);
            StTyp:= Copy (S, 684, 1);
            if ((StSal=StSalBis) and (StTyp='J')) then
               S2:=Copy(S, 677, 645) ;
            StSalBis:= Copy (S, 1322, 6);
            StTyp:= Copy (S, 1328, 1);
            if ((StSal=StSalBis) and (StTyp='R')) then
               S3:=Copy(S, 1321, 658) ;
            StSalBis:= Copy (S, 1979, 6);
            StTyp:= Copy (S, 1985, 1);
            if ((StSal=StSalBis) and (StTyp='S')) then
               S4:=Copy(S, 1978, 645) ;
            if (Compl <> '-2') and (FCumul.Checked) then
               EnregHistoCum(THistoCum,THistoBull,S1,S2,S3,S4,Etab) ;
            END ;
         if (Id='***PSA') then
            BEGIN //RIB
            STRA:=Copy(S, 20, 529);
            SSiret:=Copy(S, 7, 14);
            compl := RecupCompl(SSiret, Etab);
            if (Compl <> '-2') then
               EnregRIB(TRIB,STRA,QRechSal, Etab);
            END ;
         if (Id='***PCN') then
            begin //Congés Payés
            STRA:=Copy(S, 20, 177);
            SSiret:=Copy(S, 7, 14);
            compl := RecupCompl(SSiret, Etab);
            if (Compl <> '-2') then
               EnregCP(STRA, TCP, Etab);
            end;
         end;
      end;
Ferme(QRechSal);
FiniMove ;
CloseFile(F);
FreeAndNil(THistoCum);
TMinConv.SetAllModifie (TRUE);
TMinConv.InsertOrUpdateDb(FALSE);
FreeAndNil(TMinConv);
TCP.SetAllModifie (TRUE);
TCP.InsertOrUpdateDb(FALSE);
FreeAndNil(TCP);
THistoBull.SetAllModifie (TRUE);
THistoBull.InsertOrUpdateDb(FALSE);
FreeAndNil(THistoBull);
TRIB.SetAllModifie (TRUE);
TRIB.InsertOrUpdateDb(FALSE);
FreeAndNil(TRIB);
result := ContinueT;         //PT62-4
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/04/2006
Modifié le ... :   /  /
Description .. : Traitement concernant le fichier TDS
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpIsis.Traitement_TDS(TSalaries : TOB) : boolean;
var
Code, Etab, FileName, Long, S, S1, StSal : string;
QRechSal : TQuery;
TSal : TOB;
Chariot, ContinueT : boolean;
j, Longueur : integer;
begin
ContinueT:= True;     //PT62-4
j:= 1;
FileName := 'dadspgi.log';
AssignFile(FRapport, FileName);
if FileExists(FileName) then
   Append(FRapport)
else
   ReWrite(FRapport);
Writeln(FRapport, '');
Writeln(FRapport, 'Début de traitement : '+DateTimeToStr(Now));
FReport.Items.Add('Début de traitement : '+DateTimeToStr(Now));//PT73

InitMove(1000,'') ;
//Première lecture pour la récupération des paramètres du fichier
Readln(F, S);
//Longueur de l'enregistrement
Long := Copy(S,4,5);
if (Long = '00000') then
   Longueur := 564
else
   Longueur := 128;
// Existe-t-il des retours chariot
if (Eof(F) = TRUE) then
   Chariot := FALSE
else
   Chariot := TRUE;
Reset(F) ;

StSal:='SELECT PSA_SALARIE FROM SALARIES';
QRechSal:=OpenSql(StSal,TRUE);
TSal := TOB.Create('Les salaries', NIL, -1);
TSal.LoadDetailDB('SAL','','',QRechSal,False);
Ferme(QRechSal);

//Cas où les retours chariot sont présents
if (Chariot = TRUE) then
   BEGIN
   while ((not Eof(F)) and (ContinueT=True)) do        //PT62-4
         BEGIN
         MoveCur(False) ;
         Readln(F, S);  //Deuxième lecture pour la récupération

//Lecture du code de l'enregistrement
         if (Longueur = 128) then
            BEGIN
            Code := Copy(S, 28, 3);
            S1 := Copy(S, 9, Longueur-8);
            END
         else
            BEGIN
            Code := Copy(S, 20, 3);
            S1 := S;
            END;
         if (Code = '   ') then
            BEGIN
            StMessage:='Erreur dans le fichier DADS';
            PGIBox(StMessage, StTitre);        //PT64
            FReport.Items.Add(StMessage); //PT73
            Writeln(FRapport, StMessage);
            break;
            END;
         case StrToInt(Code) of
              0: BEGIN
                 END;
              5: BEGIN
                 END;
              10: BEGIN
                  if (Longueur = 128) then
                     BEGIN
                     Readln(F, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     Readln(F, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     Readln(F, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     END;
                  END;
              20: BEGIN  //Etablissement
                  if (Longueur = 128) then
                     BEGIN
                     Readln(F, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     Readln(F, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     Readln(F, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     END;
                  Etab := ControlSIRETEtab(S1);
                  if (Etab = '-1') then
                     BEGIN
                     Ferme(QRechStat) ;
                     StMessage:='Traitement annulé';
                     PGIBox(StMessage, StTitre);    //PT64
                     FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now)); //PT73
                     Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
                     ContinueT:= False;            //PT62-4
                     END;
                  END;
              130: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              200: BEGIN //Salarié
                   if (Longueur = 128) then
                      BEGIN
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   if (Etab <> '-2') then
                      EnregSalTDS(TSalaries,TSal,S1,Etab);
                   END;
              202: BEGIN //Salarié
                   if (Longueur = 128) then
                      BEGIN
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   if (Etab <> '-2') then
                      EnregSalTDS(TSalaries,TSal,S1,Etab);
                   END;
              210: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              300: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              302: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              310: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              312: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      Readln(F, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              990: break; //Fin du fichier en francs
              995: break; //Fin du fichier en euros
              else //Code enregistrement non répertorié
                   BEGIN
                   StMessage:='Erreur dans le fichier DADS';
                   PGIBox(StMessage, StTitre);        //PT64
                   FReport.Items.Add(StMessage); //PT73
                   Writeln(FRapport, StMessage);
                   break;
                   END;
              end;
         END;
   CloseFile(F);
   END
else //Cas où les retours chariot sont absents
   BEGIN
   CloseFile(F);
   AssignFile(Fichier, edtFichier.Text);
   Reset(Fichier, 1);
   while ((not Eof(Fichier)) and (ContinueT=True)) do //PT62-4
         BEGIN
         MoveCur(False) ;
//Lecture pour la récupération sur la longueur d'un enregistrement
         LireFichierDADS(j, Longueur, S);
//Lecture du code de l'enregistrement
         if (Longueur = 128) then
            BEGIN
            Code := Copy(S, 28, 3);
            S1 := Copy(S, 9, Longueur-8);
            END
         else
            BEGIN
            Code := Copy(S, 20, 3);
            S1 := S;
            END;
         if (Code = '   ') then
            BEGIN
            StMessage:='Erreur dans le fichier DADS';
            PGIBox(StMessage, StTitre);        //PT64
            FReport.Items.Add(StMessage); //PT73
            Writeln(FRapport, StMessage);
            break;
            END;
         case StrToInt(Code) of
              0: BEGIN
                 END;
              5: BEGIN
                 END;
              10: BEGIN
                  if (Longueur = 128) then
                     BEGIN
                     LireFichierDADS(j, Longueur, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     LireFichierDADS(j, Longueur, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     LireFichierDADS(j, Longueur, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     END;
                  END;
              20: BEGIN //Etablissement
                  if (Longueur = 128) then
                     BEGIN
                     LireFichierDADS(j, Longueur, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     LireFichierDADS(j, Longueur, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     LireFichierDADS(j, Longueur, S);
                     S1:=S1+Copy(S, 9, Longueur-8) ;
                     END;
                  Etab := ControlSIRETEtab(S1);
                  if (Etab = '-1') then
                     BEGIN
                     Ferme(QRechStat) ;
                     StMessage:='Traitement annulé';
                     PGIBox(StMessage, StTitre);    //PT64
                     FReport.Items.Add(StMessage+' : '+DateTimeToStr(Now)); //PT73
                     Writeln(FRapport, StMessage+' : '+DateTimeToStr(Now));
                     ContinueT:= False;            //PT62-4
                     END;
                  END;
              130: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              200: BEGIN //Salarié
                   if (Longueur = 128) then
                      BEGIN
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   if (Etab <> '-2') then
                      EnregSalTDS(TSalaries,TSal,S1,Etab);
                   END;
              202: BEGIN //Salarié
                   if (Longueur = 128) then
                      BEGIN
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   if (Etab <> '-2') then
                      EnregSalTDS(TSalaries,TSal,S1,Etab);
                   END;
              210: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              300: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              302: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              310: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              312: BEGIN
                   if (Longueur = 128) then
                      BEGIN
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      LireFichierDADS(j, Longueur, S);
                      S1:=S1+Copy(S, 9, Longueur-8) ;
                      END;
                   END;
              990: break; //Fin du fichier en francs
              995: break; //Fin du fichier en euros
              else //Code enregistrement non répertorié
                   BEGIN
                   StMessage:='Erreur dans le fichier DADS';
                   PGIBox(StMessage, StTitre);        //PT64
                   FReport.Items.Add(StMessage); //PT73
                   Writeln(FRapport, StMessage);
                   break;
                   END;
              end;
         END;
   CloseFile(Fichier);
   END;
FiniMove ;
FreeAndNil(TSal);
result:= ContinueT;   //PT62-4
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/04/2006
Modifié le ... :   /  /
Description .. : Traitement concernant le fichier TDS
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpIsis.Traitement_DADSU(TSalaries : TOB) : boolean;
var
Code, FileName, S, SSiret, StSal : string;
QRechSal : TQuery;
TSal : TOB;
ContinueT : boolean;
begin
ContinueT:= True;   //PT62-4
FileName := 'dadsupgi.log';
AssignFile(FRapport, FileName);
if FileExists(FileName) then
   Append(FRapport)
else
   ReWrite(FRapport);
Writeln(FRapport, '');
Writeln(FRapport, 'Début de traitement : '+DateTimeToStr(Now));
FReport.Items.Add('Début de traitement : '+DateTimeToStr(Now)); //PT73
InitMove(1000,'') ;

StSal:='SELECT PSA_SALARIE FROM SALARIES';
QRechSal:=OpenSql(StSal,TRUE);
TSal := TOB.Create('Les salaries', NIL, -1);
TSal.LoadDetailDB('SAL','','',QRechSal,False);
Ferme(QRechSal);
Premier:= True;
Suivant:= False;

while not Eof(F) do
      BEGIN
      if (Suivant = False) then
         begin
         MoveCur(False) ;
         Readln(F, S);  //Lecture pour la récupération
         end;
      Suivant:= False;

//Lecture du code de l'enregistrement
      Code:= Copy (S, 2, 2);
//PT67-2
      if (IsNumeric (Code)) then
         begin
//FIN PT67-2
         case StrToInt(Code) of
              10: BEGIN      //emetteur
                  END;
              20: BEGIN      //déclarant
                  BoucleDADSU (TSalaries, S, SSiret, Code);
                  END;
              30: BEGIN      //salarié
//PT62-2
                  if (TEtab = nil) then
                     ChargeTOBEtab;
//FIN PT62-2
                  if (Premier <> true) then
                     EnregSalDADSU(TSalaries, TSal, SSiret);
                  BoucleDADSU (TSalaries, S, SSiret, Code);
                  END;
              41: BEGIN      //période d'activité
                  BoucleDADSU (TSalaries, S, SSiret, Code);
                  END;
              42: BEGIN      //IRCANTEC
                  END;
              43: BEGIN      //CNRACL
                  END;
              44: BEGIN      //AGIRC-ARRCO
                  END;
              45: BEGIN      //Prévoyance
                  END;
              46: BEGIN      //période d'inactivité
                  END;
              51: BEGIN      //Elements de cotisation
                  END;
              53: BEGIN      //RAFP
                  END;
              66: BEGIN      //CP BTP
                  END;
              70: BEGIN      //Honoraires
                  END;
              80: BEGIN      //Etablissements
                  if (Premier <> true) then
                     EnregSalDADSU(TSalaries, TSal, SSiret);
                  BoucleDADSU (TSalaries, S, SSiret, Code);
                  END;
              90: break; //Fin du fichier
              else //Code enregistrement non répertorié
                  BEGIN
                  StMessage:='Erreur dans le fichier DADS-U';
                  PGIBox(StMessage, StTitre);        //PT64
                  FReport.Items.Add(StMessage); //PT73
                  Writeln(FRapport, StMessage);
                  ContinueT:= False;           //PT62-4
                  break;
                  END;
              end;
//PT67-2
         end
      else
         begin
         StMessage:= 'Le fichier n''est pas au format DADS-U';
         PGIBox (StMessage, StTitre);
         Writeln (FRapport, StMessage);
         ContinueT:= False;
         break;
         end;
//FIN PT67-2
      END;
CloseFile(F);
FiniMove ;
FreeAndNil(TSal);
result:= ContinueT;  //PT62-4
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/04/2006
Modifié le ... :   /  /
Description .. : Boucle sur une structure DADS-U et récupération des éléments
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.BoucleDADSU (var TSalaries : TOB; var S, SSiret : string;
                                 CodeInit : string);
var
Code, Rubrique, Valeur : string;
begin
Code:= CodeInit;

while (Code=CodeInit) do
      begin
      Rubrique:= Copy (S, 1, Pos (',', S)-1);
      Valeur:= Copy (S, Pos ('''', S)+1, Length (S)-Length (Rubrique)-3);
      SauveRubriqueDADSU (TSalaries, SSiret, Code, Rubrique, Valeur);

      MoveCur(False) ;
      Readln(F, S);  //Lecture pour la récupération
      Code := Copy(S, 2, 2);
      if (Code <> CodeInit) then
         suivant:= True;
      end;
end;
//FIN PT59

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/01/2001
Modifié le ... :   /  /
Description .. : Lecture du fichier TDS pour la récupération des données
Suite ........ : sur la longueur d'un enregistrement
Suite ........ : j : position du début de lecture
Suite ........ : Longueur : longueur de l'enregistrement à lire
Suite ........ : S : Chaine de caractères
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.LireFichierDADS(var j:integer;Longueur:integer;var S:string);
var
i, long : integer;
Car:array[1..564] of char;
begin
long := Longueur;

for i := 1 to 564 do
    Car[i] := ' ';

for i := 1 to long do
    BEGIN
    BlockRead(Fichier, Car, 1);
    Seek(Fichier, j);
    j := j+1;
    if (i= 1) and ((Car[1]<'0') or (Car[1]>'9')) and (Car[1]<>' ') then
       BEGIN
       BlockRead(Fichier, Car, 1);
       long := long+1;
       Seek(Fichier, j);
       j := j+1;
       END;
    S[i] := Car[1];
    END;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 13/10/2005
Modifié le ... :   /  /
Description .. : IMPORT ISIS
Mots clefs ... :
*****************************************************************}
{ DEB PT56 }
procedure TFImpIsis.EnregSalDefaut(T: Tob);
begin
T.PutValue('PSA_TYPEDITORG', 'ETB');
T.PutValue('PSA_TYPNBACQUISCP', 'ETB');
T.PutValue('PSA_DATANC', 'ETB');
T.PutValue('PSA_CPACQUISMOIS', 'ETB');
T.PutValue('PSA_CPACQUISSUPP', 'ETB');
T.PutValue('PSA_TYPEDITBULCP', 'ETB');
T.PutValue('PSA_CPACQUISANC', 'ETB');
T.PutValue('PSA_CPTYPEMETHOD', 'ETB');
T.PutValue('PSA_CPTYPERELIQ', 'ETB');
T.PutValue('PSA_CPTYPEVALO', 'ETB');
T.PutValue('PSA_TYPREGLT', 'ETB');
T.PutValue('PSA_TYPDATPAIEMENT', 'ETB');
T.PutValue('PSA_PRISEFFECTIF', 'X');
T.PutValue('PSA_UNITEPRISEFF', 1);
T.PutValue('PSA_TYPACTIVITE', 'ETB'); //PT31
T.PutValue('PSA_TYPPROFIL', 'ETB');
T.PutValue('PSA_TYPPERIODEBUL', 'ETB');
T.PutValue('PSA_TYPPROFILRBS', 'ETB');
T.PutValue('PSA_TYPPROFILAFP', 'ETB');
T.PutValue('PSA_TYPPROFILAPP', 'ETB');
T.PutValue('PSA_TYPPROFILMUT', 'ETB');
T.PutValue('PSA_TYPPROFILPRE', 'ETB');
T.PutValue('PSA_TYPPROFILTSS', 'ETB');
T.PutValue('PSA_TYPPROFILCGE', 'ETB');
T.PutValue('PSA_TYPPROFILANC', 'ETB');
T.PutValue('PSA_TYPPROFILFNAL', 'DOS');
T.PutValue('PSA_TYPPROFILTRANS', 'ETB');
T.PutValue('PSA_TYPPROFILRET', 'ETB');
T.PutValue('PSA_TYPPAIACOMPT', 'ETB');
T.PutValue('PSA_TYPPAIFRAIS', 'ETB');
{PT57
T.PutValue('PSA_DATELIBRE1', null);
T.PutValue('PSA_DATELIBRE2', null);
T.PutValue('PSA_DATELIBRE3', null);
T.PutValue('PSA_DATELIBRE4', null);
T.PutValue('PSA_DATELIBRE5', null);
T.PutValue('PSA_DATELIBRE6', null);
T.PutValue('PSA_DATELIBRE7', null);
T.PutValue('PSA_DATELIBRE8', null);
}
T.PutValue('PSA_DATELIBRE1', IDate1900);
T.PutValue('PSA_DATELIBRE2', IDate1900);
T.PutValue('PSA_DATELIBRE3', IDate1900);
T.PutValue('PSA_DATELIBRE4', IDate1900);
T.PutValue('PSA_DATELIBRE5', IDate1900);
T.PutValue('PSA_DATELIBRE6', IDate1900);
T.PutValue('PSA_DATELIBRE7', IDate1900);
T.PutValue('PSA_DATELIBRE8', IDate1900);
//FIN PT57
T.PutValue('PSA_ORDREAT', '1');
T.PutValue('PSA_CATBILAN', '000');
T.PutValue('PSA_STANDCALEND', 'ETB');
T.PutValue('PSA_TYPJOURHEURE', 'DOS');
T.PutValue('PSA_JOURHEURE', 'ETB');
T.PutValue('PSA_TYPREDREPAS', 'ETB');
T.PutValue('PSA_TYPREDRTT1', 'ETB');
T.PutValue('PSA_TYPREDRTT2', 'ETB');
T.PutValue('PSA_TYPPROFILREM', 'ETB');
T.PutValue('PSA_TYPVIRSOC', 'ETB');
T.PutValue('PSA_TYPACPSOC', 'ETB');
T.PutValue('PSA_TYPFRAISSOC', 'ETB');
T.PutValue('PSA_CONFIDENTIEL', '0');
{PT57
T.PutValue('PSA_DADSDATE', null);
}
T.PutValue('PSA_DADSDATE', IDate1900);
T.PutValue('PSA_TYPDADSFRAC', 'ETB');
T.PutValue('PSA_TOPCONVERT', 'X');
{FIN PT56 }
T.PutValue ('PSA_TYPPAIEVALOMS', 'ETB');    //PT69
T.PutValue ('PSA_TYPCONVENTION', 'ETB');    // PT71
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Enregistrement d'un salarié provenant d'ISIS II (T)
Suite ........ : T : TOB salarié
Suite ........ : S : Chaine de caractères
Suite ........ : Etab : Numéro d'établissement
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregSalarie(T:Tob;S, Etab:string);
Var
okok : boolean ;
InfTiers : Info_Tiers;
CodeAuxi, CodeConv, CodeStat, LeRapport, Libell, Preno, Sexe : String;
Situation, StRech, StRegulAncien : String;
LongLib, LongPre, RegulAncien : integer;
QRech : TQuery;
TEtabD : TOB;
begin
T.InitValeurs;
T.PutFixedStValue('PSA_SALARIE',S,2,6,tcTrim,FALSE);
T.PutFixedStValue('PSA_LIBRE8', S, 2, 6, tcTrim, FALSE); // PT55
T.PutFixedStValue('PSA_NUMEROSS',S,111,15,tcTrim,FALSE);
T.PutFixedStValue('PSA_LIBELLE',S,8,30,tcTrim,FALSE);
T.PutFixedStValue('PSA_PRENOM',S,38,20,tcTrim,FALSE);
T.PutFixedStValue('PSA_NOMJF',S,60,20,tcTrim,FALSE);
T.PutValue('PSA_ETABLISSEMENT',Etab);
T.PutFixedStValue('PSA_ADRESSE1',S,129,25,tcTrim,FALSE);
T.PutFixedStValue('PSA_ADRESSE2',S,154,25,tcTrim,FALSE);
T.PutFixedStValue('PSA_ADRESSE3',S,209,25,tcTrim,FALSE);
T.PutFixedStValue('PSA_CODEPOSTAL',S,179,5,tcTrim,FALSE);
T.PutFixedStValue('PSA_VILLE',S,184,25,tcTrim,FALSE);
T.PutFixedStValue('PSA_TELEPHONE',S,234,15,tcTrim,FALSE);

DateTest := Copy(S, 80, 6);
if (DateTest <> '      ') then
   T.PutFixedStValue('PSA_DATENAISSANCE',S,80,6,tcDate6JMA,FALSE)
else
   T.PutValue('PSA_DATENAISSANCE', IDate1900);
T.PutFixedStValue('PSA_COMMUNENAISS',S,86,25,tcTrim,FALSE);
T.PutFixedStValue('PSA_DEPTNAISSANCE',S,116,2,tcTrim,FALSE);
T.PutFixedStValue ('PSA_NATIONALITE',S,126,3,tcEnum,FALSE, '=;ALG=DZA;ALL=DEU;'+
                  'AUS=AUS;AUT=AUT;BEL=BEL;BRE=BRA;CAN=CAN;CAR=CMR;CHI=CHN;'+
                  'CLI=CHL;CON=COG;DAN=DNK;EGY=EGY;ESP=ESP;EUA=USA;FIN=FIN;'+
                  'FRA=FRA;GRE=GRC;IRL=IRL;ISR=ISR;ITA=ITA;LUX=LUX;MAR=MAR;'+
                  'PBA=NLD;POR=PRT;ROY=GBR;SUE=SWE;SUI=CHE;TCH=CSK;TUN=TUN');
Sexe := Copy(S, 58, 1);
T.PutValue('PSA_SEXE', Sexe);
Situation := Copy(S, 59, 1);
T.PutFixedStValue ('PSA_SITUATIONFAMIL',S,59,1,tcEnum,FALSE,
                   '=;C=CEL;D=DIV;M=MAR;S=CON;V=VEU');

DateTest := Copy(S, 364, 6);
if (DateTest <> '      ') then
   T.PutFixedStValue('PSA_DATEENTREE',S,364,6,tcDate6JMA,FALSE)
else
   T.PutValue('PSA_DATEENTREE', IDate1900);

DateTest := Copy(S, 370, 6);
if (DateTest <> '      ') then
   T.PutFixedStValue('PSA_DATESORTIE',S,370,6,tcDate6JMA,FALSE)
else
   T.PutValue('PSA_DATESORTIE', IDate1900);

DateTest := Copy(S, 382, 6);
if (DateTest <> '      ') then
   T.PutFixedStValue('PSA_DATESORTIEPREC',S,382,6,tcDate6JMA,FALSE)
else
   T.PutValue('PSA_DATESORTIEPREC', IDate1900);

DateTestPrec := Copy(S, 376, 6);
if (DateTestPrec <> '      ') then
   T.PutFixedStValue('PSA_DATEENTREEPREC',S,376,6,tcDate6JMA,FALSE)
else
   T.PutValue('PSA_DATEENTREEPREC', IDate1900);

T.PutFixedStValue('PSA_SORTIEDEFINIT',S,483,1,tcBooleanON,FALSE);
T.PutFixedStValue('PSA_SUSPENSIONPAIE', S,388,1,tcBooleanON,FALSE);
T.PutFixedStValue('PSA_PERSACHARGE',S,249,2,tcEntier,FALSE);
T.PutFixedStValue ('PSA_PGMODEREGLE',S,253,2,tcEnum,FALSE,
                   '=;ES=001;CH=002;VI=008');
T.PutFixedStValue('PSA_COEFFICIENT', S,332,4,tcTrim,FALSE);
T.PutFixedStValue('PSA_QUALIFICATION', S,360,4,tcTrim,FALSE);
T.PutFixedStValue('PSA_CODEEMPLOI', S,336,4,tcTrim,FALSE);

DateTest := Copy(S, 364, 6);
if (DateTestPrec <> '      ') then
   begin
   T.PutFixedStValue('PSA_ANCIENPOSTE',S,376,6,tcDate6JMA,FALSE);
   T.PutFixedStValue('PSA_DATEANCIENNETE',S,376,6,tcDate6JMA,FALSE);
   end
else
   begin
   if (DateTest <> '      ') then
      begin
      T.PutFixedStValue('PSA_ANCIENPOSTE',S,364,6,tcDate6JMA,FALSE);
      T.PutFixedStValue('PSA_DATEANCIENNETE',S,364,6,tcDate6JMA,FALSE);
      end
   else
      begin
      T.PutValue('PSA_ANCIENPOSTE', IDate1900);
      T.PutValue('PSA_DATEANCIENNETE', IDate1900);
      end;
   end;
StRegulAncien:= Copy(S, 391, 3);
if ((StRegulAncien<>'   ') and (isnumeric(StRegulAncien))) then
   RegulAncien:= -StrToInt(StRegulAncien)
else
   RegulAncien:=0;
T.PutValue('PSA_REGULANCIEN', RegulAncien);
TEtabD:= TEtab.FindFirst(['ET_ETABLISSEMENT'], [Etab], TRUE);
if ((TEtabD<>nil) and (TEtabD.GetValue('ETB_CONVENTION')<>null)) then
   begin
   CodeConv := TEtabD.GetValue('ETB_CONVENTION');
   T.PutValue('PSA_CONVENTION', CodeConv);
   end;
T.PutValue('PSA_DATEMODIF', Date);
T.PutValue('PSA_UTILISATEUR', V_PGI.User);
if (Sexe = 'M')	then
   T.PutValue('PSA_CIVILITE', 'MR')
else
   begin
   if (Situation = 'C') or (Situation = 'S') then
    begin
      T.PutValue('PSA_CIVILITE', 'MLE');
      T.PutValue('PSA_NOMJF', '');  //PT75
    end
   else
      T.PutValue('PSA_CIVILITE', 'MME');
   end;
if (Copy(S, 527, 3) = '   ') then
   T.PutValue('PSA_TYPPRUDH', 'ETB')
else
   begin
   T.PutFixedStValue('PSA_PRUDHCOLL', S,527,1,tcTrim,FALSE);
   T.PutFixedStValue('PSA_PRUDHSECT' ,S,528,1,tcTrim,FALSE);
   T.PutFixedStValue('PSA_PRUDHVOTE', S,529,1,tcTrim,FALSE);
   T.PutValue('PSA_TYPPRUDH', 'PER');
   end;
if Assigned(TEtabD) then
  T.PutValue('PSA_CONGESPAYES', TEtabD.GetValue('ETB_CONGESPAYES'))
else
  T.PutValue('PSA_CONGESPAYES', '-');
CodeStat := Copy(S, 328, 4);
StRech:= 'SELECT CC_CODE'+
         ' FROM CHOIXCOD WHERE'+
         ' CC_TYPE = "PSQ" AND'+
         ' CC_ABREGE = "'+CodeStat+'"';
QRech:=OpenSQL(StRech,TRUE);
CodeStat := QRech.FindField('CC_CODE').Asstring ;
T.PutValue('PSA_CODESTAT', CodeStat);
Ferme(QRech);
T.PutFixedStValue('PSA_TRAVETRANG', S, 485, 1,tcTrim,FALSE);
//debut PT77
T.PutFixedStValue('PSA_REGIMESS', S, 390, 1,tcEnum,FALSE,'=;1=200;2=900;3=900');
//T.PutFixedStValue('PSA_REGIMESS', S, 390, 1,tcEnum,FALSE,'=;1=200');
//fin PT77
T.PutFixedStValue('PSA_LIBRE1', S, 328, 4, tcTrim,FALSE);
T.PutFixedStValue ('PSA_CATDADS',S,389,1,tcEnum,FALSE,
                   '=;A=003;C=002;D=001;E=005;O=004');
T.PutFixedStValue('PSA_CONDEMPLOI', S, 402, 1,tcTrim,FALSE);
T.PutFixedStValue('PSA_CATEGNORM', S, 533, 3,tcTrim,FALSE);

EnregSalDefaut(T);

ChangeCodeSalarie(T, 'PSA', TRUE) ;

// Chargement des infos nescessaires à la creation du compte de tiers associé
{if (VH_Paie.PgTypeNumSal = 'NUM') and (VH_Paie.PgTiersAuxiAuto = TRUE) then }//PT78
if (VH_Paie.PgTiersAuxiAuto = TRUE) then //PT78
   BEGIN
   Libell := T.GetValue('PSA_LIBELLE');
   LongLib := Length(Libell);
   Preno := T.GetValue('PSA_PRENOM');
   LongPre := Length(Preno);
   Libell := Copy(Libell, 1, LongLib)+ ' ' + Copy(Preno, 1, LongPre);
   with InfTiers do
        begin
        Libelle    := Copy(Libell, 1, 35);
        Adresse1   := T.GetValue ('PSA_ADRESSE1');
        Adresse2   := T.GetValue ('PSA_ADRESSE2');
        Adresse3   := T.GetValue ('PSA_ADRESSE3');
        Ville      := T.GetValue ('PSA_VILLE');
        Telephone  := T.GetValue ('PSA_TELEPHONE');
        CodePostal := T.GetValue ('PSA_CODEPOSTAL');
        Pays       := '';
        end;
   LeRapport := '';
// Creation du compte de tiers en automatique et recup du numéro
   CodeAuxi:= CreationTiers (InfTiers,LeRapport, 'SAL',
                             T.GetValue('PSA_SALARIE'));
   if CodeAuxi <> '' then
      T.PutValue('PSA_AUXILIAIRE',CodeAuxi)
   Else //debut PT73
    begin
     StMessage:='Vous devez contrôler le compte auxiliaire'+
                ' associé au salarié '+T.GetValue('PSA_SALARIE')+'.';
     FReport.Items.Add(StMessage); 
     Writeln(FRapport, StMessage);
     StMessage:= 'Attention : le RIB est associé au compte auxiliaire';
     FReport.Items.Add(StMessage);
     Writeln(FRapport, StMessage);
    end;  //fin PT73
   if LeRapport <> '' then
      BEGIN
      FReport.Items.Add(LeRapport);
      Writeln(FRapport,LeRapport);
      END;
   END;
okok:=T.InsertOrUpdateDb(FALSE);
CreateYRS (T.GetValue('PSA_SALARIE'), '', '');          //PT66
if (FTypeFichier.ItemIndex = 0) then
   begin
   StMessage:= 'Salarié (T) '+Copy(S,2,6)+' '+Trim(Copy(S,8,30))+' '+
               Trim(Copy(S,38,20))+Verifokok(OkOk);
   FReport.Items.Add(StMessage);
   Writeln(FRapport,StMessage);
   end
else
   begin
   StMessage:= 'Salarié (PSA) '+Copy(S,2,6)+' '+Trim(Copy(S,8,30))+' '+
               Trim(Copy(S,38,20))+Verifokok(OkOk);
   FReport.Items.Add(StMessage);
   Writeln(FRapport,StMessage);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Enregistrement d'un salarié provenant d'ISIS II (x)
Suite ........ : T : TOB salarié
Suite ........ : S : chaine de caractères
Suite ........ : Etab : Numéro d'établissement
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregSalarie2(T:TOB;S,Etab:string);
Var
okok : boolean;
FlTauxPartiel : Double;
Buffer : string;
begin
T.InitValeurs;
T.PutFixedStValue('PSA_SALARIE', S,2,6,tcTrim,FALSE);
T.PutValue('PSA_ETABLISSEMENT',Etab);
okok:=ChangeCodeSalarie(T, 'PSA', FALSE) ;
if okok then
   BEGIN
   T.PutValue('PSA_DATEMODIF', Date);
   T.PutValue('PSA_UTILISATEUR', V_PGI.User);
   T.PutFixedStValue('PSA_CARTESEJOUR',S,8,15,tcTrim,FALSE);
   DateTest := Copy(S, 23, 8);
   if (DateTest <> '        ') then
      T.PutFixedStValue('PSA_DATEXPIRSEJOUR',S,23,8,tcDate8JMA,FALSE)
   else
{PT57
      T.PutValue('PSA_DATEXPIRSEJOUR', null);
}
      T.PutValue('PSA_DATEXPIRSEJOUR', IDate1900);
   T.PutFixedStValue('PSA_DELIVPAR', S,31,30,tcTrim,FALSE);
   T.PutFixedStValue('PSA_MOTIFENTREE',S,118,2,tcEnum,FALSE,'=;EM=001');
   Buffer:= Copy(S, 279, 2);
   if (Buffer<>'  ') then
      T.PutFixedStValue('PSA_CONDEMPLOI',S,279,2,tcEnum,FALSE,
                        '01=C;02=P;04=I;05=D;06=S;07=V;08=O;90=W');
   T.PutFixedStValue('PSA_DADSPROF',S,281,2,tcTrim,FALSE);
   T.PutFixedStValue('PSA_DADSCAT',S,283,2,tcTrim,FALSE);
   Buffer:= Copy(S, 458, 13);
   if ((Buffer<>'             ') and (isnumeric(Buffer))) then
      FlTauxPartiel:= StrToFloat(Copy(S, 458, 13))
   else
      FlTauxPartiel:= 0;
   T.PutValue('PSA_TAUXPARTIEL', (FlTauxPartiel/100));
   END ;
if okok then okok:=T.InsertOrUpdateDb(FALSE);
if (FTypeFichier.ItemIndex = 0) then
   begin
   StMessage:='Salarié (x) '+Copy(S,2,6)+Verifokok(okok);
   FReport.Items.Add(StMessage);
   Writeln(FRapport,StMessage);
   end
else
   begin
   StMessage:='Salarié (PSB) '+Copy(S,2,6)+Verifokok(okok);
   FReport.Items.Add(StMessage);
   Writeln(FRapport,StMessage);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 14/05/2003
Modifié le ... :   /  /
Description .. : Enregistrement d'un salarié provenant d'un fichier TRA
Suite ........ : T : TOB salarié
Suite ........ : S : chaine de caractères
Suite ........ : Etab : Numéro d'établissement
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregSalarie3(T:TOB;S,Etab:string);
var
okok : boolean;
email : string;
TB : TOB;
begin
T.InitValeurs;
T.PutFixedStValue('PSA_SALARIE', S,2,6,tcTrim,FALSE);
T.PutValue('PSA_ETABLISSEMENT',Etab);
okok:=ChangeCodeSalarie(T, 'PSA', FALSE) ;
if okok then
   BEGIN
   T.PutValue('PSA_DATEMODIF', Date);
   T.PutValue('PSA_UTILISATEUR', V_PGI.User);
   T.PutFixedStValue('PSA_PRENOMBIS',S,8,35,tcTrim,FALSE);
   T.PutFixedStValue('PSA_SURNOM',S,43,35,tcTrim,FALSE);
   T.PutFixedStValue('PSA_PAYS',S,78,3,tcTrim,FALSE);
   T.PutFixedStValue('PSA_PAYSNAISSANCE',S,81,3,tcTrim,FALSE);
   T.PutFixedStValue('PSA_LIBELLEEMPLOI',S,84,35,tcTrim,FALSE);
   T.PutFixedStValue('PSA_HORAIREMOIS',S,119,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_TAUXHORAIRE',S,134,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_SALAIREMOIS1',S,149,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_SALAIREMOIS2',S,164,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_SALAIREMOIS3',S,179,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_SALAIREMOIS4',S,194,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_SALAIREMOIS5',S,209,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_SALAIRANN1',S,224,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_SALAIRANN2',S,239,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_SALAIRANN3',S,254,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_SALAIRANN4',S,269,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_SALAIRANN5',S,284,15,tcDouble100,FALSE);
   T.PutFixedStValue('PSA_TRAVAILN1',S,299,3,tcTrim,FALSE);
   T.PutFixedStValue('PSA_TRAVAILN2',S,302,3,tcTrim,FALSE);
   T.PutFixedStValue('PSA_TRAVAILN3',S,305,3,tcTrim,FALSE);
   T.PutFixedStValue('PSA_TRAVAILN4',S,308,3,tcTrim,FALSE);
   T.PutFixedStValue('PSA_INDICE',S,311,17,tcTrim,FALSE);
   T.PutFixedStValue('PSA_NIVEAU',S,328,17,tcTrim,FALSE);
   T.PutFixedStValue('PSA_DADSPROF',S,345,3,tcTrim,FALSE);
   T.PutFixedStValue('PSA_DADSCAT',S,348,3,tcTrim,FALSE);

   Email := Copy(S, 351, 250);
   if Email <> '' then
      begin
      TB := TOB.create('DEPORTSAL', nil, -1);
      TB.InitValeurs;
      TB.PutFixedStValue('PSE_EMAILPROF', S, 351, 250, tcTrim, FALSE);
      TB.Putvalue ('PSE_SALARIE', T.Getvalue('PSA_SALARIE'));
      end;
   okok := T.InsertOrUpdateDb(FALSE);
   CreateYRS (T.GetValue('PSA_SALARIE'), '', '');            //PT66

   if (okok) and (Email <> '') then TB.InsertOrUpdateDb(FALSE);
   end;
if Assigned(TB) then FreeAndNil(TB);

StMessage:='Salarié (PSC) '+Copy(S,2,6)+Verifokok(okok);
FReport.Items.Add(StMessage);
Writeln(FRapport,StMessage);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/10/2000
Modifié le ... :   /  /
Description .. : Enregistrement d'un salarié provenant d'un fichier TDS
Suite ........ : T : TOB salarié
Suite ........ : S : Chaine de caractères
Suite ........ : Etab : Numéro d'établissement
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregSalTDS(T,TSalPGI:TOB;S,Etab:string);
Var
okok : boolean ;
Buf, CodeAuxi, CodeSal, LeRapport, Libell, Nom, Preno : String;
InfTiers : Info_Tiers;
LongLib, LongPre, LongSal : Integer;
TSalPGID : TOB;
begin
T.InitValeurs;
T.PutFixedStValue('PSA_SALARIE',S,23,10,tcTrim,FALSE);
T.PutFixedStValue('PSA_NUMEROSS',S,33,15,tcTrim,FALSE);
T.PutValue('PSA_ETABLISSEMENT',Etab);
CodeSal:='';
//Début de calcul du code salarié pour codage numérique
if (VH_PAIE.PgTypeNumSal='NUM') then
   begin
   CodeSal:= Copy(S,23,10);
   Case FCodeSal.ItemIndex of
        2 : BEGIN //Etab+Sal
            LongSal:=Length(Trim(CodeSal));
            if (LongSal > 7) then
               CodeSal:=Trim(Etab)+Copy(Trim(CodeSal), LongSal+1-7, 7)
            else
               CodeSal:= Trim(Etab)+
                         ColleZeroDevant(StrToInt(Copy(trim(CodeSal), 1, 7)),7);
            END ;
        3 : BEGIN //Sal+Etab
            LongSal:=Length(Trim(CodeSal));
            if (LongSal > 7) then
               CodeSal:=Copy(Trim(CodeSal), LongSal+1-7, 7)+Trim(Etab)
            else
               CodeSal:= ColleZeroDevant(StrToInt(Copy(trim(CodeSal), 1, 7)),7)+
                         Trim(Etab);
            END ;
        END;

   if ((VH_Paie.PgTypeNumSal = 'NUM') and (FCodeSal.ItemIndex<>1)) then
      CodeSal:=ColleZeroDevant(StrToInt(Trim(CodeSal)), 10);
   okok := ChangeCodeSalarie(T, 'PSA', TRUE);
   CodeSal:= T.GetValue ('PSA_SALARIE');
//Fin de calcul du code salarié pour codage numérique

   if (okok= True) then
      begin
      if (((isnumeric(CodeSal)) and (CodeSal<'2147483647')) or
         (FCodeSal.ItemIndex=1)) then
         okok:=True
      else
         begin
         okok:=False;
         if (isnumeric(CodeSal)) then
            StMessage:= 'Le code salarié '+CodeSal+' est supérieur à 2147483647'
         else
            begin
            StMessage:= 'Import impossible pour le salarié '+CodeSal+' car la'+
                        ' gestion des salariés est';
            FReport.Items.Add(StMessage);
            Writeln(FRapport,StMessage);
            StMessage:= 'définie en type numérique';
            end;
         FReport.Items.Add(StMessage);
         Writeln(FRapport,StMessage);
         end;
      end;
   end
else
   begin
   okok := ChangeCodeSalarie(T, 'PSA', TRUE);
   CodeSal:= T.GetValue ('PSA_SALARIE');
   end;

if (okok= True) then
   begin
   TSalPGID := TSalPGI.FindNext(['PSA_SALARIE'], [CodeSal], TRUE);
   if (TSalPGID <> nil) then
      begin
      okok:=False;
      StMessage:='Le code salarié '+CodeSal+' existe déjà';
      FReport.Items.Add(StMessage);
      Writeln(FRapport,StMessage);
      end;
   if (CodeSal='          ') then
      begin
      okok:=False;
      StMessage:='Le code salarié '+CodeSal+' n''est pas permis';
      FReport.Items.Add(StMessage);
      Writeln(FRapport,StMessage);
      end;
   end;

if (okok=False) then
   begin
   Nom := Copy(S, 138, 30);
   if (Nom = '                              ') then
      Nom := Copy(S, 88, 30);
   end
else
   begin
   Nom := Copy(S, 138, 30);
   if (Nom <> '                              ') then
      BEGIN
      T.PutFixedStValue('PSA_LIBELLE',S,138,30,tcTrim,FALSE);
      T.PutFixedStValue('PSA_NOMJF',S,88,20,tcTrim,FALSE);
      END
   else
      BEGIN
      T.PutFixedStValue('PSA_LIBELLE',S,88,30,tcTrim,FALSE);
      Nom := Copy(S, 88, 30);
      END;

   T.PutFixedStValue('PSA_PRENOM',S,118,20,tcTrim,FALSE);
   Buf:= Copy(S, 201, 32);
   if (Buf<>'                                ') then
      begin
      T.PutFixedStValue('PSA_ADRESSE1',S,201,32,tcTrim,FALSE);
      T.PutFixedStValue('PSA_ADRESSE2',S,168,32,tcTrim,FALSE);
      end
   else
      T.PutFixedStValue('PSA_ADRESSE1',S,168,32,tcTrim,FALSE);
   T.PutFixedStValue('PSA_CODEPOSTAL',S,265,5,tcTrim,FALSE);
   T.PutFixedStValue('PSA_VILLE',S,271,26,tcTrim,FALSE);

   DateTest := Copy(S, 48, 6);
   if (DateTest <> '      ') then
      T.PutFixedStValue('PSA_DATENAISSANCE',S,48,6,tcDate6JMA,FALSE)
   else
      T.PutValue('PSA_DATENAISSANCE', IDate1900);

   T.PutFixedStValue('PSA_COMMUNENAISS',S,59,26,tcTrim,FALSE);
   T.PutFixedStValue('PSA_DEPTNAISSANCE',S,54,2,tcTrim,FALSE);
   T.PutFixedStValue('PSA_SEXE',S,33,1,tcEnum,FALSE,'=;1=M;2=F');
   T.PutValue('PSA_DATEENTREE', IDate1900);
   T.PutValue('PSA_DATESORTIE', IDate1900);
   T.PutValue('PSA_DATEENTREEPREC', IDate1900);
   T.PutValue('PSA_DATESORTIEPREC', IDate1900);
   T.PutFixedStValue('PSA_CODEEMPLOI', S,327,4,tcTrim,FALSE);
   T.PutValue('PSA_DATEANCIENNETE', IDate1900);
   T.PutFixedStValue('PSA_CIVILITE', S,85,3,tcEnum,FALSE,'MME=MME;MLE=MLE;=MR');
   //debut PT77
   //T.PutFixedStValue('PSA_REGIMESS', S, 535, 1,tcEnum,FALSE,'=;1=200');
   T.PutFixedStValue('PSA_REGIMESS', S, 535, 1,tcEnum,FALSE,'=;1=200;2=900;3=900');
   //fin PT77
   T.PutFixedStValue('PSA_CATDADS',S,332,1,tcEnum,FALSE,'=;A=A;C=C;D=D;E= ;O= ');
   T.PutFixedStValue('PSA_CONDEMPLOI', S, 351, 1,tcTrim,FALSE);
   EnregSalDefaut(T);
   end;

if okok = FALSE then
   begin
   StMessage:= 'Salarié '+Copy(S,23,10)+' '+Trim(Nom)+' '+Trim(Copy(S,118,20))+
               Verifokok(OkOk)+' matricule';
   FReport.Items.Add(StMessage);
   Writeln(FRapport,StMessage);
   Exit;
   end;

// Chargement des infos nescessaires à la creation du compte de tiers associé
{if (VH_Paie.PgTypeNumSal = 'NUM') and (VH_Paie.PgTiersAuxiAuto = TRUE) then }//PT78
if (VH_Paie.PgTiersAuxiAuto = TRUE) then //PT78
   BEGIN
   Libell := T.GetValue('PSA_LIBELLE');
   LongLib := Length(Libell);
   Preno := T.GetValue('PSA_PRENOM');
   LongPre := Length(Preno);
   Libell := Copy(Libell, 1, LongLib)+ ' ' + Copy(Preno, 1, LongPre);
   with InfTiers do
        begin
        Libelle    := Copy(Libell, 1, 35);
        Adresse1   := T.GetValue ('PSA_ADRESSE1');
        Adresse2   := T.GetValue ('PSA_ADRESSE2');
        Adresse3   := T.GetValue ('PSA_ADRESSE3');
        Ville      := T.GetValue ('PSA_VILLE');
        Telephone  := T.GetValue ('PSA_TELEPHONE');
        CodePostal := T.GetValue ('PSA_CODEPOSTAL');
        Pays       := '';
        end;
// Creation du compte de tiers en automatique et recup du numéro
   CodeAuxi:=CreationTiers (InfTiers, LeRapport, 'SAL',
                            T.GetValue('PSA_SALARIE'));
   if (CodeAuxi<>'') then
      T.PutValue ('PSA_AUXILIAIRE', CodeAuxi)
      Else //debut PT73
    begin
    StMessage:='Vous devez contrôler le compte auxiliaire'+
                ' associé au salarié '+T.GetValue('PSA_SALARIE')+'.';
     FReport.Items.Add(StMessage); 
     Writeln(FRapport, StMessage);
     StMessage:= 'Attention : le RIB est associé au compte auxiliaire';
     FReport.Items.Add(StMessage);
     Writeln(FRapport, StMessage);
    end; //fin PT73
   END;
T.PutValue ('PSA_DATEMODIF', Date);
T.PutValue ('PSA_UTILISATEUR', V_PGI.User);

if (okok=True) then
   begin
   okok:= T.InsertOrUpdateDb (FALSE);
   CreateYRS (T.GetValue('PSA_SALARIE'), '', '');      //PT66
   end;

StMessage:= 'Salarié '+Copy(S,23,10)+' '+Trim(Nom)+' '+Trim(Copy(S,118,20))+
            Verifokok(OkOk);
FReport.Items.Add(StMessage);
Writeln(FRapport,StMessage);
end;

//PT59
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/04/2006
Modifié le ... :   /  /
Description .. : Enregistrement d'une rubrique dans la Tob salarié
Suite ........ : TSalaries : TOB salarié
Suite ........ : Rubrique : Numéro de la rubrique DADS-U
Suite ........ : Valeur : Valeur de la rubrique DADS-U
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.SauveRubriqueDADSU (var TSalaries : TOB; var SSiret : string;
                                        Code, Rubrique, Valeur : string);
var
Buffer, StTablette : string;
QRechTablette : TQuery;
TEtabD : TOB;
begin
if (Code = '20') then
   begin
   if (Rubrique = 'S20.G01.00.001') then
      Siren:= Valeur;
   end
else
if (Code = '30') then
   begin
   if (Rubrique = 'S30.G01.00.001') then
      begin
      TSalaries.InitValeurs;
      Premier:= False;
      if (length (Valeur)=13) then
         begin
         Buffer:= IntToStr (CalculCleSS(Valeur));
         //debut PT76
         Valeur := Valeur + Copy('00',1,2-length(Buffer)) + Copy(buffer,1,length(Buffer))
         //fin PT76
         end;
      TSalaries.PutValue ('PSA_NUMEROSS', Valeur);
      end
   else
   if (Rubrique = 'S30.G01.00.002') then
      TSalaries.PutValue ('PSA_LIBELLE', Valeur)
   else
   if (Rubrique = 'S30.G01.00.003') then
      TSalaries.PutValue ('PSA_PRENOM', Valeur)
   else
   if (Rubrique = 'S30.G01.00.004') then
      begin
      TSalaries.PutValue ('PSA_NOMJF', TSalaries.GetValue('PSA_LIBELLE'));
      TSalaries.PutValue ('PSA_LIBELLE', Valeur);
      end
   else
   if (Rubrique = 'S30.G01.00.006') then
      TSalaries.PutValue ('PSA_SURNOM', Valeur)
   else
   if (Rubrique = 'S30.G01.00.007') then
      begin
      if (Valeur = '01') then
         begin
         TSalaries.PutValue ('PSA_CIVILITE', 'MR');
         TSalaries.PutValue ('PSA_SEXE', 'M');
         end
      else
         begin
         if (Valeur = '02') then
            TSalaries.PutValue ('PSA_CIVILITE', 'MME')
         else
         if (Valeur = '03') then
            TSalaries.PutValue ('PSA_CIVILITE', 'MLE');
         TSalaries.PutValue ('PSA_SEXE', 'F');
         end;
      end
   else
   if (Rubrique = 'S30.G01.00.008.001') then
      TSalaries.PutValue ('PSA_ADRESSE2', Valeur)
   else
   if (Rubrique = 'S30.G01.00.008.003') then
      TSalaries.PutValue ('PSA_ADRESSE1', Valeur)
   else
   if (Rubrique = 'S30.G01.00.008.004') then
      TSalaries.PutValue ('PSA_ADRESSE1', TSalaries.GetValue ('PSA_ADRESSE1')+
                                          ' '+Valeur)
   else
   if (Rubrique = 'S30.G01.00.008.006') then
      TSalaries.PutValue ('PSA_ADRESSE1', TSalaries.GetValue ('PSA_ADRESSE1')+
                                          ' '+Valeur)
   else
   if (Rubrique = 'S30.G01.00.008.010') then
      TSalaries.PutValue ('PSA_CODEPOSTAL', Valeur)
   else
   if (Rubrique = 'S30.G01.00.008.012') then
      TSalaries.PutValue ('PSA_VILLE', Valeur)
   else
   if (Rubrique = 'S30.G01.00.008.014') then
      begin
      PaysISOLibEnvers (Valeur, Buffer);
      TSalaries.PutValue ('PSA_PAYS', Buffer);
      end
   else
   if (Rubrique = 'S30.G01.00.009') then
      begin
      Buffer:= copy(Valeur, 1, 2)+'/'+copy(Valeur, 3, 2)+'/'+copy(Valeur, 5, 4);
      TSalaries.PutValue ('PSA_DATENAISSANCE', StrToDate (Buffer));
      end
   else
   if (Rubrique = 'S30.G01.00.010') then
      TSalaries.PutValue ('PSA_COMMUNENAISS', Valeur)
   else
   if (Rubrique = 'S30.G01.00.011') then
      begin
      Buffer:= Valeur;
      if (Buffer = '2A') then
         Buffer := '20A';
      if (Buffer = '2B') then
         Buffer := '20B';
      if (Buffer <> '99') then
         TSalaries.PutValue ('PSA_DEPTNAISSANCE', Buffer);
      end
   else
   if (Rubrique = 'S30.G01.00.012') then
      begin
      PaysISOLibEnvers (Valeur, Buffer);
      TSalaries.PutValue ('PSA_PAYSNAISSANCE', Buffer);
      end
   else
   if (Rubrique = 'S30.G01.00.013') then
      begin
      PaysISOLibEnvers (Valeur, Buffer);
      TSalaries.PutValue ('PSA_NATIONALITE', Buffer);
      end;
   end
else
if (Code = '41') then
   begin
   if (Rubrique = 'S41.G01.00.005') then
      begin
      SSiret:= Siren+Valeur;
{PT62-2
      StTablette:= 'SELECT ET_ETABLISSEMENT'+
                   ' FROM ETABLISS WHERE'+
                   ' ET_SIRET="'+SSiret+'"';
      QRechTablette:= OpenSQL(StTablette,TRUE);
      if (not QRechTablette.EOF) then
         TSalaries.PutValue ('PSA_ETABLISSEMENT',
                             QRechTablette.FindField('ET_ETABLISSEMENT').AsString);
      Ferme(QRechTablette);
}
      TEtabD:= TEtab.FindFirst (['ET_SIRET'], [SSiret], True);
      if (TEtabD <> nil) then
         TSalaries.PutValue ('PSA_ETABLISSEMENT',
                             TEtabD.GetValue ('ET_ETABLISSEMENT'));
//FIN PT62-2
      end
   else
   if (Rubrique = 'S41.G01.00.008.001') then
      begin
      if (Valeur='01') then
         TSalaries.PutValue ('PSA_MULTIEMPLOY', '-')
      else
      if (Valeur='02') then
         TSalaries.PutValue ('PSA_MULTIEMPLOY', 'X');
      end
   else
   if (Rubrique = 'S41.G01.00.010') then
      begin
      StTablette:= 'SELECT CC_CODE'+
                   ' FROM CHOIXCOD WHERE'+
                   ' CC_LIBELLE="'+Valeur+'" AND'+
                   ' CC_TYPE = "PSQ"';
      QRechTablette:= OpenSQL(StTablette,TRUE);
      if (not QRechTablette.EOF) then
         TSalaries.PutValue ('PSA_LIBELLEEMPLOI',
                             QRechTablette.FindField('CC_CODE').AsString);
      Ferme(QRechTablette);
      end
   else
   if (Rubrique = 'S41.G01.00.011') then
      TSalaries.PutValue ('PSA_CODEEMPLOI', Valeur)
   else
   if (Rubrique = 'S41.G01.00.013') then
      begin
      if Valeur = '01' then
         Buffer := 'C'
      else
      if Valeur = '02' then
         Buffer := 'P'
      else
      if Valeur = '04' then
         Buffer := 'I'
      else
      if Valeur = '05' then
         Buffer := 'D'
      else
      if Valeur = '06' then
         Buffer := 'S'
      else
      if Valeur = '07' then
         Buffer := 'V'
      else
      if Valeur = '08' then
         Buffer := 'O'
      else
      if Valeur = '09' then
         Buffer := 'N'
      else
      if Valeur = '10' then
         Buffer := 'F'
      else
      if Valeur = '90' then
         Buffer := 'W';

      TSalaries.PutValue ('PSA_CONDEMPLOI', Buffer);
      end
   else
   if (Rubrique = 'S41.G01.00.014') then
      TSalaries.PutValue ('PSA_DADSPROF', Valeur)
   else
   if (Rubrique = 'S41.G01.00.015') then
      TSalaries.PutValue ('PSA_DADSCAT', Valeur)
   else
   if (Rubrique = 'S41.G01.00.016') then
      begin
      StTablette:= 'SELECT PCV_CONVENTION'+
                   ' FROM CONVENTIONCOLL WHERE'+
                   ' PCV_IDCC="'+Valeur+'"';
      QRechTablette:= OpenSQL(StTablette,TRUE);
      if (not QRechTablette.EOF) then
         TSalaries.PutValue ('PSA_CONVENTION',
                             QRechTablette.FindField('PCV_CONVENTION').AsString);
      Ferme(QRechTablette);
      end
   else
   if (Rubrique = 'S41.G01.00.017') then
      begin
      StTablette:= 'SELECT PMI_CODE'+
                   ' FROM MINIMUMCONVENT WHERE'+
                   ' ##PMI_PREDEFINI##'+
                   ' PMI_LIBELLE="'+Valeur+'" AND'+
                   ' PMI_NATURE="COE"';
      QRechTablette:= OpenSQL(StTablette,TRUE);
      if (not QRechTablette.EOF) then
         TSalaries.PutValue ('PSA_COEFFICIENT',
                             QRechTablette.FindField('PMI_CODE').AsString);
      Ferme(QRechTablette);
      end
   else
   if (Rubrique = 'S41.G01.00.018.001') then
      TSalaries.PutValue ('PSA_REGIMESS', Valeur)
   else
   if (Rubrique = 'S41.G01.00.019') then
      TSalaries.PutValue ('PSA_SALARIE', Valeur)
   else
   if (Rubrique = 'S41.G01.00.020') then
      TSalaries.PutValue ('PSA_TAUXPARTIEL', Valeur)
   else
   if (Rubrique = 'S41.G01.00.034') then
      begin
      if Valeur = '01' then
         Buffer := 'F'
      else
      if Valeur = '02' then
         Buffer := 'E';
{PT61
      TSalaries.PutValue ('PSA_TAUXPARTIEL', Buffer);
}
      TSalaries.PutValue ('PSA_TRAVETRANG', Buffer);
      end
   else
   if (Rubrique = 'S41.G01.00.043') then
      begin
      if Valeur = 'P' then
         TSalaries.PutValue ('PSA_REMPOURBOIRE', 'X');
      end
   else
   if (Rubrique = 'S41.G01.00.045') then
      begin
      if Valeur = 'F' then
         TSalaries.PutValue ('PSA_ALLOCFORFAIT', 'X');
      end
   else
   if (Rubrique = 'S41.G01.00.046') then
      begin
      if Valeur = 'R' then
         TSalaries.PutValue ('PSA_REMBJUSTIF', 'X');
      end
   else
   if (Rubrique = 'S41.G01.00.047') then
      begin
      if Valeur = 'P' then
         TSalaries.PutValue ('PSA_PRISECHARGE', 'X');
      end
   else
   if (Rubrique = 'S41.G01.00.048') then
      begin
      if Valeur = 'D' then
         TSalaries.PutValue ('PSA_AUTREDEPENSE', 'X');
      end
   else
   if (Rubrique = 'S41.G02.00.009') then
      TSalaries.PutValue ('PSA_PRUDHCOLL', Copy (Valeur, 2, 1))
   else
   if (Rubrique = 'S41.G02.00.010') then
      TSalaries.PutValue ('PSA_PRUDHSECT', Copy (Valeur, 2, 1))
   else
   end
else
if (Code = '45') then
   begin
   if (Rubrique = 'S45.G01.00.001') then
      TSalaries.PutValue ('PSA_DATEENTREE', Valeur)
   else
   if (Rubrique = 'S45.G01.01.009') then
      begin
      StTablette:= 'SELECT CO_CODE'+
                   ' FROM COMMUN WHERE'+
                   ' CO_ABREGE="'+Valeur+'" AND'+
                   ' CO_TYPE="PSF"';
      QRechTablette:= OpenSQL(StTablette,TRUE);
      if (not QRechTablette.EOF) then
         TSalaries.PutValue ('PSA_SITUATIONFAMIL',
                             QRechTablette.FindField('CO_CODE').AsString);
      Ferme(QRechTablette);
      end
   else
   if (Rubrique = 'S45.G01.01.010') then
      TSalaries.PutValue ('PSA_PERSACHARGE', Valeur)
   else
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/04/2006
Modifié le ... :   /  /
Description .. : Enregistrement d'un salarié provenant d'un fichier DADS-U
Suite ........ : T : TOB salarié
Suite ........ : TSalPGI : TOB des salariés existants
Suite ........ : SIRET : Siret de l'établissement
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregSalDADSU(T, TSalPGI : TOB; SIRET : string);
Var
okok : boolean ;
CodeAuxi, CodeSal, Etab, LeRapport, Libell, Nom, Preno : String;
InfTiers : Info_Tiers;
LongLib, LongPre, LongSal : Integer;
TEtabD, TSalPGID : TOB;
begin
StMessage2 := ''; //PT72
okok := True;   //PT72
CodeSal:= T.GetValue('PSA_SALARIE');
Nom:= T.GetValue('PSA_LIBELLE');
Etab:= T.GetValue('PSA_ETABLISSEMENT');
//PT62-2
TEtabD:= TEtab.FindFirst(['ET_ETABLISSEMENT'], [Etab], TRUE);
if Assigned(TEtabD) then
  T.PutValue('PSA_CONGESPAYES', TEtabD.GetValue('ETB_CONGESPAYES'))
else
  T.PutValue('PSA_CONGESPAYES', '-');
//FIN PT62-2

//Début de calcul du code salarié pour codage numérique
if (VH_PAIE.PgTypeNumSal='NUM') then
   begin
   Case FCodeSal.ItemIndex of
        2 : BEGIN //Etab+Sal
//Debut PT72
            LongSal:=Length(Trim(CodeSal));
             If VerifType( Copy(trim(CodeSal), LongSal+1-7, 7)) then
             Begin
               If VerifType ( Trim(Etab)) then
               Begin
                  if (LongSal > 7) then
                  CodeSal:=Trim(Etab)+Copy(Trim(CodeSal), LongSal+1-7, 7)
                  else
                  CodeSal:= Trim(Etab)+
                            ColleZeroDevant(StrToInt(Copy(trim(CodeSal), 1, 7)),7);
               END
               Else
                okok := False;
             END
             Else
                okok := False;
            END ;
        3 : BEGIN //Sal+Etab
            LongSal:=Length(Trim(CodeSal));
             If VerifType( Copy(trim(CodeSal), LongSal+1-7, 7)) then
             Begin
               If VerifType ( Trim(Etab)) then
               Begin
                  if (LongSal > 7) then
                  CodeSal:=Copy(Trim(CodeSal), LongSal+1-7, 7)+Trim(Etab)
                  else
                  CodeSal:= ColleZeroDevant(StrToInt(Copy(trim(CodeSal), 1, 7)),7)+
                            Trim(Etab);
               END
               Else
                okok := False;
             END
             Else
                okok := False;
            END ;
        END;

   if ((VH_Paie.PgTypeNumSal = 'NUM') and (FCodeSal.ItemIndex<>1)) then
   begin
      LongSal:=Length(Trim(CodeSal));
             If VerifType( Copy(trim(CodeSal), 1, 10)) then
             CodeSal:=ColleZeroDevant(StrToInt(Trim(CodeSal)), 10)
             Else
                okok := False;
   end;
//fin PT72

//Fin de calcul du code salarié pour codage numérique

   if (okok= True) then
      begin
//PT72
   okok:= ChangeCodeSalarie(T, 'PSA', TRUE);
   CodeSal:= T.GetValue ('PSA_SALARIE');
//FIN PT72
      if (((isnumeric(CodeSal)) and (CodeSal<'2147483647')) or
         (FCodeSal.ItemIndex=1)) then
         okok:=True
      else
         begin
         okok:=False;
         if (isnumeric(CodeSal)) then
            StMessage:= 'Le code salarié '+CodeSal+' est supérieur à 2147483647'
         else
            begin
            StMessage:= 'Import impossible pour le salarié '+CodeSal+' car la'+
                        ' gestion des salariés est';
            FReport.Items.Add(StMessage);
            Writeln(FRapport,StMessage);
            StMessage:= 'définie en type numérique';
            end;
         FReport.Items.Add(StMessage);
         Writeln(FRapport,StMessage);
         end;
      end;
   end
else
   begin
   okok:= ChangeCodeSalarie(T, 'PSA', TRUE);
   CodeSal:= T.GetValue ('PSA_SALARIE');
   end;

if (okok= True) then
   begin
   if (CodeSal='          ') then
      begin
      okok:=False;
      StMessage:='Le code salarié '+CodeSal+' n''est pas permis';
      FReport.Items.Add(StMessage);
      Writeln(FRapport,StMessage);
      end;
   end;

if (okok=True) then
   begin
   if (T.GetValue ('PSA_PAYS')='') then
      T.PutValue('PSA_PAYS', 'FRA');
   EnregSalDefaut(T);
   end;

if okok = FALSE then
   begin
   StMessage:= 'Salarié '+CodeSal+' '+Trim(Nom)+' '+Trim(Preno)+Verifokok(OkOk)+
               ' matricule';
   FReport.Items.Add(StMessage);
   Writeln(FRapport,StMessage);
// debut PT72
   If StMessage2 <> '' then
   begin
       FReport.Items.Add(StMessage2);
       Writeln(FRapport,StMessage2);
   end;
//fin PT72
   Exit;
   end
else
   begin
   TSalPGID := TSalPGI.FindNext(['PSA_SALARIE'], [CodeSal], TRUE);
   if (TSalPGID <> nil) then
      begin
      okok:=False;
      StMessage:='Le code salarié '+CodeSal+' existe déjà';
      FReport.Items.Add(StMessage);
      Writeln(FRapport,StMessage);
      end;
   end;

// Chargement des infos nescessaires à la creation du compte de tiers associé
{if (VH_Paie.PgTypeNumSal = 'NUM') and (VH_Paie.PgTiersAuxiAuto = TRUE) then }//PT78
if (VH_Paie.PgTiersAuxiAuto = TRUE) then //PT78
   BEGIN
   Libell := T.GetValue('PSA_LIBELLE');
   LongLib := Length(Libell);
   Preno := T.GetValue('PSA_PRENOM');
   LongPre := Length(Preno);
   Libell := Copy(Libell, 1, LongLib)+ ' ' + Copy(Preno, 1, LongPre);
   with InfTiers do
        begin
        Libelle    := Copy(Libell, 1, 35);
        Adresse1   := T.GetValue ('PSA_ADRESSE1');
        Adresse2   := T.GetValue ('PSA_ADRESSE2');
        Adresse3   := T.GetValue ('PSA_ADRESSE3');
        Ville      := T.GetValue ('PSA_VILLE');
        Telephone  := T.GetValue ('PSA_TELEPHONE');
        CodePostal := T.GetValue ('PSA_CODEPOSTAL');
        Pays       := '';
        end;
// Creation du compte de tiers en automatique et recup du numéro
   CodeAuxi:= CreationTiers (InfTiers, LeRapport, 'SAL',
                             T.GetValue ('PSA_SALARIE'));
   if (CodeAuxi<>'') then
      T.PutValue ('PSA_AUXILIAIRE', CodeAuxi)
        Else //debut PT72
    begin
    StMessage:='Vous devez contrôler le compte auxiliaire'+
                ' associé au salarié '+T.GetValue('PSA_SALARIE')+'.';
     FReport.Items.Add(StMessage);
     Writeln(FRapport, StMessage);
     StMessage:= 'Attention : le RIB est associé au compte auxiliaire';
     FReport.Items.Add(StMessage);
     Writeln(FRapport, StMessage);
    end; //fin PT72
   END;
T.PutValue ('PSA_DATEMODIF', Date);
T.PutValue ('PSA_UTILISATEUR', V_PGI.User);

if (okok = true) then
   begin
   okok:=T.InsertOrUpdateDb (FALSE);
   CreateYRS (T.GetValue('PSA_SALARIE'), '', '');        //PT66
   end;
StMessage:= 'Salarié '+CodeSal+' '+Trim(Nom)+' '+Trim(Preno)+Verifokok(OkOk);
FReport.Items.Add(StMessage);
Writeln(FRapport,StMessage);
end;
//FIN PT59


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/10/2000
Modifié le ... :   /  /
Description .. : Enregistrement d'un RIB salarié
Suite ........ : TM : TOB RIB
Suite ........ : S : chaine de caractères
Suite ........ : QRechSal : TQuery sur les salariés existants
Suite ........ : Etab : Numéro d'établissement
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregRIB(TM:Tob; S:string; QRechSal:TQuery; Etab:string);
Var
TRIB_Sal : TOB;
Auxi, CodeSal, Donnee : string;
i, L : integer;
Source : PChar;
Ch: Char;
begin
TRIB_Sal := TOB.Create('RIB', TM, -1);
CodeSal := Copy(S,2,6);
if (VH_Paie.PgTypeNumSal = 'NUM') then // Calcul du code salarié
   BEGIN
   Case FCodeSal.ItemIndex of
//Sans changement du code
        0 : CodeSal:=ColleZeroDevant(StrToInt(CodeSal), 10);

//Renumérotation
        1 : BEGIN
            CodeSal := Trim(Etab)+Trim(CodeSal);
            i:=ListeCodeSal.IndexOf(CodeSal) ;
            if i<0 then
               Exit
            else
               CodeSal:=ColleZeroDevant(longint(ListeCodeSal.Objects[i]), 10);
            END ;

//Etab+Sal
        2 : BEGIN
            case FTypeFichier.ItemIndex of
                 0 : CodeSal := Trim(Etab)+Trim(CodeSal);
                 2 : CodeSal := Trim(Etab)+Trim(CodeSal);   //PT25
                 END;
            CodeSal:=ColleZeroDevant(StrToInt(CodeSal), 10);
            END ;

//Sal+Etab
        3 : BEGIN
            case FTypeFichier.ItemIndex of
                 0 : CodeSal := Trim(CodeSal)+Trim(Etab);
                 2 : CodeSal := Trim(CodeSal)+Trim(Etab);      //PT25
                 END;
            CodeSal:=ColleZeroDevant(StrToInt(CodeSal), 10);
            END ;
        END;
   END ;
QRechSal.First ;
while  not QRechSal.Eof do
    begin
    if QRechSal.FindField ('PSA_SALARIE').Asstring = CodeSal then
       break;
    QRechSal.Next;
    end;
Auxi := QRechSal.FindField('PSA_AUXILIAIRE').Asstring ;
TRIB_Sal.PutValue('R_AUXILIAIRE', Auxi) ;
TRIB_Sal.PutValue('R_NUMERORIB', '1') ;
TRIB_Sal.PutValue('R_PRINCIPAL', 'X') ;
TRIB_Sal.PutFixedStValue('R_ETABBQ',S,280,5,tcTrim,FALSE) ;
TRIB_Sal.PutFixedStValue('R_GUICHET',S,285,5,tcTrim,FALSE) ;
TRIB_Sal.PutFixedStValue('R_NUMEROCOMPTE',S,290,11,tcTrim,FALSE) ;
TRIB_Sal.PutFixedStValue('R_CLERIB',S,301,2,tcTrim,FALSE) ;
{PT14
TRIB_Sal.PutFixedStValue('R_DOMICILIATION',S,255,25,tcTrim,FALSE) ;
}
{PT18
TRIB_Sal.PutFixedStValue('R_DOMICILIATION',S,255,24,tcTrim,FALSE) ;
}

TRIB_Sal.PutValue ('R_PAYS', CodePaysDeIso(VH^.PaysLocalisation)); //PT74

Donnee := Copy(S,255,24);
if ExisteCarInter(Donnee) then
   begin
   L := Length(Donnee);
   Source := Pointer(Donnee);
   while L <> 0 do
         begin
         Ch := Source^;
         if (Ch in ['"','_','''','%','*']) then
            begin
            Ch := ' ';
            Source^ := Ch;
            end;
         Inc(Source);
         Dec(L);
         end;
   end;
TRIB_Sal.PutValue('R_DOMICILIATION',Donnee) ;
//FIN PT18

TRIB_Sal.PutValue('R_SALAIRE', 'X') ;
TRIB_Sal.PutValue('R_ACOMPTE', 'X') ;
TRIB_Sal.PutValue('R_FRAISPROF', 'X') ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Transformation d'un booléen en message
Suite ........ : okok : booléen d'origine
Suite ........ : result : chaine de caractères pour message de sortie
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
Function TFImpIsis.Verifokok( okok : boolean) : String ;
BEGIN
if okok then result:=' ok' else result:=' erreur' ;
END ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Récupération du numéro établissement
Suite ........ : S : chaine de caractères
Suite ........ : résultat : '-1' si l'établissement n'existe pas ou numéro de
Suite ........ : l'établissement sur lequel on se positionne
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
Function TFImpIsis.EnregEtab(S:string) : string;
var
Siret : String;
TEtabD : TOB;
begin
if (TEtab = nil) then
   ChargeTOBEtab;
Siret := Copy(S, 203, 14);
TEtabD := TEtab.FindFirst(['ET_SIRET'], [Siret], TRUE);
if (TEtabD = nil) then
   begin
   PGIBox ('L''établissement dont le SIRET est "'+Siret+'" n''existe pas!',
           StTitre);         //PT64
   result := '-1';
   end
else
   begin
   result := TEtabD.GetValue('ET_ETABLISSEMENT');
   if ((VH_Paie.PgTypeNumSal = 'NUM') and not(IsNumeric(result)) and
      ((FCodeSal.ItemIndex=2) or (FCodeSal.ItemIndex=3))) then
      begin
      PGIBox('L''établissement "'+result+'" a un code alphanumérique#13#10'+
             'alors que le codage des salariés est numérique !', StTitre);  //PT64
      result := '-1';
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Positionnement des indicateurs sur ceux de l'établissement
Suite ........ : en cours
Suite ........ : Comparaison horaire établissement PGI et horaire
Suite ........ : établissement fichier d'import
Suite ........ : S : Chaine de caractères
Suite ........ : Etab : Numéro d'établissement
Suite ........ : Résultat :
Suite ........ : '0' si on récupère la date de fin de paye
Suite ........ : '-1' si on ne trouve pas les informations de l'établissement
Suite ........ : complémentaire
Suite ........ : '-2' s'il n'existe pas de date de fin de paye
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpIsis.EnregEtab3(S,Etab:string) : string;
Var
Horaire, HorArch : double;
TEtabD : TOB;
begin
result := '0';
try
   begintrans;
   RecupEtab3(S);
   CommitTrans;
Except
   Rollback;
   StMessage:='Aucune session de paie n''a été ouverte dans ISIS2 pour#13#10'+
              ' l''établissement '+Etab+'. Vous devez ouvrir une session#13#10'+
              ' de paie dans ISIS2 et reconstruire le fichier d''import#13#10'+
              ' pour cet établissement';
   FReport.Items.Add(StMessage); //PT73
   PGIBox (StMessage, StTitre);         //PT64
   StMessage:='Aucune session de paie n''a été ouverte dans ISIS2 pour'+
              ' l''établissement '+Etab+'. Vous devez ouvrir une session'+
              ' de paie dans ISIS2 et reconstruire le fichier d''import'+
              ' pour cet établissement';
   Writeln(FRapport, StMessage);
   result := '-2';
END;

TEtabD := TEtab.FindFirst(['ET_ETABLISSEMENT'], [Etab], TRUE);
if (TEtabD<>nil) then
   begin
   Horaire := TEtabD.GetValue('ETB_HORAIREETABL');
   HorArch := StrToFloat(Copy(S, 6, 13))/100;
   Horaire := Arrondi(Horaire, 2);
   HorArch := Arrondi(HorArch, 2);
   if (Horaire <> HorArch) then
      begin //Message d'erreur
      StMessage:='Les informations complémentaires de#13#10'+
                 'l''établissement "'+Etab+'" ne correspondent pas!';
      FReport.Items.Add(StMessage); //PT73
      PGIBox(StMessage, StTitre);  //PT64
      StMessage:='Les informations complémentaires de'+
                 'l''établissement "'+Etab+'" ne correspondent pas!';
      Writeln(FRapport, StMessage);
      result := '-1';
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/02/2001
Modifié le ... :   /  /
Description .. : Positionnement des indicateurs sur ceux de l'établissement
Suite ........ : en cours
Suite ........ : S : chaine de caractères
Suite ........ : Etab : Numéro de l'établissement
Suite ........ : Résultat :
Suite ........ : '0' si on récupère la date de fin de paye
Suite ........ : '-2' s'il n'existe pas de date de fin de paye
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpIsis.EnregEtab3bis(S:string) : string;
begin
result := '0';
try
   begintrans;
   RecupEtab3(S);
   CommitTrans;
Except
   Rollback;
   result := '-2';
END;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 08/02/2001
Modifié le ... :   /  /
Description .. : Récupération de la fin de période, du nombre de mois traité
Suite ........ : et du flag session ouverte/fermée
Suite ........ : S : chaine de caractères
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.RecupEtab3(S:string);
begin
DateTest := Copy(S, 86, 6);
if (DateTest <> '      ') then
   begin
   FinExo:=Str6ToDate(Copy(S,86,6),90);
   NbMois:=StrToInt(Copy(S,92,2));
   Ouvert:=Copy(S,100,1);
   if (Ouvert = 'O') then
      NbMois:=NbMois-1;
   if (NbMois < 1) then
      NbMois:=1;
   end
else
   begin
{PT57
   FinExo:=null ;
}
   FinExo:=IDate1900;
   NbMois:=1 ;
   Ouvert:=Copy(S,100,1);
   end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/10/2000
Modifié le ... :   /  /
Description .. : Dans le cas du fichier TDS, contrôle du SIRET et des
Suite ........ : informations complémentaires
Suite ........ : Input : S: Chaîne de caractères contenant le numéro de
Suite ........ : SIRET
Suite ........ : Output : result : Numéro de l'établissement sur lequel
Suite ........ : on se positionne
Suite ........ :    '-1' si on ne trouve pas les informations de
Suite ........ :         l'établissement complémentaires
Suite ........ :    '-2' si l'établissement n'existe pas
Mots clefs ... : PAIE;PGIMPORT;SIRET
*****************************************************************}
function TFImpIsis.ControlSIRETEtab(S : string) : string;
var
SIRET, StEtabComp : String;
TEtabD : TOB;
begin
//PT34
if (TEtab = nil) then
   ChargeTOBEtab;
//FIN PT34
SIRET := Copy(S, 23, 14);
TEtabD := TEtab.FindFirst(['ET_SIRET'], [SIRET], TRUE);
if (TEtabD=nil) then
   begin
   StMessage:='L''établissement dont le SIRET est "'+SIRET+'" n''existe pas!';
   FReport.Items.Add(StMessage);
   Writeln(FRapport,StMessage);
   result := '-2';
   end
else
   begin
   result := TEtabD.GetValue('ET_ETABLISSEMENT');
   StEtabComp := TEtabD.GetValue('ETB_HORAIREETABL');
   if (StEtabComp='') then
      begin
      StMessage:= 'Les informations complémentaires de#13#10'+
                  ' l''établissement "'+result+'" n''existent pas!';
      FReport.Items.Add(StMessage);
      StMessage:= 'Les informations complémentaires de'+
                  ' l''établissement "'+result+'" n''existent pas!';
      Writeln(FRapport,StMessage);
      result := '-1';
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Récupération du début d'exercice du dossier
Suite ........ : St : chaine de caractères
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregDossier(St: string);
begin
DateTest := Copy(St, 26, 6);
if (DateTest <> '      ') then
   DebutExo:=Str6ToDate(Copy(St,26,6),90)
else
{PT57
   DebutExo:=null ;
}
   DebutExo:=IDate1900 ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Enregistrement cumuls individuels salariés
Suite ........ : H : TOB cumuls salariés
Suite ........ : HB : TOB historique bulletin pour récupération des bases
Suite ........ : S : Chaine de caractères N°1
Suite ........ : S2 : Chaine de caractères N°2
Suite ........ : S3 : Chaine de caractères N°3
Suite ........ : S4 : Chaine de caractères N°4
Suite ........ : Etab : Numéro d'établissement
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregHistoCum(H,HB: TOB; S,S2,S3,S4,Etab: string);
Var
i : Integer ;
Tot : Double ;
okok : boolean ;
TD : TOB ;
CodeCum, Corresp : String ;
DateD, DateEntree : TDateTime;
CatDADS, CodeSal, StSal : String ;
QRechSal : TQuery;
begin
H.InitValeurs;
H.PutFixedStValue ('PHC_SALARIE', S, 2, 6, tcTrim, FALSE);
H.PutValue ('PHC_ETABLISSEMENT', Etab);
H.PutValue ('PHC_REPRISE', 'X');
okok:=ChangeCodeSalarie (H, 'PHC', TRUE) ;

CodeSal:= H.GetValue ('PHC_SALARIE');
StSal:= 'SELECT PSA_DATEENTREE, PSA_CATDADS'+
        ' FROM SALARIES WHERE'+
        ' PSA_SALARIE="'+CodeSal+'"';
QRechSal:= OpenSQL (StSal, TRUE);
CatDADS:= QRechSal.FindField ('PSA_CATDADS').AsString;
DateEntree:= QRechSal.FindField ('PSA_DATEENTREE').AsDateTime;
Ferme (QRechSal);

if (Histo=False) then
//Si on n'a pas de cumul mois par mois
   BEGIN
   if (Ouvert='O') then
      DateD:= DebutDeMois (PlusMois (FinExo, -1))
   else
      DateD:= DebutDeMois (FinExo);
   H.PutValue ('PHC_DATEDEBUT', DateD) ;
   H.PutValue ('PHC_DATEFIN', FinDeMois (DateD)) ;
   if okok then
      BEGIN
//Cumul individuel + cumul de reprise des 50 premiers cumuls
      For i:=1 to 50 do
          BEGIN
          Tot:= 0;
          Tot:= Tot+Valeur (Copy (S, 9+13*(i-1), 13));
          Tot:= Tot+Valeur (Copy (S3,9+13*(i-1), 13));
          CodeCum:= FormatFloat ('00', i);
          TD:= TOB_Cumuls.FindFirst (['COR'], [CodeCum], TRUE);
          while TD<>NIL do
                BEGIN
                Corresp:= TD.GetValue ('PCL_CUMULPAIE');
                if (Corresp<>'') then
                   BEGIN
                   H.PutValue ('PHC_CUMULPAIE', Corresp);
                   if ((CodeCum='07') or (CodeCum='08')) then
                      Tot:= -Tot;
                   H.PutValue ('PHC_MONTANT', Tot/10000.0);
                   okok:= H.InsertOrUpdateDb (FALSE);
                   END;
                TD:= TOB_Cumuls.FindNext (['COR'], [CodeCum], TRUE);
                END;
          END;
//Cumul individuel + cumul de reprise des 49 derniers cumuls
      For i:=1 to 49 do
          BEGIN
          Tot:= 0;
          Tot:= Tot+Valeur (Copy (S2, 9+13*(i-1), 13));
          Tot:= Tot+Valeur (Copy (S4, 9+13*(i-1), 13));
          CodeCum:= FormatFloat ('00', i+50);
          TD:= TOB_Cumuls.FindFirst (['COR'], [CodeCum], TRUE);
          while TD<>NIL do
                BEGIN
                Corresp:= TD.GetValue ('PCL_CUMULPAIE');
                if (Corresp<>'') then
                   BEGIN
                   H.PutValue ('PHC_CUMULPAIE', Corresp);
                   if ((CodeCum='07') or (CodeCum='08')) then
                      Tot :=-Tot;
                   H.PutValue ('PHC_MONTANT', Tot/10000.0);
                   okok:= H.InsertOrUpdateDb (FALSE);
                   END;
                TD:= TOB_Cumuls.FindNext (['COR'], [CodeCum], TRUE);
                END;
          END;
      END;
   END
else
//Si on a des cumul mois par mois, on reprend juste les cumuls de reprise
   BEGIN
   if (DateEntree>DebutExo) then
      DateD:= DebutDeMois (DateEntree)
   else
      DateD:= DebutDeMois (DebutExo);
   H.PutValue ('PHC_DATEDEBUT', DateD);
   H.PutValue ('PHC_DATEFIN', DateD);
   if okok then
      BEGIN
//Cumul de reprise des 50 premiers cumuls
      For i:=1 to 50 do
          BEGIN
          Tot:= Valeur (Copy (S3, 9+13*(i-1), 13));
          CodeCum:= FormatFloat ('00', i);
          TD:= TOB_Cumuls.FindFirst (['COR'], [CodeCum], TRUE);
          while TD<>NIL do
                BEGIN
                Corresp:= TD.GetValue ('PCL_CUMULPAIE');
                if (Corresp<>'') then
                   BEGIN
                   H.PutValue ('PHC_CUMULPAIE', Corresp);
                   if ((CodeCum='07') or (CodeCum='08')) then
                      Tot :=-Tot;
                   H.PutValue ('PHC_MONTANT', Tot/10000.0);
                   if (Tot<>0) then
                      okok:= H.InsertOrUpdateDb (FALSE);
                   END;
                TD:= TOB_Cumuls.FindNext (['COR'], [CodeCum], TRUE);
                END;
          END;
//Cumul de reprise des 49 derniers cumuls
      For i:=1 to 49 do
          BEGIN
          Tot:= Valeur (Copy (S4, 9+13*(i-1), 13));
          CodeCum:= FormatFloat ('00', i+50);
          TD:= TOB_Cumuls.FindFirst (['COR'], [CodeCum], TRUE);
          while TD<>NIL do
                BEGIN
                Corresp:= TD.GetValue ('PCL_CUMULPAIE');
                if (Corresp<>'') then
                   BEGIN
                   H.PutValue ('PHC_CUMULPAIE', Corresp);
                   if ((CodeCum='07') or (CodeCum='08')) then
                      Tot :=-Tot;
                   H.PutValue ('PHC_MONTANT', Tot/10000.0);
                   if (Tot<>0) then
                      okok:= H.InsertOrUpdateDb (FALSE);
                   END;
                TD:= TOB_Cumuls.FindNext (['COR'], [CodeCum], TRUE);
                END;
          END;
      END;
   END;

if (FTypeFichier.ItemIndex=0) then
   begin
   StMessage:= 'Cumul Salarié (S) '+Copy (S, 2, 6)+Verifokok (okok);
   FReport.Items.Add (StMessage);
   Writeln (FRapport, StMessage);
   end
else
   begin
   StMessage:= 'Cumul Salarié (PHC) '+Copy (S, 2, 6)+Verifokok (okok);
   FReport.Items.Add (StMessage);
   Writeln (FRapport,StMessage);
   end;

//Récupération des bases de cotisation
if (TRub<>nil) then
   begin
   CreerTOBFille(HB, Etab, S, S3, CatDADS, '1000');
   CreerTOBFille(HB, Etab, S, S3, CatDADS, '1020');
   CreerTOBFille(HB, Etab, S, S3, CatDADS, '1040');
   CreerTOBFille(HB, Etab, S, S3, CatDADS, '1042');
   CreerTOBFille(HB, Etab, S, S3, CatDADS, '1060');
   CreerTOBFille(HB, Etab, S, S3, CatDADS, '1062');
   CreerTOBFille(HB, Etab, S2, S4, CatDADS, '1070');           //PT68
   CreerTOBFille(HB, Etab, S, S3, CatDADS, '1080');
   CreerTOBFille(HB, Etab, S, S3, CatDADS, '1420');
   CreerTOBFille(HB, Etab, S, S3, CatDADS, '1500');
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 16/03/2001
Modifié le ... :   /  /
Description .. : Calcul des montants au niveau des bases de cotisation
Suite ........ : HB : TOB historique bulletin pour récupération des bases
Suite ........ : Etab : Numéro d'établissement
Suite ........ : S : Chaine de caractères N°1
Suite ........ : Rub : Rubrique de cotisation de base
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.CreerTOBFille(HB: TOB; Etab, S, S2, CatDADS, Rub : string);
var
Base, Plafond, Plafond1, Plafond2, Plafond3 : Double ;
TC, TRech : TOB ;
DateD : TDateTime;
CalculTranche : boolean;
begin
CalculTranche:= True;
TRech:= TRub.FindFirst (['PCT_RUBRIQUE'], [Rub], False);
if (TRech<>nil) then
   begin
   if (((TRech.GetValue ('PCT_TYPETRANCHE1')<>'NBR') and
      (TRech.GetValue ('PCT_TYPETRANCHE1')<>'')) or
      ((TRech.GetValue ('PCT_TYPETRANCHE2')<>'NBR') and
      (TRech.GetValue ('PCT_TYPETRANCHE2')<>'')) or
      ((TRech.GetValue ('PCT_TYPETRANCHE3')<>'NBR') and
      (TRech.GetValue ('PCT_TYPETRANCHE3')<>''))) then
      CalculTranche:= False;
   TC:= TOB.Create ('HISTOBULLETIN', HB, -1);
   TC.PutValue ('PHB_ETABLISSEMENT', Etab);
   TC.PutFixedStValue ('PHB_SALARIE', S, 2, 6, tcTrim, FALSE);
   ChangeCodeSalarie(TC, 'PHB', TRUE) ;
   if (Rub='1000') then
      begin //cumul 4 + cumul 3
      Base:= Valeur (Copy (S, 48, 13))/10000;
      Base:= Base+Valeur (Copy (S2, 48, 13))/10000;
      Plafond:= Valeur (Copy (S, 35, 13))/10000;
      Plafond:= Plafond+Valeur (Copy (S2, 35, 13))/10000;
      end
   else
   if (Rub='1020') then
      begin //cumul 5 + cumul 14
      Base:= Valeur (Copy (S, 61, 13))/10000;
      Base:= Base+Valeur (Copy (S2, 61, 13))/10000;
      Plafond:= Valeur (Copy (S, 178, 13))/10000;
      Plafond:= Plafond+Valeur (Copy (S2, 178, 13))/10000;
      end
   else
   if ((Rub='1040') and (CatDADS<>'001') and (CatDADS<>'002')) then
      begin //cumul 6 + cumul 15
      Base:= Valeur (Copy (S, 74, 13))/10000 ;
      Base:= Base+Valeur (Copy (S2, 74, 13))/10000;
      Plafond:= Valeur (Copy (S, 191, 13))/10000;
      Plafond:= Plafond+Valeur (Copy (S2, 191, 13))/10000;
      end
   else
   if ((Rub='1042') and (CatDADS<>'001') and (CatDADS<>'002')) then
      begin //cumul 6 + cumul 15
      Base:= Valeur (Copy (S, 74, 13))/10000 ;
      Base:= Base+Valeur (Copy (S2, 74, 13))/10000;
      Plafond:= Valeur (Copy (S, 191, 13))/10000;
      Plafond:= Plafond+Valeur (Copy (S2, 191, 13))/10000;
      end
   else
   if ((Rub='1060') and ((CatDADS='001') or (CatDADS='002'))) then
      begin //cumul 6 + cumul 15
      Base:= Valeur (Copy (S, 74,  13))/10000;
      Base:= Base+Valeur (Copy (S2,74, 13))/10000;
      Plafond:= Valeur (Copy (S, 191, 13))/10000;
      Plafond:= Plafond+Valeur (Copy (S2, 191, 13))/10000;
      end
   else
   if ((Rub='1062') and ((CatDADS='001') or (CatDADS='002'))) then
      begin //cumul 6 + cumul 15
      Base:= Valeur (Copy (S, 74, 13))/10000;
      Base:= Base+Valeur (Copy (S2, 74, 13))/10000;
      Plafond:= Valeur (Copy (S, 191, 13))/10000;
      Plafond:= Plafond+Valeur (Copy (S2, 191, 13))/10000;
      end
   else
//PT68
   if (Rub='1070') then
      begin //cumul 70 + cumul 71
      Base:= Valeur (Copy (S, 256, 13))/10000;
      Base:= Base+Valeur (Copy (S2, 256, 13))/10000;
      Plafond:= Valeur (Copy (S, 269, 13))/10000;
      Plafond:= Plafond+Valeur (Copy (S2, 269, 13))/10000;
      end
   else
//FIN PT68
   if (Rub='1080') then
      begin //cumul 4 + cumul 3
      Base:= Valeur (Copy (S, 48, 13))/10000;
      Base:= Base+Valeur (Copy (S2, 48, 13))/10000;
      Plafond:= Valeur (Copy (S, 35, 13))/10000;
      Plafond:= Plafond+Valeur (Copy (S2, 35, 13))/10000;
      end
   else
   if ((Rub='1420') and ((CatDADS='001') or (CatDADS='002'))) then
      begin //cumul 6 + cumul 15
      Base:= Valeur (Copy (S, 74, 13))/10000;
      Base:= Base+Valeur (Copy (S2, 74, 13))/10000;
      Plafond:= Valeur (Copy (S, 191, 13))/10000;
      Plafond:= Plafond+Valeur (Copy (S2, 191, 13))/10000;
      end
   else
   if (Rub='1500') then
      begin //cumul 29 + cumul 3
      Base:= Valeur (Copy (S, 373, 13))/10000;
      Base:= Base+Valeur (Copy (S2, 373, 13))/10000;
      Plafond:= Valeur (Copy (S, 35, 13))/10000;
      Plafond:= Plafond+Valeur (Copy (S2, 35, 13))/10000;
      end
   else
      begin
      Base:= 0;
      Plafond:= 0;
      end;

   if (Ouvert='O') then
      DateD:= DebutDeMois (PlusMois (FinExo, -NbMois))
   else
      DateD:= DebutDeMois (PlusMois (FinExo, -NbMois+1));
   TC.PutValue ('PHB_DATEDEBUT', DateD);
   TC.PutValue ('PHB_DATEFIN', FinDeMois (PlusMois (DateD, NbMois-1)));
   TC.PutValue ('PHB_IMPRIMABLE', '-');
   TC.PutValue ('PHB_BASEREM', '0');
   TC.PutValue ('PHB_TAUXREM', '0');
   TC.PutValue ('PHB_COEFFREM', '0');
   TC.PutValue ('PHB_MTREM', '0');
   TC.PutValue ('PHB_BASECOT', Base);
   TC.PutValue ('PHB_TAUXSALARIAL', '0');
   TC.PutValue ('PHB_MTSALARIAL', '0');
   TC.PutValue ('PHB_TAUXPATRONAL', '0');
   TC.PutValue ('PHB_MTPATRONAL', '0');
   TC.PutValue ('PHB_BASEREMIMPRIM', '-');
   TC.PutValue ('PHB_TAUXREMIMPRIM', '-');
   TC.PutValue ('PHB_COEFFREMIMPRIM', '-');
   TC.PutValue ('PHB_PLAFOND', Plafond);
   TC.PutValue ('PHB_TRANCHE3', 0);
   TC.PutValue ('PHB_CONSERVATION', 'PRO');
   TC.PutValue ('PHB_SENSBUL', 'P');
   TC.PutValue ('PHB_CONFIDENTIEL', '0');

   if ((Plafond*Valeur (TRech.GetValue ('PCT_TRANCHE1'))<Base) and
      (CalculTranche=True)) then
      begin
      TC.PutValue ('PHB_TRANCHE1', Plafond*Valeur (TRech.GetValue ('PCT_TRANCHE1')));
      TC.PutValue ('PHB_TRANCHE2', Base-(Plafond*Valeur (TRech.GetValue ('PCT_TRANCHE1'))));
      end
   else
      begin
      if (CalculTranche=True) then
         TC.PutValue ('PHB_TRANCHE1', Base)
      else
         TC.PutValue ('PHB_TRANCHE1', 0);
      TC.PutValue ('PHB_TRANCHE2', 0);
      end;
   TC.PutValue ('PHB_NATURERUB', TRech.GetValue ('PCT_THEMECOT'));
   TC.PutValue ('PHB_RUBRIQUE', TRech.GetValue ('PCT_RUBRIQUE'));
   TC.PutValue ('PHB_LIBELLE', TRech.GetValue ('PCT_LIBELLE'));
   TC.PutValue ('PHB_BASECOTIMPRIM', TRech.GetValue ('PCT_BASEIMP'));
   TC.PutValue ('PHB_TAUXSALIMPRIM', TRech.GetValue ('PCT_TXSALIMP'));
   TC.PutValue ('PHB_TAUXPATIMPRIM', TRech.GetValue ('PCT_TXPATIMP'));
   TC.PutValue ('PHB_ORGANISME', TRech.GetValue ('PCT_ORGANISME'));
   Plafond1:= 0;
   Plafond2:= 0;
   Plafond3:= 0;
   if (CalculTranche=True) then
      begin
      Plafond1:= Plafond*Valeur (TRech.GetValue ('PCT_TRANCHE1'));
      if (TRech.GetValue ('PCT_TRANCHE2')<>'') then
         Plafond2:= Plafond*Valeur (TRech.GetValue ('PCT_TRANCHE2'))-Plafond1;
      if (TRech.GetValue ('PCT_TRANCHE3')<>'') then
         Plafond3:= Plafond*Valeur (TRech.GetValue ('PCT_TRANCHE3'))-Plafond2
                    -Plafond1;
      end;
   TC.PutValue ('PHB_PLAFOND1', Plafond1);
   TC.PutValue ('PHB_PLAFOND2', Plafond2);
   TC.PutValue ('PHB_PLAFOND3', Plafond3);
   TC.PutValue ('PHB_ORDREETAT', Valeur (TRech.GetValue ('PCT_ORDREETAT')));
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Modification du code du salarié en fonction du choix de
Suite ........ : l'utilisateur
Suite ........ : T : TOB de la table à modifier
Suite ........ : Préfixe : Préfixe de la table
Suite ........ : PeutCréer : Si = True, création du salarié possible dans la
Suite ........ : table de correspondance
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpISIS.ChangeCodeSalarie (T : TOB; Prefixe : string ;
                                      PeutCreer : boolean ) : boolean;
Var
Sal, St : String ;
Code, LongSal : Integer ;
begin
Result:=TRUE ;
Case FCodeSal.ItemIndex of
   0 : BEGIN //Sans changement du code
       St := Trim(T.GetValue(Prefixe+'_SALARIE'));
       if (VH_Paie.PgTypeNumSal = 'NUM') then
          begin
          if (St <> '') then
             St := ColleZeroDevant(StrToInt(St), 10)
          else
             begin
             PGIBox ('Le matricule n''est pas renseigné. Ce salarié#10#13'+
                    'n''est pas pris en compte dans l''import.#10#13'+
                    'Vous pouvez refaire l''import en choisissant #10#13'+
                    '"Réaffectation par ordre chronologique"', StTitre);   //PT64
             Result := FALSE;
             Exit;
             end;
          end;

       Result := SalarieDansListe(St, PeutCreer, Code);
       if Result = True then
          T.PutValue(Prefixe+'_SALARIE',St);
       END;

   1 : BEGIN //Renumérotation
       if ((FTypeFichier.ItemIndex = 0) or (FTypeFichier.ItemIndex = 2)) then
          St:= Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'))+
               Trim(T.GetValue(Prefixe+'_SALARIE'))
       else
          St:= Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'))+
               Trim(T.GetValue(Prefixe+'_SALARIE'))+
               Trim(T.GetValue(Prefixe+'_NUMEROSS'));
       Result := SalarieDansListe(St, PeutCreer, Code);
       if Result = True then
          T.PutValue(Prefixe+'_SALARIE',FormatFloat('0000000000',Code));
       END ;

   2 : BEGIN //Etab+Sal
       case FTypeFichier.ItemIndex of
            0 : St:= Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'))+
                     Trim(T.GetValue(Prefixe+'_SALARIE'));

            1 : BEGIN
                Sal := Trim(T.GetValue(Prefixe+'_SALARIE'));
                LongSal:=Length(Sal);
                if (LongSal > 7) then
                   St:=Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'))+
                       Copy(Trim(Sal), LongSal+1-7, 7)
                else
                   St:=Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'))+
                       ColleZeroDevant(StrToInt(Copy(trim(Sal), 1, 7)),7);
                END;

            2 : St:= Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'))+
                     Trim(T.GetValue(Prefixe+'_SALARIE'));

//PT59
            3 : BEGIN
                Sal := Trim(T.GetValue(Prefixe+'_SALARIE'));
                LongSal:=Length(Sal);
                if (LongSal > 7) then
                   St:=Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'))+
                       Copy(Trim(Sal), LongSal+1-7, 7)
                else
                   St:=Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'))+
                       ColleZeroDevant(StrToInt(Copy(trim(Sal), 1, 7)),7);
                END;
//FIN PT59
            END;
       if (VH_Paie.PgTypeNumSal = 'NUM') then
          St := ColleZeroDevant(StrToInt(St), 10);
       Result := SalarieDansListe(St, PeutCreer, Code);
       if Result = True then
          T.PutValue(Prefixe+'_SALARIE',St);
       END ;

   3 : BEGIN //Sal+Etab
       case FTypeFichier.ItemIndex of
            0 : St:= Trim(T.GetValue(Prefixe+'_SALARIE'))+
                     Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'));

            1 : BEGIN
                Sal := Trim(T.GetValue(Prefixe+'_SALARIE'));
                LongSal:=Length(Sal);
                if (LongSal > 7) then
                   St:= Copy(Trim(Sal), LongSal+1-7, 7)+
                        Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'))
                else
                   St:= ColleZeroDevant(StrToInt(Copy(trim(Sal), 1, 7)), 7)+
                        Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'));
                END;

            2 : St:= Trim(T.GetValue(Prefixe+'_SALARIE'))+
                     Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'));

//PT59
            3 : BEGIN
                Sal := Trim(T.GetValue(Prefixe+'_SALARIE'));
                LongSal:=Length(Sal);
                if (LongSal > 7) then
                   St:= Copy(Trim(Sal), LongSal+1-7, 7)+
                        Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'))
                else
                   St:= ColleZeroDevant(StrToInt(Copy(trim(Sal), 1, 7)), 7)+
                        Trim(T.GetValue(Prefixe+'_ETABLISSEMENT'));
                END;
//FIN PT59
            END;
       if (VH_Paie.PgTypeNumSal = 'NUM') then
          St := ColleZeroDevant(StrToInt(St), 10);
       Result := SalarieDansListe(St, PeutCreer, Code);
       if Result = True then
          T.PutValue(Prefixe+'_SALARIE',St);
       END ;
   END ;
Index := Index + 1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/10/2000
Modifié le ... :   /  /
Description .. : Initialisation de la table de correspondance des codes stat
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.InitCodeStat;
Var
Q : TQuery ;
MaxCodeStat : Integer ;
begin
ListeCodeStat:=TStringList.Create ;
ListeCodeStat.Sorted:=TRUE ;
MaxCodeStat:=1 ;
Q:= OpenSQL('SELECT MAX(CC_CODE) FROM CHOIXCOD WHERE CC_TYPE = "PSQ"',TRUE) ;
if Not Q.EOF then
   try
   MaxCodeStat:= ValeurI(Q.Fields[0].AsString)+1 ;
   except
         on E: EConvertError do
            MaxCodeStat:= 1;
   end;
Ferme(Q) ;
ListeCodeStat.AddObject('_______',tobject(MaxCodeStat)) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/10/2000
Modifié le ... :   /  /
Description .. : Libération de la table de correspondance des codes stat
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.FinCodeStat;
begin
ListeCodeStat.Free ;
end;

//PT21
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 04/03/2003
Modifié le ... : 13/02/2004
Description .. : Chargement des codes salariés dans une TOB
Suite ........ : Codes retour :
Suite ........ : 0 : Tout OK
Suite ........ : -1 : Problème au niveau de l'établissement
Suite ........ : -2 : Matricule salarié=''
Suite ........ : -3 : Matricule salarié déjà existant
Suite ........ : -4 : Matricule salarié numérique > 2147483647
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpIsis.ChargeSalariesDansTOB : integer;
var
Etab, S, S1, St, StSal : string;
QRechSal : TQuery;
TSal, TSalD, TSalFille : TOB;
begin
Result := 0;
BSuivant.Enabled:=FALSE ;
BPrecedent.Enabled:=FALSE ;

StSal:='SELECT PSA_SALARIE FROM SALARIES';
QRechSal:=OpenSql(StSal,TRUE);
TSal := TOB.Create('Les salaries', NIL, -1);
TSal.LoadDetailDB('SAL','','',QRechSal,False);
Ferme(QRechSal);

AssignFile(F, edtFichier.Text);
Reset(F);

InitMove(1000,'') ;
while not Eof(F) do
      begin
      MoveCur(False) ;
      Readln(F, S);
      if Trim(S)='' then
         continue ;
      case S[1] of
           'D': begin //Etablissement 1
                Etab := EnregEtab (S);
                if (Etab = '-1') then
                   begin
                   FreeAndNil(TSal);
                   FiniMove ;
                   CloseFile(F);
                   Result := -1;
                   Exit;
                   end;
                end;
           'T': begin //Salarié
                S1:=S ;
                Readln(F, S);
                S1:=S1+S ;
                Readln(F, S);
                S1:=S1+S ;

                Case FCodeSal.ItemIndex of
                     0 : St := Trim(Copy(S1, 2, 6)); //Sans changement du code
                     2 : St := Trim(Etab)+Trim(Copy(S1, 2, 6));
                     3 : St := Trim(Copy(S1, 2, 6))+Trim(Etab);
                     end;

                if (VH_Paie.PgTypeNumSal = 'NUM') then
                   begin
//PT38-1
                   St := ColleZeroDevant(StrToInt(St), 10);    //PT41
                   if (St >'2147483647') then
                      result := -4
                   else
                      if (St = '') then
                         Result := -2;
//FIN PT38-1
                   end;

                TSalD := TSal.FindFirst(['PSA_SALARIE'], [St], TRUE);
                if (TSalD <> nil) then
                   Result := -3
                else
                   begin
                   TSalFille := TOB.Create('',TSal,-1);
                   TSalFille.AddChampSup('PSA_SALARIE', FALSE);
                   TSalFille.PutValue('PSA_SALARIE', St);
                   end;
                end;
           end;
      end;
FreeAndNil(TSal);
FiniMove ;
CloseFile(F);
END;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Changement de page
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.PChange(Sender: TObject);
var
ControleOK : integer;
begin
inherited;

ControleOK := 0;
if P.ActivePage = PCodageSal then
   BEGIN
   case FTypeFichier.ItemIndex of
        0 : BEGIN
            FCodeSal.Visible := TRUE;
            TLCodage.caption := 'Précisez la façon de coder les salariés'+
                                ' (ATTENTION en cas de codage chronologique,'+
                                ' vous ne pourrez pas importer deux fois le'+
                                ' même fichier)';
            END;

        1 : BEGIN
            if (VH_Paie.PgTypeNumSal = 'NUM') then
               BEGIN
               FCodeSal.Visible := TRUE;
               TLCodage.caption := 'Précisez la façon de coder les salariés .'+
                                   ' ATTENTION en cas de codage chronologique,'+
                                   ' si vous importez deux fois le même'+
                                   ' fichier, les salariés seront doublés';
               END
            else
               BEGIN
               FCodeSal.Visible := FALSE;
               TLCodage.caption := 'Vous avez choisi de numéroter les salariés'+
                                   ' avec un code alphanumérique . Les'+
                                   ' matricules des salariés seront inchangés.'+
                                   ' ATTENTION, chaque code salarié dans le'+
                                   ' fichier DADS doit être unique';
               END;

            END;

        2 : BEGIN
            FCodeSal.Visible := TRUE;
            TLCodage.caption := 'Précisez la façon de coder les salariés'+
                                ' (ATTENTION en cas de codage chronologique,'+
                                ' vous ne pourrez pas importer deux fois le'+
                                ' même fichier)';
            END;
//PT59
        3 : BEGIN
            if (VH_Paie.PgTypeNumSal = 'NUM') then
               BEGIN
               FCodeSal.Visible := TRUE;
               TLCodage.caption := 'Précisez la façon de coder les salariés .'+
                                   ' ATTENTION en cas de codage chronologique,'+
                                   ' si vous importez deux fois le même'+
                                   ' fichier, les salariés seront doublés';
               END
            else
               BEGIN
               FCodeSal.Visible := FALSE;
               TLCodage.caption := 'Vous avez choisi de numéroter les salariés'+
                                   ' avec un code alphanumérique . Les'+
                                   ' matricules des salariés seront inchangés.'+
                                   ' ATTENTION, chaque code salarié dans le'+
                                   ' fichier DADS-U doit être unique';
               END;

            END;
//FIN PT59
        END;
   END;

if P.ActivePage=PCumuls then
   BEGIN
   if ((FTypeFichier.ItemIndex = 0) or (FTypeFichier.ItemIndex = 2)) then
      BEGIN
      ChargeCumuls;
      if (FCodeSal.ItemIndex <> 1) then
         ControleOK := ChargeSalariesDansTOB;
      if (ControleOK = -1) then
         begin
         BPrecedent.Click;
         Exit;
         end;
      if (ControleOK = -2) then
         begin
         PGIBox ('Attention, ce type de codage n''est pas permis car il#10#13'+
                 'existe un code salarié non renseigné', StTitre);   //PT64
         BPrecedent.Click;
         Exit;
         end;

      if (ControleOK = -3) then
         begin
         if (FTypeFichier.ItemIndex = 0) then
            PGIBox ('Attention, ce type de codage n''est pas permis car:#10#13'+
                    '- Soit il existe plusieurs salariés ayant le même#10#13'+
                    '  matricule dans le fichier provenant d''ISIS II,#10#13'+
                    '#10#13'+
                    '- Soit il existe déjà des salariés ayant le même#10#13'+
                    '  matricule dans la base PGI.#10#13'+
                    '  Si vous voulez réimporter des données provenant#10#13'+
                    '  d''ISIS II pour des salariés déjà importés dans#10#13'+
                    '  la base PGI, vous devez refaire, dans ISIS II,#10#13'+
                    '  un fichier d''archivage pour migration vers#10#13'+
                    '  Paie PGI en répondant "N" à la question#10#13'+
                    '  "Reprise des Salariés"', StTitre)              //PT64
         else
            PGIBox ('Attention, ce type de codage n''est pas permis car:#10#13'+
                    '- Soit il existe plusieurs salariés ayant le même#10#13'+
                    '  matricule dans le fichier,#10#13'+
                    '#10#13'+
                    '- Soit il existe déjà des salariés ayant le même#10#13'+
                    '  matricule dans la base PGI.#10#13', StTitre);         //PT64
         BPrecedent.Click;
         Exit;
         end;

      if (ControleOK = -4) then
         begin
         PGIBox ('Attention, ce type de codage n''est pas permis car un#10#13'+
                 'matricule numérique ne doit pas excéder "2147483647".#10#13'+
                 'Sélectionnez une autre façon de coder les salariés', StTitre);  //PT64
         BPrecedent.Click;
         Exit;
         end;
      BSuivant.Enabled:=TRUE ;
      BPrecedent.Enabled:=TRUE ;
      FCumul.Checked := TRUE;
      FCumul.Visible := TRUE;
      FCumul.Enabled := TRUE;
      BCumuls.Visible := TRUE;
      BCumuls.Enabled := TRUE;
      GCumuls.Visible := TRUE;
      GCumuls.Enabled := TRUE;
      FCumulDest.Enabled := TRUE;
      Label_cumul.Visible := FALSE;
      GCumuls.Row := 1;
      GCumuls.ShowCombo(2,1);
      END;
{PT59
   if (FTypeFichier.ItemIndex = 1) then
}
   if ((FTypeFichier.ItemIndex = 1) or (FTypeFichier.ItemIndex = 3)) then
      BEGIN
      FCumul.Checked := FALSE;
      FCumul.Visible := FALSE;
      FCumul.Enabled := FALSE;
      BCumuls.Visible := FALSE;
      BCumuls.Enabled := FALSE;
      GCumuls.Visible := FALSE;
      GCumuls.Enabled := FALSE;
      FCumulDest.Enabled := FALSE;
      Label_cumul.Visible := TRUE;
      if (VH_Paie.PgTypeNumSal = 'NUM') then
         BEGIN
         case FCodeSal.ItemIndex of
              0 : Label_cumul.caption := 'ATTENTION, vous avez choisi de ne'+
                                         ' pas changer les matricules des'+
                                         ' salariés . Chaque matricule salarié'+
                                         ' doit être unique dans tout le'+
                                         ' dossier . Cliquez sur <Suivant>'+
                                         ' pour passer à la récupération des'+
                                         ' données';

              1 : Label_cumul.caption := 'Vous avez choisi de renuméroter les'+
                                         ' matricules . Si des salariés ont'+
                                         ' déjà été créés, la numérotation des'+
                                         ' salariés récupérés se fait par'+
                                         ' incrémentation du numéro de'+
                                         ' matricule le plus élevé . Cliquez'+
                                         ' sur <Suivant> pour passer à la'+
                                         ' récupération des données';

              2 : Label_cumul.caption := 'ATTENTION, vous avez choisi de coder'+
                                         ' les matricules des salariés avec le'+
                                         ' code établissement qui doit être'+
                                         ' sur 3 caractères NUMERIQUES . Dans'+
                                         ' le fichier DADS, chaque matricule'+
                                         ' salarié doit être unique sur les 7'+
                                         ' derniers caractères . Cliquez sur'+
                                         ' <Suivant>';

              3 : Label_cumul.caption := 'ATTENTION, vous avez choisi de coder'+
                                         ' les matricules des salariés avec le'+
                                         ' code établissement qui doit être'+
                                         ' sur 3 caractères NUMERIQUES . Dans'+
                                         ' le fichier DADS, chaque matricule'+
                                         ' salarié doit être unique sur les 7'+
                                         ' derniers caractères . Cliquez sur'+
                                         ' <Suivant>';
              END;
         END
      else
         Label_cumul.caption := 'Cliquez sur <Suivant> pour passer à la'+
                                ' récupération des données';
      END;
   END;

if P.ActivePage = PLancer then
   BEGIN
   if ((FTypeFichier.ItemIndex = 0) or (FTypeFichier.ItemIndex = 2)) then
      VerifCumuls;
   END;

//PT67-1
if (P.ActivePage = PReport) then
   begin
   bFin.Visible:= True;
   bAnnuler.Visible:= False;
   end;
//FIN PT67-1
END;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... : 26/09/2002
Description .. : Chargement de la table de correspondance des cumuls
Suite ........ : PT13-2 : Fonction modifiée dans sa totalité
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.ChargeCumuls;
var
TCor, TD, TMereCumul : TOB ;
okok : boolean ;
i : integer ;
PourVoir, St, StCumul : String;
QRechCumul : TQuery;
begin
if TOB_Cumuls<>Nil then
   FreeAndNil(TOB_Cumuls);
{PT13-2
TOB_Cumuls := TOB.Create('les cumuls',Nil,-1);
}
TMereCumul := TOB.Create('les cumuls',Nil,-1);
AssignFile(F, edtFichier.Text);
Reset(F);
InitMove(100, '');
okok := FALSE;
while not Eof(F) do
      begin
      MoveCur(False) ;
      Readln(F, St);
{PT25
      if (Trim(St)<> '') and (St[1]='R') and (Copy(St,2,2)<>'00') then
}
      if (((FTypeFichier.ItemIndex = 0) and (Trim(St)<> '') and
         (St[1]='R') and (Copy(St,2,2)<>'00')) or
         ((FTypeFichier.ItemIndex = 2) and (Trim(St)<> '') and
         (Copy(St, 1, 6) ='***PCL') and (Copy(St,7,2)<>'00'))) then
         BEGIN
         TD:=TOB.Create ('COMMUN', TMereCumul, -1);
         TD.PutValue('CO_TYPE','PAM');
//PT25
         if (FTypeFichier.ItemIndex = 0) then
            begin
            TD.PutFixedStValue('CO_CODE',St,2,2,tcTrim,FALSE);
            TD.PutFixedStValue('CO_LIBELLE',St,4,25,tcTrim,FALSE);
            end
         else
            begin
            TD.PutFixedStValue('CO_CODE',St,7,2,tcTrim,FALSE);
            TD.PutFixedStValue('CO_LIBELLE',St,9,25,tcTrim,FALSE);
            end;
//FIN PT25
            
{
         TD:=TOB.Create('le cumul ISIS',TOB_Cumuls,-1) ;
         TD.AddChampSup('COD',FALSE) ; TD.PutFixedStValue('COD',St,2,2,tcTrim,FALSE) ;
         TD.AddChampSup('LIB',FALSE) ; TD.PutFixedStValue('LIB',St,4,25,tcTrim,FALSE) ;
         TD.AddChampSup('ABR',FALSE) ; TD.PutFixedStValue('ABR',St,29,11,tcTrim,FALSE) ;
         TD.AddChampSup('TYP',FALSE) ; TD.PutFixedStValue('TYP',St,40,1,tcTrim,FALSE) ;
         TD.AddChampSup('RAZ',FALSE) ; TD.PutFixedStValue('RAZ',St,46,2,tcTrim,FALSE) ;
         TD.AddChampSup('QTE',FALSE) ; TD.PutFixedStValue('QTE',St,86,1,tcEnum,FALSE,'**=99') ;
         TD.AddChampSup('COR',FALSE) ; TD.PutFixedStValue('COR',St,2,2,tcEnum,FALSE,
         '=;01=01;02=02;03=30;04=31;05=32;07=07;08=08;09=09;10=10;13=20;14=36;15=38') ;
}
         okok:=TRUE ;
         END
      else
         BEGIN
         if (okok) and (St<>'') then
            Break ;
         END ;
      end;
TMereCumul.InsertOrUpdateDB(TRUE);
AvertirTable('PGTABLETEMPO');
FCumulDest.ReLoad;
FiniMove ;
CloseFile(F);
TCor := TOB.Create('les correspondances',Nil,-1);
CreeFilleCumul(TCor,'Correspondance01','01','01');
CreeFilleCumul(TCor,'Correspondance02','02','02');
CreeFilleCumul(TCor,'Correspondance07','07','07');
CreeFilleCumul(TCor,'Correspondance08','08','08');
CreeFilleCumul(TCor,'Correspondance09','09','09');
CreeFilleCumul(TCor,'Correspondance10','10','10');
CreeFilleCumul(TCor,'Correspondance11','11','11');
CreeFilleCumul(TCor,'Correspondance13','13','60');
CreeFilleCumul(TCor,'Correspondance15','15','59');
CreeFilleCumul(TCor,'Correspondance20','20','13');
CreeFilleCumul(TCor,'Correspondance21','21','13');
CreeFilleCumul(TCor,'Correspondance22','22','13');
CreeFilleCumul(TCor,'Correspondance30','30','03');
CreeFilleCumul(TCor,'Correspondance31','31','04');
CreeFilleCumul(TCor,'Correspondance32','32','05');
CreeFilleCumul(TCor,'Correspondance33','33','06');
CreeFilleCumul(TCor,'Correspondance34','34','06');
CreeFilleCumul(TCor,'Correspondance35','35','58');
CreeFilleCumul(TCor,'Correspondance36','36','14');
CreeFilleCumul(TCor,'Correspondance37','37','15');
CreeFilleCumul(TCor,'Correspondance38','38','15');
CreeFilleCumul(TCor,'Correspondance40','40','04');
{
TOB_Cumuls.PutGridDetail(GCumuls,FALSE,FALSE,'COD;LIB',TRUE) ;
}
StCumul := 'SELECT PCL_CUMULPAIE, PCL_LIBELLE'+
           ' FROM CUMULPAIE WHERE'+
           ' ##PCL_PREDEFINI## PCL_CUMULPAIE <> "A"';
QRechCumul:=OpenSQL(StCumul,TRUE);
TOB_Cumuls:=TOB.Create('les cumuls',Nil,-1) ;
TOB_Cumuls.LoadDetailDB('CUMUL_PAIE', '', '', QRechCumul, FALSE);
Ferme(QRechCumul);
TOB_Cumuls.Detail.Sort('PCL_CUMULPAIE');  //PT19
TOB_Cumuls.PutGridDetail(GCumuls,FALSE,FALSE,'PCL_CUMULPAIE;PCL_LIBELLE',TRUE) ;
for i:=1 to GCumuls.RowCount-1 do
    BEGIN
{
    TD:=TOB_Cumuls.Detail[i-1] ;
    PourVoir := TD.GetValue('COR');
    GCumuls.Cells[2,i]:=RechDom('PGCUMULPAIE',TD.GetValue('COR'),FALSE) ;
}
    PourVoir := GCumuls. Cells[0,i];
    TD:=TCor.FindFirst(['PGI'], [PourVoir], FALSE);
    if TD <> nil then
       GCumuls.Cells[2,i]:=TD.GetValue('ISIS');
    END ;
FreeAndNil(TMereCumul);
FreeAndNil(TCor);
//FIN PT13-2
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 25/09/2002
Modifié le ... :   /  /
Description .. : Création de e la tob fille des correspondances entre
Suite ........ : cumul PGI et cumul ISIS
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.CreeFilleCumul (var Mere : TOB; NomFille, CumulPGI,
                                    CumulISIS:string);
var
TDCor : TOB;
Libelle : string;
begin
TDCor := TOB.Create(NomFille, Mere, -1);
TDCor.AddChampSup('PGI',FALSE);
TDCor.PutValue('PGI',CumulPGI);
TDCor.AddChampSup('ISIS',FALSE);
Libelle := Rechdom('PGTABLETEMPO',CumulISIS,FALSE);
TDCor.PutValue('ISIS',Libelle);
end;


procedure TFImpIsis.FormCreate(Sender: TObject);
var
StEtab : string;
QRechEtab : TQuery ;
begin
inherited;
TOB_Cumuls:= Nil;
FCodeSal.ItemIndex:= 0;
if (VH_Paie.PgTypeNumSal <> 'NUM') then
   begin
   StEtab:= 'SELECT ET_ETABLISSEMENT FROM ETABLISS';
   QRechEtab:= OpenSQL(StEtab,TRUE);
   QRechEtab.First ;
   while not QRechEtab.Eof do
         begin
         if (((QRechEtab.FindField ('ET_ETABLISSEMENT').Asstring < '000') or
            (QRechEtab.FindField ('ET_ETABLISSEMENT').Asstring > '999')) and
            (FCodeSal.Items.Count = 4)) then
            begin
            FCodeSal.Items.Delete(3);
            FCodeSal.Items.Delete(2);
            end;
         QRechEtab.Next;
         end;
   Ferme(QRechEtab);
   end;
StTitre:= 'Import Fichier CEGID ISIS II';   //PT64
end;

procedure TFImpIsis.FormDestroy(Sender: TObject);
var
StDelete : string;
begin
  inherited;
FreeAndNil(TOB_Cumuls);

//PT13-2
StDelete := 'DELETE FROM COMMUN WHERE'+
            ' CO_TYPE="PAM"';
ExecuteSQL(StDelete);
//FIN PT13-2
end;

procedure TFImpIsis.BCumulsClick(Sender: TObject);
begin
  inherited;
AGLLanceFiche('PAY','CUMUL','','', ' ') ;
// PT2-2 AGLLanceFiche('PAY','FCUMUL','','', ' ') ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Vérification : table de correspondance alimentée pour tous
Suite ........ : les cumuls
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.VerifCumuls;
var
i: Integer;
GaffeVide: Boolean;
TD: TOB;
begin
GaffeVide:=FALSE ;
for i:=1 to GCumuls.RowCount-1 do
   BEGIN
   FCumulDest.Libelle:=GCumuls.Cells[2,i] ;
   TD:=TOB_Cumuls.Detail[i-1] ;
   TD.AddChampSup('COR',FALSE) ;          //PT13-2
   TD.PutValue('COR',FCumulDest.Value) ;
   if FCumulDest.Value='' then GaffeVide:=TRUE ;
   END ;
//PGVisuUnObjet(TOB_Cumuls, '', '');

if GaffeVide and FCumul.Checked then
   BEGIN
   if Msg.Execute(2,'','')<>mrYes then
      BEGIN
      bPrecedentClick(Nil) ;
      END ;
   END ;
end;

procedure TFImpIsis.FormShow(Sender: TObject);
begin
inherited;
bAnnuler.Visible:= True;	//PT67-1
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/08/2000
Modifié le ... :   /  /
Description .. : Enregistrement cumuls mois par mois
Suite ........ : H : TOB cumuls salariés
Suite ........ : S : Chaine de caractères
Suite ........ : Etab : Numéro d'établissement
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregHisto(H: TOB; S,Etab: string);
Var
i : Integer ;
Tot : Double ;
okok : boolean ;
TD : TOB ;
CodeCum, Corresp, Cumul : String ;
DebutMois, FinMois : TDateTime;
begin
H.InitValeurs;
H.PutFixedStValue('PHC_SALARIE', S,2,6,tcTrim,FALSE);
H.PutValue('PHC_ETABLISSEMENT',Etab);
H.PutValue('PHC_REPRISE','X');
okok:=ChangeCodeSalarie(H, 'PHC', TRUE) ;
if okok then
   BEGIN
   Cumul := Copy(S,9,2);
   if (Cumul = 'CS') then
      BEGIN
      For i:=1 to 12 do
          BEGIN
          Tot:=Valeur(Copy(S,14+13*(i-1),13)) ;
          if (Tot <> 0) then
             BEGIN
             CodeCum:=Copy(S,11,2) ;
             TD:=TOB_Cumuls.FindFirst(['COR'],[CodeCum],TRUE) ;
             while TD<>NIL do
                   BEGIN
                   Corresp:=TD.GetValue('PCL_CUMULPAIE') ;
                   if Corresp<>'' then
                      BEGIN
{PT57
                      if (DebutExo <> null) then
}
                      if (DebutExo <> IDate1900) then
                         BEGIN
                         DebutMois := DebutDeMois(PlusMois(DebutExo, i-1));
                         FinMois := FinDeMois (PlusMois(DebutExo, i-1));
                         END
                      else
                         BEGIN
{PT57
                         DebutMois := null;
                         FinMois := null;
}
                         DebutMois := IDate1900;
                         FinMois := IDate1900;
//FIN PT57
                         END;
                      H.PutValue('PHC_DATEDEBUT',DebutMois) ;
                      H.PutValue('PHC_DATEFIN',FinMois) ;
                      H.PutValue('PHC_CUMULPAIE',Corresp);
                      if ((CodeCum = '07') or (CodeCum = '08')) then
                         Tot := -Tot;
                      H.PutValue('PHC_MONTANT',Tot/10000.0);
                      H.InsertOrUpdateDb(FALSE);
                      END;
                   TD:=TOB_Cumuls.FindNext(['COR'],[CodeCum],TRUE) ;
                   END ;
             END;
          END ;
      END;
   END ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/10/2000
Modifié le ... :   /  /
Description .. : Enregistrement codes statistiques
Suite ........ : St : Chaine de caractères
Suite ........ : QRechStat : Liste des codes stat existants
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregStat(St: string);
Var
okok : boolean;
Abrege, St2 : string;
TStat : TOB;
Code, i : Integer ;
begin
TStat := TOB.Create('CHOIXCOD', NIL, -1);
TStat.PutValue('CC_TYPE', 'PSQ');
TStat.PutFixedStValue('CC_CODE',St,2,4,tcTrim,FALSE);
TStat.PutFixedStValue('CC_LIBELLE',St,6,25,tcTrim,FALSE);
Abrege := Copy(St, 2, 4);
TStat.PutValue('CC_ABREGE',Abrege);

//PT63
if (not QRechStat.EOF) then
   begin
//FIN PT63

   QRechStat.First ;
   While Not QRechStat.EOF do
         BEGIN
         if (Abrege = QRechStat.FindField('CC_ABREGE').Asstring) then
            break;
         QRechStat.Next ;
         END ;
   end;	 //PT63

if Not QRechStat.EOF then
   Code:=ValeurI(QRechStat.FindField('CC_CODE').AsString)
else
   BEGIN
   St2:=Trim(TStat.GetValue('CC_TYPE'))+Trim(TStat.GetValue('CC_ABREGE')) ;
   i:=ListeCodeStat.IndexOf(St2) ;
   if i<0 then
      BEGIN
      Code:=Longint(ListeCodeStat.Objects[0]);
      ListeCodeStat.AddObject(St2,tobject(Code)) ;
      ListeCodeStat.Objects[0]:=tobject(Code+1) ;
      END
   else
      BEGIN
      Code:=longint(ListeCodeStat.Objects[i]) ;
      END ;
   END;
TStat.PutValue('CC_CODE',FormatFloat('000',Code));
okok:=TStat.InsertOrUpdateDb(FALSE);
//PT25
if (FTypeFichier.ItemIndex = 0) then
   begin
   StMessage:='Code Stat (N) '+Copy(St,2,4)+Verifokok(okok);
   FReport.Items.Add(StMessage);
   Writeln(FRapport,StMessage);
   end
else
   begin
   StMessage:='Code Stat (PCC) '+Copy(St,2,4)+Verifokok(okok);
   FReport.Items.Add(StMessage);
   Writeln(FRapport,StMessage);
   end;
//FIN PT25
FreeAndNil(TStat) ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/10/2000
Modifié le ... :   /  /
Description .. : Enregistrement codes conventions collectives
Suite ........ : St : Chaine de caractères
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregConvention(St: string);
Var
TC : TOB ;
okok, exist : boolean ;
CodeConv, StConv : String;
begin
CodeConv := Copy(St, 3, 3);
//PT20
if (CodePair(CodeConv) = True) then
   exit;
//FIN PT20
StConv:='SELECT PCV_LIBELLE'+
        ' FROM CONVENTIONCOLL WHERE'+
        ' ##PCV_PREDEFINI## PCV_CONVENTION = "'+CodeConv+'"';
exist := ExisteSQL(StConv);
if (exist=FALSE) then
   begin
   TC:=TOB.Create ('CONVENTIONCOLL', NIL,-1);
   TC.PutFixedStValue('PCV_CONVENTION',St,3,3,tcTrim,FALSE);
   TC.PutValue('PCV_PREDEFINI','STD');
   TC.PutValue('PCV_NODOSSIER','000000');
   TC.PutFixedStValue('PCV_LIBELLE',St,8,50,tcTrim,FALSE);
   TC.PutFixedStValue('PCV_LIBELLE1',St,58,50,tcTrim,FALSE);
   TC.PutFixedStValue('PCV_LIBELLE2',St,108,50,tcTrim,FALSE);
   TC.PutFixedStValue('PCV_LIBELLE3',St,158,50,tcTrim,FALSE);
   TC.PutValue('PCV_DATEMODIF', Date); //PT12

   okok:=TC.InsertOrUpdateDb(FALSE);
//PT25
   if (FTypeFichier.ItemIndex = 0) then
      begin
      StMessage:='Convention (P) '+Copy(St,2,4)+Verifokok(okok);
      FReport.Items.Add(StMessage);
      Writeln(FRapport,StMessage);
      end
   else
      begin
      StMessage:='Convention (PCV) '+Copy(St,2,4)+Verifokok(okok);
      FReport.Items.Add(StMessage);
      Writeln(FRapport,StMessage);
      end;
//FIN PT25
   FreeAndNil(TC);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/10/2000
Modifié le ... :   /  /
Description .. : Enregistrement minimum conventionnel
Suite ........ : TMC : TOB minimum conventionnel
Suite ........ : St : Chaine de caractères
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregMiniConv(TMC : TOB; St : string);
var
TMinConv: TOB;
begin
TMinConv:=TOB.Create ('MINIMUMCONVENT', TMC, -1);
TMinConv.PutValue('PMI_NATURE','QUA');
TMinConv.PutValue('PMI_CONVENTION','000');
TMinConv.PutValue('PMI_TYPENATURE','VAL');
TMinConv.PutFixedStValue('PMI_CODE', St,5,4,tcTrim,FALSE);
TMinConv.PutFixedStValue('PMI_LIBELLE', St,11,20,tcTrim,FALSE);
TMinConv.PutValue('PMI_DATECREATION', Date);
TMinConv.PutValue('PMI_DATEMODIF', Date);
TMinConv.PutValue('PMI_PREDEFINI', 'DOS');
{PT27
TMinConv.PutValue('PMI_NODOSSIER', V_PGI_Env.NoDossier);
}
TMinConv.PutValue('PMI_NODOSSIER', PgRendNoDossier());
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 11/10/2000
Modifié le ... :   /  /
Description .. : Enregistrement congés payés
Suite ........ : St : Chaine de caractères
Suite ........ : TCP : TOB congés payés
Suite ........ : Etab : Numéro d'établissement
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.EnregCP(St : string; TCP : TOB; Etab : string);
Var
TACQUISN, TPRISN, TACQUISN1, TPRISN1 : TOB ;
BufDateDeb, BufDateFin : String ;
DateBufDeb, DateBufFin : TDateTime;
NbJours : Integer;
begin
NbJours := StrToInt(Copy(St, 74, 13));
if (NbJours <> 0) then
   BEGIN
   TACQUISN := TOB.Create('ABSENCESALARIE', TCP, -1);
   TACQUISN.PutValue('PCN_TYPEMVT','CPA') ;
   TACQUISN.PutFixedStValue('PCN_SALARIE',St,2,6,tcTrim,FALSE);
   TACQUISN.PutValue('PCN_ORDRE',1);
   TACQUISN.PutValue('PCN_DATEMODIF', Date);
   TACQUISN.PutValue('PCN_DATESOLDE',0);
   TACQUISN.PutValue('PCN_ETABLISSEMENT',Etab);

   DateTest := Copy(St, 8, 8);
   if (DateTest <> '        ') then
      BEGIN
      TACQUISN.PutFixedStValue('PCN_DATEDEBUT',St,8,8,tcDate8AMJ,FALSE);
      TACQUISN.PutFixedStValue('PCN_DATEFIN',St,8,8,tcDate8AMJ,FALSE);
      TACQUISN.PutFixedStValue('PCN_DATEVALIDITE',St,8,8,tcDate8AMJ,FALSE);
      END
   else
      BEGIN
      TACQUISN.PutValue('PCN_DATEDEBUT', idate1900);     //PT52
      TACQUISN.PutValue('PCN_DATEFIN', idate1900);       //PT52
      TACQUISN.PutValue('PCN_DATEVALIDITE', idate1900);  //PT52
      END;

   TACQUISN.PutValue('PCN_PERIODECP',0);
   TACQUISN.PutValue('PCN_TYPECONGE','REP');
   TACQUISN.PutValue('PCN_SENSABS','+');
   TACQUISN.PutValue('PCN_LIBELLE','Reprise Acquis');
   TACQUISN.PutValue('PCN_CODETAPE','...'); //PT23
   TACQUISN.PutValue('PCN_TYPEIMPUTE','AC2'); { PT65 }
   TACQUISN.PutFixedStValue('PCN_JOURS',St, 74, 13, tcDouble100, FALSE);
   TACQUISN.PutFixedStValue('PCN_BASE',St, 126, 13, tcDouble100, FALSE);
   TACQUISN.PutFixedStValue('PCN_NBREMOIS',St, 152, 13, tcDouble100, FALSE);
   TACQUISN.PutFixedStValue('PCN_APAYES',St, 100, 13, tcDouble100, FALSE);
{PT17
   ChangeCodeSalarie(TACQUISN,'PCN',FALSE) ;
}
   ChangeCodeSalarie(TACQUISN, 'PCN', TRUE) ;
   END;

NbJours := StrToInt(Copy(St, 100, 13));
if (NbJours <> 0) then
   BEGIN
   TPRISN := TOB.Create('ABSENCESALARIE',TCP, -1);
   TPRISN.PutValue('PCN_TYPEMVT','CPA') ;
   TPRISN.PutFixedStValue('PCN_SALARIE',St,2,6,tcTrim,FALSE);
   TPRISN.PutValue('PCN_ORDRE',2);
   TPRISN.PutValue('PCN_DATEMODIF', Date);
{PT28
   TPRISN.PutValue('PCN_DATEVALIDITE', Date);
   TPRISN.PutValue('PCN_DATEPAIEMENT', Date);
}
   TPRISN.PutValue('PCN_DATESOLDE',0);
   TPRISN.PutValue('PCN_ETABLISSEMENT',Etab);

   DateTest := Copy(St, 8, 8);
   if (DateTest <> '        ') then
      BEGIN
      TPRISN.PutFixedStValue('PCN_DATEDEBUT',St,8,8,tcDate8AMJ,FALSE);
      TPRISN.PutFixedStValue('PCN_DATEFIN',St,8,8,tcDate8AMJ,FALSE);
//PT28
      TPRISN.PutFixedStValue('PCN_DATEVALIDITE', St,8,8,tcDate8AMJ,FALSE);
      TPRISN.PutFixedStValue('PCN_DATEPAIEMENT', St,8,8,tcDate8AMJ,FALSE);
//FIN PT28
      END
   else
      BEGIN
      TPRISN.PutValue('PCN_DATEDEBUT', idate1900);     //PT52
      TPRISN.PutValue('PCN_DATEFIN', idate1900);       //PT52
//PT28
      TPRISN.PutValue('PCN_DATEVALIDITE', idate1900);  //PT52
      TPRISN.PutValue('PCN_DATEPAIEMENT', idate1900);  //PT52
//FIN PT28
      END;

   TPRISN.PutValue('PCN_PERIODECP',0);
   TPRISN.PutValue('PCN_TYPECONGE','CPA');
   TPRISN.PutValue('PCN_SENSABS','-');
   TPRISN.PutValue('PCN_LIBELLE','Reprise Pris');
   TPRISN.PutValue('PCN_CODETAPE','P');
   TPRISN.PutValue('PCN_TYPEIMPUTE','AC2'); { PT65 }
   TPRISN.PutFixedStValue('PCN_JOURS',St, 100, 13, tcDouble100, FALSE);
   TPRISN.PutFixedStValue('PCN_BASE',St, 126, 13, tcDouble100, FALSE);
//   TPRISN.PutFixedStValue('PCN_NBREMOIS',St, 152, 13, tcDouble100, FALSE);
// Suite à Tests CEGID, 21/12/00
   TPRISN.PutValue('PCN_NBREMOIS',0);
{PT17
   ChangeCodeSalarie(TPRISN,'PCN',FALSE) ;
}
   ChangeCodeSalarie(TPRISN, 'PCN', TRUE) ;
   END;

NbJours := StrToInt(Copy(St, 87, 13));
if (NbJours <> 0) then
   BEGIN
   TACQUISN1 := TOB.Create('ABSENCESALARIE', TCP, -1);
   TACQUISN1.PutValue('PCN_TYPEMVT','CPA') ;
   TACQUISN1.PutFixedStValue('PCN_SALARIE',St,2,6,tcTrim,FALSE);
   TACQUISN1.PutValue('PCN_ORDRE',3);
   TACQUISN1.PutValue('PCN_DATEMODIF', Date);
   TACQUISN1.PutValue('PCN_DATESOLDE',0);
   TACQUISN1.PutValue('PCN_ETABLISSEMENT',Etab);

   DateTest := Copy(St, 8, 8);
   if (DateTest <> '        ') then
      BEGIN
      BufDateDeb := Copy(St, 14, 2)+'/'+Copy(St, 12, 2)+'/'+Copy(St, 8, 4);
      DateBufDeb := PlusMois(StrToDate(BufDateDeb),-12);
      TACQUISN1.PutValue('PCN_DATEDEBUT',DateBufDeb);
      TACQUISN1.PutValue('PCN_DATEFIN',DateBufDeb);
      TACQUISN1.PutValue('PCN_DATEVALIDITE',DateBufDeb);
      END
   else
      BEGIN
      TACQUISN1.PutValue('PCN_DATEDEBUT',idate1900);      //PT52
      TACQUISN1.PutValue('PCN_DATEFIN',idate1900);        //PT52
      TACQUISN1.PutValue('PCN_DATEVALIDITE',idate1900);   //PT52
      END;
   TACQUISN1.PutValue('PCN_PERIODECP',1);
   TACQUISN1.PutValue('PCN_TYPECONGE','REP');
   TACQUISN1.PutValue('PCN_SENSABS','+');
   TACQUISN1.PutValue('PCN_LIBELLE','Reprise Acquis');
   TACQUISN1.PutValue('PCN_CODETAPE','...'); //PT6
   TACQUISN1.PutValue('PCN_TYPEIMPUTE','AC2'); { PT65 }
   TACQUISN1.PutFixedStValue('PCN_JOURS',St, 87, 13, tcDouble100, FALSE);
   TACQUISN1.PutFixedStValue('PCN_BASE',St, 139, 13, tcDouble100, FALSE);
   TACQUISN1.PutFixedStValue('PCN_NBREMOIS',St, 165, 13, tcDouble100, FALSE);
   TACQUISN1.PutFixedStValue('PCN_APAYES',St, 113, 13, tcDouble100, FALSE);
{PT17
   ChangeCodeSalarie(TACQUISN1,'PCN',FALSE) ;
}
   ChangeCodeSalarie(TACQUISN1, 'PCN', TRUE) ;
   END;

NbJours := StrToInt(Copy(St, 113, 13));
if (NbJours <> 0) then
   BEGIN
   TPRISN1 := TOB.Create('ABSENCESALARIE',TCP, -1);
   TPRISN1.PutValue('PCN_TYPEMVT','CPA') ;
   TPRISN1.PutFixedStValue('PCN_SALARIE',St,2,6,tcTrim,FALSE);
   TPRISN1.PutValue('PCN_ORDRE',4);
   TPRISN1.PutValue('PCN_DATEMODIF', Date);
   TPRISN1.PutValue('PCN_DATESOLDE',0);
   TPRISN1.PutValue('PCN_ETABLISSEMENT',Etab);
   DateTest := Copy(St, 8, 8);
   if (DateTest <> '        ') then
      BEGIN
      BufDateDeb := Copy(St, 14, 2)+'/'+Copy(St, 12, 2)+'/'+Copy(St, 8, 4);
      DateBufDeb := PlusMois(StrToDate(BufDateDeb),-12);
      TPRISN1.PutValue('PCN_DATEDEBUT',DateBufDeb);
      TPRISN1.PutValue('PCN_DATEFIN',DateBufDeb);
      END
   else
      BEGIN
      TPRISN1.PutValue('PCN_DATEDEBUT',idate1900);     //PT52
      TPRISN1.PutValue('PCN_DATEFIN',idate1900);       //PT52
      END;

   DateTest := Copy(St, 16, 8);
   if (DateTest <> '        ') then
      BEGIN
      BufDateFin := Copy(St, 22, 2)+'/'+Copy(St, 20, 2)+'/'+Copy(St, 16, 4);
      DateBufFin := PlusMois(StrToDate(BufDateFin),-12);
      TPRISN1.PutValue('PCN_DATEVALIDITE',DateBufFin);
      TPRISN1.PutValue('PCN_DATEPAIEMENT',DateBufFin);
      END
   else
      BEGIN
      TPRISN1.PutValue('PCN_DATEVALIDITE',idate1900);    //PT52
      TPRISN1.PutValue('PCN_DATEPAIEMENT',idate1900);    //PT52
      END;

   TPRISN1.PutValue('PCN_PERIODECP',1);
   TPRISN1.PutValue('PCN_TYPECONGE','CPA');
   TPRISN1.PutValue('PCN_SENSABS','-');
   TPRISN1.PutValue('PCN_LIBELLE','Reprise Pris');
   TPRISN1.PutValue('PCN_CODETAPE','P');
   TPRISN1.PutValue('PCN_TYPEIMPUTE','AC2'); { PT65 }
   TPRISN1.PutFixedStValue('PCN_JOURS',St, 113, 13, tcDouble100, FALSE);
   TPRISN1.PutFixedStValue('PCN_BASE',St, 139, 13, tcDouble100, FALSE);
   TPRISN1.PutFixedStValue('PCN_NBREMOIS',St, 165, 13, tcDouble100, FALSE);
{PT17
   ChangeCodeSalarie(TPRISN1,'PCN',FALSE) ;
}
   ChangeCodeSalarie(TPRISN1, 'PCN', TRUE) ;
   END;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/09/2003
Modifié le ... :   /  /
Description .. : Controle congés payés
Suite ........ : St : Chaine de caractères
Suite ........ : Etab : Numéro d'établissement
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpIsis.ControleCP(St, Etab : string) : string;
var
reponse : integer;
TEtabD : TOB;
BufDateFin : string;
DateBufFin, DatePGIFin : TDateTime;
begin
result:='1';
TEtabD := TEtab.FindFirst(['ET_ETABLISSEMENT'], [Etab], False);
if (TEtabD <> nil) then
   begin
   if (TETabD.GetValue('ETB_CONGESPAYES') = '-') then
      begin
      reponse := PGIAskCancel('Vous ne gérez pas les congés payés de#10#13'+
                              'l''établissement '+Etab+'.#10#13'+
                              'Confirmez-vous l''import des congés payés ?',
                              StTitre);       //PT64
      if (reponse=mrNo) then
         begin
         result:='-2';
         exit;
         end
      else
         if (reponse=mrCancel) then
            begin
            result:='-1';
            exit;
            end;
      end;
   DateTest := Copy(St, 16, 8);
   if (DateTest <> '        ') then
      BEGIN
      BufDateFin := Copy(St, 22, 2)+'/'+Copy(St, 20, 2)+'/'+Copy(St, 16, 4);
      DateBufFin := StrToDate(BufDateFin);
      DatePGIFin := TETabD.GetValue('ETB_DATECLOTURECPN');
      if (DateBufFin <> DatePGIFin) then
         begin
         PGIBox('La date de clôture des congés payés ne correspond pas#10#13'+
                'pour l''établissement '+Etab+'.', StTitre);       //PT64
         result:='-1';
         end;
      END
   else
      BEGIN
      reponse:=PGIAskCancel('Dans le fichier d''import, la date de#10#13'+
                            'clôture des congés payés n''est pas#10#13'+
                            'renseignée pour l''établissement '+Etab+'.#10#13'+
                            'Confirmez-vous l''import des congés payés ?',
                            StTitre);       //PT64
      if (reponse=mrNo) then
         begin
         result:='-2';
         exit;
         end
      else
         if (reponse=mrCancel) then
            begin
            result:='-1';
            exit;
            end;
      END;
   end
else
      result:='-1';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Philippe DUMET
Créé le ...... : 11/10/2000
Modifié le ... :   /  /
Description .. : Choix du fichier d'import
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.FTypeFichierClick(Sender: TObject);
begin
  inherited;
case FTypeFichier.ItemIndex of
     0 : BEGIN
         edtFichier.DataType := 'OPENFILE(*PGI*.SAV)';
         edtFichier.Text := '';
         StTitre:= 'Import Fichier CEGID ISIS II';    //PT64
         END;

     1 : BEGIN
         edtFichier.DataType := 'OPENFILE(DADS*.*)';
         edtFichier.Text := '';
         StTitre:= 'Import Fichier TDS';            //PT64
         END;

     2 : BEGIN
         edtFichier.DataType := 'OPENFILE(*.TRA)';
         edtFichier.Text := '';
         StTitre:= 'Import Fichier de mouvement TRA';    //PT64
         END;
//PT59
     3 : BEGIN
         edtFichier.DataType := 'OPENFILE(*.*)';
         edtFichier.Text := '';
         StTitre:= 'Import Fichier DADS-U';     //PT64
         END;
//FIN PT59
     END;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 04/01/2002
Modifié le ... :   /  /
Description .. : Gestion de l'aide
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic (Self);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Chargement des établissements en TOB pour fichier TRA
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.ChargeTOBEtab();
Var
StEtab : string;
QRechEtab : TQuery;
begin
//Chargement de la TOB ETABLISSEMENT
StEtab:='SELECT ET_ETABLISSEMENT, ET_SIRET, ETB_CONGESPAYES,'+
        ' ETB_DATECLOTURECPN, ETB_HORAIREETABL, ETB_CONVENTION'+
        ' FROM ETABLISS'+
        ' LEFT JOIN ETABCOMPL ON'+
        ' ET_ETABLISSEMENT=ETB_ETABLISSEMENT';

QRechEtab:=OpenSql(StEtab,TRUE);
TEtab := TOB.Create('Les etablissements', NIL, -1);
TEtab.LoadDetailDB('ETAB','','',QRechEtab,False);
Ferme(QRechEtab);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Ajout des établissements dans TOB pour fichier TRA
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.RemplitTOBEtab(S : string; var Etab : string);
Var
Ouv, Siret : string;
FinExercice : TDateTime;
NMois : integer;
TEtabD : TOB;
begin
Siret := Copy(S, 7, 14);
FinExercice := Str6ToDate(Copy(S,105,6),90);
NMois := StrToInt(Copy(S,111,2));
Ouv := Copy(S,119,1);
TEtabD := TEtab.FindFirst(['ET_SIRET'], [Siret], TRUE);
if NOT TEtabD.FieldExists ('FINEXERCICE') then
   begin
   TEtabD.AddChampSup('FINEXERCICE', FALSE);
   TEtabD.AddChampSup('NBMOIS', FALSE);
   TEtabD.AddChampSup('OUVERT', FALSE);
   TEtabD.PutValue('FINEXERCICE', FinExercice);
   TEtabD.PutValue('NBMOIS', NMois);
   TEtabD.PutValue('OUVERT', Ouv);
   end;
Etab := TEtabD.GetValue('ET_ETABLISSEMENT');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Recherche de la validité des établissements pour fichier
Suite ........ : TRA
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpIsis.RecupCompl(SSiret : string; var Etab : string) : string;
Var
TEtabD : TOB;
begin
result := '0';
try
   begintrans;
   TEtabD := TEtab.FindFirst(['ET_SIRET'], [SSiret], TRUE);
   Etab := TEtabD.GetValue('ET_ETABLISSEMENT');
   FinExo := TEtabD.GetValue('FINEXERCICE');
   if (FinExo <> IDate1900) then
      begin
      NbMois:=TEtabD.GetValue('NBMOIS');
      Ouvert:=TEtabD.GetValue('OUVERT');
      if (Ouvert = 'O') then
         NbMois:=NbMois-1;
      if (NbMois < 1) then
         NbMois:=1;
      end
   else
      begin
{PT57
      FinExo:=null ;
}
      FinExo:=IDate1900 ;
      NbMois:=1 ;
      Ouvert:=TEtabD.GetValue('OUVERT');
      end;
      CommitTrans;
Except
   Rollback;
   result := '-2';
END;
end;

//PT67-1
procedure TFImpIsis.bAnnulerClick(Sender: TObject);
begin
inherited;
if (P.ActivePage<>PReport) then
   FreeAndNil (TEtab);
frmLitSav.Close;
FMEnuG.ChoixModule; // Pour afficher la liste des modules
end;

procedure TFImpIsis.bFinClick(Sender: TObject);
begin
inherited;
//frmLitSav.Close;
FMEnuG.ChoixModule; // Pour afficher la liste des modules
end;
//FIN PT67-1

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /
Description .. : Initialisation de la table de correspondance
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.InitCodeSal;
Var
Q : TQuery ;
MaxCodeSal : Integer ;
begin
ListeCodeSal:=TStringList.Create ;
ListeCodeSal.Sorted:=TRUE ;
MaxCodeSal:=1 ;
if (FCodeSal.ItemIndex=0) then
   BEGIN
   Q:=OpenSQL('SELECT MAX(PSA_SALARIE) FROM SALARIES',TRUE) ;
   if Not Q.EOF then
      try
      MaxCodeSal:=ValeurI(Q.Fields[0].AsString)+1 ;
      except
            on E: EConvertError do
               MaxCodeSal:= 1;
      end;
   Ferme(Q) ;
   END ;
ListeCodeSal.AddObject('_______',tobject(MaxCodeSal)) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /
Description .. : Libération de la table de correspondance
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
procedure TFImpIsis.FinCodeSal;
begin
ListeCodeSal.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/08/2007
Modifié le ... :   /  /
Description .. : Vérification de la présence d'un salarié pour permettre la
Suite ........ : création
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpIsis.SalarieAbsent (St : string) : boolean;
Var
StSal : string;
begin
StSal:= 'SELECT PSA_SALARIE FROM SALARIES WHERE PSA_SALARIE="'+St+'"';
Result := ExisteSQL (StSal);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 23/08/2007
Modifié le ... :   /  /
Description .. : Gestion de la liste des salariés
Suite ........ : PeutCréer : Si = True, création du salarié possible dans la
Suite ........ : table de correspondance
Mots clefs ... : PAIE;PGIMPORT
*****************************************************************}
function TFImpIsis.SalarieDansListe (St : string; PeutCreer : boolean;
                                     var Code : Integer) : boolean;
Var
i : Integer ;
begin
result := TRUE;
i:=ListeCodeSal.IndexOf(St) ;
if i<0 then
   BEGIN
   if Not PeutCreer then
      BEGIN
      result:=FALSE ;
      Exit ;
      END ;
   if (Index = 0) then
      Code:=Longint(ListeCodeSal.Objects[0]) + Num
   else
      Code:=Longint(ListeCodeSal.Objects[0]);
   ListeCodeSal.AddObject(St,tobject(Code)) ;
   ListeCodeSal.Objects[0]:=tobject(Code+1) ;
   END
else
   BEGIN
   Code:=longint(ListeCodeSal.Objects[i]) ;
   END ;
end;

//debut PT72
function TFImpIsis.VerifType(Deftype : string):boolean;
begin
 if not (isnumeric(Deftype)) then
 begin
    StMessage2:= 'Import impossible pour un code salarié en alphanumérique';
    result := False;
 end
 else
 result := True;
end;
// fin PT72
end.

