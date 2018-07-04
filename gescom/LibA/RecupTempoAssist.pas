
unit RecupTempoAssist;



// programme de récupération d'un dossier SCOT ou TEMPO en PGI.
// rechercher "MCD" pour tous les cas non encore traité ou en suspend.

                                     
// penser de forcer en majuscule le code assistant récupérer dans les nouvelles
// focntion : ligne affaire, facture, facture dans les temps ...

{ mcd 16/11/00 :
N'accepte plus apporteur dans les choix stat client ,car l'apporteur est un client
Met le champ T_TIERS dans les zones GP_TIERSFACTURE  de piece (attention si reprise faite avant 1.02)
Alimente tierspayeur et facture même dans les pieces CC et DE
Affiche sur libelle stat si lien assistant
}
{mcd 28/11/00
- Pour création facture dans les temps ou stat fatcure, alimente bine les partie de
l'affaire.
- Idem pour table LIGNE
- Pour ressource, reprend emploi dans table libre 3 et dans texte libre 1  +
si 1 seul stat 2 reprise, mise en libellé libre dans stat 1
- pour code affaire des fichiers annexes: si code non trouvé dans table affaire,
on ne traite pas l'enrgt. Il est écrit dans le fichier .log
- sur facture FRE, mise zone vivante à -
+ le 19/12/00 alimente table ANNUAIRE pour GI
}

{ mcd 29/12/00
- renseigne aff_datedebgener si date 1ere facturation renseignée dans fichier
- n'alimente plus ACT_DESCRIPTIF
- OK pour alimenter tablette AFCOMPTAAFFAIRE  si recup zone comptable
- si annule les fichiers, recrée article par défaut du profil fatcure
}
                              

{
mcd 08/01/01
-affaire : si mode facturation ACTivité, on force le mode de reprise activité à tout
- contact tiers: ecriture du code auxi sur la bonne longueur     + OK avec nouveaux champs ! type et nature
- adresse tiers: mise en clé T au lieu de TIE !!!
}

{mcd 22/03/01
tout revu pour bien libérer la mémoire (^lus modif factcomm
}
{ mcd 26/03/01
- Affiche dans .log le nbre enrgt traité plus le temps.
- Ok pour fournisseurs en GA
- Ok pour commentaire ligne
- Ok pour mission et mode reprise activité par defaut + Ok dnas factaff
}
{mcd 16/05/01
- alimente afa_tyepeche
- prise en compte de l'option checkreprisesup pour ne pas reprendre fiche supprimée
+ message si Clt/Art/res inexistant (sauf pour table ressourcepr, car restaurer avant la reprise des fiches de bases !!!)
 + n'efface pas toutes les missions si option effacement non coché.
- si affaire terminée, met CLO dans etataffaire
- Récupération valorisation en PV
- Récupère les stat clit,aff,pres des stats glob et det dans piece
- Si efface temps, fait jusqu'à une date
- Récupère zone clt filiale si existe
- Pour les stats affaire et ressource, propose une liste de choix + traite les 4 stat possibles
}

{mcd 08/11/01  ajout boni/mali  }

{mcd 27/11/01 peut traiter n° compte sur 12c pour SST (prend les c qui suive, sur 2 maxi) }
{mcd 31/03/03 pour nouvelle clé activité}
{mcd 23/10/03
NB laisse addlessupligne (au lieu de NewTobLigne de FactTOB, car trop dangereux de tou changer maintenant
a voir plus tard.. mais on n'utilise pas les fct de claul piece std?? on écrit directement la TOB .. à priori OK
+ modif pour metre date du jour, si date deb facture et deb facturation à blanc pour tempo
}

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, HSysMenu, hmsgbox, HTB97, StdCtrls, Hctrls, ComCtrls, ExtCtrls,HPanel,
  Mask, Buttons,Hent1,UIutil,HStatus,RecupTempoParamEntite,RecupTempoParamFjur,RecupTempoParam, RecupTempoParamCli,HeureUtil,
{$IFDEF EAGLCLIENT}

{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
   Utilressource,
   ActiviteUtil,CheckLst,UTob,UTom,ParamSoc,EntGC,UAFO_Ressource, Ent1,Dicobtp, AffaireUtil,FactUtil,
   FactComm, UtilPgi,TiersUtil,UTOMArticle, UtilArticle,
   FactTob, FactPiece, FactTiers // PL le 14/08/03 : suite modifs JLD facturation
{$IFDEF DP}
   ,Annoutils, DpJurOutils
{$ENDIF}
	 ,uEntCommun,UtilTOBPiece
   ;
type
  TFAsRecupTempo = class(TFAssist)
    TabParam: TTabSheet;
    HLabel3: THLabel;
    HLabel8: THLabel;
    LibNumDossier: THLabel;
    Repfich: THCritMaskEdit;
    NumDossier: TEdit;                                     
    Label1: TLabel;
    TabFicBase: TTabSheet;
    StaticText1: TStaticText;
    HLabelBase: THLabel;
    TabTemps: TTabSheet;
    TabFacture: TTabSheet;
    TabPlanning: TTabSheet;
    HLabelTemps: THLabel;
    HLabelFact: THLabel;
    HLabelPlan: THLabel;
    NomFicBase: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    NomFicTemps: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    DatePlanning: TEdit;
    TabAchat: TTabSheet;
    Label9: TLabel;
    NomFicFact: TLabel;
    Label4: TLabel;
    NomFicPlan: TLabel;
    HLabelAchat: THLabel;
    Label10: TLabel;
    NomFicAchat: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    FicTempsLabelDate: TLabel;
    FicTempsDate: TEdit;
    StaticText3: TStaticText;
    CheckListFicAchat: TCheckListBox;
    CheckListFicBase: TCheckListBox;
    TabReprise: TTabSheet;
    LabelRepise1: TLabel;
    HLabelReprise: THLabel;
    LabelReprise2: TLabel;
    CheckListFicFacture: TCheckListBox;
    StaticText2: TStaticText;
    CheckListFicTemps: TCheckListBox;
    StaticText4: TStaticText;
    LabelFicTempsValor: TLabel;
    FicTempsValor: TComboBox;
    StaticText5: TStaticText;
    CheckListFicPlan: TCheckListBox;
    GroupBoxTrace: TGroupBox;
    Labelconseil1: TLabel;
    LabelConseil2: TLabel;
    MemoTrace: TMemo;
    LabelConseil3: TLabel;
    BStop: TBitBtn;
    CheckSupBase: TCheckBox;
    CheckSupTemps: TCheckBox;
    ChecksupPlan: TCheckBox;
    LabelConseil4: TLabel;
    LabelConseil5: TLabel;
    CheckRepriseSup: TCheckBox;
    BParamBase: TBitBtn;
    BParamClient: TBitBtn;
    LabelmonnaieTempo: TLabel;
    LabelMonTempo: TLabel;
    LabelMonnaiePGI: TLabel;
    LabelMonPGI: TLabel;
    Shape1: TShape;
    Shape3: TShape;
    Shape4: TShape;
    Label7: TLabel;
    Shape2: TShape;
    Shape5: TShape;
    Shape6: TShape;
    Labelconseilmonnaie: TLabel;
    CheckSupCLi: TCheckBox;
    Exercice: TEdit;
    TExercice: TLabel;
    ExerciceDec: TEdit;
    BParamFJur: TBitBtn;
    BEntite: TBitBtn;
    GAMME2: TCheckBox;
    PRODUC: TCheckBox;
    CheckSupAchat: TCheckBox;
    FicTempsDateEff: TEdit;
    FicFacDateEff: TEdit;
    TFicFacDateEff: TLabel;

    procedure bSuivantClick(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BParamBaseClick(Sender: TObject);
    procedure BParamClientClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure BParamFJurClick(Sender: TObject);
    procedure BEntiteClick(Sender: TObject);
    procedure FicTempsDateEffExit(Sender: TObject);
    procedure FicFacDateEffExit(Sender: TObject);
    procedure FicTempsDateExit(Sender: TObject);
    procedure DatePlanningExit(Sender: TObject);
    procedure TabPlanningExit(Sender: TObject);
    procedure TabTempsExit(Sender: TObject);
    procedure PChange(Sender: TObject); override ;
   private                   
    { Déclarations privées }
    ListeAffaire: TstringList; //mcd 31/03/03
    LigneMemo : Integer;
    T_PIECESAUV : TOB;  // stock valeur par défaut de la ligne
    TobArt : Tob;   // pour mise à jour zone qualif dans ligne aff
    FParam: TFRecupTempoParam;
    FParamCli : TFParamCli;
    FParamFJur : TFParamFjur;
    FParamEntite : TFParamEntite;
    FichierLog : textfile;
    FamAffaire : Integer;  // 1= Stat1  2= Stat2  3= Stat3  4= Fam.comptable
    nbtva : Integer;        // nbre tva gérée
    TotTaxe,  TotTTC, totqte : double;    // pour cumul tva et ttc de la piece/affaire
    nbjour : Integer;       // permet de renuméroter les jours fériés
    debutpre, debutfour, debutfrais: string  ;
    Nbtot,Nbint, Nbmax,NbComm : Integer;   // pour comptage nbre de lignes
    gt: array[1..27]  of Integer;  // 0 ou 1 si les 26 GT  sont traité
                          // en table libre : utilisée pour générer les libellés tables
    NbPage, NoPage: Integer ;   // compte le nbre de pages réelle de la reprise
    NbMode:Integer;     // stock le n° du mode de règlement
        // liste des options de SCOT/TEMPO prise en compte dans la reprise
    Stat1Cli : Integer;   // De 1 à 4 = Stat, 5 =  Apporteur Aff n'est plus géré au 16/11/00  6= Secteur activité
    Stat2Cli : Integer;
    Stat3Cli : Integer;
    Stat4Cli : Integer;
    Stat1Aff : Integer;   // 1 à 4 = table libre, 5 = etablissement, 6 = Départemetn, 7 = Type reconduction
    Stat2Aff : Integer;
    Stat3Aff : Integer;
    Stat4Aff : Integer;
    Stat1Res : Integer;   //1 à 4 = table libre, 5 = Etablissement, 6 = Département, 7 = Type salarié
    Stat2Res : Integer;
    Stat3Res : Integer;
    Stat4Res : Integer;
    noaff : Integer;    // pour TEMPO stock le dernier n° affaire traité
    nodev : Integer;    // pour TEMPO stock le dernier n° affaire traité pour les devis
    EuroDefaut: string; //stock t_eurodefaut du tiers
    FamPrime : String;   // TEMPO, permet de connaitre le code famille associé à la prime
    S7 :String;     // permet de savoir si on est en lien S7 ou pas
    Stat :String;     // permet de savoir mode stat gérée
    DebEnc : string;    // permet de savoir si facture débit ou encaissement
    Exer : String;  //permet de savoir pour SCOT si lien exercice ou pas
    TauxTva : Array[1..20] of double;   //stock les taux de tva
    CodeTva : Array[1..20] of string;  //stock les codes tva
    Multi: string ;       // permet de savoir si on est en multi ou pas
    NoEntite:integer;       // n° stat géré en entité
    LiaisonIsis : Boolean;        // O/N si liaison ISIS dans SCOT/TEMPO
    LiaisonCompta: Boolean;        // O/N si liaison avec SISCO dans SCOT/TEMPO
    LngCpte:integer;               // longueur des comptes pour SST peut être > à 10
    FactureGlobale : Boolean;   // O/N si facture globale en paramètre
    FraisCompta : Boolean;      //O/N si  lien frais compta.
    HeurePlanning : String;
    Debexer, FinExer : string ; // Date début et fin exercice SCOT
    CleDocLigaf : R_CLEDOC; // stock n° cde en cours pour reprise ligne
    NoOrdre : Integer;      // pour reprise ligne
    Precedent : string;     // stock N° affaire en cours
    Valpr,Valpv : string;   // stock mode valorisation PV/PR
    lng_gene : integer;     // longueur du code client à récupérer
    Imptra, Implcr: string; // Impression oui non traite/lcr sur papier
    MttMin : Integer;   // mtt minimum pour traite/LCR
    MRRempl : String;   // mde règlement remplacement si mini
        // fin option reprise de SCOT/TEMPO

    procedure InitArticle(TypeArt:string);
    Procedure RecupCreerPiece ( Var CleDoc : R_CleDoc ; CodeTiers,CodeAffaire,part1,part2,part3,Entite,etbs,nofac : String;VAr TobPiece:TOB;Var NoOrdre:integer)  ;
    Procedure RecupFrais( CleDoc : R_CleDoc ; StTot : String;Monpres:boolean)  ;
    Procedure CreerTablette(TypeTable,TypeCode,code,Libelle:string;LngCode:integer);
    procedure RecupChampLibre ();
    procedure MajPieceAff ( );
    procedure ImportTob ( TobImport : TOB; BUpdate:boolean );
    procedure RecupAnnuaire (StTot, champ:string);
    procedure RecupAffaire (StTot:string);
    procedure RecupLigneAffaire (StTot:string);
    procedure RecupPieceAffaire (StTot:string;CleDocAffaire : R_CLEDOC; MonPres:Boolean);
    procedure RecupAffAdresse (StTot,aff:string);
    procedure RecupTiersAdresse (StTot,clt:string);
    procedure RecupTiersContact (StTot:string; champ:string;lng_auxi : integer);
    procedure RecupTiersContactFou (StTot:string; champ:string;lng_auxi : integer);
    procedure RecupRib (StTot, clt:string; condi:Boolean);
    procedure RecupRibFou (StTot, clt:string; condi:Boolean);
    procedure RecupFactAff (StTot, Aff, datdeb, datfin,GenerAuto:string; Var Periodicite:string; Var Interval : Integer; Monpres : boolean);
    procedure RecupAffTiers (stTot,TypeTiers:string;Rang:Integer; Aff:string);
    procedure RecupArticle (stTot:string);
    procedure RecupPort (stTot:string);
    procedure RecupModePaie (stTot:string);
    procedure RecupHisto (stTot:string);
    procedure RecupRessource (enreg:string);
    procedure RecupTablette (enreg:string);
    procedure RecupJourFerie (StTot:string);
    procedure RecupFonction (StTot:string);
    procedure RecupValor (StTot:string);
    procedure RecupNotes (StTot:string; Nbligne:integer);
    procedure RecupTemps (StTot:string);
    procedure RecupStatGlob (StTot:string);
    procedure RecupStatDet (StTot:string);
    procedure RecupTempsFact (StTot:string);
    procedure RecupFourAttente (StTot:string);
    procedure RecupPlanning (StTot:string);
    procedure Recupclient (StTot:string);
    procedure RecupFour (StTot:string);
    procedure RecupArtFour (StTot:string);
    procedure Recupclientcompl (StTot,auxi,tiers:string);
    procedure RecupclientcomplFou (StTot,auxi,tiers:string);
    procedure AffectStat;
    Procedure ChangeClientFacture ;
    function ConvertMtt (Mtt,ecart:string ):double;
    Function EntiteVersEtbs(Entite:string):string;
    //function ConvertMttbis (Mtt,ecart:string ):double;
    function Tablelibre (StTot:string; ii:integer ):string ;
    function RecupModeRegl(FinMOis, ModePaie,Cheque:string; Duree, Jour, Nbech, Separe, MontMin:integer ):string ;
    function RecupModeReglFou(FinMOis, ModePaie,Cheque:string; Duree, Jour:integer ):string ;
    function AffectTiers ( Clt:string):string ;
    function AffectAff (Aff,  ZExer, Clt, typ,avenant:string;Var Part0, Part1, Part2, Part3 : STring; Creation : Boolean;Var MoisClot :Integer ):string ;
    function TestExistance(Zchamp,Art,TypArt,Ress,Aff,  ZExer, Clt, typ,avenant:string;Var Part0, Part1, Part2, Part3 ,affaire: STring; Var MoisClot :Integer ):boolean ;
    function RecupTypeTable ( ii:integer ):string ;
    function LectureParamBase :Boolean;

  public
    { Déclarations publiques }
    ModuleBase :boolean;
    ModuleTemps :boolean;
    ModuleFacture : boolean;
    ModulePlanning : boolean;
    ModuleAchat : boolean;
    StopRecup :  Boolean;
    function  PreviousPage : TTabSheet ; override ;
    function  NextPage : TTabSheet ; override ;
    end;

// Procédure appelée depuis Menudisp
Procedure AssistRecupTempo ;

implementation
{$R *.DFM}
Procedure AssistRecupTempo ;
Var
    FAsRecupTempo  : TFAsRecupTempo ;
begin

if ToutSeulAff  then  exit; // quelqu'un d'autre travaille sur la base
FAsRecupTempo:=TFAsRecupTempo.Create(Application) ;
FAsRecupTempo.FParam := TFRecupTempoParam.Create(Application);
FAsRecupTempo.FParamCli := TFParamCli.Create(Application);
FAsRecupTempo.FParamFjur := TFParamFjur.Create(Application);
FAsRecupTempo.FParamEntite := TFParamEntite.Create(Application);
FAsRecupTempo.FParamFjur.TobFjur := TOB.Create('Liste des Fjur', Nil,-1);
FAsRecupTempo.FParamEntite.TobEntite := TOB.Create('Liste des Entités', Nil,-1);

FAsRecupTempo.ShowModal ;
// initialisation des modules
FAsRecupTempo.ModuleBase := False;
FAsRecupTempo.ModuleTemps := False;
FAsRecupTempo.ModuleFacture := False;
FAsRecupTempo.ModulePlanning := False;
    // MCD module achat traité en écran, mais  traité uniquement pour les fournisseurs
FAsRecupTempo.ModuleAchat := False;
FAsRecupTempo.StopRecup := False;

FAsRecupTempo.FParam.Free;
FAsRecupTempo.FParamCli.Free;
FAsRecupTempo.FParamFJur.TobfJur.Free;
FAsRecupTempo.FParamFJur.Free;
FAsRecupTempo.FParamEntite.TobEntite.Free;
FAsRecupTempo.FParamEntite.Free;
FAsRecupTempo.Free;
Bloqueur ('AffToutSeul',False);
end;

function  TFAsRecupTempo.PreviousPage : TTabSheet ;
var tt : TTabsheet ;
    boucle : Integer;
begin
tt := inherited PreviousPage;
boucle := -1;
if (tt <> nil) Then
   Begin
   if ((tt.name = 'TabAchat') And (not ModuleAchat)) then Begin  tt:=P.Pages[P.ActivePage.PageIndex-1+boucle]; dec(boucle); End;
   if ((tt.name = 'TabPlanning') And (not ModulePlanning)) then Begin  tt:=P.Pages[P.ActivePage.PageIndex-1+boucle]; dec(boucle);End;
   if ((tt.name = 'TabFacture') And (not ModuleFacture)) then Begin  tt:=P.Pages[P.ActivePage.PageIndex-1+boucle]; dec(boucle);    End;
   if ((tt.name = 'TabTemps') And (not ModuleTemps)) then Begin  tt:=P.Pages[P.ActivePage.PageIndex-1+boucle]; End;
   End;
result :=tt;
end;

function  TFAsRecupTempo.NextPage : TTabSheet ;
var tt : TTabsheet ;
    boucle : Integer;
begin
tt := inherited NextPage;
if (tt <> nil) Then
   Begin
   boucle := 1;
   if ((tt.name = 'TabTemps') And (not ModuleTemps)) then Begin  tt:=P.Pages[P.ActivePage.PageIndex+1+boucle]; inc(boucle); End;
   if ((tt.name = 'TabFacture') And (not ModuleFacture)) then Begin  tt:=P.Pages[P.ActivePage.PageIndex+1+boucle]; inc(boucle); End;
   if ((tt.name = 'TabPlanning') And (not ModulePlanning)) then Begin  tt:=P.Pages[P.ActivePage.PageIndex+1+boucle]; inc(boucle);   End;
   if ((tt.name = 'TabAchat') And (not ModuleAchat)) then Begin  tt:=P.Pages[P.ActivePage.PageIndex+1+boucle];  End;
   End;
result :=tt;
end;


procedure TFAsRecupTempo.bSuivantClick(Sender: TObject);
var st: string;
    ii,err: integer;

begin
Inc (Nopage);
if (P.ActivePage.PageIndex = 0) then
   Begin
        // traitement de la touche suivant sur le 1er écran
        // prépare les seuls onglet à traiter
   NbPage:=3;
   If ctxScot in V_PGI.PGIContexte Then st := RepFich.text+'\scb'+NumDossier.text+'.sav'
                               Else st := RepFich.text+'\sgb'+NumDossier.text+'.sav';


   if  Not FileExists(st) then
       Begin
       PGIBoxAf('Fichier des données de base inexistant',TitreHalley);bSuivant.Enabled := False;
       TWinControl(Numdossier).SetFocus;
       exit;
        End Else
       Begin
       ModuleBase:=True; NomFicBase.Caption:=st;
       if Not (LectureParamBase) then
            Begin
            PGIBoxAf('Fichier généré avec une version de Tempo < V6.0 ou Dossier en Francs',TitreHalley);bSuivant.Enabled := False;
            TWinControl(Numdossier).SetFocus;
            exit;
            End;
       End;

     // test des modules présents
   If ctxScot in V_PGI.PGIContexte Then   st := RepFich.text+'\sct'+NumDossier.text+'.sav'
                               Else   st := RepFich.text+'\sgt'+NumDossier.text+'.sav';
   if  FileExists(st) then Begin ModuleTemps := True; Inc(NbPage); NomFicTemps.Caption:= st; End;
   If ctxScot in V_PGI.PGIContexte Then   st := RepFich.text+'\scg'+NumDossier.text+'.sav'
                               Else   st := RepFich.text+'\sgg'+NumDossier.text+'.sav';
   if  FileExists(st) then Begin ModuleTemps := True; Inc(NbPage);NomFicTemps.Caption:= st; End;
   If ctxScot in V_PGI.PGIContexte Then   st := RepFich.text+'\scf'+NumDossier.text+'.sav'
                               Else   st := RepFich.text+'\sgf'+NumDossier.text+'.sav';
   if  FileExists(st) then Begin ModuleFacture := True; Inc(NbPage); NomFicFact.Caption:= st; End;
   If ctxScot in V_PGI.PGIContexte Then   st := RepFich.text+'\scp'+NumDossier.text+'.sav'
                               Else   st := RepFich.text+'\sgp'+NumDossier.text+'.sav';
   if  FileExists(st) then Begin ModulePlanning := True;Inc(NbPage); NomFicPlan.Caption:= st; End;
   If ctxScot in V_PGI.PGIContexte Then   st := RepFich.text+'\sca'+NumDossier.text+'.sav'
                               Else   st := RepFich.text+'\sga'+NumDossier.text+'.sav';
   if  FileExists(st) then Begin ModuleAchat := True; Inc(NbPage);NomFicAchat.Caption:= st; End;
   //MCD PAS DE RECUP PLANNING pour l'instant. A retester quand
  // le module planning sera avancé ==> forcer à False temporairement
   if (ModulePlanning = TRUE) then
    begin
    ModulePlanning:=False;
    Dec (NbPage);
    End;
        // Si tempo, pas de saisie exercice par defaut
   If not(ctxScot in V_PGI.PGIContexte) then
      begin
      TEXERCICE.visible :=FALSE;
      EXERCICE.visible :=FALSE;
      EXERCICEDEC.visible :=FALSE;
      end
  else
      begin
        //sytéamtique poru SCOT car utiliser si gestion exercice et
        // exercice à blanc ou à *****
      TEXERCICE.visible :=TRUE;
      EXERCICE.visible :=TRUE;
      EXERCICEDEC.visible :=TRUE;
      end;
   if (ModuleFacture =true) and  (ModuleTemps =true) then CheckListFicFacture.Checked[1] := False;
   end;
if (P.ActivePage.PageIndex = 1) and (ctxScot in V_PGI.PGIContexte) then
   Begin
   err:=0;
   if (COpy(Exercice.text,3,1) <>'/') then err:=1;
   if (COpy(Exercice.text,1,2)< '00') or (COpy(Exercice.text,1,2)>'99')then err:=1;
   if (COpy(Exercice.text,4,2)< '00') or (COpy(Exercice.text,4,2)>'99')then err:=1;
   if (err = 1) then
        begin
        PGIInfo(TraduitGA('Le code exercice est incorrect'),TitreHalley);
        TWinControl(Exercice).SetFocus;
        exit;
        end;
   err:=0;
   if (COpy(ExerciceDec.text,3,1) <>'/') then err:=1;
   if (COpy(ExerciceDec.text,1,2)< '00') or (COpy(ExerciceDec.text,1,2)>'99')then err:=1;
   if (COpy(ExerciceDec.text,4,2)< '00') or (COpy(ExerciceDec.text,4,2)>'99')then err:=1;
   if (err = 1) then
        begin
        PGIInfo(TraduitGA('Le code exercice est incorrect'),TitreHalley);
        TWinControl(ExerciceDec).SetFocus;
        exit;
        end;
   END;
   // si module non géré, il faut remettre à faux les champs
if (ModuleTemps = False) then begin
    For ii := 0 to  (CheckListFicTemps.items.Count - 1)  do begin CheckListFicTemps.Checked[ii] := False;     end;
    end;
if (ModulePlanning = False) then begin
    For ii := 0 to  (CheckListFicPlan.items.Count - 1)  do begin CheckListFicPlan.Checked[ii] := False;     end;
    end;
if (ModuleFacture = False) then begin
    For ii := 0 to  (CheckListFicFacture.items.Count - 1)  do begin CheckListFicFacture.Checked[ii] := False;     end;
    end;
    // si facture et temps, on ne peut récupérer que facture dans les temps ou stat.
if (CheckListFicFacture.Checked[1] =true) and  (CheckListFicTemps.Checked[4] =true) then begin
  PGIInfo(TraduitGA('Vous ne pouvez pas reprendre le facturé dans les temps et les statistiques. Suppression Statistiques'),TitreHalley);
  CheckListFicFacture.Checked[1] := False;
  end;
if (ModuleAchat = False) then begin
    For ii := 0 to  (CheckListFicAchat.items.Count - 1)  do begin CheckListFicAchat.Checked[ii] := False;     end;
    end;
    // on ne fait la récupération que des fournisseurs
    // a changer si recup facture,relaiquat ...


if (ModuleBase = False) then begin
    For ii := 0 to  (CheckListFicBase.items.Count - 1)  do begin CheckListFicBase.Checked[ii] := False;     end;
    end;

inherited;
if (P.ActivePage.PageIndex = P.PageCount - 1)
    then bFin.enabled := True ;
End;

procedure TFAsRecupTempo.bFinClick(Sender: TObject);
var Fichier : textfile;
    NomFicLog,
    st,stTot,clenote  : string;
    NbtotRes,NbtotClt,Nbtotpro,nbmax1,   nblig,trace, NbligneNote, ii, Nbmis : integer;
    TraiteNote : boolean;
    QQ : Tquery;
begin
  inherited;
EnableControls(Self, False);
NbLigneNote:=1;
nbmax1 := 100; LigneMemo := 0; trace := 0;
NbtotRes :=0; NBint :=0; NbtotClt :=0;NbtotPro :=0;
NbComm:=0;
ListeAffaire := TStringList.create;  // mcd 32/03/03
T_PIECESAUV := TOB.CREATE ('PIECE', NIL, -1);
NbMis:=0;
Nbmax:=100;   // a changer en fct affichage message voulu
TraiteNote := False;
AffectStat; // fonction d'affectation des différentes zones paramétrables
If ctxScot in V_PGI.PGIContexte then NomFicLog := RepFich.text + '\RecupScot.Log'
   else NomFicLog := RepFich.text + '\RecupTempo.Log';
AssignFile(FichierLog, NomFicLog);
InitParpieceAffaire;
   // on cree ou réouvre le fichier de log
if (FileExists(NomFicLog) = TRUE) then   Append (FichierLog)
  else  Rewrite (FichierLog);
      // on regarde si les différents modules sont à traiter ou non,
      // en fonction des type de récup fichier coché.
ModuleBase := False;
ModuleTemps := False;
ModuleFacture := False;
ModulePlanning := False;
ModuleAchat := False;

For ii := 0 to  (CheckListFicBase.items.Count - 1)  do begin
  if (CheckListFicBase.Checked[ii] = TRUE) then ModuleBase :=TRUE;
  end;
For ii := 0 to  (CheckListFicTemps.items.Count - 1)  do begin
  if (CheckListFicTemps.Checked[ii] = TRUE) then ModuleTemps :=TRUE;
  end;
For ii := 0 to  (CheckListFicFacture.items.Count - 1)  do begin
  if (CheckListFicFacture.Checked[ii] = TRUE) then ModuleFacture :=TRUE;
  end;
For ii := 0 to  (CheckListFicPlan.items.Count - 1)  do begin
  if (CheckListFicPlan.Checked[ii] = TRUE) then ModulePlanning :=TRUE;
  end;
For ii := 0 to  (CheckListFicAchat.items.Count - 1)  do begin
  if (CheckListFicAchat.Checked[ii] = TRUE) then ModuleAchat :=TRUE;
  end;
writeln (FichierLog, TraduitGA('****** Début de reprise des données Tempo *****  ')+FormatDateTime('dd/mm/yyyy ttttt',CurrentDate));
InitMove (nbmax1, '');
If ((ModuleBase) and (Not Stoprecup)) Then
   Begin
        // traitement du fichier de base
    MemoTrace.Lines[LigneMemo] := 'Ouverture du fichier de base'; Inc(LigneMemo);
    AssignFile(Fichier, NomFicBase.Caption);
    Reset (Fichier); //ouvre le fichier
    writeln (FichierLog, TraduitGA('****** Début de reprise des Fichiers de Base *****'));
    RecupChampLibre; //remplit titre des tables libres
    If (CheckSupBase.Checked = True)  then
        Begin
        MemoTrace.Lines[LigneMemo] := 'Suppression des données existantes'; Inc(LigneMemo);
        if (CheckListFicBase.Checked[0] = TRUE) then
            begin
            ExecuteSQL ('delete from ARTICLE where GA_TYPEARTICLE ="PRE"');
            writeln (FichierLog, TraduitGA('Suppression des prestations'));
            InitArticle('PRE');
            end;
        if (CheckListFicBase.Checked[1] = TRUE) then
            begin
            ExecuteSQL ('delete from ARTICLE where GA_TYPEARTICLE ="FRA"');
            ExecuteSQL ('delete from ARTICLE where GA_TYPEARTICLE ="PRI"');
            writeln (FichierLog, TraduitGA('Suppression des frais'));
            InitArticle('FRA');
            end;
        if (CheckListFicBase.Checked[2] = TRUE) then
            begin
            ExecuteSQL ('delete from ARTICLE where GA_TYPEARTICLE ="MAR"');
           writeln (FichierLog, TraduitGA('Suppression des fournitures'));
            InitArticle('FOU');
            end;
      if (CheckListFicBase.Checked[3] = TRUE) or (CheckListFicBase.Checked[6] = TRUE)then
            begin
            ExecuteSQL ('delete from AFFAIRE');
            writeln (FichierLog, TraduitGA('Suppression des affaires'));
              // on détruit aussi le fichier afftiers, factaff, affadresse pour AFF
            ExecuteSQL ('delete from AFFTIERS');
            ExecuteSQL ('delete from FACTAFF');
            ExecuteSQL ('delete from ADRESSES where ADR_TYPEADRESSE = "INT"');
            ExecuteSQL ('delete from ADRESSES where ADR_TYPEADRESSE = "AFA"');
           if CheckListFicBase.Checked[3] = TRUE then
              begin
              ExecuteSQL ('delete from PIEDPORT where GPT_NATUREPIECEG = "' + VH_GC.AFNatAffaire+'"');
              ExecuteSQL ('delete from LIGNE where GL_NATUREPIECEG = "' + VH_GC.AFNatAffaire+'"');
              ExecuteSQL ('delete from PIECE where GP_NATUREPIECEG = "'+ VH_GC.AFNatAffaire+'"');
              end;
            if CheckListFicBase.Checked[6] = TRUE then
              begin
              ExecuteSQL ('delete from PIEDPORT where GPT_NATUREPIECEG = "' + VH_GC.AFNatProposition+ '"');
              ExecuteSQL ('delete from LIGNE where GL_NATUREPIECEG = "' + VH_GC.AFNatProposition+ '"');
              ExecuteSQL ('delete from PIECE where GP_NATUREPIECEG = "'+ VH_GC.AFNatProposition+ '"');
              end;
                  // il faut ensuite remettre le compteur à 0 pour SCOT, à la valeur Lue pour Tempo
            If ctxScot in V_PGI.PGIContexte then begin Noaff :=0; Nodev :=0; end;
            if ((VH_GC.CleAffaire.Co1Type = 'CPT')and(VH_GC.CleAffaire.NbPartie >=1)) then
              begin
              SetParamsoc('SO_AFFCO1VALEUR',noaff );
              VH_GC.CleAffaire.Co1valeur := IntToStr(noaff);
              SetParamsoc('SO_AFFCO1VALEURPRO',nodev );
              //A remettre qd suppr GEtParamSoc VH_GC.CleAffaire.Co1valeurPro := IntToStr(nodev);
              end;
            if ((VH_GC.CleAffaire.Co2Type = 'CPT')and(VH_GC.CleAffaire.NbPartie >=2)) then
              begin
              SetParamsoc('SO_AFFCO2VALEUR',Noaff );
              VH_GC.CleAffaire.Co2valeur := IntToStr(noaff);
              SetParamsoc('SO_AFFCO2VALEURPRO',Nodev );
              //A remettre qd suppr GEtParamSoc VH_GC.CleAffaire.Co2valeurPro := IntToStr(nodev) ;
               end;
           if ((VH_GC.CleAffaire.Co3Type = 'CPT')and(VH_GC.CleAffaire.NbPartie >=3)) then
              begin
              SetParamsoc('SO_AFFCO3VALEUR',Noaff );
              VH_GC.CleAffaire.Co3valeur := IntToStr(noaff);
              SetParamsoc('SO_AFFCO3VALEURPRO',Nodev );
              //A remettre qd suppr GEtParamSoc VH_GC.CleAffaire.Co3valeurPro :=intToStr(nodev);
              end;
            end;
         if (CheckListFicBase.Checked[4] = TRUE) then
            begin
            ExecuteSQL ('delete from RESSOURCE where ARS_TYPERESSOURCE="SAL"');
            writeln (FichierLog, TraduitGA('Suppression des employés'));
            end;
        If (CheckSupCli.Checked = True)  then
          begin
          if (CheckListFicBase.Checked[5] = TRUE) then
              begin
                  // on détruit les RIB client ..
              ExecuteSQL ('delete from RIB where R_AUXILIAIRE In (SELECT T_AUXILIAIRE from TIERS where T_NATUREAUXI="CLI")');
              ExecuteSQL ('delete from TIERS where T_NATUREAUXI="CLI"');
              ExecuteSQL ('delete from TIERSCOMPL');
              writeln (FichierLog, TraduitGA('Suppression des clients'));
              end;
            end;
         if (CheckListFicBase.Checked[7] = TRUE) then
            begin
            ExecuteSQL ('delete from RESSOURCEPR ');
            writeln (FichierLog, TraduitGA('Suppression des valorisations ressources'));
            end;
      End;
Application.ProcessMessages;
TobArt := Tob.create('Fiche article',NIL,-1);
while ((not EOF (Fichier)) and (Not Stoprecup)) do
          begin
          MoveCur (False);
          If (not Traitenote) Then readln(Fichier,st);
          stTot:=St;
                    // cas ou depsui NT enrgt avec retour seulement
          if (Length(ST)<6) then continue;  //mcd 14/01/02
                            // cas ou suite au retoru, l'enrgt n'est pas correct
          if (copy(st,1,2) <> '*D') then continue;  //mcd 14/01/02
          if (St[5] = 'N') then
             Begin
             NbLigneNote:=1;
             // traitement spécifique aux bloc notes
             CleNote:=Copy(St,6,16); TraiteNote := True;
             While (Copy(St,6,16) = CleNote) do
                   Begin
                   Readln(Fichier,st);
                   if (Copy(St,6,16)=CleNote)Then Begin StTot:=StTot+#13#10+Copy(st,24,60);Inc(NbLigneNote); End;
                   End;
             End
          Else
             Begin
             TraiteNote := False;
             if (copy(st,1,2) = '*D') then
                 Begin
                 For NbLig := StrToInt(st[4]) downto 2 do
                     Begin
                     Readln(Fichier,st);
                     StTot:= stTot + st;
                     End;
                 End;
             end;
              // Enregistrement tablettes
             if (StTot[5] = '1') then
                Begin
                       If (trace <>1) then
                           Begin
                           MemoTrace.Lines[LigneMemo] := 'Traitement des tablettes'; Inc(LigneMemo);trace := 1;
                           writeln (FichierLog, '->Traitement Des tablettes');
                           Application.ProcessMessages;
                           End;
                       RecupTablette(StTot);
                End Else
            // Enregistrement valorisation assistant/prestations
             if (StTot[5] = '3') then
                Begin
                 if CheckListFicBase.Checked[7] = TRUE then begin
                       If (trace <>3) then
                           Begin
                           MemoTrace.Lines[LigneMemo] := 'Traitement des valorisations'; Inc(LigneMemo);trace := 3;
                           writeln (FichierLog, '->Traitement Des valorisations');
                           Application.ProcessMessages;
                           End;
                           // on ne traite que les enrgt ass/prestation
                       if (Sttot[6]= 'P') then RecupValor(StTot);
                       end;
                End Else
            // Enregistrement prestation
             if (StTot[5] = '4') then
                Begin
                  if ((((CheckListFicBase.Checked[0] = TRUE) And (stTot[6] = 'P')) Or
                    ((CheckListFicBase.Checked[1] = TRUE) And (stTot[6] = 'f')) Or
                    ((CheckListFicBase.Checked[2] = TRUE) And (stTot[6] = 'F')))) Then
                        begin
                        If (trace <>4) then
                           Begin
                           if (CheckListFicBase.Checked[0] = TRUE) then
                             begin
                             MemoTrace.Lines[LigneMemo] := 'Traitement des prestations'; Inc(LigneMemo);trace := 4;
                             writeln (FichierLog, '->Traitement des prestations');
                             end;
                           if (CheckListFicBase.Checked[1] = TRUE) then
                             begin
                             MemoTrace.Lines[LigneMemo] := 'Traitement des frais'; Inc(LigneMemo);trace := 4;
                             writeln (FichierLog, '->Traitement des frais');
                              end;
                           if (CheckListFicBase.Checked[2] = TRUE) then
                             begin
                             MemoTrace.Lines[LigneMemo] := 'Traitement des fournitures'; Inc(LigneMemo);trace := 4;
                             writeln (FichierLog, '->Traitement des fournitures');
                             end;
                          Application.ProcessMessages;
                           End;
                        RecupArticle (StTot);
                        end;
                    End Else
             // Enregistrements clients
             if (StTot[5] = '7') then
                Begin                               
                  if (NbtotRes <>0) then begin
                     writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbtotRes]) +TraduitGA(' Ressources à ')+FormatDateTime('ttttt',CurrentDate));
                     NbtotRes:=0;
                     end;
                  if ((CheckListFicBase.Checked[5] = TRUE)And (StTot[6] = 'C')) Then
                  begin
                  If (trace <>7) then
                       Begin
                       MemoTrace.Lines[LigneMemo] := 'Traitement des Clients'; Inc(LigneMemo);trace := 7;
                       writeln (FichierLog, '->Traitement des clients');
                       Application.ProcessMessages;
                       nbtotClt:=0;   nbint:=0;
                       End;
                  Inc (NbtotClt); Inc (NBint);
                  if (Nbint >= NbMax) then
                      begin
                      Nbint:=0;
                      MemoTrace.Lines[LigneMemo-1]:= 'Traitement des Clients : ' +Format ('%d', [nbtotClt]) +' clients à '+FormatDateTime('ttttt',CurrentDate);
                      Application.ProcessMessages;
                      end;
                    RecupClient(StTot);
                  end;
              End Else
              // Enregistrement Employés
              if (StTot[5] = '5') then
                Begin
                if (CheckListFicBase.Checked[4] = TRUE) then
                   begin
                   if (trace <>5) then
                       Begin
                       MemoTrace.Lines[LigneMemo]:= TraduitGA('Traitement des ressources'); Inc(LigneMemo);trace := 5;
                       writeln (FichierLog, TraduitGA('->Traitement Des ressources'));
                       Application.ProcessMessages;
                       nbtotRes:=0;nbint:=0;
                       End;
                  Inc (NbtotRes); Inc (NBint);
                  if (Nbint >= NbMax) then
                      begin
                      Nbint:=0;
                      MemoTrace.Lines[LigneMemo-1]:= TraduitGa('Traitement des ressources : ') +Format ('%d', [nbtotRes]) +traduitga(' employés à ')+FormatDateTime('ttttt',CurrentDate);
                      Application.ProcessMessages;
                      end;
                    RecupRessource (StTot);
                    end;
                End Else
              // Enregistrement Affaire  && Devis
              if (StTot[5] = '8') then
                Begin
                if (NbtotClt <>0) then begin
                     writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbtotClt]) +TraduitGA(' Clients à ')+FormatDateTime('ttttt',CurrentDate));
                     NbtotClt:=0;
                     end;
                 If ((CheckListFicBase.Checked[3] = TRUE) AND (StTot[6] = 'A'))  then
                     begin
                      if (trace <>8) then
                           Begin
                           MemoTrace.Lines[LigneMemo]:=TraduitGA('Traitement des affaires');
                           Inc(LigneMemo);trace := 8;
                           writeln (FichierLog, TraduitGA('->Traitement Des Affaires'));
                           Application.ProcessMessages;
                           NbInt :=0;
                            // appel recuptablette pour créer la table AFFLIENTIERS
                           St:='     **RES   Responsable technique                    ';
                           RecupTablette(St);
                           St:='     **APP   Apporteur Interne                        ';
                           RecupTablette(St);
                           End;
                       Inc (NbMis); Inc (NBint);
                      if (Nbint >= NbMax) then
                        begin
                        Nbint:=0;
                        MemoTrace.Lines[LigneMemo-1]:= TraduitGA('Traitement des affaires : ') +Format ('%d', [nbmis]) +TraduitGa(' affaires à ') +FormatDateTime('ttttt',CurrentDate);
                        Application.ProcessMessages;
                        end;
                       RecupAffaire (StTot);
                       end;
                // Devis    MCD on ne traite que les devis sur client. VOir + tard pour devis sur prospect
                If ((CheckListFicBase.Checked[6] = TRUE)  AND (StTot[6] = 'D')  AND (StTot[7] = 'C')) then
                    begin
                     if (trace <>11) then
                           Begin
                           writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbmis]) +TraduitGA(' Affaires à ')+FormatDateTime('ttttt',CurrentDate));
                           MemoTrace.Lines[LigneMemo]:=TraduitGA('Traitement prop. affaire');
                           Inc(LigneMemo);trace := 11;
                           writeln (FichierLog, TraduitGA('->Traitement des propositions d''Affaire'));
                           Application.ProcessMessages;
                           // appel recuptablette pour créer la table AFFLIENTIERS
                           St:='     **RES   Responsable technique                    ';
                           RecupTablette(St);
                           St:='     **APP   Apporteur Interne                        ';
                           RecupTablette(St);
                           nbmis:=0; nbint:=0;
                           End;
                       Inc (NbtotPro); Inc (NBint);
                      if (Nbint >= NbMax) then
                        begin
                        Nbint:=0;
                        MemoTrace.Lines[LigneMemo-1]:= TraduitGA('Traitement des propositions d''affaire : ') +Format ('%d', [nbTotPro]) +TraduitGa(' affaires à ') +FormatDateTime('ttttt',CurrentDate);
                        Application.ProcessMessages;
                        end;
                     RecupAffaire (StTot);
                     end;
                  End Else
               // Enregistrement LIgne affaires et Devis
              if (StTot[5] = '9') then
                Begin
                if (NbMis <>0) then begin
                   writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbmis]) +TraduitGA(' affaires à ')+FormatDateTime('ttttt',CurrentDate));
                   Nbmis:=0;
                   end;
                if (NbtotPro <>0) then begin
                   writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbtotPro]) +TraduitGA(' propositions d''affaire à ')+FormatDateTime('ttttt',CurrentDate));
                   NbtotPro:=0;
                   end;
                If ((CheckListFicBase.Checked[3] = TRUE) AND (StTot[6] = 'A'))  then
                     begin
                      if (trace <>9) then
                           Begin
                           MemoTrace.Lines[LigneMemo]:=TraduitGA('Traitement des Lignes affaires'); Inc(LigneMemo);trace := 9;
                           writeln (FichierLog, TraduitGA('->Traitement Des Lignes Affaires'));
                           Application.ProcessMessages;
                           Nbtot:=0; NbInt :=0;
                            // on alimente la tob article
                            // PL le 06/03/02 : INDEX 6
                            QQ := OpenSQL('SELECT GA_ARTICLE, GA_QUALIFUNITEVTE FROM ARTICLE Where GA_TYPEARTICLE="PRE" AND GA_QUALIFUNITEVTE <> "H"',True,-1,'',true) ;
                            If Not QQ.EOF then TobArt.LoadDetailDB('ARTICLE','','',QQ,True);
                            Ferme(QQ);
                           End;
                       Inc (Nbtot); Inc (NBint);
                      if (Nbint >= NbMax) then
                        begin
                        Nbint:=0;
                        MemoTrace.Lines[LigneMemo-1]:= TraduitGA('Traitement des Lignes affaires : ') +Format ('%d', [nbtot]) +TraduitGa(' Lignes affaires à ') +FormatDateTime('ttttt',CurrentDate);
                        Application.ProcessMessages;
                        end;
                        // MCD voir par la suite pour les Tâche ....
                       if (stTot[33] <> 'T')  then RecupLigneAffaire (StTot);
                       end;
                // Devis    MCD on ne traite que les devis sur client. VOir + tard pour devis sur prospect
                If ((CheckListFicBase.Checked[6] = TRUE)  AND (StTot[6] = 'D')  AND (StTot[7] = 'C')) then
                    begin
                     if (trace <>12) then
                           Begin
                            if (Nbtot <>0) then begin
                               writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbtot]) +TraduitGA(' lignes affaire à ')+FormatDateTime('ttttt',CurrentDate));
                               Nbtot:=0;
                               end;
                           MemoTrace.Lines[LigneMemo]:=TraduitGA('Traitement Lignes prop. affaire');
                           Inc(LigneMemo);trace := 12;
                           writeln (FichierLog, TraduitGA('->Traitement Des Lignes de propositions d''Affaires'));
                           Application.ProcessMessages;
                                 // on alimente la tob article
                                 // PL le  06/03/02 : INDEX 5
                            QQ := OpenSQL('SELECT GA_ARTICLE, GA_QUALIFUNITEVTE FROM ARTICLE Where GA_TYPEARTICLE="PRE" AND GA_QUALIFUNITEVTE <> "H"',True,-1,'',true) ;
                            If Not QQ.EOF then TobArt.LoadDetailDB('ARTICLE','','',QQ,True);
                            Ferme(QQ);
                            nbint:=0;
                            nbtotclt:=0;
                           End;
                       Inc (Nbtotclt); Inc (NBint);
                      if (Nbint >= NbMax) then
                        begin
                        Nbint:=0;
                        MemoTrace.Lines[LigneMemo-1]:= TraduitGA('Traitement Lignes prop. affaire : ') +Format ('%d', [nbTotclt]) +TraduitGa(' Lignes affaires à ') +FormatDateTime('ttttt',CurrentDate);
                        Application.ProcessMessages;
                        end;
                     if (stTot[33] <> 'T') then RecupLigneAffaire (StTot);
                     end;
                End Else
              // Enregistrement Bloc note
              if (StTot[5] = 'N') then
                Begin
                if (trace <>10) then
                   Begin
                   MemoTrace.Lines[LigneMemo]:= 'Traitement des bloc notes'; Inc(LigneMemo);trace := 10;
                   writeln (FichierLog, '->Traitement des bloc notes');
                   Application.ProcessMessages;
                   End;
                   RecupNotes (StTot,NbLigneNote);
               End;
             End;
    if (NbTot <>0) then begin
       writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbTot]) +TraduitGA(' Lignes  d''affaires à ')+FormatDateTime('ttttt',CurrentDate));
       end;
    if (NbTotClt <>0) then begin
       writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbtotClt]) +TraduitGA(' Lignes de propositions d''affaires à ')+FormatDateTime('ttttt',CurrentDate));
       end;
    If ((CheckListFicBase.Checked[3] = TRUE) or (CheckListFicBase.Checked[6] = TRUE))
        then if (totttc <>0) then MajPieceAff (); // pour mettre à jour dernière pièce

    writeln (FichierLog, 'Fermeture du fichier '+NomFicBase.Caption);
    CloseFile(Fichier);
        // on relit fichier tiers pour voir si il y a des codes client à facture
    if (CheckListFicBase.Checked[5] = TRUE) then ChangeClientFacture;

        // si traitement fichier affaire ou Devis, il faut, si SCOT
        // ecrire en fichier le nouveau compteur
    If ((CheckListFicBase.Checked[3] = TRUE) and (ctxScot in V_PGI.PGIContexte) ) then
        begin
         if ((VH_GC.CleAffaire.Co1Type = 'CPT')and(VH_GC.CleAffaire.NbPartie >=1)) then
            begin
            SetParamsoc('SO_AFFCO1VALEUR',VH_GC.CleAffaire.Co1valeur  );
            // A remettre GetParamSoc SetParamsoc('SO_AFFCO1VALEURPRO',VH_GC.CleAffaire.Co1valeurPro  );
            end;
          if ((VH_GC.CleAffaire.Co2Type = 'CPT')and(VH_GC.CleAffaire.NbPartie >=2)) then
            begin
            SetParamsoc('SO_AFFCO2VALEUR',VH_GC.CleAffaire.Co2valeur );
            // A remettre GetParamSOc SetParamsoc('SO_AFFCO2VALEURPRO',VH_GC.CleAffaire.Co2valeurPro  );
            end;
         if ((VH_GC.CleAffaire.Co3Type = 'CPT')and(VH_GC.CleAffaire.NbPartie >=3)) then
            begin
            SetParamsoc('SO_AFFCO3VALEUR',VH_GC.CleAffaire.Co3valeur  );
            // A remettre GetParamSoc SetParamsoc('SO_AFFCO3VALEURPRO',VH_GC.CleAffaire.Co3valeurPro  );
            end;
        end;
    end;


If ((ModuleTemps) and (Not Stoprecup)) then
   Begin
   writeln (FichierLog, '****** Reprise du Module des Temps *****');
   AssignFile(Fichier, NomFicTemps.Caption);
   writeln (FichierLog, 'Ouverture du fichier'+NomFicTemps.Caption);
   MemoTrace.Lines[LigneMemo]:= 'Ouverture du fichier des temps';
   Inc(LigneMemo);
   Application.ProcessMessages;
   Nbtot:=0; Nbint:=0;
   Reset (Fichier); //ouvre le fichier
   If (CheckSupTemps.Checked = True) then
    Begin
    // suppression des lignes de temps
    MemoTrace.Lines[LigneMemo] := 'Suppression des données existantes'; Inc(LigneMemo);
    Application.ProcessMessages;
    if (CheckListFicTemps.Checked[0] = TRUE) then
       begin
       ExecuteSQL ('delete from ACTIVITE where ACT_TYPEACTIVITE ="REA" And ACT_TYPEARTICLE ="PRE" AND ACT_DATEACTIVITE <"'+UsDateTime(StrToDate(FicTempsDateEff.text))+'"');
        writeln (FichierLog, TraduitGA('Suppression du réalisé prestations'));
        end;
     if (CheckListFicTemps.Checked[2] = TRUE) then
        begin
        ExecuteSQL ('delete from ACTIVITE where ACT_TYPEACTIVITE ="REA" And ACT_TYPEARTICLE ="FRA" AND ACT_DATEACTIVITE <"'+UsDateTime(StrToDate(FicTempsDateEff.text))+'"');
        end;
    if (CheckListFicTemps.Checked[2] = TRUE) then
       begin
       ExecuteSQL ('delete from ACTIVITE where ACT_TYPEACTIVITE ="REA" And ACT_TYPEARTICLE ="PRI" AND ACT_DATEACTIVITE <"'+UsDateTime(StrToDate(FicTempsDateEff.text))+'"');
       writeln (FichierLog, TraduitGA('Suppression du réalisé prime'));
       end;
    if (CheckListFicTemps.Checked[1] = TRUE) then
        begin
        ExecuteSQL ('delete from ACTIVITE where ACT_TYPEACTIVITE ="REA" And ACT_TYPEARTICLE ="MAR" AND ACT_DATEACTIVITE <"'+UsDateTime(StrToDate(FicTempsDateEff.text))+'"');
        writeln (FichierLog, TraduitGA('Suppression du réalisé fournitures'));
        end;
    if (CheckListFicTemps.Checked[5] = TRUE) then
        begin
        ExecuteSQL ('delete from ACTIVITE where ACT_TYPEACTIVITE ="BON" AND ACT_DATEACTIVITE <"'+UsDateTime(StrToDate(FicTempsDateEff.text))+'"');
        writeln (FichierLog, TraduitGA('Suppression du boni/mali'));
        end;
    if (CheckListFicTemps.Checked[4] = TRUE) then
        begin
        ExecuteSQL ('delete from LIGNE where GL_NATUREPIECEG = "FRE" AND GL_DATEPIECE <"'+UsDateTime(StrToDate(FicTempsDateEff.text))+'"');
        ExecuteSQL ('delete from PIECE where GP_NATUREPIECEG = "FRE" AND GP_DATEPIECE <"'+UsDateTime(StrToDate(FicTempsDateEff.text))+'"');
        writeln (FichierLog, TraduitGA('Suppression du facturé'));
        end;
    if (CheckListFicTemps.Checked[3] = TRUE) then
       begin
       ExecuteSQL ('delete from HISTOACTIVITE ');
        writeln (FichierLog, TraduitGA('Suppression Historique activité'));
        end;
   End;
   while ((not EOF (Fichier)) and (Not Stoprecup)) do
      begin
      MoveCur (False);
      Inc (Nbtot); Inc (NBint);
      if (Nbint >= NbMax) then
        begin
        Nbint:=0;
        MemoTrace.Lines[LigneMemo]:= 'Traitement de ' +Format ('%d', [nbtot]) +' lignes temps à '+FormatDateTime('ttttt',CurrentDate);
        Application.ProcessMessages;
        end;
      readln(Fichier,st);
      stTot := St;
      if (copy(st,1,2) = '*D') then
         Begin
         For NbLig := StrToInt(st[4]) downto 2 do
           Begin
           Readln(Fichier,st);
           StTot:= stTot + st;
           End;
         // Enregistrement lignes de temps
         if ((StTot[5] = '5')) then
            Begin
            if (trace <>20) then
               Begin
               if CheckListFicTemps.Checked[0] = TRUE then
                 begin
                 MemoTrace.Lines[LigneMemo]:= 'Traitement du réalisé prestation';
                 Inc(LigneMemo);trace := 20;
                 writeln (FichierLog, '->Traitement du réalisé prestation');
                 Application.ProcessMessages;
                 end;
              if CheckListFicTemps.Checked[2] = TRUE then
                 begin
                 MemoTrace.Lines[LigneMemo]:= 'Traitement du réalisé frais'; Inc(LigneMemo);trace := 20;
                 writeln (FichierLog, '->Traitement du réalisé frais');
                 Application.ProcessMessages;
                 end;
              if CheckListFicTemps.Checked[1] = TRUE then
                 begin
                 MemoTrace.Lines[LigneMemo]:= 'Traitement du réalisé fourniture'; Inc(LigneMemo);trace := 20;
                 writeln (FichierLog, '->Traitement du réalisé fourniture');
                 Application.ProcessMessages;
                 end;
              if CheckListFicTemps.Checked[4] = TRUE then
                 begin
                 MemoTrace.Lines[LigneMemo]:= 'Traitement du facturé'; Inc(LigneMemo);trace := 20;
                 writeln (FichierLog, '->Traitement du facturé');
                 Application.ProcessMessages;
                 end;
               End;
               If (StTot[59]='T')  then begin
                   If (((CheckListFicTemps.Checked[0] = TRUE) And (stTot[60] = 'P')) Or
                   ((CheckListFicTemps.Checked[2] = TRUE) And (stTot[60] = 'f')) Or
                   ((CheckListFicTemps.Checked[1] = TRUE) And (stTot[60] = 'F')))
                        Then  RecupTemps(StTot);
                    end
                 else  begin
                       if (StTot[59]='B')then begin
                                              if (CheckListFicTemps.Checked[5] = TRUE) then RecupTemps(StTot);
                                              end
                       else if (CheckListFicTemps.Checked[4] = TRUE) then RecupTempsFact(StTot);
                       end;
            End;
                // traitement historique cumulé historiques
         if (StTot[5] = 'H') then
            Begin
            if CheckListFicTemps.Checked[3] = TRUE then
               Begin
               if (trace <>21) then                          
                  begin
                   MemoTrace.Lines[LigneMemo]:= 'Traitement des historiques'; Inc(LigneMemo);trace := 21;
                   writeln (FichierLog, '->Traitement des historiques');
                   Application.ProcessMessages;
                   end;
               RecupHisto(StTot)
               End;
            End;
         End;
      End;
   writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbtot]) +TraduitGA(' Lignes de temps à ')+FormatDateTime('ttttt',CurrentDate));
   writeln (FichierLog, 'Fermeture du fichier '+NomFicTemps.Caption);
   CloseFile(Fichier);
   if (nbtot > NbMax) then Inc(LigneMemo);    // non fait sur message X lignes
   End;
If ((ModuleFacture)and (Not Stoprecup)) then
   Begin
   writeln (FichierLog, '****** Reprise du Module Facturation *****');
   AssignFile(Fichier, NomFicFact.Caption);
   writeln (FichierLog, 'Ouverture du fichier'+NomFicFact.Caption);
   MemoTrace.Lines[LigneMemo]:= 'Ouverture du fichier Facturation';
   Inc(LigneMemo);
   Application.ProcessMessages;
   if (CheckListFicFacture.Checked[1] = TRUE) then
        begin
        ExecuteSQL ('delete from LIGNE where GL_NATUREPIECEG = "FRE" AND GL_DATEPIECE <"'+UsDateTime(StrToDate(FicFacDateEff.text))+'"');
        ExecuteSQL ('delete from PIECE where GP_NATUREPIECEG = "FRE" AND GP_DATEPIECE <"'+UsDateTime(StrToDate(FicFacDateEff.text))+'"');
        writeln (FichierLog, TraduitGA('Suppression des statistiques'));
        end;

   Reset (Fichier); //ouvre le fichier
   Nbint:=0; Nbtot:=0;
   while ((not EOF (Fichier)) and (Not Stoprecup))  do
      begin
      MoveCur (False);
      readln(Fichier,st);
      stTot := St;
      if (copy(st,1,2) = '*D') then
         Begin
         For NbLig := StrToInt(st[4]) downto 2 do
           Begin
           Readln(Fichier,st);
           StTot:= stTot + st;
           End;
          if   (CheckListFicFacture.Checked[1] = true) then
               begin
               if ((stat = 'D' ) or (stat = 'C')) and  (StTot[5] = 'd') then
                  begin
                  if (trace <>31) then
                     Begin
                     MemoTrace.Lines[LigneMemo]:= 'Traitement stat détaillées';
                     Inc(LigneMemo);
                     trace := 31;
                     writeln (FichierLog, '->Traitement des stat détaillées');
                     Application.ProcessMessages;
                     nbint:=0; nbtotClt:=0;
                     End;
                  RecupStatDet (Sttot);
                  Inc (NbtotClt); Inc (NBint);
                 end;
               if ((stat = 'O' ) or (stat = 'E')) and  (StTot[5] = 'S') then
                  begin
                  if (trace <>32) then
                     Begin
                     MemoTrace.Lines[LigneMemo]:= 'Traitement stat globales';
                     Inc(LigneMemo);
                     trace := 32;
                     writeln (FichierLog, '->Traitement des stat globales');
                     Application.ProcessMessages;
                     nbint:=0; nbtotClt:=0;
                     End;
                  RecupStatGlob (Sttot);
                   Inc (NbtotClt); Inc (NBint);
                  end;
               end;
         // Enregistrement Fournitures en attente
         if (StTot[5] = 'F') then
            Begin
            if (trace <>33) then
               Begin
               MemoTrace.Lines[LigneMemo]:= 'Traitement des fournitures en attente';
               Inc(LigneMemo);
               trace := 33;
               writeln (FichierLog, '->Traitement des fournitures en attente');
               Application.ProcessMessages;
               nbint:=0; nbtot:=0;
               End;
            If ((CheckListFicFacture.Checked[0] = TRUE)
                and (copy (StTot ,211,6) = '      ')) Then begin
                            RecupFourAttente(StTot);
                           Inc (Nbtot); Inc (NBint);
                           end;
            End;
         End;
      if (Nbint >= NbMax) then
        begin
        Nbint:=0;
        if nbtot <>0 then MemoTrace.Lines[LigneMemo]:= 'Traitement de ' +Format ('%d', [nbtot]) +' lignes fournitures à '+FormatDateTime('ttttt',CurrentDate);
        if nbtotclt <>0 then MemoTrace.Lines[LigneMemo]:= 'Traitement de ' +Format ('%d', [nbtotclt]) +' lignes stat. à '+FormatDateTime('ttttt',CurrentDate);
       Application.ProcessMessages;
        end;
      End;
   if (nbtot > nbmax) then Inc (LigneMemo);   // non fait sur nbre de lignes
  if (NbTot <>0) then begin
     writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbtot]) +TraduitGA(' Fournitures ')+FormatDateTime('ttttt',CurrentDate));
     Nbtot:=0;
     end;
  if (NbTotClt <>0) then begin
     writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbtotClt]) +TraduitGA(' Statistiques ')+FormatDateTime('ttttt',CurrentDate));
     NbtotClt:=0;
     end;
   writeln (FichierLog, 'Fermeture du fichier '+NomFicFact.Caption);
   CloseFile(Fichier);
    End;


If ((ModulePlanning)and (Not Stoprecup)) then
   Begin
   AssignFile(Fichier, NomFicPlan.Caption);
   writeln (FichierLog, 'Ouverture du fichier'+NomFicPlan.Caption);
   MemoTrace.Lines[LigneMemo]:= 'Ouverture du fichier Planning'; Inc(LigneMemo);
   Application.ProcessMessages;
   Nbtot :=0; Nbint:=0;
   Reset (Fichier); //ouvre le fichier
   If (CheckSupPlan.Checked = True) then
    Begin
    // suppression des lignes de planning
    MemoTrace.Lines[LigneMemo] := 'Suppression des données existantes'; Inc(LigneMemo);
    Application.ProcessMessages;
    If (CheckListFicPlan.Checked[0] = TRUE) then
       begin
       ExecuteSQL ('delete from ACTIVITE where ACT_TYPEACTIVITE ="PRE"');
        writeln (FichierLog, TraduitGA('Suppression du planning'));
        end;

   End;
   while ((not EOF (Fichier))and (Not Stoprecup)) do
      begin
      MoveCur (False);
      Inc (Nbtot); Inc (NBint);
      if (Nbint >= NbMax) then
        begin
        Nbint:=0;
        MemoTrace.Lines[LigneMemo]:= 'Traitement de ' +Format ('%d', [nbtot]) +' lignes planning à '+FormatDateTime('ttttt',CurrentDate);
        Application.ProcessMessages;
        end;
      readln(Fichier,st);
      stTot := St;
      if (copy(st,1,2) = '*D') then
         Begin
         For NbLig := StrToInt(st[4]) downto 2 do
           Begin
           Readln(Fichier,st);
           StTot:= stTot + st;
           End;
         // Enregistrement planning
         if (StTot[5] = 'L')  then
            Begin
            if (trace <>41) then
               Begin
               MemoTrace.Lines[LigneMemo]:= 'Traitement du planning';
               Inc(LigneMemo);trace := 41;
               writeln (FichierLog, '->Traitement du planning');
               Application.ProcessMessages;
               End;
            if (CheckListFicPlan.Checked[0] = TRUE)  Then RecupPlanning(StTot);
            End;
         // Enregistrement Indisponibilités  MCD à faire
         if (StTot[5] = 'I')  then
            Begin
            if (trace <>42) then
               Begin
               MemoTrace.Lines[LigneMemo]:= 'Traitement des indisponibilités';
               Inc(LigneMemo);trace := 42;
               writeln (FichierLog, '->Traitement des indisponibilités');
               Application.ProcessMessages;
               End;
            if (CheckListFicPlan.Checked[1] = TRUE)  Then  ; // MCD ????

            End;
         // MCD il faudra faire la récup des semaines types
         // CheckListFicPlan.Checked[2]
         End;
      End;
   writeln (FichierLog, 'Fermeture du fichier '+NomFicPlan.Caption);
   CloseFile(Fichier);
   End;

If ((ModuleAchat)and (Not Stoprecup)) then
   Begin
   AssignFile(Fichier, NomFicAchat.Caption);
   writeln (FichierLog, 'Ouverture du fichier'+NomFicAchat.Caption);
   MemoTrace.Lines[LigneMemo]:= 'Ouverture du fichier Achat'; Inc(LigneMemo);
   Application.ProcessMessages;
   Nbtot :=0; Nbint:=0;
   Reset (Fichier); //ouvre le fichier
    // suppression des fournisseurs
    MemoTrace.Lines[LigneMemo] := 'Suppression des données existantes'; Inc(LigneMemo);
    Application.ProcessMessages;
    If (CheckListFicAchat.Checked[0] = TRUE) and (CheckSupAchat.Checked = True) then
       begin
       ExecuteSQL ('delete from TIERS where T_NATUREAUXI ="FOU"');
        writeln (FichierLog, TraduitGA('Suppression des fournisseurs'));
        end;
    If (CheckListFicAchat.Checked[1] = TRUE) and (CheckSupAchat.Checked = True) then
       begin
       ExecuteSQL ('delete from CATALOGU');
       writeln (FichierLog, TraduitGA('Suppression des art/fournisseurs'));
       end;

   while ((not EOF (Fichier))and (Not Stoprecup)) do
      begin
      MoveCur (False);
      readln(Fichier,st);
      stTot := St;
      if (copy(st,1,2) = '*D') then
         Begin
         For NbLig := StrToInt(st[4]) downto 2 do
           Begin
           Readln(Fichier,st);
           StTot:= stTot + st;
           End;
         // Enregistrement fournissuer
         if (StTot[5] = '1')  then
            Begin
            if (trace <>51) then
               Begin
               MemoTrace.Lines[LigneMemo]:= 'Traitement des fournisseurs';
               Inc(LigneMemo);trace := 51;
               writeln (FichierLog, '->Traitement des fournisseurs');
               Application.ProcessMessages;
               End;
            if (CheckListFicAchat.Checked[0] = TRUE)  Then RecupFour(StTot);
            Inc (NbtotClt); Inc (NBint);
            if (Nbint >= NbMax) then
              begin
              Nbint:=0;
              MemoTrace.Lines[LigneMemo]:= 'Traitement de ' +Format ('%d', [nbtotClt]) +' Fournisseurs à '+FormatDateTime('ttttt',CurrentDate);
              Application.ProcessMessages;
              end;
            End else
         // Enregistrement Articles /fournisseurs
         if (StTot[5] = '2')  then
            Begin
            if (trace <>52) then
               Begin
               MemoTrace.Lines[LigneMemo]:= 'Traitement des Art/Fournisseur';
               Inc(LigneMemo);trace := 52;
               writeln (FichierLog, '->Traitement des Art/Fournisseur');
               Application.ProcessMessages;
               End;
            if (CheckListFicAchat.Checked[1] = TRUE)  Then RecupArtFour(StTot);
            Inc (NbtotClt); Inc (NBint);
            if (Nbint >= NbMax) then
              begin
              Nbint:=0;
              MemoTrace.Lines[LigneMemo]:= 'Traitement de ' + Format ('%d', [nbtotClt]) +' Art/Fournisseurs à ' + FormatDateTime('ttttt',CurrentDate);
              Application.ProcessMessages;
              end;
            End;

         End;
         // voir si recup des fctures, cde fournisseurs.
      End;
    if (NbTotClt <>0) then begin
     writeln (FichierLog, 'Traitement de : ' +Format ('%d', [nbtotClt]) +TraduitGA(' Fournisseurs ')+FormatDateTime('ttttt',CurrentDate));
     end;
    writeln (FichierLog, 'Fermeture du fichier '+NomFicAchat.Caption);
   CloseFile(Fichier);
   End;


MemoTrace.Lines[LigneMemo]:= 'Traitement Terminé';
Application.ProcessMessages;

Writeln (FichierLog,TraduitGA( '****** Fin de reprise des données Tempo *****--------------------------------------------'));
CloseFile (FichierLog);
FiniMove ();
self.close ;
if (not StopRecup) then PGIInfo(TraduitGA('Traitement de reprise des données Tempo terminé'),TitreHalley);
EnableControls(Self, True);
ListeAffaire.free; // mcd 31/03/03
T_PIECESAUV.Free; T_PIECESAUV :=NIL;
TobArt.free; TobArt := NIl;
end;

procedure TFAsRecupTempo.bPrecedentClick(Sender: TObject);
begin
bFin.enabled := False;
Dec (Nopage);
inherited;
end;

procedure TFAsRecupTempo.FormShow(Sender: TObject);
var cpt : Integer;
begin
  inherited;
bFin.enabled := False;
FicTempsDateEff.text:='31/12/2099';
FicFacDateEff.text:='31/12/2099';
NbPage :=3;
for cpt := 0 to (CheckListFicBase.items.Count - 1) do begin CheckListFicBase.Checked[cpt] := True;  End;
for cpt := 0 to (CheckListFicAchat.items.Count - 1)  do begin CheckListFicAchat.Checked[cpt] := True; End;
for cpt := 0 to (CheckListFicTemps.items.Count - 1)  do begin CheckListFicTemps.Checked[cpt] := True;  End;
for cpt := 0 to (CheckListFicFacture.items.Count - 1)  do begin CheckListFicFacture.Checked[cpt] := True;  End;
for cpt := 0 to (CheckListFicPlan.items.Count - 1)  do begin CheckListFicPlan.Checked[cpt] := True;  End;
If not(ctxScot in V_PGI.PGIContexte) then
  begin
  PRODUC.visible :=FALSE;
  Gamme2.visible :=FALSE;
  Gamme2.Checked:=False;
  Produc.Checked:=True;
  end
else
  begin
  if (GetParamsoc('SO_AFLIENDP') = true) then begin
    GAMME2.visible :=TRUE;
    PRODUC.visible :=TRUE;
    Gamme2.Checked:=True;
    Produc.Checked:=false;
    end
  else begin
    GAMME2.visible :=False;
    PRODUC.visible :=False;
    Gamme2.Checked:=False;
    Produc.Checked:=True;
    end;
  end;

end;

procedure TFAsRecupTempo.BParamBaseClick(Sender: TObject);
begin
  inherited;
FParam.ShowModal;
End;

function TFAsRecupTempo.LectureParamBase:Boolean;
Var st : string;
    StTot : string;
    Fichier: textfile;
    Nb ,ano :integer;
    TobDetFJur : TOB;
    TobDetEntite: TOB;
begin
  inherited;
  Result:=True; //mcd 08/10/03
  AssignFile(Fichier, NomFicBase.Caption);
  Lng_gene :=   VH^.Cpta[fbGene].Lg;
  Reset (Fichier);
  NbMode :=0;
  FactureGlobale := False;
  LiaisonCompta := False;
  LngCpte:=6;
  readln(Fichier,stTot);
  readln(Fichier,stTot);
    // on met par défaut à vrai, toutes les listes de choix possibles
  For Nb:= StrToInt(stTot[4]) downto 2 do
             Begin
             Readln(Fichier,st);
             StTot:= stTot + st;
             End;

  Nb:= StrToInt(StTot[200]);
  FParamCli.NbStatCli.Caption := StTot[200];
  Exer := StTot[198];
  if exer =' ' then Exer :='N' ; //mcd 23/10/03 pour cas TEMPO, pas renseigné

  FParamCli.EditCollectifClient.text := VH^.DefautCli;
  if (Nb > 0) then begin
     FParamCli.CheckStat1Cli.Caption := Copy (StTot,201,32);
     if sttot[804]='O' then FParamCli.CheckStat1Cli.Caption :='O '+ Copy (StTot,201,32);
     end
  else
    Begin
    FParamCli.CheckStat1Cli.Enabled:=False;
    FParamCli.CheckStat1Cli.Checked:=False;
    End;
  if (Nb > 1) then  begin
     FParamCli.CheckStat2Cli.Caption := Copy (StTot,233,32);
     if sttot[805]='O' then FParamCli.CheckStat2Cli.Caption :='O '+ Copy (StTot,233,32);
     end
  else
    Begin
    FParamCli.CheckStat2Cli.Enabled:=False;
    FParamCli.CheckStat2Cli.Checked:=False;
    End;
  if (Nb > 2) then  begin
     FParamCli.CheckStat3Cli.Caption := Copy (StTot,265,32);
     if sttot[806]='O' then FParamCli.CheckStat3Cli.Caption :='O '+ Copy (StTot,265,32);
     end
  else
    Begin
    FParamCli.CheckStat3Cli.Enabled:=False;
    FParamCli.CheckStat3Cli.Checked:=False;
    End;
  if (Nb > 3) then  begin
     FParamCli.CheckStat4Cli.Caption := Copy (StTot,297,32);
     if sttot[807]='O' then FParamCli.CheckStat4Cli.Caption :='O '+ Copy (StTot,297,32);
     end
  else
    Begin
    FParamCli.CheckStat4Cli.Enabled:=False;
    FParamCli.CheckStat4Cli.Checked:=False;
    End;
  if GetParamSocSecur('SO_GCVENTCPTAART', False) = FALSE then
     begin
     FParam.PreCOmpt.Enabled:=False;
     FParam.ComboPreCompt.Enabled:=False;
     end;
  if GetParamSocSecur('SO_GCVENTCPTAAFF', False) = FALSE then
     begin
     FParam.AffCOmpt.Enabled:=False;
     FParam.ComboAffCompt.Enabled:=False;
     end;
  if GetParamSocSecur('SO_GCVENTCPTATIERS', False) = FALSE then
     begin
     FParamCli.LabelComptaTiers.Enabled:=False;
     FParamCli.ComboFamComptaTiers.Enabled:=False;
     end;

  DebExer := Copy (Sttot, 171,8);
  FinExer := Copy (Sttot, 179,8);
  Nb:= StrToInt(StTot[361]);
  FParam.NbStatPrest.Caption := StTot[361];
  if (Nb > 0) then FParam.CheckStat1Prest.Caption :=  Copy (StTot,362,32)
  else Begin FParam.CheckStat1Prest.Enabled:=False; FParam.CheckStat1Prest.Checked:=False; End;
  if (Nb > 1) then FParam.CheckStat2Prest.Caption :=  Copy (StTot,394,32)
  else Begin FParam.CheckStat2Prest.Enabled:=False; FParam.CheckStat2Prest.Checked:=False; End;

  Nb:= StrToInt(StTot[490]);
  FParam.NbStatAff.Caption := StTot[490];
  FParam.CheckApport.Checked:=False;
  if (Nb > 0) then FParam.CheckStat1Aff.Caption := Copy (StTot,491,32)
  else Begin FParam.CheckStat1Aff.Enabled:=False; FParam.CheckStat1Aff.Checked:=False; End;
  if (Nb > 1) then FParam.CheckStat2Aff.Caption := Copy (StTot,523,32)
  else Begin FParam.CheckStat2Aff.Enabled:=False; FParam.CheckStat2Aff.Checked:=False; End;
   if (Nb > 2) then FParam.CheckStat3Aff.Caption := Copy (StTot,555,32)
  else Begin FParam.CheckStat3Aff.Enabled:=False; FParam.CheckStat3Aff.Checked:=False; End;
   if (Nb > 3) then FParam.CheckStat4Aff.Caption := Copy (StTot,587,32)
  else Begin FParam.CheckStat4Aff.Enabled:=False; FParam.CheckStat4Aff.Checked:=False; End;
  multi := Sttot[971];
  NoEntite := StrToInt(Copy(Sttot,972,2));
  if multi = 'O' then begin
     if Strtoint (Copy(Sttot,972,2)) = 1
        then begin FParam.CheckStat1Aff.Enabled:=False; FParam.CheckStat1Aff.Checked:=False; End
        else  if Strtoint (Copy(Sttot,972,2)) = 2 then begin FParam.CheckStat2Aff.Enabled:=False; FParam.CheckStat2Aff.Checked:=False; End
        else  if Strtoint (Copy(Sttot,972,2)) = 3 then begin FParam.CheckStat3Aff.Enabled:=False; FParam.CheckStat3Aff.Checked:=False; End
        else  begin FParam.CheckStat4Aff.Enabled:=False; FParam.CheckStat4Aff.Checked:=False; End;
     Bentite.enabled := True;
     Bentite.visible := True;
     end
  else begin
       Bentite.enabled := False;
       Bentite.visible := False;
       end;


  Nb:= StrToInt(StTot[619]);
  FParam.NbStatEmp.Caption := StTot[619];
  if (Nb > 0) then FParam.CheckStat1Emp.Caption := Copy (StTot,620,32)
  else Begin FParam.CheckStat1Emp.Enabled:=False; FParam.CheckStat1Emp.Checked:=False; End;
  if (Nb > 1) then FParam.CheckStat2Emp.Caption := Copy (StTot,652,32)
  else Begin FParam.CheckStat2Emp.Enabled:=False; FParam.CheckStat2Emp.Checked:=False; End;
  if (Nb > 2) then FParam.CheckStat3Emp.Caption := Copy (StTot,684,32)
  else Begin FParam.CheckStat3Emp.Enabled:=False; FParam.CheckStat3Emp.Checked:=False; End;
  if (Nb > 3) then FParam.CheckStat4Emp.Caption := Copy (StTot,716,32)
  else Begin FParam.CheckStat4Emp.Enabled:=False; FParam.CheckStat4Emp.Checked:=False; End;

  // initialisation des affectations dans les combo
     //client
  FParamCli.ComboStat1Cli.ItemIndex := 0;FParamCli.ComboStat2Cli.ItemIndex := 0;
  FParamCli.ComboStat3Cli.ItemIndex := 0;FParamCli.ComboStat4Cli.ItemIndex := 0;
  if (FParamCli.CheckStat1Cli.Enabled = False) then
    Begin FParamCli.ComboStat1Cli.Enabled := False; FParamCli.ComboStat1Cli.ItemIndex := 5; end;
  if (FParamCli.CheckStat2Cli.Enabled = False) then
    Begin FParamCli.ComboStat2Cli.Enabled := False; FParamCli.ComboStat2Cli.ItemIndex := 5; end;
  if (FParamCli.CheckStat3Cli.Enabled = False) then
    Begin FParamCli.ComboStat3Cli.Enabled := False; FParamCli.ComboStat3Cli.ItemIndex := 5; end;
  if (FParamCli.CheckStat4Cli.Enabled = False) then
    Begin FParamCli.ComboStat4Cli.Enabled := False; FParamCli.ComboStat4Cli.ItemIndex := 5; end;
          //affaire
  fparam.ComboStat1aff.ItemIndex := 0;fparam.ComboStat2aff.ItemIndex := 0;
  fparam.ComboStat3aff.ItemIndex := 0;fparam.ComboStat4aff.ItemIndex := 0;
  if (fparam.CheckStat1aff.Enabled = False) then
    Begin fparam.ComboStat1aff.Enabled := False; fparam.ComboStat1aff.ItemIndex := 5; end;
  if (fparam.CheckStat2aff.Enabled = False) then
    Begin fparam.ComboStat2aff.Enabled := False; fparam.ComboStat2aff.ItemIndex := 5; end;
  if (fparam.CheckStat3aff.Enabled = False) then
    Begin fparam.ComboStat3aff.Enabled := False; fparam.ComboStat3aff.ItemIndex := 5; end;
  if (fparam.CheckStat4aff.Enabled = False) then
    Begin fparam.ComboStat4aff.Enabled := False; fparam.ComboStat4aff.ItemIndex := 5; end;
    //ressource
  fparam.ComboStat1emp.ItemIndex := 0;fparam.ComboStat2emp.ItemIndex := 0;
  fparam.ComboStat3emp.ItemIndex := 0;fparam.ComboStat4emp.ItemIndex := 0;
  if (fparam.CheckStat1emp.Enabled = False) then
    Begin fparam.ComboStat1emp.Enabled := False; fparam.ComboStat1emp.ItemIndex := 5; end;
  if (fparam.CheckStat2emp.Enabled = False) then
    Begin fparam.ComboStat2emp.Enabled := False; fparam.ComboStat2emp.ItemIndex := 5; end;
  if (fparam.CheckStat3emp.Enabled = False) then
    Begin fparam.ComboStat3emp.Enabled := False; fparam.ComboStat3emp.ItemIndex := 5; end;
  if (fparam.CheckStat4emp.Enabled = False) then
    Begin fparam.ComboStat4emp.Enabled := False; fparam.ComboStat4emp.ItemIndex := 5; end;

  Nb := StrToInt(Copy (StTot,812,3));
  // test de la version des fichiers de base
  If (Nb < 17) then result :=False else result :=True;
  // Affichage de la monnaie de tenue du dossier pour tester la cohérence
  If (Copy (Sttot,974,1)='F') Then LabelMonTempo.Caption:='Francs' else LabelMonTempo.Caption:='Euros';
  If (VH^.TenueEuro = False) Then LabelMonPGI.Caption:='Francs' else LabelMonPGI.Caption:='Euros';
  if (LabelMonTempo.Caption <> LabelMonPGI.Caption) then
   begin
   PGIBoxAf('Monnaies de tenue différentes',TitreHalley);
   Result:=False; //mcd 10/03/02
   exit;  //mcd 08/10/03 plus de bascule euro
   end;
  // récupération du dernier n° d'affaire TEMPO
  noaff := StrToint (Copy (sttot, 925,6));

  FamPrime := '';
  // récupération de la famille de prime  sur enrgt param temps
  While not EOF(Fichier) do
        begin
        If (copy(st,1,5)='*D01t')then
           Begin
                // mcd 15/06/01 ajout trim
           If ctxTempo in V_PGI.PGIContexte then FamPrime:=trim(copy(st,222,3));
           If(St[168]= 'O')Then Begin LiaisonIsis:=True; FParam.LabelIsis.Caption := 'Oui'; FParam.CheckChangementIsis.Enabled:= True; End
           Else Begin LiaisonIsis := False; FParam.LabelIsis.Caption := 'Non';FParam.CheckChangementIsis.Enabled:= False;  End;
           If(St[225]= 'O')Then  FraisCompta:=True
           Else  FraisCompta := False;
           ValPr := St[210];    // récupère mode valorisation PR
           ValPv := St[7];    // récupère mode valorisation PV
           ano :=0;
           if ((Valpr = 'E') and (VH_GC.AFValoActPR <> 'RES')) then ano :=1;
           if ((Valpr = 'P') and (VH_GC.AFValoActPR <> 'ART')) then ano :=1;
           if ((Valpv = 'E') and (VH_GC.AFValoActPV <> 'RES')) then ano :=1;
           if ((Valpv = 'P') and (VH_GC.AFValoActPV <> 'ART')) then ano :=1;
           if (ano = 1) then PGIBoxAf (TraduitGA(Format('Valorisation TEMPO PR : %s, PV : %s . PGI PR : %s, PV: %s. Il faut abandonner le traitement',[valpr,valpv,VH_GC.AFValoActPR,VH_GC.AFValoActPV])),TitreHalley);
           Break;
           End;
        readln(Fichier,st);
        end;
        // force à une valeur autre si blanc, pour ne pas mettre prime
        // sur tous les codes frais n'ayant pas de code famille
  if (FamPRime = '') then FamPrime := '(((';
        // on récupère si lien S7 ou pas  sur enrgt param facture
  While not EOF(Fichier) do
        begin
        If (copy(st,1,5)='*D04f')then
            Begin
            StTot:= st;
             For Nb := 4 downto 2 do
               Begin
               Readln(Fichier,st);
               StTot:= stTot + st;
               End;
            S7 :=copy(sttot,978,1);
            stat :=copy(sttot,605,1);
            if stat = 'N' then CheckListFicFacture.Checked[1] := False;
            Imptra :=copy(sttot,63,1);
            ImpLcr :=copy(sttot,64,1);
            DebEnc :=copy(sttot,503,1);
            MttMin :=StrToInt (copy(sttot,68,4));
            MRrempl :=trim(copy(sttot,102,2));
            if (MRRempl <> '') then
                begin
                    // on crée le mode de remplacement par defaut
                MRRempl := RecupModeRegl ('N', MRRempl,'', 0, 0, 1, 0, 0)
                end;
            If(Sttot[61]= 'O')Then  FactureGlobale:=True
            Else  FactureGlobale := False;
            If(Sttot[493]= 'O')Then  LiaisonCompta:=True
            Else  LiaisonCompta := False;
            LngCpte:= StrToInt(Copy(Sttot,494,2));
            Break;
            End;
        readln(Fichier,st);
        end;
        // pour récupération info module lettre mission/devis
while not EOF(Fichier) do
        begin
        If (copy(st,1,5)='*D01d')then Begin Nodev:=StrToInt(copy(st,6,6)); Break; End;
        readln(Fichier,st);
        end;
        // récupération si planning en Heure ou centième
While not EOF(Fichier) do
        begin
        If (copy(st,1,5)='*D01p')then Begin HeurePlanning:=copy(st,50,1); Break; End;
        readln(Fichier,st);
        end;

while not EOF (Fichier)  do
    begin
    readln(Fichier,sttot);
    if (copy(sttot,1,5)='*D011') then begin   // enrgt table
      // Récupération des formes juridique. création TOB
      if (copy(sttot,6,2) = 'FJ') then
          Begin
          TobdetFjur := Tob.create('detail Fjur', FparamFjur.TobFjur, -1);
          TobdetFJur.AddChampSup ('Code',False);
          TobdetFJur.AddChampSup ('Libelle',False);
          TobdetFJur.AddChampSup ('Nouveau code',False);
          TobdetFJur.PutValue ('Code', copy (sttot, 8,5));
          TobdetFJur.PutValue ('Libelle', copy (sttot,14 ,35));
          TobdetFJur.PutValue ('Nouveau code', '');
          End;
      if (copy(sttot,6,2) = 'EJ') then     
          Begin
          TobdetEntite := Tob.create('detail Entite', FparamEntite .TobEntite , -1);
          TobdetEntite .AddChampSup ('Code',False);
          TobdetEntite .AddChampSup ('Libelle',False);
          TobdetEntite .AddChampSup ('Etablissement ',False);
          TobdetEntite .PutValue ('Code', copy (sttot, 8,1));
          TobdetEntite .PutValue ('Libelle', copy (sttot,14 ,35));
          TobdetEntite .PutValue ('Etablissement ', '');
          End;
       end
    else if (copy(sttot,1,5)='*D012')then break;
  end;
  CloseFile(Fichier);
End;


procedure TFAsRecupTempo.ImportTob ( TobImport : TOB;BUpdate:boolean );
begin
if (Bupdate) then TobImport.InsertOrUpdateDB
else TobImport.InsertDB(nil);
End;

procedure TFAsRecupTempo.AffectStat;
Var Rang, Nb : Integer;
begin

     // alimente les options famille affaire
FamAffaire := 0; Rang:=1;
NbTva :=0;
nbjour:=0;

If (FParam.RadioFamAffStat.Checked = True) then
   Begin
   If (FParam.CheckStat1Aff.Checked = True) then Inc(Rang);
   If (FParam.CheckStat2Aff.Checked = True) then Inc(Rang);
   FamAffaire := Rang;
   End;
   // on alimente la zone de début du code prestation/frais/fourniture
debutpre :='';
debutfour :='';
debutfrais :='';
If (FParam.CheckCodePres.Checked = True) then
  begin
  debutpre :='P';
  debutfour :='F';
  debutfrais :='FR';
  end;

   // alimente les options des stat cleint
Stat1cli := 0;Stat2cli := 0;Stat3cli := 0;Stat4cli:= 0;Rang:= 1;
Stat1Aff :=0; Stat1Res:=0;
Stat2Aff :=0; Stat2Res:=0;
Stat3Aff :=0; Stat3Res:=0;
Stat4Aff :=0; Stat4Res:=0;
If (FParamCli.CheckStat1Cli.Checked = True) then
   Begin
   If (FParamCli.ComboStat1Cli.ItemIndex =0)then Begin Stat1Cli:= Rang; Inc(Rang); End Else
   If (FParamCli.ComboStat1Cli.ItemIndex =1)then Stat1Cli:=6 Else
   If (FParamCli.ComboStat1Cli.ItemIndex =2)then Stat1Cli:=7 Else
   If (FParamCli.ComboStat1Cli.ItemIndex =3)then Stat1Cli:=8 Else
   If (FParamCli.ComboStat1Cli.ItemIndex =4)then Stat1Cli:=9  ;
    End;
If (FParamCli.CheckStat2Cli.Checked = True) then
   Begin
   If (FParamCli.ComboStat2Cli.ItemIndex =0)then Begin Stat2Cli:= Rang; Inc(Rang); End Else
   // If (FParamCli.ComboStat2Cli.ItemIndex =1)then  Stat2Cli:=5  Else
   If (FParamCli.ComboStat2Cli.ItemIndex =1)then Stat2Cli:=6 Else
   If (FParamCli.ComboStat2Cli.ItemIndex =2)then Stat2Cli:=7 Else
   If (FParamCli.ComboStat2Cli.ItemIndex =3)then Stat2Cli:=8 Else
   If (FParamCli.ComboStat2Cli.ItemIndex =4)then Stat2Cli:=9  ;
   End;
If (FParamCli.CheckStat3Cli.Checked = True) then
   Begin
   If (FParamCli.ComboStat3Cli.ItemIndex =0)then Begin Stat3Cli:= Rang; Inc(Rang); End Else
   //If (FParamCli.ComboStat3Cli.ItemIndex =1)then  Stat3Cli:=5  Else
   If (FParamCli.ComboStat3Cli.ItemIndex =1)then Stat3Cli:=6 Else
   If (FParamCli.ComboStat3Cli.ItemIndex =2)then Stat3Cli:=7 Else
   If (FParamCli.ComboStat3Cli.ItemIndex =3)then Stat3Cli:=8 Else
   If (FParamCli.ComboStat3Cli.ItemIndex =4)then Stat3Cli:=9  ;
  End;
If (FParamCli.CheckStat4Cli.Checked = True) then
   Begin
   If (FParamCli.ComboStat4Cli.ItemIndex =0)then Begin Stat4Cli:= Rang; End Else
   //If (FParamCli.ComboStat4Cli.ItemIndex =1)then  Stat4Cli:=5  Else
   If (FParamCli.ComboStat4Cli.ItemIndex =1)then Stat4Cli:=6 Else
   If (FParamCli.ComboStat4Cli.ItemIndex =2)then Stat4Cli:=7 Else
   If (FParamCli.ComboStat4Cli.ItemIndex =3)then Stat4Cli:=8 Else
   If (FParamCli.ComboStat4Cli.ItemIndex =4)then Stat4Cli:=9  ;
   End;
   // on alimente les GT et origine si gérer pour créer les tables

 For Nb:= 1 to 27 do
     Begin
     gt[Nb]:=0;
     End;
if FParamCli.ComboTableLibre1.ItemIndex > 0 then gt[FParamCli.ComboTableLibre1.ItemIndex]:=1;
if FParamCli.ComboTableLibre2.ItemIndex > 0 then gt[FParamCli.ComboTableLibre2.ItemIndex]:=1;
if FParamCli.ComboTableLibre3.ItemIndex > 0 then gt[FParamCli.ComboTableLibre3.ItemIndex]:=1;
if FParamCli.ComboTableLibre4.ItemIndex > 0 then gt[FParamCli.ComboTableLibre4.ItemIndex]:=1;
if FParamCli.ComboTableLibre5.ItemIndex > 0 then gt[FParamCli.ComboTableLibre5.ItemIndex]:=1;
if FParamCli.ComboTableLibre6.ItemIndex > 0 then gt[FParamCli.ComboTableLibre6.ItemIndex]:=1;
if FParamCli.ComboTableLibre7.ItemIndex > 0 then gt[FParamCli.ComboTableLibre7.ItemIndex]:=1;
if FParamCli.ComboTableLibre8.ItemIndex > 0 then gt[FParamCli.ComboTableLibre8.ItemIndex]:=1;
if FParamCli.ComboTableLibre9.ItemIndex > 0 then gt[FParamCli.ComboTableLibre9.ItemIndex]:=1;
if FParamCli.ComboTableLibre10.ItemIndex > 0 then gt[FParamCli.ComboTableLibre10.ItemIndex]:=1;

  // pour affaire
Rang:=1;
If (FParam.CheckStat1Aff.Checked = True) then
   Begin
   If (FParam.ComboStat1aff.ItemIndex =0)then Begin Stat1aff:= Rang; Inc(Rang); End Else
   If (FParam.ComboStat1aff.ItemIndex =1)then Stat1aff:=5 Else
   If (FParam.ComboStat1aff.ItemIndex =2)then Stat1aff:=6 Else
   If (Fparam.ComboStat1aff.ItemIndex =3)then Stat1aff:=7  ;
    End;
If (Fparam.CheckStat2aff.Checked = True) then
   Begin
   If (Fparam.ComboStat2aff.ItemIndex =0)then Begin Stat2aff:= Rang; Inc(Rang); End Else
   If (Fparam.ComboStat2aff.ItemIndex =1)then Stat2aff:=5 Else
   If (Fparam.ComboStat2aff.ItemIndex =2)then Stat2aff:=6 Else
   If (Fparam.ComboStat2aff.ItemIndex =3)then Stat2aff:=7 ;
   End;
If (Fparam.CheckStat3aff.Checked = True) then
   Begin
   If (Fparam.ComboStat3aff.ItemIndex =0)then Begin Stat3aff:= Rang; Inc(Rang); End Else
   If (Fparam.ComboStat3aff.ItemIndex =1)then Stat3aff:=5 Else
   If (Fparam.ComboStat3aff.ItemIndex =2)then Stat3aff:=6 Else
   If (Fparam.ComboStat3aff.ItemIndex =3)then Stat3aff:=7 ;
  End;
If (Fparam.CheckStat4aff.Checked = True) then
   Begin
   If (Fparam.ComboStat4aff.ItemIndex =0)then Begin Stat4aff:= Rang; End Else
   If (Fparam.ComboStat4aff.ItemIndex =1)then Stat4aff:=5 Else
   If (Fparam.ComboStat4aff.ItemIndex =2)then Stat4aff:=6 Else
   If (Fparam.ComboStat4aff.ItemIndex =3)then Stat4aff:=7 ;
   End;
// Pour ressource
Rang:=1;
If (FParam.CheckStat1emp.Checked = True) then
   Begin
   If (FParam.ComboStat1Emp.ItemIndex =0)then Begin Stat1res:= Rang; Inc(Rang); End Else
   If (FParam.ComboStat1Emp.ItemIndex =1)then Stat1res:=5 Else
   If (FParam.ComboStat1Emp.ItemIndex =2)then Stat1res:=6 Else
   If (Fparam.ComboStat1Emp.ItemIndex =3)then Stat1res:=7  ;
    End;
If (Fparam.CheckStat2Emp.Checked = True) then
   Begin
   If (Fparam.ComboStat2Emp.ItemIndex =0)then Begin Stat2res:= Rang; Inc(Rang); End Else
   If (Fparam.ComboStat2Emp.ItemIndex =1)then Stat2res:=5 Else
   If (Fparam.ComboStat2Emp.ItemIndex =2)then Stat2res:=6 Else
   If (Fparam.ComboStat2Emp.ItemIndex =3)then Stat2res:=7 ;
   End;
If (Fparam.CheckStat3Emp.Checked = True) then
   Begin
   If (Fparam.ComboStat3Emp.ItemIndex =0)then Begin Stat3res:= Rang; Inc(Rang); End Else
   If (Fparam.ComboStat3Emp.ItemIndex =1)then Stat3res:=5 Else
   If (Fparam.ComboStat3Emp.ItemIndex =2)then Stat3res:=6 Else
   If (Fparam.ComboStat3EMp.ItemIndex =3)then Stat3res:=7 ;
  End;
If (Fparam.CheckStat4Emp.Checked = True) then
   Begin
   If (Fparam.ComboStat4Emp.ItemIndex =0)then Begin Stat4res:= Rang; End Else
   If (Fparam.ComboStat4Emp.ItemIndex =1)then Stat4res:=5 Else
   If (Fparam.ComboStat4Emp.ItemIndex =2)then Stat4res:=6 Else
   If (Fparam.ComboStat4Emp.ItemIndex =3)then Stat4res:=7 ;
   End;

end ;

//********************Traitement Article ****************
procedure TFAsRecupTempo.RecupArticle (stTot:string);
var T_ARTICLE : TOB;
    TOM_ARTICLEVAR : TOM;
    champ,TypeArt, valeur:string;
    ii : integer;
Begin
if (CheckRepriseSup.checked=False) then  begin
    champ :=Copy (Sttot, 310,8);
    if (champ = '00000000') then champ:='        ';
    If (Trim(Copy (champ,1,8)) <>'') Then exit;
   end;
     // cas frais en %. on crée un enrgt port
If (Sttot[6]= 'f') and (Sttot[331]= 'O') then begin
   RecupPort(sttot);
   exit;
   end;
T_ARTICLE := TOB.CREATE ('ARTICLE', NIL, -1);
TOM_ARTICLEVAR := CreateTom('ARTICLE', Nil, False, False);

If (stTot[6] = 'P') then
   begin
   TypeARt:='PRE';
   champ := debutpre;
   end
Else If (stTot[6] = 'F') then
   begin
   TypeArt:='MAR';
   champ := debutfour;
   end
Else If  (stTot[6] = 'f') then
    Begin
    If (FamPrime <>'') and (FamPrime = Copy(stTot,122,3)) Then TypeArt:='PRI'
    Else TypeArt:='FRA';
    champ := debutfrais;
    End;

TOM_ARTICLE(TOM_ARTICLEVAR).SetTypeArticle( TypeArt );

TOM_ARTICLEVAR.InitTob(T_ARTICLE);

T_Article.PutValue ('GA_TYPEARTICLE',Typeart);
champ := champ + copy (stTot,7,15);
// GM le 14/09/2000
Valeur := FindEtReplace (champ,'"',' ',True);
T_ARTICLE.PutValue ('GA_ARTICLE',CodeArticleUnique2(valeur,''));
///////////////////
valeur :='';
Valeur := FindEtReplace (Copy(champ,1,17),'"',' ',True);
T_Article.PutValue ('GA_CODEARTICLE',Trim(valeur));
Valeur := FindEtReplace (Copy(StTot,22,70),'"',' ',True);
T_Article.PutValue ('GA_LIBELLE',valeur);
// PA le 19/07/2000 pour contourner pb de "" dans les noms d'articles voir si possiblité de l'inclure dans le putfixedstvalue
// T_ARTICLE.PutFixedStValue ('GA_CODEARTICLE',champ,1,17,tcTrimR,True);
// T_ARTICLE.PutFixedStValue ('GA_LIBELLE',stTot,22,70,tcTrim,False);
T_ARTICLE.PutFixedStValue ('GA_COMMENTAIRE',stTot,22,90,tcTrim,False);
T_ARTICLE.PutFixedStValue ('GA_FAMILLENIV1',stTot,122,3,tcTrim,False);
    // zone comptable alimenté par famille presttaion ou nature ou rien
If (FParam.ComboPreCompt.ItemIndex = 1) then
    T_ARTICLE.PutFixedStValue ('GA_COMPTAARTICLE',stTot,122,3,tcChaine,True);
If (FParam.ComboPreCompt.ItemIndex = 2) then
    T_ARTICLE.PutFixedStValue ('GA_COMPTAARTICLE',stTot,125,1,tcChaine,True);
if (sttot[266] <> '0') then begin
   for ii :=1 to nbtva+1 do begin
      if CodeTva[ii]= Sttot[266] then begin
          if (TauxTva[ii])< 18.0 then    T_ARTICLE.PutValue ('GA_FAMILLETAXE1','RED')
          else    T_ARTICLE.PutValue ('GA_FAMILLETAXE1','NOR');
          end;
      end;

    end
    else T_ARTICLE.PutValue ('GA_FAMILLETAXE1','EXO');
T_ARTICLE.PutFixedStValue ('GA_PVHT',stTot,143,11,tcDouble100,False);
T_ARTICLE.PutFixedStValue ('GA_PRHT',stTot,173,11,tcDouble100,False);
T_ARTICLE.PutFixedStValue ('GA_DPR',stTot,173,11,tcDouble100,False);
T_ARTICLE.PutFixedStValue ('GA_PMRP',stTot,173,11,tcDouble100,False);

If (stTot[6] = 'P') then
   Begin
   if (sttot[252]=' ') then sttot[252]:='H';
   T_ARTICLE.PutFixedStValue ('GA_QUALIFUNITEACT',stTot,252,1,tcChaine,True);
   T_ARTICLE.PutFixedStValue ('GA_QUALIFUNITEVTE',stTot,252,1,tcChaine,True);
   End;


If (FParam.CheckStat1Prest.Checked = TRUE) then
  T_ARTICLE.PutFixedStValue ('GA_LIBREART1',stTot,268,4,tcTrim,True);
If (FParam.CheckStat2Prest.Checked = TRUE) then
  Begin
  If (FParam.CheckStat1Prest.Checked = TRUE) then
      T_ARTICLE.PutFixedStValue ('GA_LIBREART2',stTot,272,4,tcTrim,True)
  Else T_ARTICLE.PutFixedStValue ('GA_LIBREART1',stTot,272,4,tcTrim,True);
  End;
if (copy (stTot,291,1) ='N') then T_ARTICLE.PutValue('GA_ACTIVITEREPRISE','N')
else     T_ARTICLE.PutValue('GA_ACTIVITEREPRISE','F');
T_ARTICLE.PutFixedStValue ('GA_DATECREATION',stTot,294,8,tcDate8AMJ,False);
T_ARTICLE.PutFixedStValue ('GA_DATEMODIF',stTot,302,8,tcDate8AMJ,False);
champ :=Copy (Sttot, 310,8);
if (champ = '00000000') then champ:='        ';
If (Trim(Copy (champ,1,8)) ='') Then T_ARTICLE.PutValue ('GA_FERME', '-')
         Else begin
         T_ARTICLE.PutValue ('GA_FERME', 'X');
         T_ARTICLE.PutFixedStValue ('GA_DATESUPPRESSION', champ,1,8, tcdate8amj, false);
         end;

                          //recup des info module achat pour fourniture tempo
If ctxTempo in V_PGI.PGIContexte then
   begin
  if stTot[6] = 'F' then
    Begin
    T_ARTICLE.PutFixedStValue ('GA_QUALIFUNITESTO',stTot,344,2,tcChaine,True);
    T_ARTICLE.PutFixedStValue ('GA_QUALIFUNITEVTE',stTot,332,2,tcChaine,True);
    T_ARTICLE.PutFixedStValue ('GA_PAHT',stTot,357,11,tcDouble100,False);
    // reprise du fournissuer principal - PA le 05/04/2001 Attention l'nereg catalogue doit exister...
    //T_ARTICLE.PutFixedStValue ('GA_FOURNPRINC',stTot,334,10,tcChaine,True);
    End;
  End;
T_ARTICLE.PutFixedStValue ('GA_COEFFG',stTot,381,6,tcDouble,False);

                          // initialisation par défaut
T_ARTICLE.PutValue('GA_TENUESTOCK','-');
T_ARTICLE.PutValue('GA_LOT','-');
T_ARTICLE.PutValue('GA_NUMEROSERIE','-');
T_ARTICLE.PutValue('GA_CONTREMARQUE','-');
T_ARTICLE.PutValue('GA_UTILISATEUR',V_PGI.User);
T_ARTICLE.PutValue('GA_CREATEUR',V_PGI.User);
T_ARTICLE.PutValue('GA_REMISEPIED','X');
T_ARTICLE.PutValue('GA_ESCOMPTABLE','X');
T_ARTICLE.PutValue('GA_COMMISSIONNABLE','X');
T_ARTICLE.PutValue('GA_REMISELIGNE','-');
T_ARTICLE.PutValue('GA_CREERPAR','IMP');
If (TOM_ARTICLEVAR.VerifTOB (T_ARTICLE)) then  ImportTob(T_ARTICLE,True)
Else writeln (FichierLog,T_ARTICLE.GetValue('GA_ARTICLE')+' '+T_ARTICLE.GetValue('GA_LIBELLE')+' Erreur '+inttostr(TOM_ARTICLEVAR.LastError)+' '+TOM_ARTICLEVAR.LastErrorMsg);

T_ARTICLE.Free; TOM_ARTICLEVAR.Free;
End;

//*************** Traitement Ressource ********************
procedure TFAsRecupTempo.RecupRessource (enreg:string);
            // fct qui récupère depuis le fichier de sauvegarde SCOT/TEMPO
            // les enrgt 5 pour creer les RESSOURCE
            // en attente
           // voir pour affecter la zone Tiers (à faire qd règle fichier tiers connu)
           
var TRess : TOB;
    TOM_Ressource : TOM;
    coef : Double;
    champ:string;
    rang:integer;

begin
if (CheckRepriseSup.checked=False) then  begin
    champ :=Copy (Enreg, 134,8);
    if (champ = '00000000') then champ:='        ';
    If (Trim(Copy (champ,1,8)) <>'') Then exit;
   end;
TRess := TOB.CREATE ('RESSOURCE', NIL, -1);      // cree la tob
TOM_Ressource := CreateTom('RESSOURCE',Nil,False, False);
TOM_Ressource.InitTob (TRess);
TRess.PutValue ('ARS_TYPERESSOURCE', 'SAL');
TRess.PutValue ('ARS_RESSOURCE', trim(AnsiUppercase(Copy(enreg, 6,6))));
TRess.PutFixedStValue ('ARS_LIBELLE', enreg, 12,30, tcTrimR, false);
TRess.PutFixedStValue ('ARS_LIBELLE2', enreg, 86,20, tcTrimR, false);
TRess.PutFixedStValue ('ARS_TELEPHONE', enreg, 106,20, tcTrimR, false);

        // recupération table libre
Case Stat1Res of
  1:  TRESS.PutFixedStValue ('ARS_LIBRERES1', enreg, 42,4, tcTrim, false);
  2:  TRESS.PutFixedStValue ('ARS_LIBRERES2', enreg, 42,4, tcTrim, false);
  3:  TRESS.PutFixedStValue ('ARS_LIBRERES3', enreg, 42,4, tcTrim, false);
  4:  TRESS.PutFixedStValue ('ARS_LIBRERES4', enreg, 42, 4,tcTrim, false);
  5:  TRESS.PutFixedStValue ('ARS_ETABLISSEMENT', enreg, 42,3, tcTrim, false);
  6:  TRESS.PutFixedStValue ('ARS_DEPARTEMENT', enreg, 42,4, tcTrim, false);
  7:  TRESS.PutFixedStValue ('ARS_TYPERESSOURCE', enreg, 42,3, tcTrim, false);
  End;
Case Stat2Res of
  1:  TRESS.PutFixedStValue ('ARS_LIBRERES1', enreg, 46,4, tcTrim, false);
  2:  TRESS.PutFixedStValue ('ARS_LIBRERES2', enreg, 46,4, tcTrim, false);
  3:  TRESS.PutFixedStValue ('ARS_LIBRERES3', enreg, 46,4, tcTrim, false);
  4:  TRESS.PutFixedStValue ('ARS_LIBRERES4', enreg, 46,4, tcTrim, false);
  5:  TRESS.PutFixedStValue ('ARS_ETABLISSEMENT', enreg, 46,3, tcTrim, false);
  6:  TRESS.PutFixedStValue ('ARS_DEPARTEMENT', enreg, 46,4, tcTrim, false);
  7:  TRESS.PutFixedStValue ('ARS_TYPERESSOURCE', enreg, 46,3, tcTrim, false);
  End;
Case Stat3Res of
  1:  TRESS.PutFixedStValue ('ARS_LIBRERES1', enreg, 50,4, tcTrim, false);
  2:  TRESS.PutFixedStValue ('ARS_LIBRERES2', enreg, 50,4, tcTrim, false);
  3:  TRESS.PutFixedStValue ('ARS_LIBRERES3', enreg, 50,4, tcTrim, false);
  4:  TRESS.PutFixedStValue ('ARS_LIBRERES4', enreg, 50,4, tcTrim, false);
  5:  TRESS.PutFixedStValue ('ARS_ETABLISSEMENT', enreg, 50,3, tcTrim, false);
  6:  TRESS.PutFixedStValue ('ARS_DEPARTEMENT', enreg, 50,4, tcTrim, false);
  7:  TRESS.PutFixedStValue ('ARS_TYPERESSOURCE', enreg, 50,3, tcTrim, false);
  End;
Case Stat4REs of
  1:  TRESS.PutFixedStValue ('ARS_LIBRERES1', enreg, 54,4, tcTrim, false);
  2:  TRESS.PutFixedStValue ('ARS_LIBRERES2', enreg, 54,4, tcTrim, false);
  3:  TRESS.PutFixedStValue ('ARS_LIBRERES3', enreg, 54,4, tcTrim, false);
  4:  TRESS.PutFixedStValue ('ARS_LIBRERES4', enreg, 54,4, tcTrim, false);
  5:  TRESS.PutFixedStValue ('ARS_ETABLISSEMENT', enreg, 54,3, tcTrim, false);
  6:  TRESS.PutFixedStValue ('ARS_DEPARTEMENT', enreg, 54,4, tcTrim, false);
  7:  TRESS.PutFixedStValue ('ARS_TYPERESSOURCE', enreg, 54,3, tcTrim, false);
  End;
If (Stat4Res <5) and (STat4res >0) then Rang:=STat4Res +1
else  If (Stat3res <4) and (STat3res >0) then Rang:=STat3res +1
else  If (Stat2res <4) and (STat2res >0) then Rang:=STat2res +1
else  If (Stat1res <4) and (STat1res >0) then Rang:=STat1res +1
else rang:=1;
//TRess.PutFixedStValue ('ARS_LIBRERES3', enreg, 62,4, tcTrimR, false);
TRess.PutFixedStValue ('ARS_LIBRERES'+IntToStr(Rang), enreg, 62,4, tcTrimR, false);
TRess.PutFixedStValue ('ARS_CHARLIBRE1', enreg, 66,20, tcTrimR, false);

TRess.PutFixedStValue ('ARS_PVHT', enreg, 155,11, tcdouble100, false);
If ctxTempo in V_PGI.PGIContexte then
TRess.PutFixedStValue ('ARS_TAUXUNIT',enreg,259,11,tcdouble100,false);
TRess.PutFixedStValue ('ARS_TAUXREVIENTUN', enreg, 185,11, tcdouble100, false);
if (copy (enreg,300,5) <>  '    0') then
  begin
  Coef := 1+ (StrToFloat(Copy(enreg,300,5))/10000);
  TRess.PutValue ('ARS_TAUXCHARGEPAT', Coef);
  end ;
if ( copy (enreg,305,6) <> '     0') then TRess.PutFixedStValue ('ARS_TAUXFRAISGEN1', enreg, 305,6, tcdouble1000, false);
// MCD  PA le 16/08/2000 supprimé en attente de lien paie pb de cohérence si la fiche salariée n'existe pas ...
// TRess.PutFixedStValue ('ARS_SALARIE', enreg, 445,6, tcTrimR, false);
    // on ne récupère le compte auxiliaire que si la gestion frais dans la compta était géré
if (FraisCompta = TRUE) then TRess.PutFixedStValue ('ARS_AUXILIAIRE', enreg, 460,10, tcTrimR, false);

if (enreg[441]= 'O') then TRess.PutValue ('ARS_ACTIVITEPAIE','X')
else TRess.PutValue ('ARS_ACTIVITEPAIE','-');
TRess.PutValue ('ARS_UNITETEMPS','H');
TRess.PutValue ('ARS_FERME','-');
if (Trim(copy(enreg,429,4)) <> '') then
  begin
  TRess.PutFixedStValue ('ARS_FONCTION4',enreg,429,4,tcTrimr,false);
  if (Trim(copy(enreg,433,8)) = '') then TRess.PutValue ('ARS_DATEFONC4',  idate1900)
  else TRess.PutFixedStValue ('ARS_DATEFONC4', enreg,433,8, tcdate8amj, false);
  End;
if (Trim(copy(enreg,417,4))<>'') then
   begin
   TRess.PutFixedStValue ('ARS_FONCTION3', enreg,417,4, tcTrimr, false);
  if (Trim(copy(enreg,431,8)) = '') or (Trim(copy(enreg,431,8)) = '00000000')
  then TRess.PutValue ('ARS_DATEFONC3',  idate1900)
  else    TRess.PutFixedStValue ('ARS_DATEFONC3', enreg,421,8, tcdate8amj, false);
   End;
if (Trim(copy(enreg,405,4))<> '') then
  begin
      TRess.PutFixedStValue ('ARS_FONCTION2', enreg,405,4, tcTrimr, false);
    if (Trim(copy(enreg,409,8)) = '') or (Trim(copy(enreg,409,8)) = '00000000')
    then TRess.PutValue ('ARS_DATEFONC2',  idate1900)
    else       TRess.PutFixedStValue ('ARS_DATEFONC2', enreg,409,8, tcdate8amj, false);
  End;
if (Trim(copy(enreg,58,4))<> '') then
  begin
  TRess.PutFixedStValue ('ARS_FONCTION1', enreg,58,4, tcTrimr, false);
  if (Trim(copy(enreg,397,8)) = '') or (Trim(copy(enreg,397,8)) = '00000000')
  then TRess.PutValue ('ARS_DATEFONC1',  idate1900)
  else   TRess.PutFixedStValue ('ARS_DATEFONC1', enreg,397,8, tcdate8amj, false);
  End;

if (Trim(copy(enreg,126,8))<> '') and (Trim(copy(enreg,126,8))<> '00000000') then
    begin
    TRess.PutFixedStValue ('ARS_DEBUTDISPO', enreg, 126,8, tcdate8amj, false);
    TRess.PutFixedStValue ('ARS_DATECREATION', enreg, 126,8, tcdate8amj, false);
    end
else
    begin
    TRess.PutValue ('ARS_DATECREATION',idate1900);
    TRess.PutValue ('ARS_DEBUTDISPO',idate1900);
    end;
if (Trim(copy(enreg,134,8))<> '') and (Trim(copy(enreg,134,8))<> '00000000') then
    begin
    TRess.PutFixedStValue ('ARS_FINDISPO', enreg, 134,8, tcdate8amj, false);
    TRess.PutValue ('ARS_FERME','X');
    end
    else     TRess.PutValue ('ARS_FINDISPO',idate2099);

TRess.PutFixedStValue ('ARS_DATEMODIF', enreg, 389,8, tcdate8amj, false);
TRess.PutValue ('ARS_CREERPAR','IMP');
TRess.PutValue('ARS_CREATEUR',V_PGI.User);
TRess.PutValue ('ARS_UTILISATEUR',V_PGI.user);
TRess.PutValue ('ARS_UTILASSOCIE',V_PGI.user);
If (TOM_Ressource.VerifTOB (TRess)) then  ImportTob(TRess,True)
Else writeln (FichierLog,TRess.GetValue('ARS_RESSOURCE')+' '+TRess.GetValue('ARS_LIBELLE')+' Erreur '+inttostr(TOM_RESSOURCE.LastError)+' '+TOM_RESSOURCE.LastErrorMsg);


TRess.Free; TOM_Ressource.Free;
End;
//*************** Traitement Tablettes ********************
procedure TFAsRecupTempo.RecupTablette (enreg:string);
var TTablette : TOB;
    TypeTable, TypeCode : string;
    Reprise : Boolean;
    LngCode ,nb,Rang: Integer;
    GT_alpha ,champ: string ;
begin
Reprise := False;
GT_Alpha :='ABCDEFGHIJKLMNOPQRSTUVWXYZ';
lngcode:=0;
if ( (CheckListFicBase.Checked[3] = TRUE) or(CheckListFicBase.Checked[6] = TRUE)) then
begin
// création tabllete lien afftiers pour les 2 tables utilisées
if ((copy(enreg,6,2) = '**') And (FParam.CheckStat1Aff.Checked = TRUE)) then
   Begin TypeTable :='CC'; LngCode := 3; TypeCode := 'LAT'; Reprise:= TRUE; End;
// stat affaires
if (copy(enreg,6,2) = 'A1')then
    begin
    if (FParam.CheckStat1Aff.Checked = TRUE) then
        Begin
        TypeTable :='YX'; LngCode := 4;
        Reprise :=True;
        Case Stat1Aff of
          1: TypeCode := 'LF1';
          2: TypeCode := 'LF2';
          3: TypeCode := 'LF3';
          4: TypeCode := 'LF4';
          6: TypeCode := 'ADP';
          else Reprise :=False;
            End;
        End;
    end;
 if (copy(enreg,6,2) = 'A2')then
    begin
    if (FParam.CheckStat2Aff.Checked = TRUE) then
        Begin
        TypeTable :='YX'; LngCode := 4;
        Reprise :=True;
        Case Stat2Aff of
          1: TypeCode := 'LF1';
          2: TypeCode := 'LF2';
          3: TypeCode := 'LF3';
          4: TypeCode := 'LF4';
          6: TypeCode := 'ADP';
          else Reprise :=False;
            End;
        End;
    end;
if (copy(enreg,6,2) = 'A3')then
    begin
    if (FParam.CheckStat3Aff.Checked = TRUE) then
        Begin
        TypeTable :='YX'; LngCode := 4;
        Reprise :=True;
        Case Stat3Aff of
          1: TypeCode := 'LF1';
          2: TypeCode := 'LF2';
          3: TypeCode := 'LF3';
          4: TypeCode := 'LF4';
          6: TypeCode := 'ADP';
          else Reprise :=False;
            End;
        End;
    end;
if (copy(enreg,6,2) = 'A4')then
    begin
    if (FParam.CheckStat4Aff.Checked = TRUE) then
        Begin
        TypeTable :='YX'; LngCode := 4;
        Reprise :=True;
        Case Stat4Aff of
          1: TypeCode := 'LF1';
          2: TypeCode := 'LF2';
          3: TypeCode := 'LF3';
          4: TypeCode := 'LF4';
          6: TypeCode := 'ADP';
          else Reprise :=False;
            End;
        End;
    end;
// pour SCOT, on récupère le libelle de la mission
If (ctxScot in V_PGI.PGIContexte) then
   begin
   if (copy(enreg,6,2) = 'AF') then
     Begin
     TypeTable :='CC';
     LngCode := 3;
     if VH_GC.CleAffaire.Co1Type = 'LIS' then TypeCode := 'AP1'; ;
     if VH_GC.CleAffaire.Co2Type = 'LIS' then TypeCode := 'AP2'; ;
     if VH_GC.CleAffaire.Co3Type = 'LIS' then TypeCode := 'AP3'; ;
     Reprise:= TRUE;
     End;
   end;

// Famille d'affaire
if (copy(enreg,6,2) = 'FA') then
   Begin
   TypeTable :='YX'; LngCode := 3;
   Reprise :=True;
   Case FamAffaire of
        1: TypeCode := 'LF1';
        2: TypeCode := 'LF2';
        3: TypeCode := 'LF3';
        else Reprise :=False;
        end;
        // On regarde si la famille est aussi à écrire en zone comptable
  if (FParam.ComboAffCompt.ItemIndex = 1) then
    Begin
    If (Reprise)  then
       Begin
       If (TypeTable = 'YX') then TTablette := TOB.CREATE ('CHOIXEXT', NIL, -1)
       Else If (TypeTable = 'CC') then TTablette := TOB.CREATE ('CHOIXCOD', NIL, -1)
       Else TTablette := TOB.CREATE ('COMMUN', NIL, -1);  // en principe CO ...
       TTablette.initValeurs;
       TTablette.PutValue (TypeTable+'_TYPE', TypeCode);
       TTablette.PutFixedStValue (TypeTable+'_CODE', enreg, 8,LngCode, tctrim, false);
       TTablette.PutFixedStValue (TypeTable+'_LIBELLE', enreg, 14,35, tctrim, false);
       TTablette.PutFixedStValue (TypeTable+'_ABREGE', enreg, 14,17, tctrim, false);
       ImportTob(TTablette,True);
       TTablette.Free;
       End;
    TypeCode := 'GCF';
    TypeTable :='CC';
    reprise :=True;
    End;
   End;
end;
    // alimentation table TVA pour test affaire et article
if ( (CheckListFicBase.Checked[3] = TRUE) or(CheckListFicBase.Checked[6] = TRUE) or  (CheckListFicBase.Checked[0] = TRUE)) then
begin
// Récupération des taux de TVA
if (copy(enreg,6,2) = 'TV') then
    Begin
    inc (nbTva);
    if (Nbtva > 20) then begin
       PGIInfo(TraduitGA('Plus de 20 taux TVA. Arrêt'),TitreHalley);
       exit;
        end;
    codeTva[NBTva] := Copy(Enreg,8,1);
    TauxTva[NbTva] := StrToFLoat(trim(Copy(Enreg,62,4)))/100.0;
    reprise :=False;
    End;
end;

if  (CheckListFicBase.Checked[0] = TRUE)  then
begin
// stat prestations
if ((copy(enreg,6,2) = 'P1') And (FParam.CheckStat1Prest.Checked = TRUE)) then
   Begin TypeTable :='YX'; LngCode := 4; TypeCode := 'LA1'; Reprise:= TRUE; End;
if ((copy(enreg,6,2) = 'P2') And (FParam.CheckStat1Prest.Checked = TRUE) And(FParam.CheckStat2Prest.Checked = TRUE)) then
   Begin TypeTable :='YX'; LngCode := 4; TypeCode := 'LA2'; Reprise:= TRUE; End;
if ((copy(enreg,6,2) = 'P2') And (FParam.CheckStat1Prest.Checked = FALSE) And(FParam.CheckStat2Prest.Checked = TRUE)) then
   Begin TypeTable :='YX'; LngCode := 4; TypeCode := 'LA1'; Reprise:= TRUE; End;
end;
if  ((CheckListFicBase.Checked[0] = TRUE) or (CheckListFicBase.Checked[1] = TRUE)
    or (CheckListFicBase.Checked[2] = TRUE)) then
begin
// Nature comptable article si traité en zone comptable PGI
if ((copy(enreg,6,2) = 'NA') and (FParam.ComboPreCompt.ItemIndex = 2)) then
   Begin TypeTable :='CC'; LngCode := 1; TypeCode := 'GCA'; Reprise:= TRUE; End;
// Famille de prestation
if (copy(enreg,6,2) = 'PR')  then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'FN1';
   Reprise:=True;
        // si la famille prestation est aussi à mettre dans la zone comptable,
        // il faut la traiter 2 fois.
   if (FParam.ComboPreCompt.ItemIndex = 1) then
     Begin
     If (Reprise)  then
       Begin
       If (TypeTable = 'YX') then TTablette := TOB.CREATE ('CHOIXEXT', NIL, -1)
       Else If (TypeTable = 'CC') then TTablette := TOB.CREATE ('CHOIXCOD', NIL, -1)
       Else TTablette := TOB.CREATE ('COMMUN', NIL, -1); // en principe CO
       TTablette.initValeurs;
       TTablette.PutValue (TypeTable+'_TYPE', TypeCode);
       TTablette.PutFixedStValue (TypeTable+'_CODE', enreg, 8,LngCode, tctrim, false);
       TTablette.PutFixedStValue (TypeTable+'_LIBELLE', enreg, 14,35, tctrim, false);
       TTablette.PutFixedStValue (TypeTable+'_ABREGE', enreg, 14,17, tctrim, false);
       ImportTob(TTablette,True);
       TTablette.Free;
       End;
     TypeCode := 'GCA';
     TypeTable :='CC';
     reprise :=TRUE;
     End;
    End;
end;

if (CheckListFicBase.Checked[4] = TRUE) then
begin
// stat Ressources
if (copy(enreg,6,2) = 'E1')then
    begin
    if (FParam.CheckStat1Emp.Checked = TRUE) then
        Begin
        TypeTable :='YX'; LngCode := 4;
        Reprise :=True;
        Case Stat1Res of
          1: TypeCode := 'LR1';
          2: TypeCode := 'LR2';
          3: TypeCode := 'LR3';
          4: TypeCode := 'LR4';
          6: TypeCode := 'ADP';
          else Reprise :=False;
            End;
        End;
    end;
if (copy(enreg,6,2) = 'E2')then
    begin
    if (FParam.CheckStat2Emp.Checked = TRUE) then
        Begin
        TypeTable :='YX'; LngCode := 4;
        Reprise :=True;
        Case Stat2Res of
          1: TypeCode := 'LR1';
          2: TypeCode := 'LR2';
          3: TypeCode := 'LR3';
          4: TypeCode := 'LR4';
          6: TypeCode := 'ADP';
          else Reprise :=False;
            End;
        End;
    end;
if (copy(enreg,6,2) = 'E3')then
    begin
    if (FParam.CheckStat3Emp.Checked = TRUE) then
        Begin
        TypeTable :='YX'; LngCode := 4;
        Reprise :=True;
        Case Stat3REs of
          1: TypeCode := 'LR1';
          2: TypeCode := 'LR2';
          3: TypeCode := 'LR3';
          4: TypeCode := 'LR4';
          6: TypeCode := 'ADP';
          else Reprise :=False;
            End;
        End;
    end;
if (copy(enreg,6,2) = 'E4')then
    begin
    if (FParam.CheckStat4Emp.Checked = TRUE) then
        Begin
        TypeTable :='YX'; LngCode := 4;
        Reprise :=True;
        Case Stat4Res of
          1: TypeCode := 'LR1';
          2: TypeCode := 'LR2';
          3: TypeCode := 'LR3';
          4: TypeCode := 'LR4';
          6: TypeCode := 'ADP';
          else Reprise :=False;
            End;
        End;
    end;

// Recupération des Qualifs
if (copy(enreg,6,2) = 'QU') then Begin RecupFonction(enreg); Reprise:=False;  End;
//recup table emploi
if (copy(enreg,6,2) = 'EM') then Begin
    If (Stat4res <5) and (STat4res >0) then Rang:=STat4res +1
    else  If (Stat3res <4) and (STat3res >0) then Rang:=STat3res +1
    else  If (Stat2res <4) and (STat2res >0) then Rang:=STat2res +1
    else  If (Stat1res <4) and (STat1res >0) then Rang:=STat1res +1
    else rang:=1;
   TypeTable :='YX';
   LngCode := 4;
   TypeCode := 'LR'+IntTostr(rang);
   Reprise:= TRUE;
    End;
end;

if (CheckListFicBase.Checked[5] = TRUE) then
begin
// Statistiques Clients
// MCD ATTENTION : si secteur d'activité ne récupère que les 3 premiers caractères
// on traite aussi eventuellement la stat 2 fois si aussi en régime comptable
//mcd 21/03/03 ??? fait que plus rien ne marche LngCode := 0;
if (copy(enreg,6,2) = 'S1')then
    begin
    if (FParamCli.CheckStat1Cli.Checked = TRUE) then
        Begin
        TypeTable :='YX'; LngCode := 4;
        Reprise :=True;
        Case Stat1Cli of
          1: TypeCode := 'LT1';
          6: Begin TypeTable :='CC'; TypeCode := 'SCC';LngCode := 3; End;  // secteur d'activité
          else Reprise :=False;
            End;
        End;
    if (FParamCli.ComboFamComptaTiers.ItemIndex = 1) then
       begin
       If (Reprise)  then
         Begin
         If (TypeTable = 'YX') then TTablette := TOB.CREATE ('CHOIXEXT', NIL, -1)
         Else If (TypeTable = 'CC') then TTablette := TOB.CREATE ('CHOIXCOD', NIL, -1)
         Else TTablette := TOB.CREATE ('COMMUN', NIL, -1);  // en principe CO
         TTablette.initValeurs;
         TTablette.PutValue (TypeTable+'_TYPE', TypeCode);
         TTablette.PutFixedStValue (TypeTable+'_CODE', enreg, 8,LngCode, tctrim, false);
         TTablette.PutFixedStValue (TypeTable+'_LIBELLE', enreg, 14,35, tctrim, false);
         TTablette.PutFixedStValue (TypeTable+'_ABREGE', enreg, 14,17, tctrim, false);
         ImportTob(TTablette,True);
         TTablette.Free;
         End;
       TypeTable :='CC';
       LngCode := 3;
       TypeCode := 'GCT';
       reprise:=TRUE;
       End;
   End ;
if (copy(enreg,6,2) = 'S2')then
    begin
    if (FParamCli.CheckStat1Cli.Checked = TRUE) then
        Begin
        TypeTable :='YX'; LngCode := 4;
        Reprise:=TRUE;
          Case Stat2Cli of
          1: TypeCode := 'LT1';
          2: TypeCode := 'LT2';
          6: Begin TypeTable :='CC'; TypeCode := 'SCC';LngCode := 3;End;  // secteur d'activité
          else Reprise :=False;
          End;
        End;
    if (FParamCli.ComboFamComptaTiers.ItemIndex = 2) then
       begin
       If (Reprise)  then
         Begin
         If (TypeTable = 'YX') then TTablette := TOB.CREATE ('CHOIXEXT', NIL, -1)
         Else If (TypeTable = 'CC') then TTablette := TOB.CREATE ('CHOIXCOD', NIL, -1)
         Else  TTablette := TOB.CREATE ('COMMUN', NIL, -1); // en principe CO
         TTablette.initValeurs;
         TTablette.PutValue (TypeTable+'_TYPE', TypeCode);
         TTablette.PutFixedStValue (TypeTable+'_CODE', enreg, 8,LngCode, tctrim, false);
         TTablette.PutFixedStValue (TypeTable+'_LIBELLE', enreg, 14,35, tctrim, false);
         TTablette.PutFixedStValue (TypeTable+'_ABREGE', enreg, 14,17, tctrim, false);
         ImportTob(TTablette,True);
         TTablette.Free;
         End;
       TypeTable :='CC';
       LngCode := 3;
       TypeCode := 'GCT';
      reprise:=TRUE;
     End;
   End ;
if (copy(enreg,6,2) = 'S3')then
    begin
    if (FParamCli.CheckStat1Cli.Checked = TRUE) then
      Begin
      TypeTable :='YX'; LngCode := 4;
      Reprise := True;
        Case Stat3Cli of
          1: TypeCode := 'LT1';
          2: TypeCode := 'LT2';
          3: TypeCode := 'LT3';
          6: Begin TypeTable :='CC'; TypeCode := 'SCC';LngCode := 3;End;  // secteur d'activité
          else Reprise :=False;
          End;
      end;
    if (FParamCli.ComboFamComptaTiers.ItemIndex = 3) then
       begin
       If (Reprise)  then
         Begin
         If (TypeTable = 'YX') then TTablette := TOB.CREATE ('CHOIXEXT', NIL, -1)
         Else If (TypeTable = 'CC') then TTablette := TOB.CREATE ('CHOIXCOD', NIL, -1)
         Else TTablette := TOB.CREATE ('COMMUN', NIL, -1);   // en principe CO
         TTablette.initValeurs;
         TTablette.PutValue (TypeTable+'_TYPE', TypeCode);
         TTablette.PutFixedStValue (TypeTable+'_CODE', enreg, 8,LngCode, tctrim, false);
         TTablette.PutFixedStValue (TypeTable+'_LIBELLE', enreg, 14,35, tctrim, false);
         TTablette.PutFixedStValue (TypeTable+'_ABREGE', enreg, 14,17, tctrim, false);
         ImportTob(TTablette,True);
         TTablette.Free;
         End;
       TypeTable :='CC';
       LngCode := 3;
       TypeCode := 'GCT';
       reprise:=TRUE;
     End;
   End ;
if (copy(enreg,6,2) = 'S4')then
    begin
    if (FParamCli.CheckStat1Cli.Checked = TRUE) then
    begin
        TypeTable :='YX'; LngCode := 4;
        Reprise:=True;
        Case Stat4Cli of
          1: TypeCode := 'LT1';
          2: TypeCode := 'LT2';
          3: TypeCode := 'LT3';
          4: TypeCode := 'LT4';
          6: Begin TypeTable :='CC'; TypeCode := 'SCC';LngCode := 3; End;  // secteur d'activité
          else Reprise :=False;
          End;
    End;
    if (FParamCli.ComboFamComptaTiers.ItemIndex = 4) then
       begin
       If (Reprise)  then
         Begin
         If (TypeTable = 'YX') then TTablette := TOB.CREATE ('CHOIXEXT', NIL, -1)
         Else If (TypeTable = 'CC') then TTablette := TOB.CREATE ('CHOIXCOD', NIL, -1)
         Else TTablette := TOB.CREATE ('COMMUN', NIL, -1); // en principe CO
         TTablette.initValeurs;
         TTablette.PutValue (TypeTable+'_TYPE', TypeCode);
         TTablette.PutFixedStValue (TypeTable+'_CODE', enreg, 8,LngCode, tctrim, false);
         TTablette.PutFixedStValue (TypeTable+'_LIBELLE', enreg, 14,35, tctrim, false);
         TTablette.PutFixedStValue (TypeTable+'_ABREGE', enreg, 14,17, tctrim, false);
         ImportTob(TTablette,True);
         TTablette.Free;
         End;
       TypeTable :='CC';
       LngCode := 3;
       TypeCode := 'GCT';
       reprise:=TRUE;
      End;
   End;
// traitement de l'origine
if (copy(enreg,6,2) = 'OR')
   then Begin
   TypeTable :='CC';
   LngCode := 1;
   TypeCode := 'GOR';
   Reprise:= TRUE;
   if (TypeCode ='') then Reprise:=FALSE;
   End;
// traitement de modes de règlements
if (copy(enreg,6,2) = 'MR')
   then Begin
   RecupModePaie (enreg);
   Reprise:=FALSE;
   End;
// traitement des grilles techniques
if (copy(enreg,6,2) = 'GR')
   then Begin
   TypeTable :='YX';
   LngCode := 3;
   nb:=1;
   While (enreg[8] <> GT_alpha[nb]) And (nb <= 26) do
       begin
        inc(nb);
        end;
   if (nb <=26) then begin TypeCode := RecupTypeTable (nb);  Reprise:= TRUE; End;
   if (TypeCode ='') then Reprise:=FALSE;
   End ;
end;

if (ModulePlanning = TRUE) then
begin
// Recupération des jours fériés
if (copy(enreg,6,2) = 'JO') then Begin RecupJourFerie(enreg); Reprise:=False; End;
end;

// traitement commentaire ligne
if (copy(enreg,6,2) = 'CO')
   then Begin
   inc(NbComm);
   TypeTable :='YX';
   LngCode := 3;
   TypeCode := 'GCI';
   Reprise:= TRUE;
   champ :=Format ('%3.3d', [NbComm]);
   End ;

// MCD Code TVA Tempo = Famille de taxe article
// A voir quelle tablette ???

If (Reprise)  then
   Begin
   If (TypeTable = 'YX') then TTablette := TOB.CREATE ('CHOIXEXT', NIL, -1)
   Else If (TypeTable = 'CC') then TTablette := TOB.CREATE ('CHOIXCOD', NIL, -1)
   Else  TTablette := TOB.CREATE ('COMMUN', NIL, -1);   // CO
   TTablette.initValeurs;
   TTablette.PutValue (TypeTable+'_TYPE', TypeCode);
   TTablette.PutFixedStValue (TypeTable+'_CODE', enreg, 8,LngCode, tctrim, false);
   if Typecode='GCI' then TTablette.PutFixedStValue (TypeTable+'_CODE', champ, 1,LngCode, tctrim, false);
   TTablette.PutFixedStValue (TypeTable+'_ABREGE', enreg, 14,17, tctrim, false);
   TTablette.PutFixedStValue (TypeTable+'_LIBELLE', enreg, 14,35, tctrim, false);
   ImportTob(TTablette,True);
   TTablette.Free;
   End;

End;

procedure TFAsRecupTempo.RecupFonction (StTot:string);
var T_FONCTION : TOB;
    taux: double;
Begin
T_FONCTION := TOB.CREATE ('FONCTION', NIL, -1);
T_FONCTION.InitValeurs;

T_FONCTION.PutFixedStValue ('AFO_FONCTION', StTot, 8,4, tcTrim, false);
T_FONCTION.PutFixedStValue ('AFO_LIBELLE', StTot, 14,35, tcTrim, false);
Taux := StrToFloat(copy (StTot,52,9))/100;
T_FONCTION.PutValue ('AFO_TAUXREVIENTUN',taux);
Taux:= StrToFloat(copy (StTot,61,9))/100;
T_FONCTION.PutValue ('AFO_PVHT',taux);
T_FONCTION.PutValue ('AFO_UNITETEMPS','H');


ImportTob(T_FONCTION,True);
T_FONCTION.Free;
End;


    // fct de création de la table valorisation Assistant/prestation
procedure TFAsRecupTempo.RecupValor (StTot:string);
var T_RESSOURCEPR : TOB;
     QQ : TQuery ;
     champ, champ1,Article : string;
     Imax : integer;
     mtt: double;
Begin
     //on traite valo en PR
T_RESSOURCEPR := TOB.CREATE ('RESSOURCEPR', NIL, -1);
T_RESSOURCEPR.InitValeurs;

(* mcd 31/10/02 ??? imax jalais utiliser ... mis en commentaireQQ:=OpenSQL('SELECT MAX(ARP_RANG) FROM RESSOURCEPR',TRUE) ;
if Not QQ.EOF then imax:=QQ.Fields[0].AsInteger+1 else iMax:=1;
Ferme(QQ) ;  *)

T_RESSOURCEPR.PutValue ('ARP_TYPEVALO', 'R');
T_RESSOURCEPR.PutValue ('ARP_RESSOURCE', trim(AnsiUppercase(Copy(Sttot,11,6))));
T_RESSOURCEPR.PutFixedStValue ('ARP_FONCTION', StTot, 7,4, tcTrim, false);
T_RESSOURCEPR.PutFixedStValue ('ARP_FAMILLENIV1', StTot, 17,3, tcTrim, false);
if ( trim(copy (stTot,20,15)) <> '') then
    begin
    T_RESSOURCEPR.PutValue('ARP_TYPEARTICLE','PRE');
    champ := debutpre;
    champ := champ + copy (stTot,20,15);
    end
else champ :='';

if (trim(champ) ='') then article:=''
else Article:=   CodeArticleUnique2(trim(champ),'');
T_RESSOURCEPR.PutValue ('ARP_ARTICLE',  Article);
      // mcd 25/05/01 on ne peut pas faire le test, les enrt article et ress sont après en fichier...
{champ :=TraduitGA(format('Enrgt Valorisation prix non traité : Ressource:%4.4s, Article:%15.15s, ',[Copy(Sttot,11,6),Article]));
if (TestExistance(champ,Article,'',T_RESSOURCEPR.GetValue ('ARP_RESSOURCE'),
   '','', '', '','',Part0,Part1, Part2, Part3,aff ,  mois)=False)
     then begin
     T_RESSOURCEPR.Free;
     exit;
     end; }
    //On calcul la formule pour le PR.
if (valpr = 'P') then champ := '[TAUXUNITPRES]'
else champ := '[TAUXREVIENT]' ;
mtt:= StrToFloat (Copy (Sttot, 36,6));
if (Sttot[35] = 'M') then Imax := 100 else Imax :=1000;
mtt := mtt/Imax;    // pour mettre le bon nombre de décimales
//if (Sttot[35] = 'P') then  mtt := (mtt+100.0)/100.0;  // cas %
champ1:= format ('%f',[mtt]);
if (Sttot[35] = 'M') then  champ := champ1
    else  if (Sttot[35] = 'P') then   champ := champ + '+ ((' + champ +'*'+champ1 +')/100)'
        else if (Sttot[35] = 'C') then champ := champ +'*(' + champ1 + ')';
T_RESSOURCEPR.PutValue ('ARP_FORMULE', champ);
If (Sttot[48] <>'V') then
   begin
  QQ:=OpenSQL('SELECT MAX(ARP_RANG) FROM RESSOURCEPR',TRUE,-1,'',true) ;
  if Not QQ.EOF then imax:=QQ.Fields[0].AsInteger+1 else iMax:=1;
  Ferme(QQ) ;
  T_RESSOURCEPR.PutValue ('ARP_RANG', Imax);
   ImportTob(T_RESSOURCEPR,True);
   end;
   // on traite valo en PV
If (StTot[48] = 'R') then begin
    T_RESSOURCEPR.Free;
    exit;
    end;
T_RESSOURCEPR.PutValue ('ARP_TYPEVALO', 'V');
QQ:=OpenSQL('SELECT MAX(ARP_RANG) FROM RESSOURCEPR',TRUE,-1,'',true) ;
if Not QQ.EOF then imax:=QQ.Fields[0].AsInteger+1 else iMax:=1;
Ferme(QQ) ;
T_RESSOURCEPR.PutValue ('ARP_RANG', Imax);
    //On calcul la formule pour le PV.
if (valpr = 'P') then champ := '[PVHT]'
else champ := '[TAUXVENTE]' ;
mtt:= StrToFloat (Copy (Sttot, 42,6));
if (Sttot[35] = 'M') then Imax := 100 else Imax :=1000;
mtt := mtt/Imax;    // pour mettre le bon nombre de décimales
//if (Sttot[35] = 'P') then  mtt := (mtt+100.0)/100.0;  // cas %
champ1:= format ('%f',[mtt]);
if (Sttot[35] = 'M') then  champ := champ1
    else  if (Sttot[35] = 'P') then   champ := champ + '+ ((' + champ +'*'+champ1 +')/100)'
        else if (Sttot[35] = 'C') then champ := champ +'*(' + champ1 + ')';
T_RESSOURCEPR.PutValue ('ARP_FORMULE', champ);
ImportTob(T_RESSOURCEPR,True);
T_RESSOURCEPR.Free;
end;

procedure TFAsRecupTempo.RecupNotes (StTot:string; Nbligne : integer);
var T_NOTES : TOB;
    Table, prefixe, St : String;
Begin
if ((CheckListFicBase.Checked[3] = FALSE) and  (CheckListFicBase.Checked[6] = FALSE)
   and ((StTot[6] ='F') or (StTot[6] ='f'))) then exit;
if ((CheckListFicBase.Checked[5] = FALSE) and(StTot[6] ='C'))then exit;
if ((CheckListFicBase.Checked[4] = FALSE) and(StTot[6] ='E'))then exit;
if ((CheckListFicBase.Checked[0] = FALSE) and(CheckListFicBase.Checked[1] = FALSE) and(CheckListFicBase.Checked[2] = FALSE) and
 ((StTot[6] ='1')Or(StTot[6] ='2')Or(StTot[6] ='3')))then exit;


prefixe := '';
If (StTot[6] ='E') then Begin table:='RESSOURCE'; prefixe := 'ARS'; End Else
// le champ n'existe plus If ((StTot[6] ='F') or (StTot[6] ='f')) then Begin table:='AFFAIRE'; prefixe := 'AFF'; End Else
If (StTot[6] ='C') then Begin table:='TIERS'; prefixe := 'T'; End Else
If ((StTot[6] ='1')Or(StTot[6] ='2')Or(StTot[6] ='3')) then Begin table:='ARTICLE'; prefixe := 'GA'; End;


if (prefixe <>'') then
  Begin
  T_NOTES := TOB.CREATE (Table, NIL, -1);
  If (T_NOTES.SelectDB('"'+Trim(Copy(StTot,7,15))+'"',Nil,False)) Then
    Begin
    St := T_Notes.GetValue(prefixe+'_BLOCNOTE');
    If (St <> '') Then T_Notes.PutValue(prefixe+'_BLOCNOTE',St+#13#10 +Copy(StTot,24,60*NbLigne))
    Else T_Notes.PutValue(prefixe+'_BLOCNOTE',Copy(StTot,24,60*NbLigne));
    ImportTob(T_NOTES,True);
    End;
  T_NOTES.Free;
  End;
End;

procedure TFAsRecupTempo.RecupJourFerie (StTot:string);
var T_JOURFERIE : TOB;
code : string;
Begin
// récupération des jours fériés.
// la table est renumérotée. Toujours mis à jour fixe. A modifier après
// si ce n'est pas le cas.
inc (nbjour);
code := Format ('%3.3d', [nbjour]);
T_JOURFERIE := TOB.CREATE ('JOURFERIE', NIL, -1);
T_JOURFERIE.InitValeurs;
T_JOURFERIE.PutValue ('AJF_CODEFERIE', code);
T_JOURFERIE.PutFixedStValue ('AJF_LIBELLE', StTot, 14,35, tcTrim, false);
T_JOURFERIE.PutValue ('AJF_JOURFIXE','X');
T_JOURFERIE.PutFixedStValue ('AJF_JOUR', StTot, 10,2, tcEntier, false);
T_JOURFERIE.PutFixedStValue ('AJF_MOIS', StTot, 8,2, tcEntier, false);

ImportTob(T_JOURFERIE,True);
T_JOURFERIE.Free;
End;

//********************Traitement Affaire ****************
procedure TFAsRecupTempo.RecupAffaire (StTot:string);
var T_AFFAIRE : TOB;
    TOM_AFF : TOM;
    Part0, Part1, Part2, Part3, Aff, date ,mois, Entite,statut, Periodicite,datdeb, datfin: String;
    Rang, moisint , interval,nbmois : Integer;
    QQ :TQuery;
    CleDocAffaire : R_CLEDOC;
    champ :String;
    MttEch : Double;
    MonPres,SaisieContre ,Existe: Boolean;
    DebutExer, FinExe, DebutInter, FInInter, DebutFact, FinFact, Liquid,Garantie,Cloture: TDateTime;
Begin
champ :=Copy (Sttot, 460,8);
if (champ = '00000000') then champ:='        ';
if (CheckRepriseSup.checked=False) then  begin
    If (Trim(Copy (champ,1,8)) <>'') Then exit;
   end;
Rang :=1;
    // si affaire provisoire, on ne reprend pas
if sttot[1543]='P' then exit;
{$IFDEF OMI}
DSup := Str8ToDate(copy(sttot,460,8),false);
if (DSup < StrToDate('01/01/2001')) and (Dsup <> 0) then Exit; // Attention Spécif OMI ...
{$ENDIF}
T_AFFAIRE := TOB.CREATE ('AFFAIRE', NIL, -1);
TOM_AFF := CreateTom('AFFAIRE',Nil,False, False);
TOM_AFF.InitTob (T_AFFAIRE);

statut :='AFF';
if sttot[6]='D' then begin T_AFFAIRE.PutValue ('AFF_STATUTAFFAIRE','PRO'); statut:='PRO'; end
else  T_AFFAIRE.PutValue ('AFF_STATUTAFFAIRE','AFF');
if sttot[600]= 'R' then T_AFFAIRE.PutValue ('AFF_ETATAFFAIRE','REF')
else if sttot[600]= 'A' then T_AFFAIRE.PutValue ('AFF_ETATAFFAIRE','ACC')
else T_AFFAIRE.PutValue ('AFF_ETATAFFAIRE','ENC') ;
If (Trim(Copy (champ,1,8)) <>'')and (sttot[6]<>'D') Then T_AFFAIRE.PutValue ('AFF_ETATAFFAIRE','CLO');
T_AFFAIRE.PutValue ('AFF_TIERS', AffectTiers (Copy (StTot, 8,10)));
        // mcd 22/05/01,si pas destruction, on recherche le code affaire
aff:='';
Existe:=False;
If (CheckSupBase.Checked = False)  then begin
   Aff := AffectAff(Copy (StTot, 18,6),Copy (StTot , 24,5),
            AffectTiers(Copy (StTot, 8,10 )), statut,Copy(Sttot,1539,2),Part0, Part1, Part2, Part3, FAlse,MoisInt);
   end ;
if (aff ='') then Aff := AffectAff(Copy (StTot, 18,6),Copy (StTot , 24,5),
            AffectTiers(Copy (StTot, 8,10 )), statut,Copy(Sttot,1539,2),Part0, Part1, Part2, Part3, TRUE,MoisInt)
 else  Existe:=True;
if (aff ='') then begin
   T_AFFAIRE.Free; TOM_AFF.Free;
   exit ;
   end;
if (exer = 'O') and ((trim(Copy(Sttot,24,5)) = '') or (Copy(Sttot,24,2) = '**') )then
    T_AFFAIRE.PutValue ('AFF_BOOLLIBRE1','X');  // pour cas pas exercice ou **/** dans gestion exercice
T_AFFAIRE.PutValue ('AFF_AFFAIRE',Aff);
T_AFFAIRE.PutValue ('AFF_AFFAIRE0', Part0);
T_AFFAIRE.PutValue ('AFF_AFFAIRE1', Part1);
T_AFFAIRE.PutValue ('AFF_AFFAIRE2', Part2);
T_AFFAIRE.PutValue ('AFF_AFFAIRE3', Part3);
champ :=Copy(Sttot,1539,2);
if champ ='  ' then champ :='00';
T_AFFAIRE.PutValue ('AFF_AVENANT', champ);
T_AFFAIRE.PutValue ('AFF_ADMINISTRATIF', '-');
T_AFFAIRE.PutValue ('AFF_AFFAIREHT', 'X');
T_AFFAIRE.PutValue ('AFF_MODELE', '-');
T_AFFAIRE.PutValue ('AFF_AFFCOMPLETE', 'X');

if (EuroDefaut = 'X') then MonPres:= True else MonPres := False;
if (Monpres <> VH^.TenueEuro) then
   BEGIN T_AFFAIRE.PutValue ('AFF_SAISIECONTRE', 'X'); SaisieContre := True; END
else
    BEGIN T_AFFAIRE.PutValue ('AFF_SAISIECONTRE', '-') ;  SaisieContre := False; END;
if (Noentite = 1) then Entite := Sttot[29];
if (Noentite = 2) then Entite := Sttot[33];
if (Noentite = 3) then Entite := Sttot[37];
if (Noentite = 4) then Entite := Sttot[41];
T_AFFAIRE.PutValue ('AFF_ETABLISSEMENT', EntiteVersEtbs(Entite)) ;
    // il faut creer la piece correspodnante à l'affaire
If Existe=False then CreerPieceAffaire(AffectTiers(Copy (StTot, 8,10 )), aff, Statut ,T_AFFAIRE.getValue ('AFF_ETABLISSEMENT'), CleDocAffaire , True,SaisieContre)
   else SelectPieceAffaire (Aff,Statut,CleDocAffaire) ;
RecupPieceAffaire (StTot, CleDocAffaire,MonPres);
if (Sttot[668]='O') then recupfrais(CleDocaffaire,Sttot,MonPres);
    // si factureglobale = O dans les paramètres factures,
    // on met d'office le code regroupement à RE1. Modifier les mission pour lesquelles
    // ce n'est pas vrai.
if (FactureGlobale = TRUE) then T_AFFAIRE.PutValue ('AFF_REGROUPEFACT', 'RE1');
 if Copy (StTot, 668,1 )= 'O'  then T_AFFAIRE.PutValue ('AFF_CHANCELLERIE', 'X')
 else   T_AFFAIRE.PutValue ('AFF_CHANCELLERIE', '-');
        // recupération table libre
Case Stat1Aff of
  1:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF1', StTot, 29,4, tcTrim, false);
  2:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF2', StTot, 29,4, tcTrim, false);
  3:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF3', StTot, 29,4, tcTrim, false);
  4:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF4', StTot, 29,4, tcTrim, false);
  5:  T_AFFAIRE.PutFixedStValue ('AFF_ETABLISSEMENT', StTot, 29,3, tcTrim, false);
  6:  T_AFFAIRE.PutFixedStValue ('AFF_DEPARTEMENT', StTot, 29,4, tcTrim, false);
  7:  T_AFFAIRE.PutFixedStValue ('AFF_RECONDUCTION', StTot, 29,3, tcTrim, false);
  End;
Case Stat2Aff of
  1:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF1', StTot, 33,4, tcTrim, false);
  2:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF2', StTot, 33,4, tcTrim, false);
  3:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF3', StTot, 33,4, tcTrim, false);
  4:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF4', StTot, 33,4, tcTrim, false);
  5:  T_AFFAIRE.PutFixedStValue ('AFF_ETABLISSEMENT', StTot, 33,3, tcTrim, false);
  6:  T_AFFAIRE.PutFixedStValue ('AFF_DEPARTEMENT', StTot, 33,4, tcTrim, false);
  7:  T_AFFAIRE.PutFixedStValue ('AFF_RECONDUCTION', StTot, 33,3, tcTrim, false);
  End;
Case Stat3Aff of
  1:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF1', StTot, 37,4, tcTrim, false);
  2:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF2', StTot, 37,4, tcTrim, false);
  3:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF3', StTot, 37,4, tcTrim, false);
  4:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF4', StTot, 37,4, tcTrim, false);
  5:  T_AFFAIRE.PutFixedStValue ('AFF_ETABLISSEMENT', StTot, 37,3, tcTrim, false);
  6:  T_AFFAIRE.PutFixedStValue ('AFF_DEPARTEMENT', StTot, 37,4, tcTrim, false);
  7:  T_AFFAIRE.PutFixedStValue ('AFF_RECONDUCTION', StTot, 37,3, tcTrim, false);
  End;
Case Stat4Aff of
  1:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF1', StTot, 41,4, tcTrim, false);
  2:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF2', StTot, 41,4, tcTrim, false);
  3:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF3', StTot, 41,4, tcTrim, false);
  4:  T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF4', StTot, 41,4, tcTrim, false);
  5:  T_AFFAIRE.PutFixedStValue ('AFF_ETABLISSEMENT', StTot, 41,3, tcTrim, false);
  6:  T_AFFAIRE.PutFixedStValue ('AFF_DEPARTEMENT', StTot, 41,4, tcTrim, false);
  7:  T_AFFAIRE.PutFixedStValue ('AFF_RECONDUCTION', StTot, 41,3, tcTrim, false);
  End;

   // zone comptable alimenté par famille affaire ou non renseignée
If (FParam.ComboAffCompt.ItemIndex = 1) then
   T_AFFAIRE.PutFixedStValue ('AFF_COMPTAAFFAIRE', StTot, 45,3, tcTrim, false) Else
If (FParam.RadioFamAffStat.Checked = True) then
   T_AFFAIRE.PutFixedStValue ('AFF_LIBREAFF'+IntToStr(Rang), StTot, 45,3, tcTrim, false);

T_AFFAIRE.PutFixedStValue ('AFF_LIBELLE', StTot, 48,70, tcTrim, false);
T_AFFAIRE.PutFixedStValue ('AFF_DESCRIPTIF', StTot, 48,90, tcTrim, false);
if (trim(copy(StTot, 400,8)) <> '') and (trim(copy(StTot, 400,8)) <> '00000000')
  then T_AFFAIRE.PutFixedStValue ('AFF_DATESIGNE', StTot, 400,8, tcdate8amj, false)
else    T_AFFAIRE.PutValue ('AFF_DATESIGNE',idate1900);
T_AFFAIRE.PutValue ('AFF_DEVISE',V_PGI.DevisePivot);

If (ctxScot in V_PGI.PGIContexte) then
    begin
   if (trim(copy(StTot, 384,8)) <> '') and (trim(copy(StTot, 384,8)) <> '00000000')
    then T_AFFAIRE.PutFixedStValue ('AFF_DATELETTREMIS', StTot, 384,8, tcdate8amj, false)
    else    T_AFFAIRE.PutValue ('AFF_DATELETTREMIS',idate1900);
    if (Sttot[601]='O') then T_AFFAIRE.Putvalue ('AFF_LETTREMISSION','X')
    else T_AFFAIRE.Putvalue ('AFF_LETTREMISSION','-');
   end;

    // on traite date début et fin mission, debut et fin exercice, debut et fin facture

Date := trim(copy (StTot, 392,8));
if (Date = '00000000') then date :='';
   //mcd 02/11/00 on recalcule systématiquemnt les dates exerccie, même si
   // date existe dans SCOT (qui sont les dates intervention...
       //  on lit client pour mois cloture
if (ctxScot in V_PGI.PGIContexte) then
  begin
    if (VH_GC.AFFormatExer<> 'AM') then
      begin
      QQ := OpenSQL ('SELECT T_MOISCLOTURE FROM TIERS WHERE T_TIERS ="'+ AffectTiers (Copy (StTot, 8,10))+'"',TRUE,-1,'',true);
      if Not QQ.EOF then
         Begin
         Mois := QQ.Fields[0].AsString;
          Moisint:= StrtoInt (Mois);
          if (MoisInt < 1) or (Moisint >12) then MoisInt:=12;
         end
       else MoisInt :=12;
      Ferme (QQ);
      end;
   CalculDateExercice (VH_GC.AFFormatExer, MoisInt, Part2, DebutExer,FinExe, DebutInter, FinInter, DebutFact, FinFact, Liquid,Garantie,Cloture,True);
   T_AFFAIRE.PutValue ('AFF_DATEDEBEXER', debutexer);
   T_AFFAIRE.PutValue ('AFF_DATEFINEXER', FinExe );
   end;
   

if ( Date <> '') then
  begin
  T_AFFAIRE.PutFixedStValue ('AFF_DATEDEBUT', StTot, 392,8, tcdate8amj, false);
  T_AFFAIRE.PutFixedStValue ('AFF_DATEDEBGENER', StTot, 392,8, tcdate8amj, false);
  datdeb := Copy (Sttot, 392,8);
  end
else begin
    if (Exer ='N') then
       begin
        if (ctxScot in V_PGI.PGIContexte) then
          begin
              // cas  SCOT : L'exercice est devenu OBLIGATOIRE.
              // On prend celui calculé précédemment.
              // MCD A changer si on rend exercice Facultatif dans SCOT (faire idem TEMPO ...)
          T_AFFAIRE.PutValue ('AFF_DATEDEBUT', DebutInter);
          T_AFFAIRE.PutValue ('AFF_DATEDEBGENER', DebutFact);
          Datdeb := Copy (DateTostr (DebutFact),7,4) +Copy (DateTostr (DebutFact),4,2)+ Copy (DateTostr (DebutFact),1,2);
          end
          else begin
              // cas  TEMPO.
              // On met date 1900  et date debut pour calcul échéance à exercice paramètre
         T_AFFAIRE.PutValue ('AFF_DATEDEBUT',idate1900);
         T_AFFAIRE.PutValue ('AFF_DATEDEBGENER',idate1900);
         T_AFFAIRE.PutValue ('AFF_DATEDEBEXER',idate1900);
         datdeb :=DebExer;
         end;
       end
    else begin
        // cas exercice on prend les dates calculées à partir exercice
      T_AFFAIRE.PutValue ('AFF_DATEDEBUT', DebutInter);
      T_AFFAIRE.PutValue ('AFF_DATEDEBGENER', DebutFact);
      Datdeb := Copy (DateTostr (DebutFact),7,4) +Copy (DateTostr (DebutFact),4,2)+ Copy (DateTostr (DebutFact),1,2);
      end;                              
    end;

    // affecte date de fin
if (trim(copy (StTot, 408,8)) <> '') and (trim(copy (StTot, 408,8)) <> '00000000') then
    begin
    T_AFFAIRE.PutFixedStValue ('AFF_DATEFIN',StTot, 408,8, tcdate8amj, false);
    T_AFFAIRE.PutFixedStValue ('AFF_DATEFINGENER',StTot, 408,8, tcdate8amj, false);
    T_AFFAIRE.PutFixedStValue ('AFF_DATEFACTLIQUID',StTot, 408,8, tcdate8amj, false);
    datfin :=COpy (Sttot, 408,8);
    FinExe := StrTodAte (COpy (Sttot, 414,2) + '/'+  COpy (Sttot, 412,2) + '/' +  COpy (Sttot, 408,4));
    end
else    begin
        // mcd 22/02/02 If (ctxTempo in V_PGI.PGIContexte) or (Exer ='N') then
  If (ctxTempo in V_PGI.PGIContexte) then
    begin
        // cas TEMPO ou pas exercice, on met dans la date 2099 + fin paramètre pour calcul dates échéances
    T_AFFAIRE.PutValue ('AFF_DATEFIN',idate2099);
    T_AFFAIRE.PutValue ('AFF_DATEFINGENER',idate2099);
    T_AFFAIRE.PutValue ('AFF_DATEFINEXER',idate2099);
    T_AFFAIRE.PutValue ('AFF_DATEFACTLIQUID',idate2099);
    datfin :=FinExer;
     FinExe := StrTodAte ('31/12/2099');
   end
  else begin
        // cas exercice gérée, on prend les dates calculées à partir exercice
    T_AFFAIRE.PutValue ('AFF_DATEFIN', FinInter );
    T_AFFAIRE.PutValue ('AFF_DATEFINGENER', FinFact);
    T_AFFAIRE.PutValue ('AFF_DATEFACTLIQUID', Liquid);
    DatFin := Copy (DateTostr (FinFact),7,4) +Copy (DateTostr (FinFact),4,2)+ Copy (DateTostr (FinFact),1,2);
    end;
  end;


   // mcd 29/12/00 si date existe de 1ere facturaiton, on prend telle quelle
If (trim(Copy(StTot,1516,8)) <> '')Then   T_AFFAIRE.PutFixedStValue ('AFF_DATEDEBGENER',StTot, 1516,8, tcdate8amj, false);

// type de génération automatique
T_AFFAIRE.PutValue('AFF_GENERAUTO','MAN');
if (ctxScot in V_PGI.PGIContexte) then begin
  If (StTot[468] = 'F') then T_AFFAIRE.PutValue('AFF_GENERAUTO','FOR');
  If (StTot[468] = 'T') then T_AFFAIRE.PutValue('AFF_GENERAUTO','ACT');
  If (StTot[468] = 'P') then T_AFFAIRE.PutValue('AFF_GENERAUTO','POU');
  If (StTot[468] = 'A') then T_AFFAIRE.PutValue('AFF_GENERAUTO','MAN');
  If (StTot[468] = 'N') then T_AFFAIRE.PutValue('AFF_GENERAUTO','AVA');
  end
else begin
  If (StTot[468] = 'F') then T_AFFAIRE.PutValue('AFF_GENERAUTO','FOR');
  If (StTot[468] = 'T') then T_AFFAIRE.PutValue('AFF_GENERAUTO','MAN');
  If (StTot[468] = 'P') then T_AFFAIRE.PutValue('AFF_GENERAUTO','ACT');
  If (StTot[468] = 'O') then T_AFFAIRE.PutValue('AFF_GENERAUTO','POU');
  If (StTot[468] = 'N') then T_AFFAIRE.PutValue('AFF_GENERAUTO','AVA');
  end;
    // Fonction de calcul des échéances + création FACTAFF ...
RecupFactAff(StTot, Aff, datdeb, datfin,T_AFFAIRE.GetValue('AFF_GENERAUTO'), Periodicite, interval, Monpres);
T_AFFAIRE.PutValue ('AFF_PERIODICITE',Periodicite);
T_AFFAIRE.PutValue ('AFF_INTERVALGENER',Interval);
     //calculer la date fin garantie = date fin + X mois
if (trim(copy(StTot, 416,8)) <> '') and (trim(copy(StTot, 416,8)) <> '00000000')
 then T_AFFAIRE.PutFixedStValue ('AFF_DATEGARANTIE', StTot, 416,8, tcdate8amj, false)
else    begin
    If (ctxTempo in V_PGI.PGIContexte) then T_AFFAIRE.PutValue ('AFF_DATEGARANTIE',idate2099)
    else begin
      NbMois:= ValeurI(GetParamSocSecur('SO_AFCALCULGARANTI', 0));
      if NbMois <> 0 then begin
                  FinInter := PlusDate(FinExe,NbMois,'M') ;
                  T_AFFAIRE.PutValue ('AFF_DATEGARANTIE',FinInter);
                  end
          else    T_AFFAIRE.PutValue ('AFF_DATEGARANTIE',idate2099);
        end;

    end;

T_AFFAIRE.PutFixedStValue ('AFF_DATECREATION', StTot, 444,8, tcdate8amj, false);
T_AFFAIRE.PutFixedStValue ('AFF_DATEMODIF', StTot, 452,8, tcdate8amj, false);
  // si activité, on force activité reprise à tous, quelque soit l'option des paramètres
T_AFFAIRE.PutValue('AFF_PROFILGENER','DEF'); //mcd 15/05/01
T_AFFAIRE.PutValue('AFF_REPRISEACTIV',VH_GC.AFRepriseActiv);
if Vh_GC.AfrepriseActiv='' then   T_AFFAIRE.PutValue('AFF_REPRISEACTIV','NON');
if T_AFFAIRE.GetValue('AFF_GENERAUTO')= 'ACT' then T_AFFAIRE.PutValue('AFF_REPRISEACTIV','TOU');
MttEch := 0;
if (ctxScot in V_PGI.PGIContexte) then
  If (StTot[468] <>'P')Then MttEch := StrtoFloat (Copy (Sttot, 644,12))/100.0
  else T_AFFAIRE.PutFixedStValue ('AFF_POURCENTAGE', StTot, 644,12, tcDouble100, false)
else
  If (StTot[468] <>'O')Then MttEch := StrtoFloat (Copy (Sttot, 644,12))/100.0
  else T_AFFAIRE.PutFixedStValue ('AFF_POURCENTAGE', StTot, 644,12, tcDouble100, false);

T_AFFAIRE.PutValue ('AFF_DATERESIL',idate2099);

    // les mtt en monnaie pivot sont toujours pris tel quel
T_AFFAIRE.PutValue ('AFF_MONTANTECHE', MttEch);
T_AFFAIRE.PutFixedStValue ('AFF_TOTALHT', StTot, 803,12, tcDouble100, false);
T_AFFAIRE.PutFixedStValue ('AFF_TOTALHTGLO', StTot, 632,12, tcDouble100, false);
    // les mtt devise, sont égale à ce qui est en fichier SCOT si monnaie client = monnaie tenue
    // convertit sinon
If (VH^.TenueEuro = MonPres ) then begin
  T_AFFAIRE.PutValue ('AFF_MONTANTECHEDEV', MttEch);
  T_AFFAIRE.PutFixedStValue ('AFF_TOTALHTDEV', StTot, 803,12, tcDouble100, false);
  T_AFFAIRE.PutFixedStValue ('AFF_TOTALHTGLODEV', StTot, 632,12, tcDouble100, false);
  end
else begin             
  T_AFFAIRE.PutValue ('AFF_MONTANTECHEDEV', COnvertSaisieEf (MttEch, FALSE));;
  T_AFFAIRE.PutValue ('AFF_TOTALHTDEV', ConvertMtt (Copy (StTot, 803,12), Sttot[1571]));
  T_AFFAIRE.PutValue ('AFF_TOTALHTGLODEV', ConvertMtt (Copy(StTot, 632,12), Sttot[1557]));
  end;

// Non fait/ le format est très particulier !! T_AFFAIRE.PutFixedStValue ('AFF_NUMDERGENER', StTot, 782,6, tcTrim, false);
T_AFFAIRE.PutFixedStValue ('AFF_RESPONSABLE', StTot, 213,6, tcTrim, false);
if (fparam.checkApport.checked = true) then T_AFFAIRE.PutFixedStValue ('AFF_RESSOURCE1', StTot, 213,6, tcTrim, false);
If (ctxTempo in V_PGI.PGIContexte) then T_AFFAIRE.PutFixedStValue ('AFF_NUMSITUATION', StTot, 1552,4, tcTrim, false);
if (trim(copy(StTot, 1544,8)) <> '') and (trim(copy(StTot, 1544,8)) <> '00000000')
    then  T_AFFAIRE.PutFixedStValue ('AFF_DATESITUATION', StTot, 1544,8, tcdate8amj, false)
else    T_AFFAIRE.PutValue ('AFF_DATESITUATION',idate1900);

// MCD à voir si à conserver T_AFFAIRE.PutFixedStValue ('AFF_COLLECTIF', StTot, 1573,8, tctrim, false);
T_AFFAIRE.PutValue('AFF_UTILISATEUR',V_PGI.User);
T_AFFAIRE.PutValue('AFF_CREATEUR',V_PGI.User);
T_AFFAIRE.PutValue('AFF_CREERPAR','IMP');
If (TOM_AFF.VerifTOB (T_AFFAIRE)) then  ImportTob(T_AFFAIRE,True)
Else writeln (FichierLog,T_AFFAIRE.GetValue('AFF_TIERS')+' '+T_AFFAIRE.GetValue('AFF_AFFAIRE')+' '+T_AFFAIRE.GetValue('AFF_LIBELLE')+' Erreur '+inttostr(TOM_AFF.LastError)+' '+TOM_AFF.LastErrorMsg);

T_AFFAIRE.Free; TOM_AFF.Free;
// récupération des tiers sur affaires responsable d'affaire + apporteur
Rang := 1;
If (trim(Copy(StTot,165,4))<>'') then  Begin RecupAffTiers(StTot,'APP',Rang, Aff); Inc(Rang); End;
If (Trim(Copy(StTot,213,4))<>'') then  RecupAffTiers(StTot,'RES',Rang, Aff);
If (Trim(Copy(StTot,224,25))<>'') then  RecupAffAdresse(StTot, aff);
If (Trim(Copy(StTot,485,45))<>'') then
    begin
    champ := TiersAuxiliaire (AffectTiers(Copy (StTot, 8,10 )), False);
    RecupRib(Copy(StTot,485,47), champ ,FALSE);
    end;

End;

procedure TFAsRecupTempo.RecupAffTiers (StTot,TypeTiers:string;Rang:Integer; Aff :String);
var T_AFFTIERS : TOB;
    TOM_AFFTIERS : TOM;
    St:string;
Begin
T_AFFTIERS := TOB.CREATE ('AFFTIERS', NIL, -1);
TOM_AFFTIERS := CreateTom('AFFTIERS',Nil,False, False);
TOM_AFFTIERS.InitTob(T_AFFTIERS);
T_AFFTIERS.PutValue ('AFT_AFFAIRE', Aff);
T_AFFTIERS.PutValue ('AFT_RANG',Rang);
If (TypeTiers ='APP') then St:= Trim(Copy (StTot,165,6)) Else St:= Trim(Copy (StTot,213,6));
T_AFFTIERS.PutValue('AFT_RESSOURCE',AnsiUppercase(St));
T_AFFTIERS.PutValue('AFT_LIENAFFTIERS', TypeTiers);


If (TOM_AFFTIERS.VerifTOB (T_AFFTIERS)) then  ImportTob(T_AFFTIERS,True)
Else writeln (FichierLog,T_AFFTIERS.GetValue('AFT_TIERS')+' '+T_AFFTIERS.GetValue('AFT_AFFAIRE')+' '+T_AFFTIERS.GetValue('AFT_LIBELLE')+' Erreur '+inttostr(TOM_AFFTIERS.LastError)+' '+TOM_AFFTIERS.LastErrorMsg);
T_AFFTIERS.Free;
TOM_AFFTIERS.Free;
End;

procedure TFAsRecupTempo.RecupHisto (StTot:string);
var T_HISTOACT : TOB;
        Part0, part1, part2, part3,champ,aff : string;
        mois : integer;
Begin
T_HISTOACT := TOB.CREATE ('HISTOACTIVITE', NIL, -1);

T_HISTOACT.PutValue ('AHI_TIERS',AffectTiers (Copy (StTot, 7,10)));
T_HISTOACT.PutValue ('AHI_RANG',StrToint (Copy (Sttot, 29,1) ));
aff:=AffectAff(Copy(StTot, 18,6),Copy(StTot,24,5), AffectTiers(Copy (StTot, 7,10)),'AFF','00',Part0, Part1, Part2, Part3, False, mois);
T_HISTOACT.PutValue ('AHI_AFFAIRE', aff);
if (aff ='') then begin
   champ :=TraduitGA(format('Enrgt HISTO non traité : client %10.10s Affaire: %6.6s ',[AffectTiers(Copy (StTot, 7,10)),Copy(StTot, 18,6)]));
   If ctxscot  in V_PGI.PGIContexte then champ := champ + format('Exercice:%5.5s',[Copy(StTot,24,5)]);
   writeln (FichierLog, champ);
   T_HISTOACT.Free; 
   exit;
   end;
T_HISTOACT.PutValue ('AHI_AFFAIRE0', Part0);
T_HISTOACT.PutValue ('AHI_AFFAIRE1', Part1);
T_HISTOACT.PutValue ('AHI_AFFAIRE2', Part2);
T_HISTOACT.PutValue ('AHI_AFFAIRE3', Part3);
T_HISTOACT.PutValue ('AHI_DATEMODIF', V_PGI.dateEntree);
T_HISTOACT.PutValue ('AHI_DATECREATION',  V_PGI.dateEntree);

T_HISTOACT.PutFixedStValue ('AHI_NOMTIERS', StTot, 30,32, tcTrimR, false);
T_HISTOACT.PutFixedStValue ('AHI_QTEPRES', StTot, 78,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_TOTPRPRES', StTot, 90,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_TOTPVPRES', StTot, 102,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_TOTFACTPRES', StTot, 114,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_QTEFRAI', StTot, 126,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_TOTPRFRAI', StTot, 138,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_TOTPVFRAI', StTot, 150,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_TOTFACTFRAI', StTot, 162,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_QTEFOUR', StTot, 174,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_TOTPRFOUR', StTot, 186,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_TOTPVFOUR', StTot, 198,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_TOTFACTFOUR', StTot, 210,12, tcDouble100, false);
T_HISTOACT.PutFixedStValue ('AHI_TOTFACTGLOB', StTot, 222,12, tcDouble100, false);

   // on met les stat dans bonne zone en fct paramétrage
Case Stat1Cli of
  1:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS1', StTot, 62,4, tcTrim, false);
  2:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS2', StTot, 62,4, tcTrim, false);
  3:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS3', StTot, 62,4, tcTrim, false);
  4:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS4', StTot, 62,4, tcTrim, false);
  7:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE1', StTot, 62,4, tcTrim, false);
  8:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE2', StTot, 62,4, tcTrim, false);
  9:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE3', StTot, 62,4, tcTrim, false);
  End;
Case Stat2Cli of
  1:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS1', StTot, 66,4, tcTrim, false);
  2:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS2', StTot, 66,4, tcTrim, false);
  3:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS3', StTot, 66,4, tcTrim, false);
  4:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS4', StTot, 66,4, tcTrim, false);
  7:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE1', StTot, 66,4, tcTrim, false);
  8:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE2', StTot, 66,4, tcTrim, false);
  9:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE3', StTot, 66,4, tcTrim, false);
  End;
Case Stat3Cli of
  1:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS1', StTot, 70,4, tcTrim, false);
  2:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS2', StTot, 70,4, tcTrim, false);
  3:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS3', StTot, 70,4, tcTrim, false);
  4:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS4', StTot, 70,4, tcTrim, false);
  7:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE1', StTot, 70,4, tcTrim, false);
  8:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE2', StTot, 70,4, tcTrim, false);
  9:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE3', StTot, 70,4, tcTrim, false);
  End;
Case Stat4Cli of
  1:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS1', StTot, 74,4, tcTrim, false);
  2:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS2', StTot, 74,4, tcTrim, false);
  3:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS3', StTot, 74,4, tcTrim, false);
  4:  T_HISTOACT.PutFixedStValue ('AHI_LIBRETIERS4', StTot, 74,4, tcTrim, false);
  7:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE1', StTot, 74,4, tcTrim, false);
  8:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE2', StTot, 74,4, tcTrim, false);
  9:  T_HISTOACT.PutFixedStValue ('AHI_RESSOURCE3', StTot, 74,4, tcTrim, false);
  End;



ImportTob(T_HISTOACT,True) ;
T_HISTOACT.Free;

End;


procedure TFAsRecupTempo.RecupAffAdresse (StTot, aff:string);
var T_ADRESSE : TOB;
    St: String;
 QQ : TQuery ;
    IMax :integer ;
Begin
    // fct qui crée les adresses éventuelles de facturations mission
T_ADRESSE := TOB.CREATE ('ADRESSES', NIL, -1);
T_ADRESSE.initValeurs;

T_ADRESSE.PutValue ('ADR_REFCODE', aff);
    // on lit le ficheir Adresse pour connaitre le plus grans numéro
QQ:=OpenSQL('SELECT MAX(ADR_NUMEROADRESSE) FROM ADRESSES',TRUE,-1,'',true) ;
if Not QQ.EOF then imax:=QQ.Fields[0].AsInteger+1 else iMax:=1;
Ferme(QQ) ;
T_ADRESSE.PutValue ('ADR_NUMEROADRESSE',InttoStr(IMax));

If ctxTempo in V_PGI.PGIContexte then
  begin
  St:= StTot[1591];
  If (St<>'O')then T_ADRESSE.PutValue('ADR_TYPEADRESSE','INT') Else T_ADRESSE.PutValue ('ADR_TYPEADRESSE','FAC');
  end
else
  T_ADRESSE.PutValue('ADR_TYPEADRESSE','AFA');

T_ADRESSE.PutFixedStValue ('ADR_LIBELLE', StTot, 219,30, tcTrim, false);
T_ADRESSE.PutFixedStValue ('ADR_LIBELLE2', StTot, 249,32, tcTrim, false);
T_ADRESSE.PutFixedStValue ('ADR_ADRESSE1', StTot, 286,32, tcTrim, false);
T_ADRESSE.PutFixedStValue ('ADR_ADRESSE2', StTot, 318,32, tcTrim, false);
T_ADRESSE.PutFixedStValue ('ADR_CODEPOSTAL', StTot, 350,5, tcTrim, false);
T_ADRESSE.PutFixedStValue ('ADR_VILLE', StTot, 355,26, tcTrim, false);
T_ADRESSE.PutValue ('ADR_PAYS', 'FRA');

ImportTob(T_ADRESSE,True);
T_ADRESSE.Free;
End;

procedure TFAsRecupTempo.RecupTiersAdresse (StTot, clt:string);
var T_ADRESSE : TOB;
 QQ : TQuery ;
    IMax :integer ;
Begin
    // fct qui crée adresses fatcurations tiers
T_ADRESSE := TOB.CREATE ('ADRESSES', NIL, -1);
T_ADRESSE.initValeurs;

T_ADRESSE.PutValue ('ADR_REFCODE', clt);
    // on lit le ficheir Adresse pour connaitre le plus grans numéro
QQ:=OpenSQL('SELECT MAX(ADR_NUMEROADRESSE) FROM ADRESSES',TRUE,-1,'',true) ;
if Not QQ.EOF then imax:=QQ.Fields[0].AsInteger+1 else iMax:=1;
Ferme(QQ) ;
T_ADRESSE.PutValue ('ADR_NUMEROADRESSE',InttoStr(IMax));

// mcd 31/05/01 !!! TIE en saisie !!! T_ADRESSE.PutValue('ADR_TYPEADRESSE','T');
T_ADRESSE.PutValue('ADR_TYPEADRESSE','TIE');

T_ADRESSE.PutFixedStValue ('ADR_LIBELLE', StTot, 429,30, tcTrim, false);
T_ADRESSE.PutFixedStValue ('ADR_LIBELLE2', StTot, 459,32, tcTrim, false);
T_ADRESSE.PutFixedStValue ('ADR_ADRESSE1', StTot, 496,32, tcTrim, false);
T_ADRESSE.PutFixedStValue ('ADR_ADRESSE2', StTot, 528,32, tcTrim, false);
T_ADRESSE.PutFixedStValue ('ADR_CODEPOSTAL', StTot, 560,5, tcTrim, false);
T_ADRESSE.PutFixedStValue ('ADR_VILLE', StTot, 565,26, tcTrim, false);
T_ADRESSE.PutFixedStValue ('ADR_TELEPHONE', StTot, 841,20, tcTrim, false);
T_ADRESSE.PutValue ('ADR_PAYS', 'FRA');

ImportTob(T_ADRESSE,True);
T_ADRESSE.Free;
End;

procedure TFAsRecupTempo.RecupLigneAffaire (StTot:string);
var T_LIGNE , TobDet: TOB;
    TypArt, Part0,Part1, Part2, Part3, Aff, statut, champ,puht,Article: String;
    QQ :TQuery;
    NbPiece, mois ,ii: Integer;
    ttc,taxe, ttccon, taxecon ,qte: double;
Begin
T_LIGNE := TOB.CREATE ('LIGNE', NIL, -1);
Taxe:=0;
statut :='AFF';
if sttot[6]='D' then statut:='PRO';

    // on affecte les informations article
If (stTot[33] = 'P') then
    begin
    champ := debutpre;
    end
 Else If (stTot[33] = 'F') then
    begin
    champ := debutfour;
    end
Else  If  (stTot[33] = 'f')then
    Begin
     champ := debutfrais;
   End;
champ := champ + copy (stTot,34,15);
If  (stTot[33] = 'C') then begin
    champ :='';   // cas commentaire  pas de code article
    end;
If  (stTot[33] = 't') then begin
    champ :='';   // cas sous total  pas de code article
    end;

If  (stTot[33] = 'A')then champ :=VH_GC.AFACOMPTE;   // cas acompte. Prise paramètre par défaut
if trim(champ) ='' then article:=''
else Article:=CodeArticleUnique2( Champ,'');
T_LIGNE.PutFixedStValue ('GL_RESSOURCE', StTot, 260,6, tcTrimR, false);
champ :=TraduitGA(format('Enrgt LIGNE AFF: clt %10.10s Affaire: %6.6s ',[AffectTiers(Copy (StTot, 8,10)),Copy(StTot, 18,6)]));
If ctxscot  in V_PGI.PGIContexte then champ := champ + format('Exercice:%5.5s',[Copy(StTot,24,5)]);
Champ:=champ + TraduitGA(format(' Art %10.10s ',[Article]));
TypArt:='PRE'; //mcd 25/02/03
If (stTot[33] = 'P') then  TypArt:='PRE'
 Else If (stTot[33] = 'F') then  TypArt:='MAR'
 Else  If  (stTot[33] = 'f')then TypArt:='FRA';
if (TestExistance(champ,Article,TypArt,T_LIGNE.GetValue('GL_RESSOURCE'),
   Copy (StTot, 18,6),Copy (StTot , 24,5),
   AffectTiers(Copy (StTot, 8,10 )), statut,Copy(Sttot,366,2),Part0,Part1, Part2, Part3,
   aff ,  mois)=False) then begin
 T_LIGNE.Free;
 exit;
 end;
// ne marche pas si avenant sur même affaire !! if (Precedent <> Copy (Sttot,6,23)) then
if (Precedent <> aff) then
    begin
        // on met à jour la piéce précedente pour cumul tva et ttc
    if precedent <>'' then MajPieceAff();
        // on regarde sous quel n° stocker les lignes
    NbPiece := SelectPieceAffaire(Aff, Statut, CleDocLigaf);
    if (NbPiece =0) then
        begin   // cas affaire provisoire on ne traite pas
        T_LIGNE.Free;
        exit;
        end;
    Precedent := aff;
    NoOrdre :=0;
    TotTTC :=0;  TotTaxe :=0;  totqte:=0;
           //ajout test mcd 22/03/01
    if (T_PieceSauv <>Nil) then T_PIECESAUV.Free;
    T_PIECESAUV :=NIL;
    T_PIECESAUV := TOB.CREATE ('PIECE', NIL, -1);
    QQ:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(CleDocLigaf,ttdPiece,False),True,-1,'',true) ;
    If Not QQ.EOF then T_PIECESAUV.SelectDB('',QQ) ;
    Ferme(QQ) ;
    end;

AddLesSupLigne(T_LIGNE,False) ;
InitLigneVide(T_PIECESAUV,T_LIGNE, Nil,Nil, 1,0) ;   // initialise ligne piece
     // mcd 13/05/02 remis ici (au lieu d'avant le test existance, car info effacer par init
T_LIGNE.PutValue('GL_REMISABLELIGNE','-');
T_LIGNE.PutValue('GL_REMISABLEPIED','X');
T_LIGNE.PutValue('GL_ESCOMPTABLE','X');
     // on affecte les informations article
T_LIGNE.PutValue('GL_TYPELIGNE','ART');
T_LIGNE.PutValue('GL_TYPEARTICLE','PRE');
If (stTot[33] = 'P') then
    begin
    T_LIGNE.PutValue('GL_TYPEARTICLE','PRE');
    end
 Else If (stTot[33] = 'F') then
    begin
    T_LIGNE.PutValue('GL_TYPEARTICLE','MAR');
    end
Else  If  (stTot[33] = 'f')then
    Begin
    T_LIGNE.PutValue('GL_TYPEARTICLE','FRA');
   End;
If  (stTot[33] = 'C') then begin
    T_LIGNE.PutValue('GL_TYPELIGNE','COM');
    T_LIGNE.PutValue('GL_TYPEARTICLE','');
    end;
If  (stTot[33] = 't') then begin
    T_LIGNE.PutValue('GL_TYPEARTICLE','');
    T_LIGNE.PutValue('GL_TYPELIGNE','TOT');
    T_LIGNE.PutValue('GL_TVAENCAISSEMENT','-');
    T_LIGNE.PutValue('GL_ESCOMPTABLE','-');
    T_LIGNE.PutValue('GL_REMISABLELIGNE','-');
    T_LIGNE.PutValue('GL_REMISABLEPIED','-');
    end;

    // on affecte le n° de lignes dans la piece
Inc (NoOrdre);
T_LIGNE.PutValue ('GL_NUMLIGNE',NoOrdre);
T_LIGNE.PutValue ('GL_NUMORDRE',NoOrdre);

if trim(Article) <>'' then T_LIGNE.PutValue ('GL_ARTICLE',Article);
T_LIGNE.PutFixedStValue ('GL_CODEARTICLE', Article, 1,17, tcTrimr, false);
T_LIGNE.PutFixedStValue ('GL_REFARTSAISIE', Article, 1,17, tcTrimr, false);
T_LIGNE.PutFixedStValue ('GL_LIBELLE', StTot, 49,70, tcTrimr, false);
T_LIGNE.PutFixedStValue ('GL_BLOCNOTE', StTot, 49,90, tcTrimr, false);
if (T_LIGNE.GetValue('GL_TYPELIGNE') = 'COM') and (Trim(Copy(StTot,49,70)) = '*')
    then   T_LIGNE.PutValue ('GL_LIBELLE',' ');

ttc := StrToFloat(copy(sttot, 177,12))/100.0;
for ii :=1 to nbtva+1 do begin
    if CodeTva[ii]= Sttot[167] then begin
        taxe := (ttc * TauxTva[ii])/100;
        if (VH^.TenueEuro) then taxe:=Arrondi(taxe ,V_PGI.OkDecV)
        else taxe:=Arrondi(taxe ,V_PGI.OkDecV);
        ttc := ttc+taxe;
        totttc := totttc + ttc;
        TotTaxe := TotTaxe + taxe;
        if (TauxTva[ii])< 18.0 then    T_LIGNE.PutValue ('GL_FAMILLETAXE1','RED')
          else    T_LIGNE.PutValue ('GL_FAMILLETAXE1','NOR');
        if (Sttot[167]='0') then  T_LIGNE.PutValue ('GL_FAMILLETAXE1','EXO');
        end;
    end;
          // le code PUHT est obligatoire, on le met = total ht si blanc
PuhT :=Copy (StTot,168,9);
if (trim(puht) = '0') then  PuhT :=Copy (StTot,180,9);  // zone 177 sur 9c au lieu de 12 ..
if (trim(puht)='') then puht:='        0';
T_LIGNE.PutFixedStValue ('GL_TOTALHT', StTot, 177,12, tcDouble100, false);
T_LIGNE.PutFixedStValue ('GL_MONTANTHT', StTot, 177,12, tcDouble100, false);
T_LIGNE.PutFixedStValue ('GL_PUHT', Puht,1,9, tcDouble100, false);
T_LIGNE.PutFixedStValue ('GL_PMRP', StTot, 251,9, tcDouble100, false);
T_LIGNE.PutFixedStValue ('GL_PMRPACTU', StTot, 251,9, tcDouble100, false);
T_LIGNE.PutFixedStValue ('GL_DPR', StTot, 251,9, tcDouble100, false);
// Mis à zéro en saisie manuelle ... T_LIGNE.PutFixedStValue ('GL_PUHTBASE', StTot, 168,9, tcDouble100, false);
T_LIGNE.PutFixedStValue ('GL_PUHTNET', Puht,1,9, tcDouble100, false);
T_LIGNE.PutValue ('GL_MONTANTTTC', ttc);
T_LIGNE.PutValue ('GL_TOTALTTC', ttc);
T_LIGNE.PutValue ('GL_TOTALTAXE1', taxe);
T_LIGNE.PutValue ('GL_PUTTC', ttc);
T_LIGNE.PutValue ('GL_PUTTCNET', ttc);
if (VH^.TenueEuro) then begin
    ttccon:=  EuroToFranc(TTC);
    taxecon:=  EuroToFranc(Taxe);
    end
else begin
    ttccon :=  FrancToEuro(TTC);
    taxecon:=  FrancToEuro(Taxe);
    end;

    // les mtt devise, sont égale à ce qui est en fichier SCOT si monnaie client = monnaie tenue
    // convertit sinon
    // fait différemment de MONpres alimenté à partir du client: dans ce cas, on sait si on est en contrevaleur ou pas,
    // sans avoir besoin de faire une comparaison entre monnaie présentation et monnaie tenue dossier
if (T_PIECESAUV.GetValue ('GP_SAISIECONTRE') =  '-')then begin
  T_LIGNE.PutFixedStValue ('GL_MONTANTHTDEV', StTot, 177,12, tcDouble100, false);
  T_LIGNE.PutFixedStValue ('GL_TOTALHTDEV', StTot, 177,12, tcDouble100, false);
  T_LIGNE.PutFixedStValue ('GL_PUHTDEV', Puht,1,9, tcDouble100, false);
  T_LIGNE.PutFixedStValue ('GL_PUHTNETDEV',Puht,1,9, tcDouble100, false);
  T_LIGNE.PutValue ('GL_MONTANTTTCDEV', ttc);
  T_LIGNE.PutValue ('GL_TOTALTTCDEV', ttc);
  T_LIGNE.PutValue ('GL_TOTALTAXEDEV1', taxe);
  T_LIGNE.PutValue ('GL_PUTTC', ttc);
  T_LIGNE.PutValue ('GL_PUTTCNET', ttc);
  end
else begin
  T_LIGNE.PutValue ('GL_MONTANTHTDEV', COnvertMtt(Copy(StTot, 177,12), Sttot[330]));
  T_LIGNE.PutValue ('GL_TOTALHTDEV', COnvertMtt(Copy(StTot, 177,12), Sttot[330]));
  T_LIGNE.PutValue ('GL_PUHTDEV', COnvertMtt(Copy(Puht,1,9), Sttot[328]));
  T_LIGNE.PutValue ('GL_PUHTNETDEV', COnvertMtt(Copy(Puht,1,9), Sttot[328]));
  T_LIGNE.PutValue ('GL_MONTANTTTCDEV', ttccon);
  T_LIGNE.PutValue ('GL_TOTALTTCDEV', ttccon);
  T_LIGNE.PutValue ('GL_TOTALTAXEDEV1', taxecon);
  T_LIGNE.PutValue ('GL_PUTTC', ttccon);
  T_LIGNE.PutValue ('GL_PUTTCNET', ttccon);
  end;
T_LIGNE.PutValue ('GL_PUTTC', ttccon);
T_LIGNE.PutValue ('GL_PUTTCNET', ttccon);

T_LIGNE.PutValue ('GL_COTATION', 1);
T_LIGNE.PutValue ('GL_PCB', 1);

if (T_LIGNE.GetValue('GL_TYPELIGNE')<>'COM') and (T_LIGNE.GetValue('GL_TYPELIGNE')<>'TOT')    then
  begin
  qte :=StrtoFloat(Copy(Sttot,139,6));
  if (qte = 0) then qte :=1 else qte := qte/100;
  T_LIGNE.PutValue ('GL_QTEFACT', qte);
  T_LIGNE.PutValue ('GL_QTESTOCK', qte);
  totqte :=totqte + qte;
  end;
T_LIGNE.PutFixedStValue ('GL_DPR', StTot, 251,9, tcDouble100, false);
T_LIGNE.PutFixedStValue ('GL_RESSOURCE', StTot, 260,6, tcTrimR, false);
T_LIGNE.PutFixedStValue ('GL_QUALIFQTEVTE', StTot, 266,1, tcTrimR, false);
if (sttot[266] = ' ') and (stTot[33] = 'P') then begin
    T_LIGNE.PutValue ('GL_QUALIFQTEVTE', 'H');
    TobDet := TobArt.FindFirst(['GA_ARTICLE'],[T_LIGNE.GetValue('GL_ARTICLE')],False);
    if TobDet <> NIL Then
      Begin
        T_LIGNE.PutValue ('GL_QUALIFQTEVTE', TobDet.GetValue('GA_QUALIFUNITEVTE'));
      End;
    end;

// MCD quand l'avancement sera géré, voir pour reprendre les zones correspodnantes
// dans le ficheir de sauvegare : % avancement, mtt avancement , qté facturée, % facturé, mtt facturé

T_LIGNE.PutValue('GL_VALIDECOM','AFF');
T_LIGNE.PutValue('GL_UTILISATEUR',V_PGI.User);
T_LIGNE.PutValue('GL_CREERPAR','IMP');
T_LIGNE.PutValue('GL_TENUESTOCK','-');

T_LIGNE.PutValue ('GL_PERIODE',GetPeriode(T_LIGNE.GetValue('GL_DATEPIECE')));
T_LIGNE.PutValue ('GL_SEMAINE',NumSemaine(T_LIGNE.GetValue('GL_DATEPIECE')));

if (DebEnc = 'E') then T_LIGNE.PutValue('GL_TVAENCAISSEMENT','X')
else    T_LIGNE.PutValue('GL_TVAENCAISSEMENT','-');


ImportTob(T_LIGNE,True);

T_LIGNE.Free;

End;



procedure TFAsRecupTempo.RecupTiersCOntact (StTot:string; champ:string;lng_auxi:integer);
var T_CONTACT: TOB;

Begin
T_CONTACT:= TOB.CREATE ('CONTACT', NIL, -1);
T_CONTACT.initValeurs;

T_CONTACT.PutFixedStValue ('C_AUXILIAIRE', champ, 1,lng_auxi, tcTrim, false);
T_CONTACT.PutFixedStValue ('C_NOM', StTot, 390,32, tcTrim, false);
T_CONTACT.PutValue ('C_NUMEROCONTACT',1);
T_CONTACT.PutValue ('C_NATUREAUXI','CLI');
T_CONTACT.PutValue ('C_TYPECONTACT','T');

ImportTob(T_CONTACT,True);
T_CONTACT.Free;
End;


procedure TFAsRecupTempo.RecupFactAff (StTot, Aff, datdeb, datfin,GenerAuto:string; Var Periodicite:string; Var Interval : Integer; MonPres: boolean);
var T_FACTAFF : TOB;
    TOM_FACTAFF : TOM;
    St,StMois,StFourniture : string;
    Rang, Posit,boucle, nbech :integer;
    DateDeb,DateFin ,DatTrav: TDateTime;
    a,m,j ,aa,mm,jj: Word ;
Begin
// type de génération automatique
Periodicite :='M';
Interval :=1;
boucle:=0;
nbech:=0;
    // on ne traite pas les missions manuelles
If (((ctxTempo in V_PGI.PGIContexte) and (StTot[468] <>'T'))
or ((ctxScot in V_PGI.PGIContexte) and (StTot[468] <>'A'))) then
  Begin
  Rang := 1;
  StMois := Copy (StTot,656,12);
  StFourniture := Copy (StTot,580,12);
  T_FACTAFF := TOB.CREATE ('FACTAFF', Nil, -1);
  TOM_FACTAFF := CreateTom('FACTAFF',Nil,False, False);
  DateDeb := Str8ToDate(Datdeb,False);
  DateFin := (Str8ToDate(Datfin,False));
  // prise en compte de la date de fin réelle si < date de fin prévue
  If (trim(Copy(StTot,416,8)) <> '' ) Then DateFin :=(Str8ToDate(Copy(StTot,416,8),False));
  // prise en compte de la date de 1ere facturation
  If (trim(Copy(StTot,1516,8)) <> '')Then
    begin
    DateDeb :=(Str8ToDate(Copy(StTot,1516,8),False));
    if (Copy(StTot,1516,8)  <>Datdeb) then
        begin
            // dans le cas ou la date de facture n'est pas la même
            // que la date de debut, on met le même écart dans la
            // date de fin ...
        DatTrav :=   Str8ToDate(Datdeb,False);
        If (trim(Copy(StTot,416,8)) <> '' ) then DateFin := DateFin + (Datedeb - DatTrav);
        end;
    end;
//select aff_generauto from affaire where aff_affaire=factaff.afa_affaire)
DecodeDate(DateDeb,aa,mm,jj) ; //mcd 11/10/01 stock date orifigne

  If ctxTempo in V_PGI.PGIContexte then      // gm 28/10/03
  Begin
    If (DateFin >= GetParamSoc('SO_AFDATEFINGENER')) then
    begin
     If GetParamSoc('SO_AFDATEFINGENER')<> iDate2099 then
        DateFin := GetParamSoc('SO_AFDATEFINGENER')
     else
        DateFin := PlusDate (DateDeb,1,'A');
    end;
  End;

  While (FinDeMois(DateDeb) <= DateFin) do
      Begin
      DecodeDate(DateDeb,a,m,j) ;
      Inc(Boucle);  // on regarde sur une boucle de 12 mois, combien on
                    // a d'échéances, pour la périodicité
      If (StMois[m] = 'O') Then
        Begin
         // génération de l'échéance
        TOM_FACTAFF.InitTob (T_FACTAFF);
        // mcd 19/06/02 le champ afa_generauto doit être renseigné
        T_FACTAFF.PutValue ('AFA_GENERAUTO',GenerAuto);
        T_FACTAFF.PutValue ('AFA_TIERS', AffectTiers (Copy (StTot, 8,10)));
        T_FACTAFF.PutValue ('AFA_TYPECHE', 'NOR');
        T_FACTAFF.PutValue ('AFA_AFFAIRE', Aff);
            // cas facture au pourcentage ou en montant
        If (((ctxTempo in V_PGI.PGIContexte) and (StTot[468] <>'O'))
            or ((ctxScot in V_PGI.PGIContexte) and (StTot[468] <>'P')))Then
            begin
            T_FACTAFF.PutFixedStValue ('AFA_MONTANTECHE', StTot,1364+(12*(m-1)),12, tcDouble100, false) ;
            If (VH^.TenueEuro = MonPres ) then T_FACTAFF.PutFixedStValue ('AFA_MONTANTECHEDEV', StTot,1364+(12*(m-1)),12, tcDouble100, false)
            else T_FACTAFF.PutValue ('AFA_MONTANTECHEDEV',ConvertMtt (Copy ( StTot,1364+(12*(m-1)),12) , Sttot[(1559+ m-1)]));
               // les mtt contre valeur, sont toujours convertit
            T_FACTAFF.PutValue ('AFA_DEVISE',V_PGI.DevisePivot);
            end
        Else T_FACTAFF.PutFixedStValue ('AFA_POURCENTAGE', StTot, 1364+(12*(m-1)),12, tcDouble100, false);

        if (FinDeMois(DateDeb) < CurrentDate) then T_FACTAFF.PutValue ('AFA_ECHEFACT','X')
                 else T_FACTAFF.PutValue ('AFA_ECHEFACT','-');
        St:=VH_GC.AFFLIBFACTAFF;
        // PL le 27/06/02 : gestion du numéro d'échéance dans le libellé
        T_FACTAFF.PutValue ('AFA_NUMECHE',Rang);
        T_FACTAFF.PutValue ('AFA_NUMECHEBIS',Rang);
        Posit := Pos ('**',St);
        If (Posit <>0) then Begin delete(St,Posit,2);Insert(T_FACTAFF.GetValue('AFA_NUMECHEBIS'),St,Posit); End;
        Posit := Pos ('$$',St);
        // mcd 20/09/01 If (Posit <>0) then Begin delete(St,Posit,2);Insert(DateToStr(FinDeMois(DateDeb)),St,Posit); End;
        If (Posit <>0) then Begin delete(St,Posit,2);Insert(DateToStr(DateDeb),St,Posit); End;
        // mcd 20/09/01 If (St='') then St:=('Echéance du '+DateToStr(FinDeMois(DateDeb)));
        If (St='') then St:=(T_FACTAFF.GetValue('AFA_NUMECHEBIS')+' : Echéance du '+DateToStr(DateDeb));
        // Fin PL le 27/06/02

        T_FACTAFF.PutValue ('AFA_LIBELLEECHE',St);

        // mcd 08/10/01 T_FACTAFF.PutValue ('AFA_DATEECHE',FinDeMois(DateDeb));
        T_FACTAFF.PutValue ('AFA_DATEECHE',DateDeb);
                           // modifier pour mettre valeur par défaut, et en fct fourniture ou pas
                           // change la valeur
        T_FACTAFF.PutValue ('AFA_REPRISEACTIV',VH_GC.AFRepriseActiv);
        // En facturation aux temps passés on reprend systématiquement tous les éléments.
        If (((ctxTempo in V_PGI.PGIContexte) and (StTot[468] ='P'))
                or ((ctxScot in V_PGI.PGIContexte) and (StTot[468] ='T')))Then T_FACTAFF.PutValue ('AFA_REPRISEACTIV','TOU');
                // mcd 03/10/02  If (StFourniture[m] = '1') Then begin
        If (StFourniture[m] <> '1') Then begin
           if T_FACTAFF.GetValue ('AFA_REPRISEACTIV') = 'ARF' then T_FACTAFF.PutValue ('AFA_REPRISEACTIV','FRA');
           if T_FACTAFF.GetValue ('AFA_REPRISEACTIV') = 'ARP' then T_FACTAFF.PutValue ('AFA_REPRISEACTIV','PRE') ;
           if T_FACTAFF.GetValue ('AFA_REPRISEACTIV') = 'ART' then T_FACTAFF.PutValue ('AFA_REPRISEACTIV','NON');
           if T_FACTAFF.GetValue ('AFA_REPRISEACTIV') = 'TOU' then T_FACTAFF.PutValue ('AFA_REPRISEACTIV','FRP');
           end;

        If (TOM_FACTAFF.VerifTOB (T_FACTAFF)) then  ImportTob(T_FACTAFF,True)
        Else writeln (FichierLog,T_FACTAFF.GetValue('AFA_TIERS')+' '+T_FACTAFF.GetValue('AFA_AFFAIRE')+' '+T_FACTAFF.GetValue('AFA_LIBELLE')+' Erreur '+inttostr(TOM_FACTAFF.LastError)+' '+TOM_FACTAFF.LastErrorMsg);
        Inc (Rang);
        if (boucle <=12) then Inc(Nbech);
        End;
      DateDeb:=PlusDate(DateDeb,1,'M');
      if (JJ = DaysPerMonth(aa,mm)) then DateDeb:=FinDeMois(DateDeb);    //mcd 11/01/02
      End;
        // en fonction du nbre d'échéances trouvées sur une année, on met l'intervalle voulu
  if (NbEch =1) then Periodicite :='A'
    else If (NbEch=3)  then Interval:=4
        else If (NbEch=4) then Interval:=3
            else If (NbEch=2) then Interval:=6
                else If (NbEch=6) then Interval:=2;
  T_FACTAFF.Free; TOM_FACTAFF.Free;
  End;
End;

procedure TFAsRecupTempo.RecupClient (StTot:string);
var T_CLIENT : TOB;
    champ, champ1 , ModeRempl :string;
    Lng_auxi,  i, nbech, duree, jour, separe , MontantMin:integer;

Begin
if (CheckRepriseSup.checked=False) then   begin
    champ :=Copy (Sttot, 719,8);
    if (champ = '00000000') then champ:='        ';
    If (Trim(Copy (champ,1,8)) <>'') Then exit;
   end;

T_CLIENT := TOB.CREATE ('TIERS', NIL, -1);
T_CLIENT.initValeurs;
Lng_auxi:=VH^.Cpta[fbAux].Lg;


// il faut copier le compte sur la longueur en fichier
// fait avec le caractère de bourrage des paramètres
champ := StringOfChar (VH^.Cpta[fbAux].Cb, Lng_auxi);
if (LngCpte > 10) then   champ1 := TRim (copy (StTot, 727, 12)) // cas SST
  else champ1 := TRim (copy (StTot, 727, 10));
if (LiaisonCompta = FALSE) then
    begin
        // dans le cas de non liaison compta, on copie le code client
        // précéder de 9
    champ1 := '9' + trim (Copy (StTot, 7,10));
    end;
champ := champ1 + champ;
T_CLIENT.PutFixedStValue ('T_AUXILIAIRE', champ,1,lng_auxi, tcChaine, false);
   // si on est en lien avec le DP, et que dossier gamme 2 ou
   // dossiers de production pas repris, on alimente annuaire
if (GetParamsoc('SO_AFLIENDP') = true) and
   ((Gamme2.checked = true ) or(produc.checked = False))
      then RecupAnnuaire (Sttot,T_CLIENT.GetValue ('T_AUXILIAIRE'));

    // on récupère le RIB
if (Trim(Copy (Sttot, 613,45)) <> '') then RecupRib (Copy (Sttot, 613,47), Copy (champ, 1 , Lng_Auxi), TRUE);
         // on récupère le contact = nom correspondant
If (Trim(Copy (Sttot, 390,32)) <>'') then
    begin
    RecupTiersContact (Sttot,champ,VH^.Cpta[fbAux].Lg);
    end;


T_CLIENT.PutValue ('T_NATUREAUXI','CLI');
       //forme juridique. Celles ci sont renumérotée dans une combo à 3c
for i:=0 to FparamFjur.TobFjur.detail.count-1 do
   begin
   if (copy (StTot, 157,5) = FparamFjur.TobFjur.Detail[i].GetValue('Code')) then
      begin
      T_CLIENT.PutValue ('T_FORMEJURIDIQUE',FparamFjur.TobFjur.Detail[i].GetValue('Nouveau code'));
      break;
      end;
    end;
if (FparamCli.ConcatAbv.checked=True) then T_CLIENT.PutFixedStValue ('T_LIBELLE', StTot, 95,30, tcTrim, false)
else begin
          // mcd 15/03/02 possibilité de ne pas concatener ABV + enseigne de SCOT
          // plus renseigne zone ABrvaition GI si valeur connu
     T_CLIENT.PutFixedStValue ('T_LIBELLE', StTot, 100,25, tcTrim, false);
     if Copy(Sttot,157,5) = COpy(StTot,95,5) then T_CLIENT.PutValue('T_JURIDIQUE',T_CLIENT.Getvalue('T_FORMEJURIDIQUE'))
     else if COpy(Sttot,95,2) ='DR' then T_CLIENT.putValue('T_JURIDIQUE','DR')
     else if COpy(Sttot,95,2) ='ME' then T_CLIENT.putValue('T_JURIDIQUE','ME')
     else if COpy(Sttot,95,2) ='MM' then T_CLIENT.putValue('T_JURIDIQUE','MME')
     else if COpy(Sttot,95,2) ='ML' then T_CLIENT.putValue('T_JURIDIQUE','MLE')
     else if COpy(Sttot,95,1) ='M' then T_CLIENT.putValue('T_JURIDIQUE','MR') ;
     end;
T_CLIENT.PutValue ('T_EAN',  AffectTiers (Copy (Sttot,7,10)));
T_CLIENT.PutValue ('T_TIERS',  AffectTiers (Copy (Sttot,7,10)));

Lng_auxi:=VH^.Cpta[fbGene].Lg;
// il faut copier le compte collectif sur la longueur en fichier
// fait avec le caractère de bourrage des paramètres
champ := StringOfChar (VH^.Cpta[fbGene].Cb, Lng_auxi);
champ1 := TRim (copy (StTot, 727, 10));
If ctxScot in V_PGI.PGIContexte
   Then begin
    If (Trim(Copy (Sttot,1985,10)) <>'') then champ1 :=  trim (Copy (StTot, 1985,10))
    else champ1 :=FParamCli.EditCollectifClient.Text;
    end
    else begin
    If (Trim(Copy (Sttot,1033,10)) <>'') then champ1 := trim (Copy (StTot, 1033,10))
    else champ1 :=FParamCli.EditCollectifClient.Text;
    end;
champ := champ1 + champ;
T_CLIENT.PutFixedStValue ('T_ORIGINETIERS', sttot,428,1, tctrim, false);
T_CLIENT.PutFixedStValue ('T_COLLECTIF', champ,1,lng_auxi, tcChaine, false);
T_CLIENT.PutFixedStValue ('T_ABREGE', StTot, 17,10, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_ADRESSE1', StTot, 162,32, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_ADRESSE2', StTot, 194,32, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_CODEPOSTAL', StTot, 226,5, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_VILLE', StTot, 231,26, tcTrim, false);
    // le code pays est systématiquement mis à FRA (= France PGI)
    // les code pays étranger ne sont pas traité dans la reprise.
    // ce sont des cas exceptionnels avec modif fiche après.
T_CLIENT.PutValue ('T_PAYS', 'FRA');
T_CLIENT.PutValue ('T_NATIONALITE', 'FRA');
T_CLIENT.PutValue ('T_DEVISE',V_PGI.DevisePivot);
T_CLIENT.PutValue ('T_LANGUE','FRA');
T_CLIENT.PutValue ('T_MULTIDEVISE','-');
                  //identification CEE
T_CLIENT.PutFixedStValue ('T_NIF', StTot, 821,9, tcTrim, false);
              //siren + siret
T_CLIENT.PutFixedStValue ('T_SIRET', StTot, 340,14, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_APE', StTot, 386,4, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_TELEPHONE', StTot, 260,20, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_RVA', StTot, 280,20, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_FAX', StTot, 300,20, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_TELEX', StTot, 320,20, tcTrim, false);
If ctxScot  in V_PGI.PGIContexte then T_CLIENT.PutFixedStValue ('T_SOCIETEGROUPE', StTot, 1652,10, tcTrim, false);

// traitement de l'adresse de facturation
if (trim (Copy (StTot, 676,10)) <> '') then
    begin
    T_CLIENT.PutValue ('T_FACTURE', AffectTiers (Copy (StTot, 676,10)));
    end
else begin
        // il faut à priori (11/07/00) mettre le code tiers dans le code client à facturer si identique
    T_CLIENT.PutValue ('T_FACTURE',  AffectTiers (Copy (Sttot,7,10)));
    end;
if (trim (Copy (StTot, 429, 30)) <> '') then
    begin
    RecupTiersAdresse (Sttot,AffectTiers (Copy (Sttot,7,10)));
    end;

If ctxTempo in V_PGI.PGIContexte then T_CLIENT.PutFixedStValue ('T_TARIFTIERS', StTot, 1021,2, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_ESCOMPTE', StTot, 609,4, tcDouble100, false);
T_CLIENT.PutValue ('T_QUALIFESCOMPTE', 'DED');
T_CLIENT.PutValue ('T_FACTUREHT', 'X');
T_CLIENT.PutValue ('T_LETTRABLE', 'X');
    // on renseigne monnaie de présentation client
If (Copy(StTot,1010,1)='E') Then T_CLIENT.PutValue ('T_EURODEFAUT', 'X')
              Else T_CLIENT.PutValue ('T_EURODEFAUT', '-');

        // si lien S7, le code régime TVA peut avoir des valeurs bien particulières,
        // on reprend tel quel
if (S7 = 'O')
    then begin
    T_CLIENT.PutFixedStValue ('T_REGIMETVA', StTot, 660,1, tcTrim, false);
    end
else
  begin
  If (Copy(StTot,660,1)='T') Then T_CLIENT.PutValue ('T_REGIMETVA', 'FRA') Else
  If (Copy(StTot,660,1)='E') Then T_CLIENT.PutValue ('T_REGIMETVA', 'EXO') Else
  If (Copy(StTot,660,1)='S') Then T_CLIENT.PutValue ('T_REGIMETVA', 'EXO');
  end;
if (DebEnc = 'E') then T_CLIENT.PutValue('T_TVAENCAISSEMENT','TE')
else    T_CLIENT.PutValue('T_TVAENCAISSEMENT','TD');

T_CLIENT.PutValue ('T_SOUMISTPF', '-');
T_CLIENT.PutFixedStValue ('T_DATECREATION', StTot, 703,8, tcdate8amj, false);
T_CLIENT.PutFixedStValue ('T_DATEMODIF', StTot,711,8, tcdate8amj, false);
champ :=Copy (Sttot, 719,8);
if (champ = '00000000') then champ:='        ';
If (Trim(Copy (champ,1,8)) ='') Then T_CLIENT.PutValue ('T_FERME', '-')
         Else  begin
         T_CLIENT.PutFixedStValue ('T_DATEFERMETURE', champ, 1, 8, tcdate8amj, false);
         T_CLIENT.PutValue ('T_FERME', 'X');
         end;
T_CLIENT.PutValue ('T_CONFIDENTIEL','-');
T_CLIENT.PutValue ('T_RELEVEFACTURE','-');
T_CLIENT.PutValue ('T_CREERPAR','IMP');
T_CLIENT.PutValue ('T_UTILISATEUR',V_PGI.User);
         // traitement des Stats paramétrables
Case Stat1Cli of
  5:  T_CLIENT.PutFixedStValue ('T_APPORTEUR', StTot, 687,4, tcTrim, false);
  6:  T_CLIENT.PutFixedStValue ('T_SECTEUR', StTot, 687,4, tcTrim, false);
  End;
Case Stat2Cli of
  5:  T_CLIENT.PutFixedStValue ('T_APPORTEUR', StTot, 691,4, tcTrim, false);
  6:  T_CLIENT.PutFixedStValue ('T_SECTEUR', StTot, 691,4, tcTrim, false);
  End;
Case Stat3Cli of
  5:  T_CLIENT.PutFixedStValue ('T_APPORTEUR', StTot, 695,4, tcTrim, false);
  6:  T_CLIENT.PutFixedStValue ('T_SECTEUR', StTot, 695,4, tcTrim, false);
  End;
Case Stat4Cli of
  5:  T_CLIENT.PutFixedStValue ('T_APPORTEUR', StTot, 699,4, tcTrim, false);
  6:  T_CLIENT.PutFixedStValue ('T_SECTEUR', StTot, 699,4, tcTrim, false);
  End;

If ctxScot  in V_PGI.PGIContexte then T_CLIENT.PutFixedStValue ('T_CREDITPLAFOND', StTot, 1975,10, tcdouble100, false)
                                else T_CLIENT.PutFixedStValue ('T_CREDITPLAFOND', StTot, 1023,10, tcdouble100, false);
If ctxScot  in V_PGI.PGIContexte then begin
   T_CLIENT.PutFixedStValue ('T_MOISCLOTURE', StTot, 1011,2, tcEntier, false) ;
   i:= StrtoInt (T_CLIENT.getvalue('T_MOISCLOTURE'));
   if (i < 1) or (i >12) then T_CLIENT.PutValue ('T_MOISCLOTURE','12');
   end;
T_CLIENT.PutValue ('T_AVOIRRBT','-');
T_CLIENT.PutValue ('T_ISPAYEUR','-');
        // traitement des codes régime comptable .
        // MCD ATTENTION ne prend que les 3 premiers caractères.
Case FParamCli.ComboFamComptaTiers.ItemIndex of
   1:  T_CLIENT.PutFixedStValue ('T_COMPTATIERS', StTot, 687,3, tcTrim, false);
   2:  T_CLIENT.PutFixedStValue ('T_COMPTATIERS', StTot, 691,3, tcTrim, false);
   3:  T_CLIENT.PutFixedStValue ('T_COMPTATIERS', StTot, 695,3, tcTrim, false);
   4:  T_CLIENT.PutFixedStValue ('T_COMPTATIERS', StTot, 699,3, tcTrim, false);
   End;

  // traitement des modes de règlement
champ :=Copy (Sttot, 600,1);
if champ =' '  then champ :='1';        // mcd 13/05/02 pour eviter plantage
if (champ = 'A') then  Nbech  := 10 else
if (champ = 'B') then  Nbech  := 11 else
if (champ = 'C') then  Nbech  := 12 else
if (champ = 'D') then  Nbech  := 13 else
if (champ = 'E') then  Nbech  := 14 else
//if (champ = 'F') then  Nbech  := 15 else nbech := StrtoInt (Copy (Sttot, 600,1));
if (champ = 'F') then  Nbech  := 15 else nbech := StrtoInt (Champ);  // mcd 13/05/02


separe:= StrtoInt (Copy (Sttot, 601,3));
duree := StrtoInt (Copy (Sttot, 604,3));
jour := StrtoInt (Copy (Sttot, 607,2));
MontantMin := 0;
ModeRempl :='';
If ((Sttot [597] = 'T') or (Sttot [597] = 'L') or (Sttot [597] = 'P')) then
    begin
    MontantMin := MttMin;
    Moderempl := MRRempl;
    end;
T_CLIENT.PutValue ('T_MODEREGLE', RecupModeRegl (Copy (Sttot, 599,1), Copy (Sttot, 597,2),
    ModeRempl, duree, jour, nbech, separe, MontantMin));


if (trim(copy(sttot,1520,40)))<> '' then
   If ctxScot  in V_PGI.PGIContexte then T_CLIENT.PutFixedStValue ('T_BLOCNOTE', StTot, 1520,100, tcTrim, false);
REcupCLientCOmpl(sttot, T_CLIENT.GetValue('T_AUXILIAIRE'),T_CLIENT.GetValue('T_TIERS'));
ImportTob(T_CLIENT,True);
T_CLIENT.Free;
End;

procedure TFAsRecupTempo.RecupClientcompl (StTot,auxi,tiers:string);
var T_CLIENTCOMPL : TOB;
    champ :string;
Begin
T_CLIENTCOMPL := TOB.CREATE ('TIERSCOMPL', NIL, -1);
T_CLIENTCOMPL.initValeurs;
T_CLIENTCOMPL.PutValue ('YTC_AUXILIAIRE', auxi);
T_CLIENTCOMPL.PutValue ('YTC_TIERS', TIERS);
         // on récupère la date d'entrée cabinet pour SCOT dans la 1ere date libre
If ctxScot in V_PGI.PGIContexte then
    begin
    champ :=Copy (Sttot, 1017,8);
    if (champ = '00000000') then champ:='        ';
    If (Trim(Copy (champ,1,8)) <>'') Then T_CLIENTCOMPL.PutFixedStValue ('YTC_DATELIBRE1', champ, 1, 8, tcdate8amj, false);
    end;
        // Traitement des tables libres  de 1 à 26 = GT, 27 = Origine
champ := Tablelibre (StTot, FParamCli.ComboTableLibre1.ItemIndex);
T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS1', champ, 1,3, tcTrim, false);
champ := Tablelibre (StTot, FParamCli.ComboTableLibre2.ItemIndex);
T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS2', champ, 1,3, tcTrim, false);
champ := Tablelibre (StTot, FParamCli.ComboTableLibre3.ItemIndex);
T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS3', champ, 1,3, tcTrim, false);
champ := Tablelibre (StTot, FParamCli.ComboTableLibre4.ItemIndex);
T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS4', champ, 1,3, tcTrim, false);
champ := Tablelibre (StTot, FParamCli.ComboTableLibre5.ItemIndex);
T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS5', champ, 1,3, tcTrim, false);
champ := Tablelibre (StTot, FParamCli.ComboTableLibre6.ItemIndex);
T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS6', champ, 1,3, tcTrim, false);
champ := Tablelibre (StTot, FParamCli.ComboTableLibre7.ItemIndex);
T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS7', champ, 1,3, tcTrim, false);
champ := Tablelibre (StTot, FParamCli.ComboTableLibre8.ItemIndex);
T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS8', champ, 1,3, tcTrim, false);
champ := Tablelibre (StTot, FParamCli.ComboTableLibre9.ItemIndex);
T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS9', champ, 1,3, tcTrim, false);
champ := Tablelibre (StTot, FParamCli.ComboTableLibre10.ItemIndex);
T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERSA', champ, 1,3, tcTrim, false);

         // traitement des Stats paramétrables
Case Stat1Cli of
  1:  T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS1', StTot, 687,4, tcTrim, false);
  End;
Case Stat2Cli of
  1:  T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS1', StTot, 691,4, tcTrim, false);
  2:  T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS2', StTot, 691,4, tcTrim, false);
  End;
Case Stat3Cli of
  1:  T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS1', StTot, 695,4, tcTrim, false);
  2:  T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS2', StTot, 695,4, tcTrim, false);
  3:  T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS3', StTot, 695,4, tcTrim, false);
  End;
Case Stat4Cli of
  1:  T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS1', StTot, 699,4, tcTrim, false);
  2:  T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS2', StTot, 699,4, tcTrim, false);
  3:  T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS3', StTot, 699,4, tcTrim, false);
  4:  T_CLIENTCOMPL.PutFixedStValue ('YTC_TABLELIBRETIERS4', StTot, 699,4, tcTrim, false);
  End;
champ :=trim (AnsiUppercase(Copy(Sttot,687,4)))  ;
if (Stat1cli = 7) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE1', champ);
if (Stat1cli = 8) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE2', champ);
if (Stat1cli = 9) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE3', champ);
champ :=trim (AnsiUppercase(Copy(Sttot,691,4)));
if (Stat2cli = 7) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE1', champ);
if (Stat2cli = 8) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE2', champ);
if (Stat2cli = 9) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE3', champ);
champ :=trim (AnsiUppercase(Copy(Sttot,695,4)));
if (Stat3cli = 7) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE1', champ);
if (Stat3cli = 8) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE2', champ);
if (Stat3cli = 9) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE3', champ);
champ :=trim (AnsiUppercase(Copy(Sttot,699,4)));
if (Stat4cli = 7) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE1', champ);
if (Stat4cli = 8) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE2', champ);
if (Stat4cli = 9) then T_CLIENTCOMPL.PutValue ('YTC_RESSOURCE3', champ);

ImportTob(T_CLIENTCOMPL,True);
T_CLIENTCOMPL.Free;
End;

procedure TFAsRecupTempo.RecupClientcomplFou (StTot,auxi,tiers:string);
var T_CLIENTCOMPL : TOB;
Begin
T_CLIENTCOMPL := TOB.CREATE ('TIERSCOMPL', NIL, -1);
T_CLIENTCOMPL.initValeurs;
T_CLIENTCOMPL.PutValue ('YTC_AUXILIAIRE', auxi);
T_CLIENTCOMPL.PutValue ('YTC_TIERS', TIERS);

ImportTob(T_CLIENTCOMPL,True);
T_CLIENTCOMPL.Free;
End;

function TFAsRecupTempo.Tablelibre (StTot:string; ii:integer ):string ;

          // fct qui permet d'alimenter dans les tables libres les GT.
begin
result :='   ';
Case ii of
  1:  result:=Copy( StTot,740,3);
  2:  result:=Copy( StTot,743,3);
  3:  result:=Copy( StTot,746,3);
  4:  result:=Copy( StTot,749,3);
  5:  result:=Copy( StTot,752,3);
  6:  result:=Copy( StTot,755,3);
  7:  result:=Copy( StTot,758,3);
  8:  result:=Copy( StTot,761,3);
  9:  result:=Copy( StTot,764,3);
  10:  result:=Copy( StTot,767,3);
  11:  result:=Copy( StTot,770,3);
  12:  result:=Copy( StTot,773,3);
  13:  result:=Copy( StTot,776,3);
  14:  result:=Copy( StTot,779,3);
  15:  result:=Copy( StTot,782,3);
  16:  result:=Copy( StTot,785,3);
  17:  result:=Copy( StTot,788,3);
  18:  result:=Copy( StTot,791,3);
  19:  result:=Copy( StTot,794,3);
  20:  result:=Copy( StTot,797,3);
  21:  result:=Copy( StTot,800,3);
  22:  result:=Copy( StTot,803,3);
  23:  result:=Copy( StTot,806,3);
  24:  result:=Copy( StTot,809,3);
  25:  result:=Copy( StTot,812,3);
  26:  result:=Copy( StTot,815,3);
  End;
End;

function TFAsRecupTempo.RecupTypeTable ( ii:integer ):string ;

          // fct qui permet d'alimenter dans les tables libres les GT.
begin
result:='';
If ii = FParamCli.ComboTableLibre1.ItemIndex then result:='LT1';
If ii  = FParamCli.ComboTableLibre2.ItemIndex then result:='LT2';
If ii  = FParamCli.ComboTableLibre3.ItemIndex then result:='LT3';
If ii  = FParamCli.ComboTableLibre4.ItemIndex then result:='LT4';
If ii  = FParamCli.ComboTableLibre5.ItemIndex then result:='LT5';
If ii  = FParamCli.ComboTableLibre6.ItemIndex then result:='LT6';
If ii  = FParamCli.ComboTableLibre7.ItemIndex then result:='LT7';
If ii  = FParamCli.ComboTableLibre8.ItemIndex then result:='LT8';
If ii  = FParamCli.ComboTableLibre9.ItemIndex then result:='LT9';
If ii  = FParamCli.ComboTableLibre10.ItemIndex then result:='LTA';

End;

procedure TFAsRecupTempo.RecupTempsFact (StTot:string);
    //pour récupérer les factures dans les temps
    // cette fct crée des factures de reprise FRE, avec le n° de facture = AAMOIS + N°
var T_LIGNE , T_Piece: TOB;
    Part0,Part1, Part2, Part3, Aff, champ,Etbs,Article: String;
    mois : Integer;
    QQ : TQuery;
    ttc,taxe, qte, puht,puttc, taux,totht: double;
Begin
T_LIGNE := TOB.CREATE ('LIGNE', NIL, -1);
T_PIECE := TOB.CREATE ('PIECE', NIL, -1);
Aff := AffectAff( Copy (StTot, 36,6),Copy (StTot , 50,5),
  AffectTiers(Copy (StTot, 25,10 )), 'AFF','00',Part0,Part1, Part2, Part3,
  False,  mois); // il faut lire affaire avant de traiter la piece
if (aff ='') then begin
    champ :=TraduitGA(format('Enrgt FACT/TEMPS:clt %10.10s Affaire: %6.6s ',[AffectTiers(Copy (StTot, 25,10)),Copy(StTot, 36,6)]));
    If ctxscot  in V_PGI.PGIContexte then champ := champ + format('Exer:%5.5s',[Copy(StTot,50,5)]);
    champ := champ + traduitga(format(' Ressource:%4.4s,Art:%15.15s,mois %4.4s,jour %2.2s',[copy(Sttot,6,4),Article,copy(Sttot,12,4),copy(Sttot,22,2)]));
   writeln (FichierLog, champ);
   T_LIGNE.Free; T_PIECE.Free;
   exit;
   end;

   // on recherche la piece si existe , on crée sinon
taux :=20.60;
if (Copy(Sttot,12,2) < '50') then
   begin
   Champ := COpy(Sttot,22,2)+'/'+ Copy(Sttot,14,2)+'/' + '20'+ Copy(Sttot,12,2);
    if (copy(Sttot,12,4) >= '0004') then taux :=19.60;
   end
else begin
     Champ := COpy(Sttot,22,2)+'/'+ Copy(Sttot,14,2)+'/' + '19'+ Copy(Sttot,12,2);
     end;
if Sttot[172]='$' then taux :=0;         // pour traiter les cas de frais non taxable de SST
if Not(IsvalidDAte(Champ)) then Champ:='01/01/1900';//mcd 07/02/02
CleDocLigaf.datepiece:=Strtodate(champ);
CleDocLIgaf.NaturePiece:='FRE';
 // il faut récupérer l'entite si multi géré
Etbs:=VH^.EtablisDefaut;
if multi = 'O' then begin
     QQ := OpenSQL('SELECT  AFF_ETABLISSEMENT From AFFAIRE WHERE AFF_AFFAIRE="'+ aff+'"',TRUE,-1,'',true);
      If Not QQ.EOF then etbs:=QQ.Fields[0].AsString ;
      Ferme(QQ);
   end;
RecupCreerPiece(CledocLigaf, AffectTiers(Copy (StTot, 25,10 )),aff,part1,part2,part3,'',Etbs ,
         Copy(Sttot,129,6),T_PIECE,NoOrdre);
AddLesSupLigne(T_LIGNE,False) ;
InitLigneVide(T_PIECE,T_LIGNE, Nil,Nil, 1,0) ;   // initialise ligne piece

    // on affecte les informations article
Champ:='';
T_LIGNE.PutValue('GL_TYPELIGNE','ART');
If (stTot[60] = 'P') then
    begin
    T_LIGNE.PutValue('GL_TYPEARTICLE','PRE');
    champ := debutpre;
    if trim(copy (stTot,69,15)) ='' then Champ:='PRGLOBAL'; // mcd 19/06/02 pour fact globale
    end
 Else If (stTot[60] = 'F') then
    begin
    T_LIGNE.PutValue('GL_TYPEARTICLE','MAR');
    champ := debutfour;
    if trim(copy (stTot,69,15)) ='' then Champ:='FOGLOBAL'; // mcd 19/06/02 pour fact globale
    end
Else  If  (stTot[60] = 'f')then
    Begin
    T_LIGNE.PutValue('GL_TYPEARTICLE','FRA');
     champ := debutfrais;
     if trim(copy (stTot,69,15)) ='' then Champ:='FRGLOBAL'; // mcd 19/06/02 pour fact globale
  End ;
champ := champ + copy (stTot,69,15);

If  (stTot[60] = '*')then begin
        T_LIGNE.PutValue('GL_TYPEARTICLE','PRE');
        champ :=VH_GC.AFACOMPTE;   // cas acompte. Prise paramètre par défaut
        end;
if trim(champ) <>'' then T_LIGNE.PutValue ('GL_ARTICLE',CodeArticleUnique2( champ,''));
T_LIGNE.PutFixedStValue ('GL_CODEARTICLE', champ, 1,17, tcTrimr, false);
T_LIGNE.PutFixedStValue ('GL_REFARTSAISIE', champ, 1,17, tcTrimr, false);
if trim(champ) ='' then article:=''
else Article:=  CodeArticleUnique2( champ,'');
if (Copy(Sttot,6,6)<>'******')then T_LIGNE.PutFixedStValue ('GL_RESSOURCE', StTot, 6,6, tcTrimR, false);
champ :=TraduitGA(format('Enrgt FACT/TEMPS:clt %10.10s Affaire: %6.6s ',[AffectTiers(Copy (StTot, 25,10)),Copy(StTot, 36,6)]));
If ctxscot  in V_PGI.PGIContexte then champ := champ + format('Exer:%5.5s',[Copy(StTot,50,5)]);
champ := champ + traduitga(format(' Ressource:%4.4s,Art:%15.15s,mois %4.4s,jour %2.2s',[copy(Sttot,6,4),Article,copy(Sttot,12,4),copy(Sttot,22,2)]));
if (TestExistance(champ,Article,T_LIGNE.GetValue('GL_TYPEARTICLE'),T_LIGNE.getValue ('GL_RESSOURCE'),
  '',Copy (StTot , 50,5),
  AffectTiers(Copy (StTot, 25,10 )), 'AFF','00',Part0,Part1, Part2, Part3,
  aff ,  mois)=False) then
  begin
  T_PIECE.Free; T_LIGNE.Free;
 exit;
 end;
if (trim(copy(sttot,36,6)) ='') and (trim(copy(sttot,50,5)) ='') then aff:=''; // cas affaire à blanc...


T_LIGNE.PutValue ('GL_NUMLIGNE',NoOrdre);
T_LIGNE.PutValue ('GL_NUMORDRE',NoOrdre);
T_LIGNE.PutFixedStValue ('GL_LIBELLE', StTot, 99,30, tcTrimr, false);
T_LIGNE.PutFixedStValue ('GL_BLOCNOTE', StTot, 99,30, tcTrimr, false);
T_LIGNE.PutValue ('GL_AFFAIRE', aff);
T_LIGNE.PutValue ('GL_AFFAIRE1', part1);
T_LIGNE.PutValue ('GL_AFFAIRE2', part2);
T_LIGNE.PutValue ('GL_AFFAIRE3', part3);
T_LIGNE.PutValue ('GL_AVENANT', '00');
T_LIGNE.PutValue ('GL_TIERSLIVRE', T_LIGNE.Getvalue('GL_TIERS'));
T_LIGNE.PutValue ('GL_TIERSPAYEUR', T_LIGNE.Getvalue('GL_TIERS'));
T_LIGNE.PutValue ('GL_TIERSFACTURE', T_LIGNE.Getvalue('GL_TIERSPAYEUR'));
T_PIECE.Putvalue ('GP_TIERSLIVRE',T_LIGNE.GetValue('GL_TIERSLIVRE'));
T_PIECE.Putvalue ('GP_TIERSFACTURE',T_LIGNE.GetValue('GL_TIERSFACTURE'));
T_PIECE.Putvalue ('GP_TIERSPAYEUR',T_LIGNE.GetValue('GL_TIERSPAYEUR'));
ttc := StrToFloat(copy(sttot, 159,12))/100.0;
taxe := (ttc * taux)/100;
if (VH^.TenueEuro) then taxe:=Arrondi(taxe ,V_PGI.OkDecV)
        else taxe:=Arrondi(taxe ,V_PGI.OkDecV);
ttc := ttc+taxe;
T_LIGNE.PutValue ('GL_FAMILLETAXE1','NOR');
          // le code PUHT est obligatoire, on le met = total ht si blanc
PuhT :=StrTofloat(Copy (StTot,159,12));
puht := puht/100;
totht:=puht;
Qte:= StrTofloat(Copy(Sttot,137,10)) ;
qte := qte/100;
if Qte <>0 then puht := (puht) / qte else qte:=1.0;
puttc := puht+ (puht * taux)/100;

T_LIGNE.PutValue ('GL_TOTALHT',totht);
T_LIGNE.PutValue ('GL_MONTANTHT', totht);
T_LIGNE.PutValue ('GL_PUHT', Puht);
T_LIGNE.PutValue ('GL_PMRP',puht);
T_LIGNE.PutValue ('GL_PMRPACTU', puht);
T_LIGNE.PutValue ('GL_DPR', puht);
T_LIGNE.PutValue ('GL_PUHTNET', Puht);
T_LIGNE.PutValue ('GL_MONTANTTTC', ttc);
T_LIGNE.PutValue ('GL_TOTALTTC', ttc);
T_LIGNE.PutValue ('GL_TOTALTAXE1', taxe);
T_LIGNE.PutValue ('GL_PUTTC', puttc);
T_LIGNE.PutValue ('GL_PUTTCNET', puttc);
T_LIGNE.PutValue ('GL_PUHTBASE', puht);
T_LIGNE.PutValue ('GL_PUTTCBASE', puttc);
    // on est en monnaie de tenue du dossier. Mtt=devsie, contrevaleur traduite.
T_LIGNE.PutValue ('GL_MONTANTHTDEV', totht);
T_LIGNE.PutValue ('GL_TOTALHTDEV', totht);
T_LIGNE.PutValue ('GL_PUHTDEV', Puht);
T_LIGNE.PutValue ('GL_PUHTNETDEV',Puht);
T_LIGNE.PutValue ('GL_MONTANTTTCDEV', ttc);
T_LIGNE.PutValue ('GL_TOTALTTCDEV', ttc);
T_LIGNE.PutValue ('GL_TOTALTAXEDEV1', taxe);
T_LIGNE.PutValue ('GL_PUTTCDEV', puttc);
T_LIGNE.PutValue ('GL_PUTTCNETDEV', puttc);

T_LIGNE.PutValue ('GL_COTATION', 1);
T_LIGNE.PutValue ('GL_PCB', 1);

if (T_LIGNE.GetValue('GL_TYPELIGNE')<>'COM') and (T_LIGNE.GetValue('GL_TYPELIGNE')<>'TOT')    then
  begin
  qte :=StrtoFloat(Copy(Sttot,137,10));
  if (qte = 0) then qte :=1 else qte := qte/100;
  T_LIGNE.PutValue ('GL_QTEFACT', qte);
  T_LIGNE.PutValue ('GL_QTESTOCK', qte);
  totqte :=totqte + qte;
  end;
T_LIGNE.PutFixedStValue ('GL_DPR', StTot, 147,12, tcDouble100, false);
T_LIGNE.PutFixedStValue ('GL_QUALIFQTEVTE', StTot, 135,1, tcTrimR, false);
if (sttot[135] = ' ')then T_LIGNE.PutValue ('GL_QUALIFQTEVTE', 'H');
T_LIGNE.PutValue('GL_VALIDECOM','AFF');
T_LIGNE.PutValue('GL_UTILISATEUR',V_PGI.User);
T_LIGNE.PutValue('GL_CREERPAR','IMP');
T_LIGNE.PutValue('GL_TENUESTOCK','-');
T_LIGNE.PutValue('GL_ESCOMPTABLE','X');
T_LIGNE.PutValue('GL_REMISABLELIGNE','-');
T_LIGNE.PutValue('GL_REMISABLEPIED','X');

T_LIGNE.PutValue ('GL_PERIODE',GetPeriode(T_LIGNE.GetValue('GL_DATEPIECE')));
T_LIGNE.PutValue ('GL_SEMAINE',NumSemaine(T_LIGNE.GetValue('GL_DATEPIECE')));
T_LIGNE.PutValue('GL_TVAENCAISSEMENT','-');
ImportTob(T_LIGNE,True);
T_LIGNE.Free;
             // on cumul les mtt dans la piece
T_PIECE.PutValue ('GP_TOTALTTC',T_PIECE.GetValue('GP_TOTALTTC')+ TTC);
T_PIECE.PutValue ('GP_TOTALTTCDEV',T_PIECE.GetValue('GP_TOTALTTCDEV')+TTC);
T_PIECE.PutValue ('GP_TOTALQTEFACT',T_PIECE.GetValue('GP_TOTALQTEFACT')+qte);
T_PIECE.PutValue ('GP_TOTALQTESTOCK',T_PIECE.GetValue('GP_TOTALQTESTOCK')+qte);
T_PIECE.PutValue ('GP_EDITEE', 'X') ;
           // affecte le total des lignes
T_PIECE.PutValue ('GP_TOTALHT',T_PIECE.GetValue('GP_TOTALHT')+totht);
T_PIECE.PutValue ('GP_TOTALBASEREM',T_PIECE.GetValue('GP_TOTALBASEREM')+totht);
T_PIECE.PutValue ('GP_TOTALBASEESC',T_PIECE.GetValue('GP_TOTALBASEESC')+totht);
T_PIECE.PutValue ('GP_TOTALHTDEV',T_PIECE.GetValue('GP_TOTALHTDEV')+totht);
T_PIECE.PutValue ('GP_TOTALBASEREMDEV',T_PIECE.GetValue('GP_TOTALBASEREMDEV')+totht);
T_PIECE.PutValue ('GP_TOTALBASEESCDEV',T_PIECE.GetValue('GP_TOTALBASEESCDEV')+totht);
T_PIECE.PutValue('GP_TVAENCAISSEMENT','TE') ;
T_PIECE.PutValue ('GP_MODEREGLE', GetParamSoc('SO_GcModeRegleDefaut'));
T_PIECE.PutValue ('GP_PERIODE',GetPeriode(T_PIECE.GetValue('GP_DATEPIECE')));
T_PIECE.PutValue ('GP_SEMAINE',NumSemaine(T_PIECE.GetValue('GP_DATEPIECE')));
T_PIECE.PutValue('GP_CREEPAR','IMP');
// mcd 03/12/02 pour mise X T_PIECE.PutValue('GP_VIVANTE','-');

ImportTob(T_PIECE,True);
T_PIECE.Free;

end;

procedure TFAsRecupTempo.RecupTemps (StTot:string);
var T_ACTIVITE : TOB;
    StDate,Stdate1,Article,Ress:String;
    QQ : TQuery;
    Famille, champ,Part0, Part1, Part2, Part3,aff: string;
    PU ,Qte: double;
    Ressource:TAFO_Ressource;
    slFonction:TStringList;
    mois :integer;
Begin
StDate:=copy (StTot,22,2)+ Copy(StTot,14,2)+Copy(StTot,12,2);
if (Copy(Sttot,12,2)>'50') then StDate1:=copy (StTot,22,2)+'/'+ Copy(StTot,14,2)+'/19'+Copy(StTot,12,2) //mcd 07/02/02
else  StDate1:=copy (StTot,22,2)+'/'+ Copy(StTot,14,2)+'/20'+Copy(StTot,12,2); //mcd 07/02/02
if Not(IsvalidDate(StDATE1)) then StDate:='010150';  //mcd 07/02/02
   // toutes les lignes closes coté SCOT sont prises
    // il est donc difficile de se pointer: journal assistant sur même période doit être OK
    // GL ==> plus dur, prend en compte les lignes closes client ...
//if (Copy(sttot , 63,2) = '~~') then exit;

// Test date d'activité incluse dans la reprise
If ((FicTempsDate.Text ='') Or (Str6ToDate(StDate,90) >= StrToDate(FicTempsDate.Text))) Then
   Begin
  T_ACTIVITE := TOB.CREATE ('ACTIVITE', NIL, -1);
  T_ACTIVITE.InitValeurs;


   If (stTot[60] = 'P') then
      begin
      T_ACTIVITE.PutValue('ACT_TYPEARTICLE','PRE');
      T_ACTIVITE.PutValue('ACT_ACTIVITEEFFECT','X');  // PL le 07/01/02 travail effectif pour les prestations
      champ := debutpre;
      end
   Else If (stTot[60] = 'F') then
      begin
      T_ACTIVITE.PutValue('ACT_TYPEARTICLE','MAR');
      champ := debutfour;
      end
  Else  If  (stTot[60] = 'f')then
      Begin
       champ := debutfrais;
     End;
  champ := champ + copy (stTot,69,15);
  if trim(champ)='' then Article:=''
  else Article:=  CodeArticleUnique2(champ, '');
  T_ACTIVITE.PutValue ('ACT_ARTICLE',Article );
  T_ACTIVITE.PutFixedStValue ('ACT_CODEARTICLE', champ, 1,17, tcTrimr, false);

  If  (stTot[60] = 'f')then
    begin
     // Test s'il s'agit d'une prime
    famille:='';
    if (FamPrime <> '(((') then
        begin
       // on ne fait pas requête si pas familleprime
       QQ := OpenSQL('SELECT  GA_FAMILLENIV1 From ARTICLE WHERE GA_CODEARTICLE="'+ Trim(champ)+'"',TRUE,-1,'',true);
        If Not QQ.EOF then Famille:=QQ.Fields[0].AsString else Famille := '';
        Ferme(QQ);
        end;
    If (FamPrime = Famille) Then T_ACTIVITE.PutValue('ACT_TYPEARTICLE','PRI')
     Else T_ACTIVITE.PutValue('ACT_TYPEARTICLE','FRA');
    end;

  if (Sttot[59]='B') then T_ACTIVITE.PutValue ('ACT_TYPEACTIVITE','BON')
  else T_ACTIVITE.PutValue ('ACT_TYPEACTIVITE','REA');
  T_ACTIVITE.PutValue ('ACT_TYPERESSOURCE','SAL');
  Ress:=AnsiUppercase(trim(Copy(StTot,6,6)));
  if (Ress = '******') then Ress:=''
   else  T_ACTIVITE.PutValue ('ACT_RESSOURCE', trim(Ress));
        // si saise en folio, on met n° folio de SCOT,sinon, 1 par défaut
  if (VH_GC.AFTYPESAISIEACT = 'FOL') then T_ACTIVITE.PutFixedStValue ('ACT_FOLIO',StTot,16,2,tcEntier,False)
    else T_ACTIVITE.PutValue ('ACT_FOLIO',1);
  T_ACTIVITE.PutFixedStValue ('ACT_DATEACTIVITE', StDate, 0,6, tcDate6JMA, false);
  T_ACTIVITE.PutValue ('ACT_PERIODE',GetPeriode(T_ACTIVITE.GetValue('ACT_DATEACTIVITE')));
  T_ACTIVITE.PutValue ('ACT_SEMAINE',NumSemaine(T_ACTIVITE.GetValue('ACT_DATEACTIVITE')));
 {mcd 31/03/03   // on prend le n° de séquence du client, mais il est en décroissant
    // depuis 9999. On remet dans l'autre sens.
  seq := StrtoInt (copy(  StTot,55,4));
  seq := 10000 - seq;
  champ := format ('%4.4d',[seq]);
  T_ACTIVITE.PutFixedStValue('ACT_NUMLIGNE',champ,1,4,tcentier,False);}

  T_ACTIVITE.PutValue ('ACT_TIERS',AffectTiers (Copy (StTot, 25,10)));
   champ :=TraduitGA(format('Enrgt TEMPS: clt %10.10s Affaire: %6.6s ',[AffectTiers(Copy (StTot, 25,10)),Copy(StTot, 36,6)]));
   If ctxscot  in V_PGI.PGIContexte then champ := champ + format('Exer:%5.5s',[Copy(StTot,50,5)]);
   champ := champ + traduitga(format(' Ressource:%4.4s,Art:%15.15s,mois %4.4s,jour %2.2s',[copy(Sttot,6,4),Article,copy(Sttot,12,4),copy(Sttot,22,2)]));
  if (TestExistance(champ,Article,T_ACTIVITE.getValue('ACT_TYPEARTICLE'),Ress,Copy(StTot, 36,6),Copy(StTot,50,5), AffectTiers(Copy (StTot, 25,10)), 'AFF','00',Part0,Part1, Part2, Part3,aff ,  mois)=False) then begin
     T_ACTIVITE.Free;
     exit;
     end;
  T_ACTIVITE.PutValue ('ACT_AFFAIRE',aff );
  T_ACTIVITE.PutValue ('ACT_AFFAIRE0', Part0);
  T_ACTIVITE.PutValue ('ACT_AFFAIRE1', Part1);
  T_ACTIVITE.PutValue ('ACT_AFFAIRE2', Part2);
  T_ACTIVITE.PutValue ('ACT_AFFAIRE3', Part3);
  T_ACTIVITE.PutValue ('ACT_AVENANT', '00');
  T_ACTIVITE.PutValue ('ACT_ACTORIGINE', 'REP');
  T_ACTIVITE.putValue ('ACT_NUMLIGNEUNIQUE', ProchainIndiceAffaires (  T_ACTIVITE.GetValue('ACT_TYPEACTIVITE'),
     T_ACTIVITE.GetValue('ACT_AFFAIRE'), ListeAffaire, false, false)); // mcd 31/03/03
  T_ACTIVITE.PutFixedStValue ('ACT_LIBELLE', StTot, 99,30, tcTrim, false);
  If (stTot[60] = 'P') then T_ACTIVITE.PutFixedStValue ('ACT_UNITE', StTot, 136,1, tcChaine, false);
  T_ACTIVITE.PutFixedStValue ('ACT_QTE', StTot, 137,10, tcDouble100, false);
    //mcd 15/05/03 pour alimente unité ref
  if (T_ACTIVITE.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
    T_ACTIVITE.PutValue ('ACT_QTEUNITEREF', ConversionUnite (T_ACTIVITE.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, T_ACTIVITE.GetValue ('ACT_QTE')))
  else
    T_ACTIVITE.PutValue ('ACT_QTEUNITEREF', '0');
  If (stTot[60] = 'P') then T_ACTIVITE.PutFixedStValue ('ACT_UNITEFAC', StTot, 136,1, tcChaine, false);
  T_ACTIVITE.PutFixedStValue ('ACT_QTEFAC', StTot, 137,10, tcDouble100, false);
  T_ACTIVITE.PutValue ('ACT_DEVISE',V_PGI.DevisePivot);
  // Gestion des prix de revient
  // recalcul des P. unitaire non stockés dans Tempo
  Qte := (StrToFloat(Copy (StTot, 137,10))/100);
  If ((FicTempsValor.ItemIndex = 0) Or (FicTempsValor.ItemIndex = 1)) Then
    Begin
    If ((stTot[60] = 'f') and (FamPrime = Famille)) Then
       Begin
       T_ACTIVITE.PutFixedStValue ('ACT_TOTPR', StTot, 185,12, tcDouble100, false);
       if (Qte <>0) then PU := ((StrToFloat(Copy (StTot, 185,12))/100)/Qte)
        else   PU := (StrToFloat(Copy (StTot, 185,12))/100) ;
       T_ACTIVITE.PutValue ('ACT_PUPR',Arrondi(PU,2));
       if (Qte <>0) then PU := ((StrToFloat(Copy (StTot, 185,12))/100)/Qte)
       else  PU := (StrToFloat(Copy (StTot, 185,12))/100);
       T_ACTIVITE.PutValue ('ACT_PUPRCHARGE',Arrondi(PU,2));
       End
    Else
        Begin
        T_ACTIVITE.PutFixedStValue ('ACT_TOTPR', StTot, 147,12, tcDouble100, false);
        If (Qte <>0) then PU := ((StrToFloat(Copy (StTot, 147,12))/100)/Qte)
        else   PU := (StrToFloat(Copy (StTot, 147,12))/100);
        T_ACTIVITE.PutValue ('ACT_PUPR',Arrondi(PU,2));
        T_ACTIVITE.PutValue ('ACT_PUPRCHARGE',Arrondi(PU,2));
        End;
    T_ACTIVITE.PutValue ('ACT_PUPRCHINDIRECT',Arrondi(PU,2));
    T_ACTIVITE.PutFixedStValue ('ACT_TOTPRCHARGE', StTot, 147,12, tcDouble100, false);
    T_ACTIVITE.PutFixedStValue ('ACT_TOTPRCHINDI', StTot, 147,12, tcDouble100, false);
    If (FicTempsValor.ItemIndex = 1) then
       Begin
            // cas ou PV non récupérer. on le met égal au PR
       T_ACTIVITE.PutValue ('ACT_PUVENTE',Arrondi(PU,2));
       T_ACTIVITE.PutFixedStValue ('ACT_TOTVENTE', StTot, 147,12, tcDouble100, false);
       T_ACTIVITE.PutValue ('ACT_PUVENTEDEV',Arrondi(PU,2));
       T_ACTIVITE.PutFixedStValue ('ACT_TOTVENTEDEV', StTot, 147,12, tcDouble100, false);

       End;
    End;

  // Gestion des prix de vente
  If ((FicTempsValor.ItemIndex = 0) Or (FicTempsValor.ItemIndex = 2)) Then
   Begin
   if (Qte <>0) then PU := ((StrToFloat(Copy (StTot, 159,12))/100)/Qte)
   else PU := (StrToFloat(Copy (StTot, 159,12))/100) ;
   T_ACTIVITE.PutValue ('ACT_PUVENTE',Arrondi(PU,2));
   T_ACTIVITE.PutFixedStValue ('ACT_TOTVENTE', StTot, 159,12, tcDouble100, false);
   T_ACTIVITE.PutValue ('ACT_PUVENTEDEV',Arrondi(PU,2));
   T_ACTIVITE.PutFixedStValue ('ACT_TOTVENTEDEV', StTot, 159,12, tcDouble100, false);
   If (FicTempsValor.ItemIndex = 2) then
       Begin
            // cas ou PR non récupérer. on le met égal au PV
       T_ACTIVITE.PutValue ('ACT_PUPR',Arrondi(PU,2));
       T_ACTIVITE.PutValue ('ACT_PUPRCHARGE'  ,Arrondi(PU,2));
       T_ACTIVITE.PutValue ('ACT_PUPRCHINDIRECT' ,Arrondi(PU,2));
       T_ACTIVITE.PutFixedStValue ('ACT_TOTPR', StTot, 159,12, tcDouble100, false);
       T_ACTIVITE.PutFixedStValue ('ACT_TOTPRCHARGE', StTot, 159,12, tcDouble100, false);
       T_ACTIVITE.PutFixedStValue ('ACT_TOTPRCHINDI', StTot, 159,12, tcDouble100, false);
       End;
   End;
    // mcd 30/01/03 pour résoudre le pb de qtefac=0 et prix # 0 . sinon facture temps passé pas ok
  if (T_ACTIVITE.Getvalue('ACT_QTE')= 0) and (T_ACTIVITE.Getvalue('ACT_QTEFAC')= 0)
      and (T_ACTIVITE.Getvalue('ACT_TOTVENTE')<>0)  then
      begin
      T_ACTIVITE.PutValue('ACT_QTE',1);
      T_ACTIVITE.PutValue('ACT_QTEFAC',1);
      if (T_ACTIVITE.GetValue ('ACT_TYPEARTICLE') = 'PRE') then  T_ACTIVITE.PutValue('ACT_QTEUNITEREF',1); //mcd 29/07/03
      end;

  if (sttot[171] = '*') then  T_ACTIVITE.PutValue ('ACT_ACTIVITEREPRIS', 'FAC')
  else   if (sttot[171] = 'R') or (sttot[171] = 'W') then  T_ACTIVITE.PutValue ('ACT_ACTIVITEREPRIS', 'A')
  else   if (sttot[171] = 'A')  then  T_ACTIVITE.PutValue ('ACT_ACTIVITEREPRIS', 'F')
  else T_ACTIVITE.PutFixedStValue ('ACT_ACTIVITEREPRIS', StTot, 171,1, tcChaine, false);
  T_ACTIVITE.PutFixedStValue ('ACT_DATECREATION', StDate, 0,6, tcDate6JMA, false);
  T_ACTIVITE.PutFixedStValue ('ACT_DATEMODIF', StDate, 0,6, tcDate6JMA, false);
  T_ACTIVITE.PutFixedStValue ('ACT_TYPEHEURE', StTot, 184,1, tcChaine, false);
  // T_ACTIVITE.PutValue ('ACT_REALISABLE','X');
    // on recherche la focntion de l'assistant valide le jour de la ligne d'acitivté.
    // on ne traite pas les asisstant ***** utilisé pour les fournitures

slFonction := nil;
if (AnsiUppercase(trim(Copy(StTot,6,6))) <> '******') then
  begin
  try
  Ressource := TAFO_Ressource.Create(AnsiUppercase(trim(Copy(StTot,6,6))));
  if (Ressource<>nil) then
     if (Ressource.tob_Champs<>nil) then
         begin
         slFonction := Ressource.FonctionDeLaRessource(Str6ToDate(StDate ,90));
                    // ajout test mcd le 22/03/01
         if (slFonction<>nil)and (slFonction.count <>0) then T_ACTIVITE.PutValue('ACT_FONCTIONRES', slFonction[0]);
         end;

  finally
  if (slFonction<>nil) then   slFonction.Free;

  end;
Ressource.Free;
end;

    // mcd 02/08/01 ajout initialisation zone par défaut
T_Activite.PutValue('ACT_ETATVISA', 'VIS');
T_Activite.PutValue('ACT_VISEUR', V_PGI.User);
T_Activite.PutValue('ACT_DATEVISA', NowH);
T_Activite.PutValue('ACT_ETATVISAFAC', 'VIS');
T_Activite.PutValue('ACT_VISEURFAC', V_PGI.User);
T_Activite.PutValue('ACT_DATEVISAFAC', NowH);

  ImportTob(T_ACTIVITE,True);
  T_ACTIVITE.Free;
  End;
End;

// Récuperation du module planning
procedure TFAsRecupTempo .RecupPlanning (StTot:string);
var T_ACTIVITE : TOB;
    QQ : TQuery;
    StDate,St , champ , Part0,Part1, Part2, Part3: String;
    Mois:integer;
    heure : extended;

Begin

// Test date de planning incluse dans la reprise
If ((DatePlanning.Text='') Or (Str8ToDate(Copy(StTot,33,8),False) >= StrToDate(DatePlanning.Text))) Then
  Begin
  T_ACTIVITE := TOB.CREATE ('ACTIVITE', NIL, -1);
  T_ACTIVITE.InitValeurs;

  T_ACTIVITE.PutValue ('ACT_TYPEACTIVITE','PRE');
  T_ACTIVITE.PutValue ('ACT_TYPERESSOURCE','SAL');
  T_ACTIVITE.PutValue ('ACT_RESSOURCE', Trim(AnsiUppercase(Copy(StTot, 147,6))));
  T_ACTIVITE.PutValue ('ACT_FOLIO',1);
  T_ACTIVITE.PutFixedStValue ('ACT_DATEACTIVITE', StTot, 33,8, tcDate8AMJ, false);
  T_ACTIVITE.PutValue ('ACT_PERIODE',GetPeriode(T_ACTIVITE.GetValue('ACT_DATEACTIVITE')));
  T_ACTIVITE.PutValue ('ACT_SEMAINE',NumSemaine(T_ACTIVITE.GetValue('ACT_DATEACTIVITE')));
  StDate:=Copy(StTot,39,2)+DateSeparator+Copy(StTot,37,2)+DateSeparator+Copy(StTot,33,4);
  champ := AffectAff(Copy( StTot, 18,6),Copy(StTot,24,5),
            AffectTiers(Copy (StTot, 8,10)), 'AFF','00',Part0,Part1, Part2, Part3, False, Mois);
  if (champ ='') then begin
     champ :=TraduitGA(format('Enrgt PLANNING non traité : client %10.10s Affaire: %6.6s ',[AffectTiers(Copy (StTot, 8,10)),Copy(StTot, 18,6)]));
     If ctxscot  in V_PGI.PGIContexte then champ := champ + format('Exercice:%5.5s',[Copy(StTot,24,5)]);
        //penser d'ajouter TestExistance ....
     writeln (FichierLog, champ);
     T_ACTIVITE.free;        //mcd 22/03/01
     exit;
     end;
  {mcd 31/03/03QQ:=OpenSQL('SELECT MAX(ACT_NUMLIGNE) FROM ACTIVITE WHERE ACT_TYPEACTIVITE="PRE" And ACT_AFFAIRE="'+ champ+'" And ACT_DATEACTIVITE="'+USDateTime(StrToDate(StDate))+'"',TRUE);
  if Not QQ.EOF then iMax:=QQ.Fields[0].AsInteger+1 else iMax:=1;
  Ferme(QQ);
  T_ACTIVITE.PutValue ('ACT_NUMLIGNE',iMax);}

  T_ACTIVITE.PutValue ('ACT_TIERS',AffectTiers (Copy (StTot, 8,10)));
  T_ACTIVITE.PutValue ('ACT_AFFAIRE',champ);
  T_ACTIVITE.PutValue ('ACT_AFFAIRE0', Part0);
  T_ACTIVITE.PutValue ('ACT_AFFAIRE1', Part1);
  T_ACTIVITE.PutValue ('ACT_AFFAIRE2', Part2);
  T_ACTIVITE.PutValue ('ACT_AFFAIRE3', Part3);
  T_ACTIVITE.PutValue ('ACT_AVENANT', '00');
  T_ACTIVITE.PutValue ('ACT_ACTORIGINE', 'REP');
  T_ACTIVITE.putValue ('ACT_NUMLIGNEUNIQUE', ProchainIndiceAffaires (  T_ACTIVITE.GetValue('ACT_TYPEACTIVITE'),
     T_ACTIVITE.GetValue('ACT_AFFAIRE'), ListeAffaire, false, false)); // mcd 31/03/03

  If (stTot[59] = 'P') then
      begin
      T_ACTIVITE.PutValue('ACT_TYPEARTICLE','PRE');
      champ := debutpre;
      end
  Else If (stTot[59] = 'F') then
      begin
      T_ACTIVITE.PutValue('ACT_TYPEARTICLE','MAR');
      champ := debutfour;
     end
  Else If  (stTot[59] = 'f')then
      begin
      T_ACTIVITE.PutValue('ACT_TYPEARTICLE','FRA');
      champ := debutfrais;
     end;
  champ := champ + copy (stTot,60,15);
  T_ACTIVITE.PutValue ('ACT_ARTICLE', CodeArticleUnique2(champ, ''));
  T_ACTIVITE.PutFixedStValue ('ACT_CODEARTICLE', champ, 1,17, tcTrimr, false);
  T_ACTIVITE.PutFixedStValue ('ACT_LIBELLE', StTot, 220,70, tcTrim, false);

  QQ := OpenSQL('SELECT  GA_QUALIFUNITEACT,GA_ACTIVITEREPRISE From ARTICLE WHERE GA_CODEARTICLE="'+ trim(champ) +'"',TRUE,-1,'',true);
  If Not QQ.EOF then
     Begin
     T_ACTIVITE.PutValue ('ACT_UNITE',QQ.Fields[0].AsString);
     T_ACTIVITE.PutValue ('ACT_UNITEFAC',QQ.Fields[0].AsString);
     T_ACTIVITE.PutValue ('ACT_ACTIVITEREPRIS',QQ.Fields[1].AsString);
     End;
  Ferme(QQ);
    // cas de planning gére en soixantième d'heure
  If (HeurePlanning = 'S') then
     Begin
        // MCD non testé dans ce cas de boucle!!!!!
     St:=Copy(StTot,107,5);
     If (Pos('.',St) = 4) then heure := StrToInt(Copy (St,Pos('.',St)-4,1))
       else  heure := StrToInt(Copy (St,Pos('.',St)-5,2));
     heure:= heure + StrToInt(Copy (St,Pos('.',St)-2,2))*100/6000.0;
     T_ACTIVITE.PutValue ('ACT_QTE',heure);
     T_ACTIVITE.PutValue ('ACT_QTEFAC',heure);
     St:=Copy(StTot,136,5);
     If (Pos('.',St) = 4) then heure := StrToInt(Copy (St,Pos('.',St)-4,1))
       else  heure := StrToInt(Copy (St,Pos('.',St)-5,2));
     heure:= heure + StrToInt(Copy (St,Pos('.',St)-2,2))*100/6000.0;
     T_ACTIVITE.PutValue ('ACT_HEUREDEBUT',heure);
     St:=Copy(StTot,141,5);
     If (Pos('.',St) = 4) then heure := StrToInt(Copy (St,Pos('.',St)-4,1))
       else  heure := StrToInt(Copy (St,Pos('.',St)-5,2));
     heure:= heure + StrToInt(Copy (St,Pos('.',St)-2,2))*100/6000.0;
     T_ACTIVITE.PutValue ('ACT_HEUREFIN',heure);
     End
  Else
     Begin
     T_ACTIVITE.PutFixedStValue ('ACT_QTE', StTot, 107,5, tcDouble100, false);
     T_ACTIVITE.PutFixedStValue ('ACT_QTEFAC', StTot, 107,5, tcDouble100, false);
     T_ACTIVITE.PutFixedStValue ('ACT_HEUREDEBUT', StTot, 136,5, tcDouble100, false);
     T_ACTIVITE.PutFixedStValue ('ACT_HEUREFIN', StTot, 141,5, tcDouble100, false);
     End;

      //mcd 15/05/03 pour alimente unité ref
   if (T_ACTIVITE.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
      T_ACTIVITE.PutValue ('ACT_QTEUNITEREF', ConversionUnite (T_ACTIVITE.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, T_ACTIVITE.GetValue ('ACT_QTE')))
   else
     T_ACTIVITE.PutValue ('ACT_QTEUNITEREF', '0');
  T_ACTIVITE.PutValue ('ACT_DEVISE',V_PGI.DevisePivot);
  T_ACTIVITE.PutFixedStValue ('ACT_DATECREATION', StTot, 33,8, tcDate8AMJ, false);
  T_ACTIVITE.PutFixedStValue ('ACT_DATEMODIF', StTot, 33,8, tcDate8AMJ, false);
  //T_ACTIVITE.PutValue ('ACT_REALISABLE','X');

  If (trim(Copy(StTot,165,4)) <> '') Then T_ACTIVITE.PutFixedStValue ('ACT_FONCTIONRES', StTot, 165,4, tcTrimR, false)
  Else T_ACTIVITE.PutFixedStValue ('ACT_FONCTIONRES', StTot, 328,4, tcTrimR, false);

     // mcd 02/08/01 ajout initialisation zone par défaut
T_Activite.PutValue('ACT_ETATVISA', 'VIS');
T_Activite.PutValue('ACT_VISEUR', V_PGI.User);
T_Activite.PutValue('ACT_DATEVISA', NowH);
T_Activite.PutValue('ACT_ETATVISAFAC', 'VIS');
T_Activite.PutValue('ACT_VISEURFAC', V_PGI.User);
T_Activite.PutValue('ACT_DATEVISAFAC', NowH);

  ImportTob(T_ACTIVITE,True);
  T_ACTIVITE.Free;
  End;
End;
// Module Facturation
// Fournitures en attente de facturation
procedure TFAsRecupTempo.RecupFourAttente (StTot:string);
var T_ACTIVITE : TOB;
    Mois: integer;
    StDate:String;
    champ, ass ,Part0, Part1, Part2, Part3,aff,Art:string;
    Ressource:TAFO_Ressource;
    slFonction:TStringList;


Begin
 T_ACTIVITE := TOB.CREATE ('ACTIVITE', NIL, -1);
  T_ACTIVITE.InitValeurs;

  T_ACTIVITE.PutValue ('ACT_TYPEACTIVITE','REA');
  T_ACTIVITE.PutValue ('ACT_TYPERESSOURCE','SAL');
  If (stTot[40] = 'f')then begin
    ass := trim(AnsiUppercase(Copy(StTot,205,6))) ;
         // on recherche la focntion de l'assistant valide le jour de la ligne d'acitivté.
        // on ne traite pas les asisstant ***** utilisé pour les fournitures
    slFonction := nil;
    if (ass <> '******') then
      begin
        try
        Ressource := TAFO_Ressource.Create(AnsiUppercase(trim(ass)));
        if (Ressource<>nil) then
           if (Ressource.tob_Champs<>nil) then
              begin
              StDate:=copy (StTot,38,2)+ Copy(StTot,36,2)+Copy(StTot,34,2);
             slFonction := Ressource.FonctionDeLaRessource(Str6ToDate(StDate,90));
                        // ajout test mcd le 22/03/01
             if (slFonction<>nil) and (slFonction.count <>0) then T_ACTIVITE.PutValue('ACT_FONCTIONRES', slFonction[0]);
             end;
        finally
        if (slFonction<>nil) then   slFonction.Free;
        end;
       Ressource.Free;
       end;
    end
  else ass:= StringOfChar (' ',6);
  T_ACTIVITE.PutValue ('ACT_RESSOURCE',ass );
  T_ACTIVITE.PutValue ('ACT_FOLIO',1);
  T_ACTIVITE.PutFixedStValue ('ACT_DATEACTIVITE', StTot, 32,8, tcDate8AMJ, false);
  T_ACTIVITE.PutValue ('ACT_PERIODE',GetPeriode(T_ACTIVITE.GetValue('ACT_DATEACTIVITE')));
  T_ACTIVITE.PutValue ('ACT_SEMAINE',NumSemaine(T_ACTIVITE.GetValue('ACT_DATEACTIVITE')));
  If (stTot[40] = 'F') then
      begin
      T_ACTIVITE.PutValue('ACT_TYPEARTICLE','MAR');
      champ := debutfour;
      end
  Else If  (stTot[40] = 'f')then
       begin
       T_ACTIVITE.PutValue('ACT_TYPEARTICLE','FRA');
      champ := debutfrais;
      end;
  champ := champ + copy (stTot,41,15);
  T_ACTIVITE.PutValue ('ACT_ARTICLE', CodeArticleUnique2(champ, ''));
  T_ACTIVITE.PutFixedStValue ('ACT_CODEARTICLE', champ, 1,17, tcTrimr, false);
  T_ACTIVITE.PutFixedStValue ('ACT_LIBELLE', StTot, 56,70, tcTrim, false);
  If trim(champ) ='' then Art:=''
  else Art:=CodeArticleUnique2(champ, '');
  champ :=TraduitGA(format('Enrgt FOURNITURE : client %10.10s Affaire: %6.6s ',[AffectTiers(Copy (StTot, 8,10)),Copy(StTot, 18,6)]));
  If ctxscot  in V_PGI.PGIContexte then champ := champ + format('Exercice:%5.5s',[Copy(StTot,24,5)]);
  Champ:=champ +  TraduitGA(format(',Art %10.10s',[Art]));
               // mcd 07/02/02 ne traiteplus les fourn if (TestExistance(champ,Art,Ass,
  if (TestExistance(champ,Art,T_ACTIVITE.GetValue('ACT_TYPEARTICLE'),trim(Ass),
      Copy(StTot, 18,6), Copy(StTot,24,5),
      AffectTiers( Copy (StTot, 8,10)),'AFF','00',Part0,
      Part1, Part2, Part3,aff ,  mois)=False) then begin
     T_ACTIVITE.Free;
     exit;
     end;
  { mcd 31/03/03 QQ:=OpenSQL('SELECT MAX(ACT_NUMLIGNE) FROM ACTIVITE WHERE ACT_TYPEACTIVITE="REA" And ACT_AFFAIRE="'+ champ +'" And ACT_DATEACTIVITE="'+USDateTime(Str8ToDate(Copy(StTot,32,8),False))+'"',TRUE);
  if Not QQ.EOF then iMax:=QQ.Fields[0].AsInteger+1 else iMax:=1;
  Ferme(QQ);
  T_ACTIVITE.PutValue ('ACT_NUMLIGNE',iMax);  }
  T_ACTIVITE.PutValue ('ACT_DEVISE',V_PGI.DevisePivot);

  T_ACTIVITE.PutValue ('ACT_TIERS', AffectTiers (Copy (StTot, 8,10)));
  T_ACTIVITE.PutValue ('ACT_AFFAIRE', aff);
  T_ACTIVITE.PutValue ('ACT_AFFAIRE0', Part0);
  T_ACTIVITE.PutValue ('ACT_AFFAIRE1', Part1);
  T_ACTIVITE.PutValue ('ACT_AFFAIRE2', Part2);
  T_ACTIVITE.PutValue ('ACT_AFFAIRE3', Part3);
  T_ACTIVITE.PutValue ('ACT_AVENANT', '00');
  T_ACTIVITE.PutValue ('ACT_ACTORIGINE', 'REP');
  T_ACTIVITE.putValue ('ACT_NUMLIGNEUNIQUE', ProchainIndiceAffaires (  T_ACTIVITE.GetValue('ACT_TYPEACTIVITE'),
     T_ACTIVITE.GetValue('ACT_AFFAIRE'), ListeAffaire, false, false)); // mcd 31/03/03

  If (stTot[40] = 'P') then T_ACTIVITE.PutFixedStValue ('ACT_UNITE', StTot, 146,1, tcChaine, false);
  T_ACTIVITE.PutFixedStValue ('ACT_QTE', StTot, 147,7, tcDouble100, false);
    //mcd 15/05/03 pour alimente unité ref
  if (T_ACTIVITE.GetValue ('ACT_TYPEARTICLE') = 'PRE') then
    T_ACTIVITE.PutValue ('ACT_QTEUNITEREF', ConversionUnite (T_ACTIVITE.GetValue ('ACT_UNITE'), VH_GC.AFMESUREACTIVITE, T_ACTIVITE.GetValue ('ACT_QTE')))
  else
    T_ACTIVITE.PutValue ('ACT_QTEUNITEREF', '0');
  If (stTot[40] = 'P') then T_ACTIVITE.PutFixedStValue ('ACT_UNITEFAC', StTot, 146,1, tcChaine, false);
  T_ACTIVITE.PutFixedStValue ('ACT_QTEFAC', StTot, 147,7, tcDouble100, false);
  T_ACTIVITE.PutValue ('ACT_DEVISE',V_PGI.DevisePivot);
  // Gestion des prix de revient. On met la même valeur de partout
  T_ACTIVITE.PutFixedStValue ('ACT_PUPR', StTot, 154,12, tcDouble100, false);
  T_ACTIVITE.PutFixedStValue ('ACT_PUPRCHARGE', StTot, 154,12, tcDouble100, false);
  T_ACTIVITE.PutFixedStValue ('ACT_PUPRCHINDIRECT', StTot, 154,12, tcDouble100, false);
  T_ACTIVITE.PutFixedStValue ('ACT_TOTPR', StTot, 178,12, tcDouble100, false);
  T_ACTIVITE.PutFixedStValue ('ACT_TOTPRCHARGE', StTot, 178,12, tcDouble100, false);
  T_ACTIVITE.PutFixedStValue ('ACT_TOTPRCHINDI', StTot, 178,12, tcDouble100, false);

  T_ACTIVITE.PutFixedStValue ('ACT_PUVENTE', StTot, 166,12, tcDouble100, false);
  T_ACTIVITE.PutFixedStValue ('ACT_TOTVENTE', StTot, 190,12, tcDouble100, false);
  T_ACTIVITE.PutFixedStValue ('ACT_PUVENTEDEV', StTot, 166,12, tcDouble100, false);
  T_ACTIVITE.PutFixedStValue ('ACT_TOTVENTEDEV', StTot, 190,12, tcDouble100, false);

    // mcd 30/01/03 pour résoudre le pb de qtefac=0 et prix # 0 . sinon facture temps passé pas ok
  if (T_ACTIVITE.Getvalue('ACT_QTE')= 0) and (T_ACTIVITE.Getvalue('ACT_QTEFAC')= 0)
      and (T_ACTIVITE.Getvalue('ACT_TOTVENTE')<>0)  then
      begin
      T_ACTIVITE.PutValue('ACT_QTE',1);
      T_ACTIVITE.PutValue('ACT_QTEFAC',1);
      if (T_ACTIVITE.GetValue ('ACT_TYPEARTICLE') = 'PRE') then T_ACTIVITE.PutValue('ACT_QTEUNITEREF',1); //mcd 29/07/03
      end;

  if (sttot[202] = '*') then  T_ACTIVITE.PutValue ('ACT_ACTIVITEREPRIS', 'FAC')
  else   if (sttot[202] = 'R') or (sttot[202] = 'W') then  T_ACTIVITE.PutValue ('ACT_ACTIVITEREPRIS', 'A')
  else   if (sttot[202] = 'A') then  T_ACTIVITE.PutValue ('ACT_ACTIVITEREPRIS', 'F')       // on perd la notion d'AN ..
  else T_ACTIVITE.PutFixedStValue ('ACT_ACTIVITEREPRIS', StTot, 202,1, tcChaine, false);
  T_ACTIVITE.PutFixedStValue ('ACT_DATECREATION', StTot, 32,8, tcDate8AMJ, false);
  T_ACTIVITE.PutFixedStValue ('ACT_DATEMODIF', StTot, 32,8, tcDate8AMJ, false);
  //T_ACTIVITE.PutValue ('ACT_REALISABLE','X');
                                          
    // mcd 02/08/01 ajout initialisation zone par défaut
T_Activite.PutValue('ACT_ETATVISA', 'VIS');
T_Activite.PutValue('ACT_VISEUR', V_PGI.User);
T_Activite.PutValue('ACT_DATEVISA', NowH);
T_Activite.PutValue('ACT_ETATVISAFAC', 'VIS');
T_Activite.PutValue('ACT_VISEURFAC', V_PGI.User);
T_Activite.PutValue('ACT_DATEVISAFAC', NowH);
  ImportTob(T_ACTIVITE,True);
  T_ACTIVITE.Free;

End;

procedure TFAsRecupTempo.BStopClick(Sender: TObject);
begin
  inherited;
StopRecup := True;
Application.ProcessMessages;
writeln (FichierLog, '****** Traitement interrompu par l''utilisateur *****');
PGIBoxAf('Traitement interrompu par l''utilisateur',TitreHalley);
end;

procedure TFAsRecupTempo.BParamClientClick(Sender: TObject);
begin
  inherited;
FParamCli.ShowModal;
end;

procedure TFAsRecupTempo.BParamFJurClick(Sender: TObject);
begin
  inherited;
FParamFjur.ShowModal;
end;
procedure TFAsRecupTempo.BEntiteClick(Sender: TObject);
begin
  inherited;
FParamEntite.ShowModal;
end;

procedure TFAsRecupTempo.FicTempsDateExit(Sender: TObject);
begin
  inherited;
If ((Not IsValidDate(FicTempsDate.Text)) And (FicTempsDate.Text<>'')) Then
   Begin
   PGIBoxAf('Date Invalide', TraduitGa('Reprise des données Temps Tempo'));
   TWinControl(FicTempsDate).SetFocus;
   End;
end;

procedure TFAsRecupTempo.FicTempsDateEffExit(Sender: TObject);
begin
  inherited;
If ((Not IsValidDate(FicTempsDateEff.Text)) And (FicTempsDateEff.Text<>'')) Then
   Begin
   PGIBoxAf('Date Invalide', TraduitGa('Pour Effacement des données Temps Tempo'));
   TWinControl(FicTempsDateEff).SetFocus;
   End;
end;

procedure TFAsRecupTempo.FicFacDateEffExit(Sender: TObject);
begin
  inherited;
If ((Not IsValidDate(FicFacDateEff.Text)) And (FicFacDateEff.Text<>'')) Then
   Begin
   PGIBoxAf('Date Invalide', TraduitGa('Pour Effacement des données stat Tempo'));
   TWinControl(FicFacDateEff).SetFocus;
   End;
end;

procedure TFAsRecupTempo.DatePlanningExit(Sender: TObject);
begin
  inherited;
If ((Not IsValidDate(DatePlanning.Text)) And (Dateplanning.Text<>'')) Then
   Begin
   PGIBoxAf('Date Invalide', TraduitGa('Reprise des données Planning Tempo'));
   TWinControl(Dateplanning).SetFocus;
   End;
end;
       // cette fct permet, en fct du paramétrage de rendre le
       // code affaire PGI correct.
       // Si création = TRUE, calcule le compteur, si False, accède
       // au fichier affaire pour connaitre le code exact de l'affaire
       // rend aussi le mois de clôture trouver en cas de gestion exercice par mois
       // pour éviter relecture client dans fichier affaire.
function TFAsRecupTempo.AffectAff (Aff, ZExer, Clt,Typ,Avenant:string; Var Part0,Part1, Part2, Part3 : STring; Creation : Boolean ;Var Moisclot:Integer):string ;
var   QQ : TQuery;
    champ, mois :string;
Begin
//Le paramétrage du code affaire, en cas de reprise de fichiers est figé :
//TEMPO : compteur et une seule partie dans le code
//SCOT : code à 3 partie si gestion exercice : liste définie + exercice + n°
//ou code à 2 parties : liste définie + compteur (n'est plus gérée !!!!) exercice obligatoire

if (typ = 'PRO') then part0 :='P' else Part0 :='A';
Part1 :='';
Part2:='';
Part3:='';
if avenant = '  ' then avenant :='00';
if avenant = '' then avenant :='00';
If ctxTempo in V_PGI.PGIContexte then
    begin
        // cette fct pour TEMPO, prend le code affaire de TEMPO sur la longueur définie en PGI
{$ifdef OMI}
          part1:='REP';
          part2:=Copy (aff,2,VH_GC.CleAffaire.Co2Lng );    // sur 5 car
{$else}
    part1:=Copy (aff,1,VH_GC.CleAffaire.Co1Lng );
{$endif}
    Result:= RegroupePartiesAffaire (Part0,Part1, Part2, Part3, Avenant);
    if creation = False then begin
       // cas de fichier autre que affaire. on regarde si enrgt exsiet
       // dans ficheir affaire. sinon on ne traite pas enrgt
        champ :=  'SELECT AFF_AFFAIRE, AFF_AFFAIRE2, AFF_AFFAIRE3 FROM AFFAIRE WHERE AFF_TIERS = "'+ Clt +'" AND AFF_AFFAIRE = "'+ result+'"'   ;
        QQ := OpenSQL (champ,TRUE,-1,'',true);
        if QQ.EOF then
            Begin
            result:='';
            end ;
        Ferme(QQ);
        end
    else begin
              //on regarde si client existe, sinon, erreur
        QQ := OpenSQL ('SELECT T_TIERS,T_EURODEFAUT FROM TIERS WHERE T_TIERS ="'+ Clt+'"',TRUE,-1,'',true);
        if QQ.EOF then
           Begin
             champ:=TraduitGA(format('Enrgt Affaire non traité : client %10.10s Affaire: %6.6s ',[clt,aff])) + TraduitGa(' Client non trouvé');
             if creation = True then writeln (FichierLog, champ);
             result:='';
             Ferme (QQ);
             exit;
             end 
        else EuroDefaut:= QQ.Fields[1].AsString;
        Ferme (QQ);
         end;
    end
else //cas SCOT
    begin
        // la partie 1 doit obligatoireement être le code mission (en table)
    part1:=Copy (aff,1,VH_GC.CleAffaire.Co1Lng );
        // la partie 2 doit être renseignée en fct du paramétrage exercice
        // le format exercice ne peut fonctionner que si le choix est saisie libre.
    if (exer = 'N') then zexer:=Exercice.text;
    if (trim(zexer) = '') or (copy(zexer,1,2)='**') then
        begin
        zexer:=Exercice.text;
        end;
    Mois :='12';
    MoisClot :=12;
    QQ := OpenSQL ('SELECT T_MOISCLOTURE,T_EURODEFAUT FROM TIERS WHERE T_TIERS ="'+ Clt +'"',TRUE,-1,'',true);
    if Not QQ.EOF then
       Begin
       Mois := QQ.Fields[0].AsString;
       EuroDefaut:= QQ.Fields[1].AsString;
       MoisClot:= StrtoInt (Mois);
       if (MoisClot < 1) or (Moisclot >12) then MoisClot:=12;
       Mois := Format ('%2.2d',[MoisClot]);
       end
    else begin
         champ:=TraduitGA(format('Enrgt Affaire non traité : client %10.10s Affaire: %6.6s Exercice: %5.5s',[clt,aff,Zexer])) + TraduitGa(' Client non trouvé');
         if creation = True then writeln (FichierLog, champ);
         result:='';
         Ferme (QQ);
         exit;
         end;
    Ferme (QQ);
    if (VH_GC.CleAffaire.Co2Type <> 'SAI') then  part2:=Zexer
    else begin
     if (VH_GC.AFFormatExer = 'AUC') then  part2:=Zexer
        else if (VH_GC.AFFormatExer = 'AM') then
            begin
            if (Exer = 'N') then
                begin
                if MoisCLot < 12 then Zexer := ExerciceDec.text else Zexer := Exercice.text;
                end;
            if (Copy (ZExer,4,2) > '50') then Part2 := '19' + Copy (ZExer, 4,2)
                else  Part2 := '20' + Copy (ZExer, 4,2);
            Part2 := Part2 + Mois;
            end
          else if (VH_GC.AFFormatExer = 'AA') then
            begin
              // cas de génération exercice sous la forme DDDDFFFF : année début  et  fin
            if (Exer = 'N') then
               begin
               if MoisCLot < 12 then Zexer := ExerciceDec.text else Zexer := Exercice.text;
               end;
            if (Copy (ZExer,1,2) > '50') then Part2 := '19' + Copy (ZExer, 1,2)
                  else  Part2 := '20' + Copy (ZExer, 1,2);
            if (Copy (ZExer,4,2) > '50') then Part2 := Part2 +'19' + Copy (ZExer, 4,2)
                  else  Part2 := Part2 + '20' + Copy (ZExer, 4,2);

             end
           else if (VH_GC.AFFormatExer = 'A') then
            begin
                // cas ou génération millesime seulement
             if (Exer = 'N') then Zexer := ExerciceDec.text ;
             if (Copy (ZExer,4,2) > '50') then Part2 := '19' + Copy (ZExer, 4,2)
                else  Part2 := '20' + Copy (ZExer, 4,2);
           end;
        end;
       // cas de création de code, on incrémente la clé
    if (Creation = TRUE) then
       begin
       if VH_GC.CleAffaire.NbPartie =2 then
           begin
              // la partie 2 est obligatoirement un compteur
            if (typ = 'PRO')  and (GetParamSoc ('So_AFFPRODIFFERENT'))
             then begin
                 part2:= CodeAffaireSuivant (GetParamSoc('SO_AFFCO2VALEURPRO'),VH_GC.CleAffaire.Co2Lng) ;
                 if (Part2 = '') then begin result:=''; exit; end;
                 SetParamSoc('SO_AFFCO2VALEURPRO', part2) ; // GetparamSoc a supprimer
                // A remettre GEtParamSOc VH_GC.CleAffaire.Co2valeurPro := Part2  ;
                end
             else begin
                part2:= CodeAffaireSuivant (VH_GC.CleAffaire.CO2VALEUR,VH_GC.CleAffaire.Co2Lng) ;
                 if (Part2 = '') then begin result:=''; exit; end;
                VH_GC.CleAffaire.Co2valeur := part2;
                end;
            end
       else begin
            // la partie 2 est obligatoirement l'exercice et la partie 3 un compteur
            if (typ = 'PRO')  and (GetParamSoc ('So_AFFPRODIFFERENT'))
             then begin
                 part3:= CodeAffaireSuivant (GetParamSoc('SO_AFFCO3VALEURPRO'),VH_GC.CleAffaire.Co3Lng) ;
                 if (Part3 = '') then begin result:=''; exit; end;
                 SetParamSoc('SO_AFFCO3VALEURPRO', part3) ; // GetparamSoc a supprimer
                // A remettre GEtParmaSOc VH_GC.CleAffaire.Co3valeurPro := Part3 ;
                end
             else begin
                part3:= CodeAffaireSuivant (VH_GC.cleaffaire.CO3VALEUR,VH_GC.CleAffaire.Co3Lng) ;
                if (Part3 = '') then begin result:=''; exit; end;
                VH_GC.CleAffaire.Co3valeur := part3;
                end;
         end;
        Result:= RegroupePartiesAffaire (Part0,Part1, Part2, Part3, Avenant);
       end
            // cas d'appel du fichier pour avoir le n° exact affaire
       else begin
        champ :=  'SELECT AFF_AFFAIRE, AFF_AFFAIRE2, AFF_AFFAIRE3 FROM AFFAIRE WHERE AFF_TIERS = "'+ Clt +'" AND AFF_AFFAIRE1 = "'+ Part1 +'" AND AFF_STATUTAFFAIRE = "'+ Typ +'"'   ;
        if VH_GC.CleAffaire.NbPartie =3 then champ := champ + 'AND AFF_AFFAIRE2 = "'+ Part2+'"';
        QQ := OpenSQL (champ,TRUE,-1,'',true);
        if Not QQ.EOF then
            Begin
            result:=QQ.Fields[0].AsString;
            Part2:=QQ.Fields[1].AsString;
            Part3:=QQ.Fields[2].AsString;
            end
        else begin
                // cas ou mission non trouvée,
            Result:= '';
            end;
        Ferme(QQ);
        end;
    end;
end;

procedure TFAsRecupTempo.TabPlanningExit(Sender: TObject);
begin
  inherited;
    DatePlanningExit(Sender);
end;

procedure TFAsRecupTempo.TabTempsExit(Sender: TObject);
begin
  inherited;
 FicTempsDateExit (Sender);
end;


procedure TFAsRecupTempo.PChange(Sender: TObject);
begin
  inherited;
        //fct qui permet de surcharger la fct PChange, afin d'afficher le n°
        // réelles de pages, en fct du nbre de fichier traiter.
        // En Entrant dans la fct, affiche sustématique page 1/3, car il y a toujours
        // 3 onglets au minimum géré. Le nbre fin est ensuite actualisé en fct des
        // fichiers à traiter.
   if (NbPage =0) then NbPage := 3;
   If (NoPage =0) then Nopage :=1;
   lEtape.Caption := Msg.Mess[0] + ' ' + IntToStr(Nopage) + '/' + IntToStr(NbPage) ;
end;

       // fct générique qui met le code tiers sur le bon nombre de caractères
function TFAsRecupTempo.AffectTiers ( Clt:string):string ;
begin
{ le code tiers est d'une longueur libre
champ := StringOfChar (VH^.Cpta[fbgene].Cb, Lng_gene);
result:= Copy ( (trim(clt) + champ) ,1,lng_gene); }
result:= trim(clt);
end;

      // fct qui permet d'alimenter la piece associée à l'affaire
procedure TFAsRecupTempo.RecupPieceAffaire (StTot:string;CleDocAffaire : R_CLEDOC; MonPres:Boolean);
var T_PIECE : TOB;
    QQ :TQuery;
    ModeRempl ,champ:string;
    nbech, duree, jour, separe , MontantMin : integer;
Begin
T_PIECE := TOB.CREATE ('PIECE', NIL, -1);
QQ:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(CleDocAffaire,ttdPiece,False),True,-1,'',true) ;
If Not QQ.EOF then T_PIECE.SelectDB('',QQ)
else begin
    Ferme(QQ) ;
    T_PIECE.free;  //mcd 22/03/01
    exit;
    end ;

Ferme(QQ) ;

if (Monpres <> VH^.TenueEuro) then T_PIECE.PutValue ('GP_SAISIECONTRE', 'X')
else T_PIECE.PutValue ('GP_SAISIECONTRE', '-') ;
T_PIECE.PutValue ('GP_EDITEE', 'X') ;
           // affecte le total des lignes
T_PIECE.PutFixedStValue ('GP_TOTALHT',StTot, 803,12, tcDouble100, false);
T_PIECE.PutFixedStValue ('GP_TOTALBASEREM',StTot, 803,12, tcDouble100, false);
T_PIECE.PutFixedStValue ('GP_TOTALBASEESC',StTot, 803,12, tcDouble100, false);
T_PIECE.PutFixedStValue ('GP_TOTALTTC',StTot, 1132,12, tcDouble100, false);

    // les mtt devise, sont égale à ce qui est en fichier SCOT si monnaie client = monnaie tenue
    // convertit sinon
If (VH^.TenueEuro = MonPres ) then begin
  T_PIECE.PutFixedStValue ('GP_TOTALHTDEV',StTot, 803,12, tcDouble100, false);
  T_PIECE.PutFixedStValue ('GP_TOTALBASEREMDEV',StTot, 803,12, tcDouble100, false);
  T_PIECE.PutFixedStValue ('GP_TOTALBASEESCDEV',StTot, 803,12, tcDouble100, false);
  T_PIECE.PutFixedStValue ('GP_TOTALTTCDEV',StTot, 1132,12, tcDouble100, false);
  end
else begin
  T_PIECE.PutValue ('GP_TOTALHTDEV',COnvertMtt(Copy(StTot, 803,12), Sttot[1571]));
  T_PIECE.PutValue ('GP_TOTALBASEREMDEV',COnvertMtt(Copy(StTot, 803,12), Sttot[1571]));
  T_PIECE.PutValue ('GP_TOTALBASEESCDEV',COnvertMtt(Copy(StTot, 803,12), Sttot[1571]));
  T_PIECE.PutValue ('GP_TOTALTTCDEV',COnvertMtt(Copy(StTot,1132,12), '3'));
  end;

if (DebEnc = 'E') then T_PIECE.PutValue('GP_TVAENCAISSEMENT','TE')
else    T_PIECE.PutValue('GP_TVAENCAISSEMENT','TD');
        // affect RIB
T_PIECE.PutValue ('GP_RIB',EncodeRib (Trim (Copy (StTot,509,5)), Trim (Copy (StTot,514,5)),
    Trim (Copy (StTot,519,12)),Trim (Copy (StTot,530,2)),Trim (Copy (StTot,485,24))));

        // création mode de reglmt
champ :=Copy (Sttot, 472,1);
if (champ = 'A') then  Nbech  := 10 else
if (champ = 'B') then  Nbech  := 11 else
if (champ = 'C') then  Nbech  := 12 else
if (champ = 'D') then  Nbech  := 13 else
if (champ = 'E') then  Nbech  := 14 else
if (champ = 'F') then  Nbech := 15 else nbech := StrtoInt (Copy (Sttot, 472,1));
separe:= StrtoInt (Copy (Sttot, 473,3));
duree := StrtoInt (Copy (Sttot, 476,3));
jour := StrtoInt (Copy (Sttot, 479,2));
MontantMin := 0;
ModeRempl :='';
If ((Sttot [469] = 'T') or (Sttot [469] = 'L') or (Sttot [469] = 'P')) then
    begin
    MontantMin := MttMin;
    Moderempl := MRRempl;
    end;
T_PIECE.PutValue ('GP_MODEREGLE', RecupModeRegl (Copy (Sttot, 471,1), Copy (Sttot, 469,2),
    ModeRempl, duree, jour, nbech, separe, MontantMin));

T_PIECE.Putvalue ('GP_TIERSLIVRE',T_PIECE.GetValue('GP_TIERS'));
T_PIECE.Putvalue ('GP_TIERSFACTURE',T_PIECE.GetValue('GP_TIERS'));
T_PIECE.Putvalue ('GP_TIERSPAYEUR',T_PIECE.GetValue('GP_TIERS'));
T_PIECE.PutValue ('GP_PERIODE',GetPeriode(T_PIECE.GetValue('GP_DATEPIECE')));
T_PIECE.PutValue ('GP_SEMAINE',NumSemaine(T_PIECE.GetValue('GP_DATEPIECE')));

T_PIECE.PutValue('GP_CREEPAR','IMP');
ImportTob(T_PIECE,True);
T_PIECE.Free;
end;

procedure TFAsRecupTempo.RecupModePaie (StTot:string);
var T_MODEPAIE : TOB;

Begin
// Alimentation de la table mode paiement de PGI, à partir
// table mode règlement

T_MODEPAIE := TOB.CREATE ('MODEPAIE', NIL, -1);
T_MODEPAIE.InitValeurs;
T_MODEPAIE.PutFixedStValue ('MP_MODEPAIE', StTot, 8,2, tcTrimr, false);
T_MODEPAIE.PutFixedStValue ('MP_LIBELLE', StTot, 14,35, tcTrimr, false);
T_MODEPAIE.PutFixedStValue ('MP_ABREGE', StTot, 14,17, tcTrimr, false);
T_MODEPAIE.PutValue ('MP_POINTABLE', 'X');
// mode de règlement à encaissement et décaissemnt ==> mixte
T_MODEPAIE.PutValue ('MP_ENCAISSEMENT' ,'MIX');
// on regarde si impression traite/LCR sur papier préimprimé
if ((StTot[8] = 'T') and (ImpTra <>'N')) then T_MODEPAIE.PutValue ('MP_LETTRETRAITE', 'X');
if ((StTot[8] = 'L') and (ImpLCR <>'N')) then T_MODEPAIE.PutValue ('MP_LETTRETRAITE','X');
// dans le cas de traite LCR, BO, il faut renseigner l'option soumis on non à l'acceptation
if (Copy (Sttot,8,2) = 'BO') then  T_MODEPAIE.PutValue ('MP_CODEACCEPT','BOR');
if ((Copy (Sttot,8,2) = 'LN') or(Copy (Sttot,8,2) = 'TN') or(Copy (Sttot,8,2) = 'LW') or(Copy (Sttot,8,2) = 'TW') )
    then  T_MODEPAIE.PutValue ('MP_CODEACCEPT','TRA');
if ((Copy (Sttot,8,2) = 'LV') or(Copy (Sttot,8,2) = 'TV') or(Copy (Sttot,8,2) = 'LR') or(Copy (Sttot,8,2) = 'TR') )
    then  T_MODEPAIE.PutValue ('MP_CODEACCEPT','NON');
// on alimente la catégorie de paiement.
// si non connu, mis CHQ
if ((StTot[8] = 'B') or(StTot[8] = 'T') or(StTot[8] = 'L') )
    then  T_MODEPAIE.PutValue ('MP_CATEGORIE','LCR');
if (Copy (Sttot,8,2) = 'PR') then  T_MODEPAIE.PutValue ('MP_CATEGORIE','PRE');
if (Copy (Sttot,8,2) = 'VI') then  T_MODEPAIE.PutValue ('MP_CATEGORIE','VIR');
if ((Copy (Sttot,8,2) = 'CH') or(Copy (Sttot,8,2) = 'CR') ) then  T_MODEPAIE.PutValue ('MP_CATEGORIE','CHQ');


ImportTob(T_MODEPAIE,True);
T_MODEPAIE.Free;
End;

procedure TFAsRecupTempo.RecupRib (StTot, Clt:string; condi : boolean);
var T_RIB : TOB;
    QQ :TQuery;
    etb, guich, cpt, cle : string;
    max : integer;
Begin
// Alimentation de la table RIB. à parti client ou affaire

Etb := Copy (StTot,  25,5);
guich := Copy (StTot , 30,5);
cpt := Copy (StTot,  35,11);
cle := Copy (StTot,  46,2);
// PL le 09/01/02 pour optimiser le nombre d'enregistrements lus en eagl
QQ:=OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="' + clt+ '" and R_ETABBQ="' + etb+
 '" and R_GUICHET="' + guich+ '" and R_NUMEROCOMPTE="' + cpt+ '" and R_CLERIB="' + cle+ '"' ,True,1,'',true) ;
//QQ:=OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="' + clt+ '" and R_ETABBQ="' + etb+
// '" and R_GUICHET="' + guich+ '" and R_NUMEROCOMPTE="' + cpt+ '" and R_CLERIB="' + cle+ '"' ,True) ;
if  QQ.EOF
    then begin   // le RIB n'existe pas, on le crée
        Ferme(QQ) ;
        QQ:=OpenSQL('SELECT MAX(R_NUMERORIB) FROM RIB WHERE R_AUXILIAIRE="' + clt+ '"'  ,TRUE,-1,'',true);
          if Not QQ.EOF then Max:=QQ.Fields[0].AsInteger+1 else Max:=1;
        T_RIB := TOB.CREATE ('RIB', NIL, -1);
        T_RIB.InitValeurs;
        T_RIB.PutValue ('R_AUXILIAIRE', Clt);
        T_RIB.PutValue ('R_NUMERORIB', Max);
        if (condi = TRUE) then T_RIB.PutValue ('R_PRINCIPAL', 'X')
            else T_RIB.PutValue ('R_PRINCIPAL', '-') ;
        T_RIB.PutValue ('R_ETABBQ', etb);
        T_RIB.PutValue ('R_GUICHET', guich);
        T_RIB.PutValue ('R_NUMEROCOMPTE',cpt);
        T_RIB.PutValue ('R_CLERIB',  cle);
        T_RIB.PutFixedStValue ('R_DOMICILIATION', StTot, 1,24, tcTrimr, false);
        T_RIB.PutValue ('R_DEVISE',V_PGI.DevisePivot);

        ImportTob(T_RIB,True);
        T_RIB.Free;
        end;
Ferme(QQ) ;

End;


Procedure TFAsRecupTempo.ChangeClientFacture ;
Var Q : TQuery;
    facture, auxi, retour : String;
    Tobtrav, tobdet : Tob;
    i : integer;
Begin
    // cette fct permet de relire le ficheir tiers, pour voir si il
    // existe des renseignements dna sla zone T_FACTURE.
    // si c'est le cas, on change cette valeur = T_TIERs, par la valeur
    // T_AUXILIAIRE correspondante
Q:=OPENSQL('SELECT T_FACTURE, T_AUXILIAIRE From TIERS WHERE T_FACTURE <> ""',true,-1,'',true);
If Not Q.EOF Then begin
    TobTrav:=TOB.Create('liste des tiers',Nil,-1) ;
    TobTrav.LoadDetailDB('ChampsTB TIERS','','',Q,False,False);
    for i:=0 to TOBTrav.Detail.Count-1 do
        BEGIN
        TOBDet:=TOBTrav.Detail[i] ;
        facture := TobDet.GetValue('T_FACTURE');
        Auxi := TobDet.GetValue('T_AUXILIAIRE');
        Retour := TiersAuxiliaire (facture, False);
        ExecuteSQL('UPDATE TIERS SET T_FACTURE = "' + Retour + '" WHERE T_AUXILIAIRE = "' + Auxi +'" ');
        End;
    TobTrav.free; //mcd 22/03/01
    end;

Ferme(Q);

End;

{***********A.G.L.***********************************************
Auteur  ...... : DESSEIGNET
Créé le ...... : 13/04/2000
Modifié le ... :   /  /    
Description .. : Permet reprise table moderegl
Mots clefs ... : MOREREGL
*****************************************************************}
function TFAsRecupTempo.RecupModeRegl(FinMois, ModePaie, Cheque:string; Duree, Jour, Nbech, Separe, MontMin:integer ):string ;
var T_MODEREGL : TOB;
    Q : TQuery;
    max, Apartirde, Separepar,Arrondijour,  Snbech, Sduree, champ ,Ecartjour : string;
    ii ,reste : integer;
    taux : double;
Begin
    { Cette fonction permet de rendre le mode de règlement associé
    aux modes gérés dans l'ancienne gamme.
    Si le mode n'existe pas, le crée.
    FinMois : O/N si calcul fin de mois ou pas
    ModePaie : CH, TR, LR ... Doit être en lien avec table MODEPAIE
    Durée : à X jours (30, 60 ...)
    Jour : le X (jour de l'échénace)
    Nbech : nbre d'échéance
    séparé : nbre de jours entre  les échéances }

    // on accède d'abord à la table pour voir si le mode de règlement existe déjà
Sduree:= InttoStr (duree);
Snbech:= InttoStr (nbech);
if (FinMois = 'O') then Apartirde := 'FIN' else Apartirde :='FAC';
case jour of
0 : Arrondijour := 'PAS';
1 : Arrondijour := 'DEB';
5 : Arrondijour := '05M';
10 : Arrondijour := '10M';
15 : Arrondijour := '15M';
20: Arrondijour := '20M';
25: Arrondijour := '25M';
30: Arrondijour := 'FIN';
31: Arrondijour := 'FIN';
else  Arrondijour := 'PAS';
end ;

ii:= separe div 30;
reste := separe mod 30;
champ :='';
Ecartjour :='';
if (reste = 0) then champ := format ('%dM',[ii])
    else begin
    if (separe = 15) then champ :='QUI'
        else if (separe = 7) then champ := 'SEM'
            else begin
                For ii := 1 to nbech do begin
                  EcartJour := EcartJour + format ('%d;',[separe]);                    end
                 end;
    end;
SeparePar := champ;
if (Nbech = 1) then SeparePar := 'QUI';
if Nbech > 12 then Nbech := 12;


Q:=OPENSQL('SELECT MR_MODEREGLE From MODEREGL WHERE MR_APARTIRDE = "' + Apartirde +
    '" AND MR_PLUSJOUR= ' + sduree + ' And MR_ARRONDIJOUR = "' + Arrondijour +
    '" And MR_NOMBREECHEANCE = ' + sNbech + ' And MR_SEPAREPAR = "'+ Separepar +
    '" And MR_MP1 = "' + ModePaie + '"' ,true,-1,'',true);
If Not Q.EOF Then begin
    result :=  Q.Fields[0].asstring;
    end
Else begin
        // il faut créer le nouveau mode. on garde en memoire les n°
        // affecté aux modes de règlements. On regarde si existe avant
        // de le créer
    Ferme(Q);
    Max := Format ('%3.3d',[(NBMode+1)]);
    Q:=OpenSQL('SELECT MR_MODEREGLE FROM MODEREGL WHERE MR_MODEREGLE = "'+ max+ '"',TRUE,-1,'',true) ;
    while (Not Q.EOF) do begin
           Ferme(Q);
           Inc (Nbmode);
           Max := Format ('%3.3d',[(NBMode+1)]);
           Q:=OpenSQL('SELECT MR_MODEREGLE FROM MODEREGL WHERE MR_MODEREGLE = "'+ max +'"',TRUE,-1,'',true) ;
           end ;
    T_MODEREGL := TOB.CREATE ('MODEREGL', NIL, -1);
    T_MODEREGL.InitValeurs;

    T_MODEREGL.PutValue ('MR_MODEREGLE', max);
    champ := Format ('%s à %d jours le %d, Fin mois %s, %d échéances séparées de %d jours', [ModePaie, duree, jour, Finmois,nbech, separe]) ;
    T_MODEREGL.PutValue ('MR_LIBELLE', champ);
    T_MODEREGL.PutValue ('MR_ABREGE', copy(champ,1,17));
    T_MODEREGL.PutValue ('MR_APARTIRDE', Apartirde);
    T_MODEREGL.PutValue ('MR_PLUSJOUR', duree);
    T_MODEREGL.PutValue ('MR_ARRONDIJOUR', Arrondijour);
    T_MODEREGL.PutValue ('MR_NOMBREECHEANCE', Nbech);
    T_MODEREGL.PutValue ('MR_SEPAREPAR', Separepar);
    T_MODEREGL.PutValue ('MR_ECARTJOURS', EcartJour);
    T_MODEREGL.PutValue ('MR_REMPLACEMIN', Cheque);
    T_MODEREGL.PutValue ('MR_MONTANTMIN', MontMin);
    For ii := 1 to Nbech do begin
        T_MODEREGL.PutValue ('MR_MP'+ IntToStr (ii), Modepaie);
        taux := (100 div nbech);
        if (ii = Nbech) then taux := (100 div nbech) + (100 mod nbech) ;
        T_MODEREGL.PutValue ('MR_TAUX'+ IntToStr (ii), taux);
        T_MODEREGL.PutValue ('MR_ESC'+ IntToStr (ii), '-');
        end;
    ImportTob(T_MODEREGL,True) ;
    T_MODEREGL.Free;
    Result := max;
    end;
Ferme(Q);
End;

function TFAsRecupTempo.ConvertMtt (Mtt,ecart:string ):double;
var ii : double;
Begin
ii := 0;
If (ecart = '0') then  ii := -0.03;
If (ecart = '1') then  ii := -0.02;
If (ecart = '2') then  ii := -0.01;
If (ecart = '3') then  ii := 0;
If (ecart = '4') then  ii := 0.01;
If (ecart = '5') then  ii := 0.02;
If (ecart = '6') then  ii := 0.03;

if (VH^.TenueEuro) then  Result := (EuroToFranc((StrToFLoat (mtt)/100.0)) + ii)
else Result := FrancToEuro((StrToFLoat (mtt)/100.0));
end;

{function TFAsRecupTempo.ConvertMttBis (Mtt,ecart:string ):double;
var ii : double;
Begin
     // idem convertmtt mais sans arrondi des mtt
ii := 0;
If (ecart = '0') then  ii := -0.03;
If (ecart = '1') then  ii := -0.02;
If (ecart = '2') then  ii := -0.01;
If (ecart = '3') then  ii := 0;
If (ecart = '4') then  ii := 0.01;
If (ecart = '5') then  ii := 0.02;
If (ecart = '6') then  ii := 0.03;

if (VH^.TenueEuro) then  Result := ((StrToFLoat (mtt)/100.0)* V_PGI.TauxEuro) + ii
else Result := ((StrToFLoat (mtt)/100.0))/ V_PGI.TauxEuro;
end; }

Procedure TFAsRecupTempo.MajPieceAff();
var T_PIECE : TOB;
    QQ :TQuery;
    totbis , taxebis, taxedev : double;
Begin
    //fct qui met à jour total TTC et TAXe dans PIECE et AFFAIRE
T_PIECE := TOB.CREATE ('PIECE', NIL, -1);
QQ:=OpenSQL('SELECT * FROM PIECE WHERE '+WherePiece(CleDocLigaf,ttdPiece,False),True,-1,'',true) ;
if Not QQ.EOF then T_PIECE.SelectDB('',QQ)
else begin
    Ferme(QQ) ;
    T_PIECE.free;  //mcd 22/03/01
    exit;
    end;

Ferme(QQ) ;

           // affecte le total des lignes
T_PIECE.PutValue ('GP_TOTALTTC',TotTTC);
if (VH^.TenueEuro) then begin
    totbis:=  EuroToFranc(TotTTC);
    taxebis:=  EuroToFranc(TotTaxe);
    end
else begin
    totbis :=  FrancToEuro(TotTTC);
    taxebis :=  FrancToEuro(TotTaxe);
    end;
    // les mtt devise, sont égale à ce qui est en fichier SCOT si monnaie client = monnaie tenue
    // convertit sinon
if (T_PIECE.GetValue ('GP_SAISIECONTRE') =  '-') then
  begin
  T_PIECE.PutValue ('GP_TOTALTTCDEV',TotTTC);
  taxedev := TotTaxe;
  end
else begin
  T_PIECE.PutValue ('GP_TOTALTTCDEV',totbis);
  taxedev :=taxebis;
  end;
T_PIECE.PutValue ('GP_TOTALQTEFACT',totqte);
T_PIECE.PutValue ('GP_TOTALQTESTOCK',totqte);
ImportTob(T_PIECE,True);


 ExecuteSQL('UPDATE AFFAIRE SET AFF_TOTALTTC= '
    + Strfpoint(T_PIECE.GetValue('GP_TOTALTTC'))+ ' ,AFF_TOTALTTCDEV =  '
    + Strfpoint(T_PIECE.GetValue('GP_TOTALTTCDEV'))+
     ',AFF_TOTALTAXE =  '
    + Strfpoint(TotTaxe)   + ',AFF_TOTALTAXEDEV =  '
    + Strfpoint(TaxeDEV)
    //+',AFF_TOTALTAXECON =  '     + Strfpoint(TAxeBis)
    + ' WHERE AFF_AFFAIRE = "' +T_PIECE.GetValue('GP_AFFAIRE') +'" ');



T_PIECE.Free;


end;



procedure TFAsRecupTempo.RecupChampLibre ();
var
    TypeTable, TypeCode,code,libelle : string;
    LngCode,Rang : Integer;
begin
             //champ libre affaire
if ( (CheckListFicBase.Checked[3] = TRUE) or(CheckListFicBase.Checked[6] = TRUE)) then
begin
if (FParam.CheckApport.Checked = TRUE) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   Code := 'MR1';
   libelle := 'Apporteur interne';
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
if (FParam.CheckStat1Aff.Checked = TRUE)and (stat1Aff<5) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   Case Stat1Aff of
      1: Code := 'MT1';
      2: Code := 'MT2';
      3: Code := 'MT3';
      4: Code := 'MT4';
      end;
   libelle := fparam.checkstat1aff.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
if (FParam.CheckStat2Aff.Checked = TRUE)and (Stat2aff <5) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   Case Stat2Aff of
      1: Code := 'MT1';
      2: Code := 'MT2';
      3: Code := 'MT3';
      4: Code := 'MT4';
      end;
   libelle := fparam.checkstat2aff.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
if (FParam.CheckStat3Aff.Checked = TRUE)and (Stat3aff <5) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   Case Stat3Aff of
      1: Code := 'MT1';
      2: Code := 'MT2';
      3: Code := 'MT3';
      4: Code := 'MT4';
      end;
    libelle := fparam.checkstat3aff.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
if (FParam.CheckStat4Aff.Checked = TRUE)and (Stat4aff <5) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   Case Stat4Aff of
      1: Code := 'MT1';
      2: Code := 'MT2';
      3: Code := 'MT3';
      4: Code := 'MT4';
      end;
    libelle := fparam.checkstat4aff.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
end;
     //prestation
if ( (CheckListFicBase.Checked[0] = TRUE) or(CheckListFicBase.Checked[1] = TRUE)or(CheckListFicBase.Checked[2] = TRUE)) then
begin
if (FParam.CheckStat1Prest.Checked = TRUE) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   code:='AT1';
   libelle := fparam.checkstat1Prest.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
if (FParam.CheckStat2Prest.Checked = TRUE) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   code:='AT2';
   libelle := fparam.checkstat2Prest.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
end;

     //ressource
if (CheckListFicBase.Checked[4] = TRUE) then
begin
if (FParam.CheckStat1Emp.Checked = TRUE)and (Stat1res <5) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   Case Stat1REs of
      1: Code := 'RT1';
      2: Code := 'RT2';
      3: Code := 'RT3';
      4: Code := 'RT4';
      end;
    libelle := fparam.checkstat1Emp.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
if (FParam.CheckStat2Emp.Checked = TRUE)and (Stat2res <5) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   Case Stat2REs of
      1: Code := 'RT1';
      2: Code := 'RT2';
      3: Code := 'RT3';
      4: Code := 'RT4';
      end;
   libelle := fparam.checkstat2Emp.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
if (FParam.CheckStat3Emp.Checked = TRUE)and (Stat3res <5) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   Case Stat3REs of
      1: Code := 'RT1';
      2: Code := 'RT2';
      3: Code := 'RT3';
      4: Code := 'RT4';
      end;
   libelle := fparam.checkstat3Emp.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
if (FParam.CheckStat4Emp.Checked = TRUE)and (Stat4res <5) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   Case Stat4REs of
      1: Code := 'RT1';
      2: Code := 'RT2';
      3: Code := 'RT3';
      4: Code := 'RT4';
      end;
   libelle := fparam.checkstat4Emp.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
   // met emploi dans table libre XXXX
TypeTable :='CC';
LngCode := 3;
TypeCode := 'ZLI';
If (Stat4res <5) and (STat4res >0) then Rang:=STat4res +1
else  If (Stat3res <4) and (STat3res >0) then Rang:=STat3res +1
else  If (Stat2res <4) and (STat2res >0) then Rang:=STat2res +1
else  If (Stat1res <4) and (STat1res >0) then Rang:=STat1res +1
else rang:=1;
Code:='RT'+IntTostr(rang);
libelle := 'Emploi';
CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   // met libellé emploi dans texte libre 1
TypeTable :='CC';
LngCode := 3;
TypeCode := 'ZLI';
code:='RC1';
libelle := 'Libellé Emploi';
CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
end;

     //client
if (CheckListFicBase.Checked[5] = TRUE) then
begin
if (FParamCli.CheckStat1Cli.Checked = TRUE ) and (stat1cli <>0) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   Case Stat1Cli of
      1: Code := 'CT1';
      2: Code := 'CT2';
      3: Code := 'CT3';
      4: Code := 'CT4';
      7: Code := 'CR1';
      8: Code := 'CR2';
      9: Code := 'CR3';
    end;
   TypeCode := 'ZLI';
   libelle := fparamCli.checkstat1Cli.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
if (FParamCli.CheckStat2Cli.Checked = TRUE ) and (stat2cli <>0) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   Case Stat2Cli of
      1: Code := 'CT1';
      2: Code := 'CT2';
      3: Code := 'CT3';
      4: Code := 'CT4';
      7: Code := 'CR1';
      8: Code := 'CR2';
      9: Code := 'CR3';
     end;
   TypeCode := 'ZLI';
   libelle := fparamCli.checkstat2cli.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
if (FParamCli.CheckStat3cli.Checked = TRUE )and (stat3cli <>0) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   Case Stat3cli of
      1: Code := 'CT1';
      2: Code := 'CT2';
      3: Code := 'CT3';
      4: Code := 'CT4';
      7: Code := 'CR1';
      8: Code := 'CR2';
      9: Code := 'CR3';
     end;
   TypeCode := 'ZLI';
   libelle := fparamCli.checkstat3cli.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
if (FParamCli.CheckStat4cli.Checked = TRUE )and (stat4cli <>0) then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   Case Stat4cli of
      1: Code := 'CT1';
      2: Code := 'CT2';
      3: Code := 'CT3';
      4: Code := 'CT4';
      7: Code := 'CR1';
      8: Code := 'CR2';
      9: Code := 'CR3';
     end;
   TypeCode := 'ZLI';
   libelle := fparamCli.checkstat4cli.caption;
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   End;
If ctxScot in V_PGI.PGIContexte then
   Begin
   TypeTable :='CC';
   LngCode := 3;
   TypeCode := 'ZLI';
   code:='CD1';
   libelle := 'Entrée cabinet';
   CreerTablette(TypeTable,TypeCode,code,Libelle,LngCode);
   end;
end;

End;

procedure TFAsRecupTempo. CreerTablette(TypeTable,TypeCode,code,Libelle:string;LngCode:integer);
var TTablette : TOB;
begin
   If (TypeTable = 'YX') then TTablette := TOB.CREATE ('CHOIXEXT', NIL, -1)
   Else If (TypeTable = 'CC') then TTablette := TOB.CREATE ('CHOIXCOD', NIL, -1)
   Else TTablette := TOB.CREATE ('COMMUN', NIL, -1); // en principe CO
   TTablette.initValeurs;
   TTablette.PutValue (TypeTable+'_TYPE', TypeCode);
   TTablette.PutFixedStValue (TypeTable+'_CODE', code, 1,LngCode, tctrim, false);
   TTablette.PutValue (TypeTable+'_ABREGE', copy(libelle,1,17));
   TTablette.PutValue (TypeTable+'_LIBELLE', libelle);
   ImportTob(TTablette,True);
   TTablette.Free;

End;

Procedure TFAsRecupTempo.RecupCreerPiece ( Var CleDoc : R_CleDoc ; CodeTiers,CodeAffaire,part1,part2,part3,Entite,etbs,nofac : String;
          Var TobPiece:TOB; Var NoOrdre : Integer)  ;
Var TOBTiers : TOB ;
    Q                 : TQuery ;
    champ, etablissement: string  ;
BEGIN
     // cette fct regarde si la piece existe, sinon, la cree.
     // celle ci est crée avec un n° de piece égal à mois + n° facture
     // mcd 03/12/02 ajout aussi dans la référence du code tiers pour résoudre le cas de N) facture identique sur même date pour 2 clt #

if nofac = '      ' then nofac :='0';
champ := DatetoStr(Cledoc.Datepiece);
champ := Copy(champ,7,4)+ Copy(Champ,4,2)+Format ('%6.6d',[StrToInt(nofac)]) +CodeTiers; //mcd 03/12/02 ajout codetiers
// Initialisations
if multi = 'N' then etablissement := VH^.EtablisDefaut
else begin
     if etbs ='' then etablissement := EntiteVersEtbs(Entite)
     else etablissement := etbs;    //cas facture dans les temps. pris dans affaire
     end;
CleDoc.Souche:=GetSoucheG(CleDoc.NaturePiece,Etablissement,'') ;
CleDoc.Indice:=0 ;
Q:=OpenSQL('SELECT * FROM PIECE WHERE GP_NUMPIECE ="'+ champ+'" and GP_ETABLISSEMENT="'+ etablissement+'"',True,-1,'',true) ;
If Not Q.EOF then begin
   TobPiece.SelectDB('',Q);
   Ferme(Q);
   CleDoc.NumeroPiece:=TobPiece.GetValue('GP_NUMERO');
   Q:=OpenSQL('SELECT MAX(GL_NUMORDRE) FROM LIGNE WHERE GL_NATUREPIECEG = "FRE" and GL_NUMERO='
                      +IntTostr(CleDoc.NumeroPiece),TRUE,-1,'',true);
   if Not Q.EOF then NoOrdre:=Q.Fields[0].AsInteger+1;
   Ferme(Q);
   exit;    
   end ;
Ferme(Q);
NoOrdre:=1;
CleDoc.NumeroPiece:=GetNumSoucheG(CleDoc.Souche,CleDoc.DatePiece) ;
IncNumSoucheG(CleDoc.Souche,CleDoc.DatePiece) ;
AddLesSupEntete (TOBPiece);  //mcd 04/04/2002

InitTOBPiece(TOBPiece) ;
// Divers
TOBPiece.PutValue('GP_NUMPIECE',Champ) ;
TOBPiece.PutValue('GP_ETABLISSEMENT',Etablissement) ;
TOBPiece.PutValue('GP_DEPOT',VH_GC.GCDepotDefaut) ;
TOBPiece.PutValue('GP_REGIMETAXE',VH^.RegimeDefaut) ;
// Tiers
TOBTiers:=TOB.Create('TIERS',Nil,-1) ;
TOBTiers.AddChampSup('RIB',False) ;
RemplirTOBTiers ( TOBTiers, CodeTiers, CleDoc.NaturePiece,False);
TOBPiece.PutValue('GP_TIERS',CodeTiers) ;
TiersVersPiece(TOBTiers,TOBPiece) ;
TOBPiece.PutValue('GP_REGIMETAXE',VH^.RegimeDefaut) ;// pour ecraser les info clients
// Entête
TOBPiece.PutValue('GP_NATUREPIECEG',CleDoc.NaturePiece) ;
TOBPiece.PutValue('GP_SOUCHE',CleDoc.Souche) ;
TOBPiece.PutValue('GP_NUMERO',CleDoc.NumeroPiece) ;
TOBPiece.PutValue('GP_DATEPIECE',CleDoc.DatePiece) ;
TOBPiece.PutValue('GP_AFFAIRE',CodeAffaire) ;
TOBPiece.PutValue('GP_AFFAIRE1',part1) ;
TOBPiece.PutValue('GP_AFFAIRE2',part2) ;
TOBPiece.PutValue('GP_AFFAIRE3',part3) ;
TOBPiece.PutValue('GP_AVENANT','00') ;
// Devise
TOBPiece.PutValue('GP_DEVISE',V_PGI.DevisePivot) ;
TOBPiece.PutValue('GP_TAUXDEV',1 ) ;
TOBPiece.PutValue('GP_DATETAUXDEV',CleDoc.DatePiece) ;
TobTiers.free;   //mcd 22/03/01

END ;

procedure TFAsRecupTempo.RecupStatDet(StTot:string);
    //pour récupérer les stat détaillées
    // cette fct crée des factures de reprise FRE, avec le n° de facture = AAMOIS + N°
var T_LIGNE , T_Piece: TOB;
    Part0,Part1, Part2, Part3, Aff, champ,Article: String;
    mois : Integer;
    ttc,taxe, qte, puht,puttc,taux,totht: double;
Begin
//MCD Penser de traiter les stat et famille quand gérée dans ficheir piece
T_LIGNE := TOB.CREATE ('LIGNE', NIL, -1);
T_PIECE := TOB.CREATE ('PIECE', NIL, -1);
Aff := AffectAff(Copy (StTot, 78,6),Copy (StTot , 84,5),
       AffectTiers(Copy (StTot, 15,10 )), 'AFF','00',Part0,Part1, Part2, Part3,
      False,  mois   ); // il faut lire affaire avant de traiter la piece
if (aff ='') then begin
    TraduitGA(format('Stat det.: clt %10.10s Affaire: %6.6s ',[AffectTiers(Copy (StTot, 15,10)),Copy(StTot, 78,6)]));
    If ctxscot  in V_PGI.PGIContexte then champ := champ + format('Exer:%5.5s',[Copy(StTot,84,5)]);
    champ := champ + format('Fact %6.6s,date %8.8s,Art:%10.10s', [Copy(StTot,25,6),Copy(StTot,6,8),Article]);
   writeln (FichierLog, champ);
   T_LIGNE.Free; T_PIECE.Free;
   exit;
   end;
    // on recherche la piece si existe , on crée sion Obligation de le faire avant article pour alimentatiopn Ok zones article dans lignes
Champ := COpy(Sttot,12,2)+'/'+ Copy(Sttot,10,2)+'/' + Copy(Sttot,6,4) ;
taux :=20.60;
if (Copy(Sttot,6,8) >= '20000401') then taux :=19.60;
if Not(Isvaliddate(champ)) then champ :='01/01/1900';  //mcd 07/02/02
CleDocLigaf.datepiece:=Strtodate(champ);
CleDocLIgaf.NaturePiece:='FRE';
RecupCreerPiece(CledocLigaf, AffectTiers(Copy (StTot, 15,10 )),aff,part1,part2, part3,StTot[215],'',
         Copy(Sttot,25,6),T_PIECE,NoOrdre);
AddLesSupLigne(T_LIGNE,False) ;
InitLigneVide(T_PIECE,T_LIGNE, Nil,Nil, 1,0) ;   // initialise ligne piece

    // on affecte les informations article
T_LIGNE.PutValue('GL_TYPELIGNE','ART');
champ :='';
T_LIGNE.PutValue('GL_TYPEARTICLE','PRE');
If (stTot[108] = 'P') then
    begin
    T_LIGNE.PutValue('GL_TYPEARTICLE','PRE');
    champ := debutpre;
    end
 Else If (stTot[108] = 'F') then
    begin
    T_LIGNE.PutValue('GL_TYPEARTICLE','MAR');
    champ := debutfour;
    end
Else  If  (stTot[108] = 'f')then
    Begin
    T_LIGNE.PutValue('GL_TYPEARTICLE','FRA');
     champ := debutfrais;
   End;
champ := champ + copy (stTot,109,15);

If  (stTot[108] = '*')then begin
        T_LIGNE.PutValue('GL_TYPEARTICLE','PRE');
        champ :=VH_GC.AFACOMPTE;   // cas acompte. Prise paramètre par défaut
        end;
if trim(champ) <>'' then T_LIGNE.PutValue ('GL_ARTICLE',CodeArticleUnique2( champ,''));
T_LIGNE.PutFixedStValue ('GL_CODEARTICLE', champ, 1,17, tcTrimr, false);
T_LIGNE.PutFixedStValue ('GL_REFARTSAISIE', champ, 1,17, tcTrimr, false);
If trim(champ)='' then Article:=''
else Article:=   CodeArticleUnique2( champ,'');
champ :=TraduitGA(format('Stat det.: clt %10.10s Affaire: %6.6s ',[AffectTiers(Copy (StTot, 15,10)),Copy(StTot, 78,6)]));
If ctxscot  in V_PGI.PGIContexte then champ := champ + format('Exer:%5.5s',[Copy(StTot,84,5)]);
champ := champ + format('Fact %6.6s,date %8.8s,Art:%10.10s', [Copy(StTot,25,6),Copy(StTot,6,8),Article]);
if (TestExistance(champ,Article, T_LIGNE.GetValue('GL_TYPEARTICLE'),'',
      '',Copy (StTot , 84,5),AffectTiers(Copy (StTot, 15,10 )), 'AFF','00',Part0,Part1, Part2, Part3,
      aff ,  mois)=False)
     then begin
     T_LIGNE.Free; T_PIECE.Free;
     exit;
     end;
if (trim(copy(sttot,78,11)) ='') then aff:=''; // cas affaire à blanc...
T_LIGNE.PutValue ('GL_AFFAIRE1', part1);
T_LIGNE.PutValue ('GL_AFFAIRE2', part2);
T_LIGNE.PutValue ('GL_AFFAIRE3', part3);
T_LIGNE.PutValue ('GL_AVENANT', '00');
T_LIGNE.PutValue ('GL_TIERSLIVRE', T_LIGNE.Getvalue('GL_TIERS'));
T_LIGNE.PutValue ('GL_TIERSPAYEUR', T_LIGNE.Getvalue('GL_TIERS'));
T_LIGNE.PutValue ('GL_TIERSFACTURE', T_LIGNE.Getvalue('GL_TIERSPAYEUR'));
T_PIECE.Putvalue ('GP_TIERSLIVRE',T_LIGNE.GetValue('GL_TIERSLIVRE'));
T_PIECE.Putvalue ('GP_TIERSFACTURE',T_LIGNE.GetValue('GL_TIERSFACTURE'));
T_PIECE.Putvalue ('GP_TIERSPAYEUR',T_LIGNE.GetValue('GL_TIERSPAYEUR'));

T_LIGNE.PutValue ('GL_NUMLIGNE',NoOrdre);
T_LIGNE.PutValue ('GL_NUMORDRE',NoOrdre);
ttc := StrToFloat(copy(sttot, 191,12))/100.0;
taxe := (ttc * taux)/100;
if (VH^.TenueEuro) then taxe:=Arrondi(taxe ,V_PGI.OkDecV)
        else taxe:=Arrondi(taxe ,V_PGI.OkDecV);
ttc := ttc+taxe;
T_LIGNE.PutValue ('GL_FAMILLETAXE1','NOR');
          // le code PUHT est obligatoire, on le met = total ht si blanc
PuhT :=StrTofloat(Copy (StTot,179,12));
puht := puht/100;
totht:=puht;
Qte:= StrTofloat(Copy(Sttot,167,12)) ;
qte := qte/100;
if Qte <>0 then puht := (puht) / qte else qte:=1.0;
puttc:=puht+ (puht * taux)/100;

T_LIGNE.PutValue ('GL_TOTALHT',totht);
T_LIGNE.PutValue ('GL_MONTANTHT', totht);
T_LIGNE.PutValue ('GL_PUHT', Puht);
T_LIGNE.PutValue ('GL_PMRP',puht);
T_LIGNE.PutValue ('GL_PMRPACTU', puht);
T_LIGNE.PutValue ('GL_DPR', puht);
T_LIGNE.PutValue ('GL_PUHTNET', Puht);
T_LIGNE.PutValue ('GL_MONTANTTTC', ttc);
T_LIGNE.PutValue ('GL_TOTALTTC', ttc);
T_LIGNE.PutValue ('GL_TOTALTAXE1', taxe);
T_LIGNE.PutValue ('GL_PUTTC', puttc);
T_LIGNE.PutValue ('GL_PUTTCNET', puttc);
T_LIGNE.PutValue ('GL_PUTTCBASE', puttc);
T_LIGNE.PutValue ('GL_PUHTBASE', puht);
    // on est en monnaie de tenue du dossier. Mtt=devsie, contrevaleur traduite.
T_LIGNE.PutValue ('GL_MONTANTHTDEV', totht);
T_LIGNE.PutValue ('GL_TOTALHTDEV', totht);
T_LIGNE.PutValue ('GL_PUHTDEV', Puht);
T_LIGNE.PutValue ('GL_PUHTNETDEV',Puht);
T_LIGNE.PutValue ('GL_MONTANTTTCDEV', ttc);
T_LIGNE.PutValue ('GL_TOTALTTCDEV', ttc);
T_LIGNE.PutValue ('GL_TOTALTAXEDEV1', taxe);
T_LIGNE.PutValue ('GL_PUTTCDEV', puttc);
T_LIGNE.PutValue ('GL_PUTTCNETDEV', puttc);

T_LIGNE.PutValue ('GL_COTATION', 1);
T_LIGNE.PutValue ('GL_PCB', 1);

if (T_LIGNE.GetValue('GL_TYPELIGNE')<>'COM') and (T_LIGNE.GetValue('GL_TYPELIGNE')<>'TOT')    then
  begin
  qte :=StrtoFloat(Copy(Sttot,167,12));
  if (qte = 0) then qte :=1 else qte := qte/100;
  T_LIGNE.PutValue ('GL_QTEFACT', qte);
  T_LIGNE.PutValue ('GL_QTESTOCK', qte);
  totqte :=totqte + qte;
  end;
T_LIGNE.PutFixedStValue ('GL_DPR', StTot, 143,12, tcDouble100, false);
T_LIGNE.PutValue ('GL_QUALIFQTEVTE', 'H');
T_LIGNE.PutValue('GL_VALIDECOM','AFF');
T_LIGNE.PutValue('GL_UTILISATEUR',V_PGI.User);
T_LIGNE.PutValue('GL_CREERPAR','IMP');
T_LIGNE.PutValue('GL_TENUESTOCK','-');
T_LIGNE.PutValue('GL_ESCOMPTABLE','X');
T_LIGNE.PutValue('GL_REMISABLELIGNE','-');
T_LIGNE.PutValue('GL_REMISABLEPIED','X');

T_LIGNE.PutValue ('GL_PERIODE',GetPeriode(T_LIGNE.GetValue('GL_DATEPIECE')));
T_LIGNE.PutValue ('GL_SEMAINE',NumSemaine(T_LIGNE.GetValue('GL_DATEPIECE')));
T_LIGNE.PutValue('GL_TVAENCAISSEMENT','-');
ImportTob(T_LIGNE,True);
T_LIGNE.Free;
             // on cumul les mtt dans la piece
T_PIECE.PutValue ('GP_TOTALTTC',T_PIECE.GetValue('GP_TOTALTTC')+ TTC);
T_PIECE.PutValue ('GP_TOTALTTCDEV',T_PIECE.GetValue('GP_TOTALTTCDEV')+TTC);
T_PIECE.PutValue ('GP_TOTALQTEFACT',T_PIECE.GetValue('GP_TOTALQTEFACT')+qte);
T_PIECE.PutValue ('GP_TOTALQTESTOCK',T_PIECE.GetValue('GP_TOTALQTESTOCK')+qte);
T_PIECE.PutValue ('GP_EDITEE', 'X') ;
           // affecte le total des lignes
T_PIECE.PutValue ('GP_TOTALHT',T_PIECE.GetValue('GP_TOTALHT')+totht);
T_PIECE.PutValue ('GP_TOTALBASEREM',T_PIECE.GetValue('GP_TOTALBASEREM')+totht);
T_PIECE.PutValue ('GP_TOTALBASEESC',T_PIECE.GetValue('GP_TOTALBASEESC')+totht);
T_PIECE.PutValue ('GP_TOTALHTDEV',T_PIECE.GetValue('GP_TOTALHTDEV')+totht);
T_PIECE.PutValue ('GP_TOTALBASEREMDEV',T_PIECE.GetValue('GP_TOTALBASEREMDEV')+totht);
T_PIECE.PutValue ('GP_TOTALBASEESCDEV',T_PIECE.GetValue('GP_TOTALBASEESCDEV')+totht);
T_PIECE.PutValue('GP_TVAENCAISSEMENT','TE') ;
T_PIECE.PutValue ('GP_MODEREGLE', GetParamSoc('SO_GcModeRegleDefaut'));
T_PIECE.PutValue ('GP_PERIODE',GetPeriode(T_PIECE.GetValue('GP_DATEPIECE')));
T_PIECE.PutValue ('GP_SEMAINE',NumSemaine(T_PIECE.GetValue('GP_DATEPIECE')));
T_PIECE.PutValue('GP_CREEPAR','IMP');
// mcd 05/02/01 T_PIECE.PutValue('GP_VIVANTE','-');

   // renseigne les stat client dans zones correspodnante de la piece
Case Stat1Cli of
  1:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS1', StTot, 35,4, tcTrim, false);
  End;
Case Stat2Cli of
  1:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS1', StTot, 39,4, tcTrim, false);
  2:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS2', StTot, 39,4, tcTrim, false);
  End;
Case Stat3Cli of
  1:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS1', StTot, 43,4, tcTrim, false);
  2:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS2', StTot, 43,4, tcTrim, false);
  3:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS3', StTot, 43,4, tcTrim, false);
  End;
Case Stat4Cli of
  1:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS1', StTot, 47,4, tcTrim, false);
  2:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS2', StTot, 47,4, tcTrim, false);
  3:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS3', StTot, 47,4, tcTrim, false);
  4:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS4', StTot, 47,4, tcTrim, false);
  End;
Case Stat1Aff of
  1:  T_PIECE.PutFixedStValue ('GP_LIBREAFF1', StTot, 92,4, tcTrim, false);
  2:  T_PIECE.PutFixedStValue ('GP_LIBREAFF2', StTot, 92,4, tcTrim, false);
  3:  T_PIECE.PutFixedStValue ('GP_LIBREAFF3', StTot, 92,4, tcTrim, false);
  4:  T_PIECE.PutFixedStValue ('GP_LIBREAFF4', StTot, 92,4, tcTrim, false);
  End;
Case Stat2Aff of
  1:  T_PIECE.PutFixedStValue ('GP_LIBREAFF1', StTot, 96,4, tcTrim, false);
  2:  T_PIECE.PutFixedStValue ('GP_LIBREAFF2', StTot, 96,4, tcTrim, false);
  3:  T_PIECE.PutFixedStValue ('GP_LIBREAFF3', StTot, 96,4, tcTrim, false);
  4:  T_PIECE.PutFixedStValue ('GP_LIBREAFF4', StTot, 96,4, tcTrim, false);
  End;
Case Stat3Aff of
  1:  T_PIECE.PutFixedStValue ('GP_LIBREAFF1', StTot, 100,4, tcTrim, false);
  2:  T_PIECE.PutFixedStValue ('GP_LIBREAFF2', StTot, 100,4, tcTrim, false);
  3:  T_PIECE.PutFixedStValue ('GP_LIBREAFF3', StTot, 100,4, tcTrim, false);
  4:  T_PIECE.PutFixedStValue ('GP_LIBREAFF4', StTot, 100,4, tcTrim, false);
  End;
Case Stat4Aff of
  1:  T_PIECE.PutFixedStValue ('GP_LIBREAFF1', StTot, 104,4, tcTrim, false);
  2:  T_PIECE.PutFixedStValue ('GP_LIBREAFF2', StTot, 104,4, tcTrim, false);
  3:  T_PIECE.PutFixedStValue ('GP_LIBREAFF3', StTot, 104,4, tcTrim, false);
  4:  T_PIECE.PutFixedStValue ('GP_LIBREAFF4', StTot, 104,4, tcTrim, false);
  End;

champ :=trim (AnsiUppercase(Copy(Sttot,35,4)))  ;
if (Stat1cli = 7) then T_PIECE.PutValue ('GP_TIERSSAL1', champ);
if (Stat1cli = 8) then T_PIECE.PutValue ('GP_TIERSSAL2', champ);
if (Stat1cli = 9) then T_PIECE.PutValue ('GP_TIERSSAL3', champ);
champ :=trim (AnsiUppercase(Copy(Sttot,39,4)));
if (Stat2cli = 7) then T_PIECE.PutValue ('GP_TIERSSAL1', champ);
if (Stat2cli = 8) then T_PIECE.PutValue ('GP_TIERSSAL2', champ);
if (Stat2cli = 9) then T_PIECE.PutValue ('GP_TIERSSAL3', champ);
champ :=trim (AnsiUppercase(Copy(Sttot,43,4)));
if (Stat3cli = 7) then T_PIECE.PutValue ('GP_TIERSSAL1', champ);
if (Stat3cli = 8) then T_PIECE.PutValue ('GP_TIERSSAL2', champ);
if (Stat3cli = 9) then T_PIECE.PutValue ('GP_TIERSSAL3', champ);
champ :=trim (AnsiUppercase(Copy(Sttot,47,4)));
if (Stat4cli = 7) then T_PIECE.PutValue ('GP_TIERSSAL1', champ);
if (Stat4cli = 8) then T_PIECE.PutValue ('GP_TIERSSAL2', champ);
if (Stat4cli = 9) then T_PIECE.PutValue ('GP_TIERSSAL3', champ);

ImportTob(T_PIECE,True);
T_PIECE.Free;

end;

procedure TFAsRecupTempo.RecupStatGlob(StTot:string);
    //pour récupérer les stat globale
    // cette fct crée des factures de reprise FRE, avec le n° de facture = AAMOIS + N°
var T_LIGNE , T_Piece: TOB;
    Part0,Part1, Part2, Part3, Aff, champ: String;
     mois : Integer;
    ttc,taxe, qte, totht: double;
Begin
T_LIGNE := TOB.CREATE ('LIGNE', NIL, -1);
T_PIECE := TOB.CREATE ('PIECE', NIL, -1);

Aff := AffectAff(Copy (StTot, 18,6),Copy (StTot , 24,5),
            AffectTiers(Copy (StTot, 7,10 )), 'AFF','00',Part0,Part1, Part2, Part3, FALSE,mois);
if (aff ='') then begin
   champ :=TraduitGA(format('Enrgt Stat glob. non traité : clt %10.10s Affaire: %6.6s ',[AffectTiers(Copy (StTot, 7,10)),Copy(StTot, 18,6)]));
   If ctxscot  in V_PGI.PGIContexte then champ := champ + format('Exer:%5.5s',[Copy(StTot,24,5)]);
   champ := champ + format(' N° fact %6.6s, date %8.8s', [Copy(StTot,37,6),Copy(StTot,29,8)]);
   writeln (FichierLog, champ);
   T_LIGNE.Free; T_PIECE.Free;
   exit;
   end;
if (trim(copy(sttot,18,11)) ='') then aff:=''; // cas affaire à blanc...
   // on recherche la piece si existe , on crée sion
Champ := COpy(Sttot,35,2)+'/'+ Copy(Sttot,33,2)+'/' + Copy(Sttot,29,4) ;
if Not(Isvaliddate(champ)) then champ :='01/01/1900';  //mcd 07/02/02
CleDocLigaf.datepiece:=Strtodate(champ);
CleDocLIgaf.NaturePiece:='FRE';
RecupCreerPiece(CledocLigaf, AffectTiers(Copy (StTot, 7,10 )),aff,part1,part2,part3,StTot[169],'',
         Copy(Sttot,37,6),T_PIECE,NoOrdre);
AddLesSupLigne(T_LIGNE,False) ;
InitLigneVide(T_PIECE,T_LIGNE, Nil,Nil, 1,0) ;   // initialise ligne piece
T_LIGNE.PutValue ('GL_AFFAIRE1', part1);
T_LIGNE.PutValue ('GL_AFFAIRE2', part2);
T_LIGNE.PutValue ('GL_AFFAIRE3', part3);
T_LIGNE.PutValue ('GL_AVENANT', '00');
T_LIGNE.PutValue ('GL_TIERSLIVRE', T_LIGNE.Getvalue('GL_TIERS'));
T_LIGNE.PutValue ('GL_TIERSPAYEUR', T_LIGNE.Getvalue('GL_TIERS'));
T_LIGNE.PutValue ('GL_TIERSFACTURE', T_LIGNE.Getvalue('GL_TIERSPAYEUR'));
T_PIECE.Putvalue ('GP_TIERSLIVRE',T_LIGNE.GetValue('GL_TIERSLIVRE'));
T_PIECE.Putvalue ('GP_TIERSFACTURE',T_LIGNE.GetValue('GL_TIERSFACTURE'));
T_PIECE.Putvalue ('GP_TIERSPAYEUR',T_LIGNE.GetValue('GL_TIERSPAYEUR'));

T_LIGNE.PutValue ('GL_NUMLIGNE',NoOrdre);
T_LIGNE.PutValue ('GL_NUMORDRE',NoOrdre);
    // on affecte les informations article
T_LIGNE.PutValue('GL_TYPELIGNE','ART');
champ :='';
T_LIGNE.PutValue('GL_TYPEARTICLE','PRE');
champ :=VH_GC.AFACOMPTE;   // ligne globale, mise sur paramètre par défaut
if trim(champ) <>'' then T_LIGNE.PutValue ('GL_ARTICLE',CodeArticleUnique2( champ,''));
T_LIGNE.PutFixedStValue ('GL_CODEARTICLE', champ, 1,17, tcTrimr, false);
T_LIGNE.PutFixedStValue ('GL_REFARTSAISIE', champ, 1,17, tcTrimr, false);
totht := StrToFloat(copy(sttot, 118,12))/100.0;
ttc:= StrToFloat(copy(sttot, 142,12))/100.0;
taxe:=ttc -totht;
T_LIGNE.PutValue ('GL_FAMILLETAXE1','NOR');
qte:=1.0;
T_LIGNE.PutValue ('GL_TOTALHT',totht);
T_LIGNE.PutValue ('GL_MONTANTHT', totht);
T_LIGNE.PutValue ('GL_PUHT', totht);
T_LIGNE.PutValue ('GL_PMRP',totht);
T_LIGNE.PutValue ('GL_PMRPACTU', totht);
T_LIGNE.PutValue ('GL_DPR', totht);
T_LIGNE.PutValue ('GL_PUHTNET', totht);
T_LIGNE.PutValue ('GL_MONTANTTTC', ttc);
T_LIGNE.PutValue ('GL_TOTALTTC', ttc);
T_LIGNE.PutValue ('GL_TOTALTAXE1', taxe);
T_LIGNE.PutValue ('GL_PUTTC', ttc);
T_LIGNE.PutValue ('GL_PUTTCNET', ttc);
T_LIGNE.PutValue ('GL_PUTTCBASE', ttc);
T_LIGNE.PutValue ('GL_PUHTBASE', totht);
    // on est en monnaie de tenue du dossier. Mtt=devsie, contrevaleur traduite.
T_LIGNE.PutValue ('GL_MONTANTHTDEV', totht);
T_LIGNE.PutValue ('GL_TOTALHTDEV', totht);
T_LIGNE.PutValue ('GL_PUHTDEV', totht);
T_LIGNE.PutValue ('GL_PUHTNETDEV',totht);
T_LIGNE.PutValue ('GL_MONTANTTTCDEV', ttc);
T_LIGNE.PutValue ('GL_TOTALTTCDEV', ttc);
T_LIGNE.PutValue ('GL_TOTALTAXEDEV1', taxe);
T_LIGNE.PutValue ('GL_PUTTCDEV', ttc);
T_LIGNE.PutValue ('GL_PUTTCNETDEV', ttc);

T_LIGNE.PutValue ('GL_COTATION', 1);
T_LIGNE.PutValue ('GL_PCB', 1);

if (T_LIGNE.GetValue('GL_TYPELIGNE')<>'COM') and (T_LIGNE.GetValue('GL_TYPELIGNE')<>'TOT')    then
  begin
  qte :=1 ;
  T_LIGNE.PutValue ('GL_QTEFACT', qte);
  T_LIGNE.PutValue ('GL_QTESTOCK', qte);
  totqte :=totqte + qte;
  end;
T_LIGNE.PutFixedStValue ('GL_DPR', StTot, 130,12, tcDouble100, false);
T_LIGNE.PutValue ('GL_QUALIFQTEVTE', 'H');
T_LIGNE.PutValue('GL_VALIDECOM','AFF');
T_LIGNE.PutValue('GL_UTILISATEUR',V_PGI.User);
T_LIGNE.PutValue('GL_CREERPAR','IMP');
T_LIGNE.PutValue('GL_TENUESTOCK','-');
T_LIGNE.PutValue('GL_ESCOMPTABLE','X');
T_LIGNE.PutValue('GL_REMISABLELIGNE','-');
T_LIGNE.PutValue('GL_REMISABLEPIED','X');

T_LIGNE.PutValue ('GL_PERIODE',GetPeriode(T_LIGNE.GetValue('GL_DATEPIECE')));
T_LIGNE.PutValue ('GL_SEMAINE',NumSemaine(T_LIGNE.GetValue('GL_DATEPIECE')));
T_LIGNE.PutValue('GL_TVAENCAISSEMENT','-');
ImportTob(T_LIGNE,True);
T_LIGNE.Free;
T_PIECE.PutValue ('GP_REFCOMPTABLE',trim(Copy(Sttot,98,20)));
             // on cumul les mtt dans la piece
T_PIECE.PutValue ('GP_TOTALTTC',T_PIECE.GetValue('GP_TOTALTTC')+ TTC);
T_PIECE.PutValue ('GP_TOTALTTCDEV',T_PIECE.GetValue('GP_TOTALTTCDEV')+TTC);
T_PIECE.PutValue ('GP_TOTALQTEFACT',T_PIECE.GetValue('GP_TOTALQTEFACT')+qte);
T_PIECE.PutValue ('GP_TOTALQTESTOCK',T_PIECE.GetValue('GP_TOTALQTESTOCK')+qte);
T_PIECE.PutValue ('GP_EDITEE', 'X') ;
           // affecte le total des lignes
T_PIECE.PutValue ('GP_TOTALHT',T_PIECE.GetValue('GP_TOTALHT')+totht);
T_PIECE.PutValue ('GP_TOTALBASEREM',T_PIECE.GetValue('GP_TOTALBASEREM')+totht);
T_PIECE.PutValue ('GP_TOTALBASEESC',T_PIECE.GetValue('GP_TOTALBASEESC')+totht);
T_PIECE.PutValue ('GP_TOTALHTDEV',T_PIECE.GetValue('GP_TOTALHTDEV')+totht);
T_PIECE.PutValue ('GP_TOTALBASEREMDEV',T_PIECE.GetValue('GP_TOTALBASEREMDEV')+totht);
T_PIECE.PutValue ('GP_TOTALBASEESCDEV',T_PIECE.GetValue('GP_TOTALBASEESCDEV')+totht);
T_PIECE.PutValue('GP_TVAENCAISSEMENT','TE') ;
T_PIECE.PutValue ('GP_MODEREGLE', GetParamSoc('SO_GcModeRegleDefaut'));
T_PIECE.PutValue ('GP_PERIODE',GetPeriode(T_PIECE.GetValue('GP_DATEPIECE')));
T_PIECE.PutValue ('GP_SEMAINE',NumSemaine(T_PIECE.GetValue('GP_DATEPIECE')));
T_PIECE.PutValue('GP_CREEPAR','IMP');
//mcd 05/02/03 T_PIECE.PutValue('GP_VIVANTE','-');
   // renseigne les stat client dans zones correspodnante de la piece
Case Stat1Cli of
  1:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS1', StTot, 82,4, tcTrim, false);
  End;
Case Stat2Cli of
  1:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS1', StTot, 86,4, tcTrim, false);
  2:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS2', StTot, 86,4, tcTrim, false);
  End;
Case Stat3Cli of
  1:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS1', StTot, 90,4, tcTrim, false);
  2:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS2', StTot, 90,4, tcTrim, false);
  3:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS3', StTot, 90,4, tcTrim, false);
  End;
Case Stat4Cli of
  1:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS1', StTot, 94,4, tcTrim, false);
  2:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS2', StTot, 94,4, tcTrim, false);
  3:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS3', StTot, 94,4, tcTrim, false);
  4:  T_PIECE.PutFixedStValue ('GP_LIBRETIERS4', StTot, 94,4, tcTrim, false);
  End;
champ :=trim (AnsiUppercase(Copy(Sttot,82,4)))  ;
if (Stat1cli = 7) then T_PIECE.PutValue ('GP_TIERSSAL1', champ);
if (Stat1cli = 8) then T_PIECE.PutValue ('GP_TIERSSAL2', champ);
if (Stat1cli = 9) then T_PIECE.PutValue ('GP_TIERSSAL3', champ);
champ :=trim (AnsiUppercase(Copy(Sttot,86,4)));
if (Stat2cli = 7) then T_PIECE.PutValue ('GP_TIERSSAL1', champ);
if (Stat2cli = 8) then T_PIECE.PutValue ('GP_TIERSSAL2', champ);
if (Stat2cli = 9) then T_PIECE.PutValue ('GP_TIERSSAL3', champ);
champ :=trim (AnsiUppercase(Copy(Sttot,90,4)));
if (Stat3cli = 7) then T_PIECE.PutValue ('GP_TIERSSAL1', champ);
if (Stat3cli = 8) then T_PIECE.PutValue ('GP_TIERSSAL2', champ);
if (Stat3cli = 9) then T_PIECE.PutValue ('GP_TIERSSAL3', champ);
champ :=trim (AnsiUppercase(Copy(Sttot,94,4)));
if (Stat4cli = 7) then T_PIECE.PutValue ('GP_TIERSSAL1', champ);
if (Stat4cli = 8) then T_PIECE.PutValue ('GP_TIERSSAL2', champ);
if (Stat4cli = 9) then T_PIECE.PutValue ('GP_TIERSSAL3', champ);

ImportTob(T_PIECE,True);
T_PIECE.Free;

end;

Function TFAsRecupTempo.EntiteVersEtbs(Entite:string):string;
var i : integer;
begin
       //A partir du code entité, on rend le code établissement voulu
result:= VH^.EtablisDefaut;
If multi = 'N' then exit;
for i:=0 to FparamEntite.TobEntite.detail.count-1 do
   begin
   if (Entite = FparamEntite.TobEntite.Detail[i].GetValue('Code')) then
      begin
      result:=  FparamEntite.TobEntite.Detail[i].GetValue('Etablissement');
      if trim(result) = '' then     result:= VH^.EtablisDefaut;
      break;
      end;
    end;
end;


procedure TFAsRecupTempo.RecupPort (stTot:string);
var T_PORT : TOB;
    TOM_PORT : TOM;
    valeur:string;
Begin
 //fct qui crée le port à partir d'un enrgt frais en % de SCOT/TEMPO
T_PORT := TOB.CREATE ('PORT', NIL, -1);
TOM_PORT := CreateTom('PORT', Nil, False, False);

TOM_PORT.InitTob(T_PORT);

valeur :='';
Valeur := FindEtReplace (Copy(stTot,7,15),'"',' ',True);
T_PORT.PutValue ('GPO_CODEPORT',valeur);
Valeur := FindEtReplace (Copy(StTot,22,70),'"',' ',True);
T_PORT.PutValue ('GPO_LIBELLE',valeur);
T_PORT.PutValue('GPO_TYPEPORT','HT');
    // zone comptable alimenté par famille presttaion ou nature ou rien
If (FParam.ComboPreCompt.ItemIndex = 1) then
    T_PORT.PutFixedStValue ('GPO_COMPTAARTICLE',stTot,122,3,tcChaine,True);
If (FParam.ComboPreCompt.ItemIndex = 2) then
    T_PORT.PutFixedStValue ('GPO_COMPTAARTICLE',stTot,125,1,tcChaine,True);
T_PORT.PutValue ('GPO_FAMILLETAXE1','NOR');

T_PORT.PutValue ('GPO_MINIMUM',0);
T_PORT.PutValue ('GPO_PVTTC',0);
T_PORT.PutValue ('GPO_PVHT',0);
T_PORT.PutFixedStValue ('GPO_COEFF',stTot,143,11,tcDouble100,False);
T_PORT.PutFixedStValue ('GPO_DATECREATION',stTot,294,8,tcDate8AMJ,False);
T_PORT.PutFixedStValue ('GPO_DATEMODIF',stTot,302,8,tcDate8AMJ,False);
T_PORT.PutValue('GPO_UTILISATEUR',V_PGI.User);
T_PORT.PutValue('GPO_CREATEUR',V_PGI.User);
T_PORT.PutValue('GPO_CREERPAR','IMP');
T_PORT.PutValue('GPO_FERME','-');
T_PORT.PutValue('GPO_FRANCO','-');
ImportTob(T_PORT,True);

T_PORT.Free; TOM_PORT.Free;
End;

Procedure TFAsRecupTempo.RecupFrais( CleDoc : R_CleDoc ;StTot : String;Monpres:boolean)  ;
var T_PORT ,T_ART, T_PIEDPORT: TOB;
    TOM_PORT: TOM;
    champ:string;
    QQ : Tquery;
    htcon:double;
Begin
 // fct qui crée la table PIEDPORT pour l'affaire en courd
QQ := OpenSQL('SELECT * FROM PORT Where GPO_CODEPORT = "'+ Copy(Sttot, 1524,15)+'"',True,-1,'',true) ;
   //   cas ou le port n'existe pas, on se trouve avec un port en mtt.
   // il faut le créer à partir de l'article.
If  QQ.EOF then begin
    ferme (QQ);
    champ := debutfrais;
    champ := champ + copy (stTot,1524,15);
    T_ART := TOB.CREATE ('ARTICLE', NIL, -1);
    QQ := OpenSQL('SELECT * FROM ARTICLE Where GA_CODEARTICLE= "'+ Champ+'"',True,-1,'',true) ;
    If  Not QQ.EOF then begin
      T_ART.SelectDB('',QQ) ;
      T_PORT := TOB.CREATE ('PORT', NIL, -1);
      TOM_PORT := CreateTom('PORT', Nil, False, False);
      TOM_PORT.InitTob(T_PORT);
      T_PORT.PutFixedStValue ('GPO_CODEPORT',stTot,1524,15,tcTrim,False);
      T_PORT.PutValue ('GPO_LIBELLE',T_ART.GetValue ('GA_LIBELLE'));
      T_PORT.PutValue('GPO_TYPEPORT','MT');
      T_PORT.PutValue ('GPO_COMPTAARTICLE',T_ART.GetValue ('GA_COMPTAARTICLE'));
      T_PORT.PutValue ('GPO_FAMILLETAXE1','NOR');

      T_PORT.PutValue ('GPO_MINIMUM',0);
      T_PORT.PutValue ('GPO_PVTTC',0);
      T_PORT.PutValue ('GPO_PVHT',T_ART.GetValue ('GA_PVHT'));
      T_PORT.PutValue ('GPO_COEFF',0);
      T_PORT.PutValue ('GPO_DATECREATION',T_ART.GetValue ('GA_DATECREATION'));
      T_PORT.PutValue ('GPO_DATEMODIF',T_ART.GetValue ('GA_DATEMODIF'));
      T_PORT.PutValue('GPO_UTILISATEUR',V_PGI.User);
      T_PORT.PutValue('GPO_CREATEUR',V_PGI.User);
      T_PORT.PutValue('GPO_CREERPAR','IMP');
      T_PORT.PutValue('GPO_FERME','-');
      T_PORT.PutValue('GPO_FRANCO','-');
      ImportTob(T_PORT,True);

      T_PORT.Free; TOM_PORT.Free;
      T_ART.Free;
      end;
    QQ := OpenSQL('SELECT * FROM PORT Where GPO_CODEPORT = "'+ Copy(Sttot, 1524,15)+'"',True,-1,'',true) ;
    end;
        // on ecrit le port dans la table PIEDPORT
        // ecrit toujours en N° 1 et n'alimente pas les mtt, même si cumul ligne existe
If  Not QQ.EOF then begin
      T_PORT := TOB.CREATE ('PORT', NIL, -1);
      T_PORT.SelectDB('',QQ) ;
      T_PIEDPORT := TOB.CREATE ('PIEDPORT', NIL, -1);

      T_PIEDPORT.PutValue ('GPT_NATUREPIECEG',cledoc.naturepiece);
      T_PIEDPORT.PutValue ('GPT_SOUCHE',cledoc.souche);
      T_PIEDPORT.PutValue ('GPT_NUMERO',cledoc.numeropiece);
      T_PIEDPORT.PutValue ('GPT_CODEPORT',T_PORT.GetValue ('GPO_CODEPORT'));
      T_PIEDPORT.PutValue ('GPT_LIBELLE',T_PORT.GetValue ('GPO_LIBELLE'));
      T_PIEDPORT.PutValue ('GPT_TYPEPORT',T_PORT.GetValue ('GPO_TYPEPORT'));
      T_PIEDPORT.PutValue ('GPT_COMPTAARTICLE',T_PORT.GetValue ('GPO_COMPTAARTICLE'));
      T_PIEDPORT.PutValue ('GPT_FAMILLETAXE1',T_PORT.GetValue ('GPO_FAMILLETAXE1'));
      T_PIEDPORT.PutValue ('GPT_POURCENT',T_PORT.GetValue ('GPO_COEFF'));
      T_PIEDPORT.PutValue ('GPT_MINIMUM',T_PORT.GetValue ('GPO_MINIMUM'));
      T_PIEDPORT.PutValue ('GPT_FRANCO',T_PORT.GetValue ('GPO_FRANCO'));
      T_PIEDPORT.PutValue ('GPT_TOTALHT',T_PORT.GetValue ('GPO_PVHT'));
      T_PIEDPORT.PutValue ('GPT_TOTALTTC',T_PORT.GetValue ('GPO_PVTTC'));
      T_PIEDPORT.PutValue ('GPT_DEVISE','FRF');
      T_PIEDPORT.PutValue ('GPT_NUMPORT',1);
      if (VH^.TenueEuro) then  htcon := EuroToFranc(T_PORT.GetValue ('GPO_PVHT'))
      else htcon := FrancToEuro(T_PORT.GetValue ('GPO_PVHT'));
      If (VH^.TenueEuro = MonPres ) then   T_PIEDPORT.PutValue ('GPT_TOTALHTDEV', T_PORT.GetValue ('GPO_PVHT'))
       else  T_PIEDPORT.PutValue ('GPT_TOTALHTDEV', htcon);
      ImportTob(T_PIEDPORT,True);
      T_PORT.Free;
      T_PIEDPORT.Free;
      end;

Ferme(QQ);

end;


procedure TFAsRecupTempo.RecupAnnuaire (StTot,champ:string);
var T_ANNUAIRE : TOB;
   QQ ,Q2: TQuery;
   i:integer;
   form : string;
Begin

T_ANNUAIRE := TOB.CREATE ('ANNUAIRE', NIL, -1);
QQ:=OPENSQL('SELECT * From ANNUAIRE WHERE ANN_TIERS="'+AffectTiers (Copy (Sttot,7,10))+'"',True,-1,'',true);
if Not QQ.EOF then
   begin   // existe, on prend les valeurs existantes
   T_ANNUAIRE.SelectDB ('', QQ) ;
    T_ANNUAIRE.PutValue ('ANN_DATEMODIF', V_PGI.DateEntree);
   end
else begin
          // n'existe pas, on recupère n° code personne
{$IFDEF DP}
     T_ANNUAIRE.PutValue ('ANN_CODEPER',CalculNumCle('ANN_CODEPER','ANNUAIRE') );
    T_ANNUAIRE.PutValue ('ANN_DATECREATION', V_PGI.DateEntree);
    T_ANNUAIRE.PutValue ('ANN_DATEMODIF', V_PGI.DateEntree);
    T_ANNUAIRE.PutValue ('ANN_CONFIDENTIEL','0');
    T_ANNUAIRE.PutValue ('ANN_PPPM','PM');
    T_ANNUAIRE.PutValue ('ANN_SEXE','M');
{$ENDIF}
     end;
Ferme(QQ);



T_ANNUAIRE.PutValue ('ANN_AUXILIAIRE', champ);
T_ANNUAIRE.PutFixedStValue ('ANN_NOM1', StTot, 95,30, tcTrim, false);
T_ANNUAIRE.PutValue ('ANN_TIERS',  AffectTiers (Copy (Sttot,7,10)));

T_ANNUAIRE.PutFixedStValue ('ANN_NOMPER', StTot, 17,10, tcTrim, false);
T_ANNUAIRE.PutFixedStValue ('ANN_ALRUE1', StTot, 162,32, tcTrim, false);
T_ANNUAIRE.PutFixedStValue ('ANN_ALRUE2', StTot, 194,32, tcTrim, false);
T_ANNUAIRE.PutFixedStValue ('ANN_ALCP', StTot, 226,5, tcTrim, false);
T_ANNUAIRE.PutFixedStValue ('ANN_ALVILLE', StTot, 231,26, tcTrim, false);
T_ANNUAIRE.PutValue ('ANN_PAYS', 'FRA');
T_ANNUAIRE.PutValue ('ANN_NATIONALITE', 'FRA');
T_ANNUAIRE.PutValue ('ANN_DEVISE',V_PGI.DevisePivot);
T_ANNUAIRE.PutValue ('ANN_LANGUE','FRA');
T_ANNUAIRE.PutFixedStValue ('ANN_CODENAF', StTot, 386,4, tcTrim, false);
T_ANNUAIRE.PutFixedStValue ('ANN_TEL1', StTot, 260,20, tcTrim, false);
T_ANNUAIRE.PutFixedStValue ('ANN_SITEWEB', StTot, 280,20, tcTrim, false);
T_ANNUAIRE.PutFixedStValue ('ANN_FAX', StTot, 300,20, tcTrim, false);
T_ANNUAIRE.PutFixedStValue ('ANN_MINITEL', StTot, 320,20, tcTrim, false);

T_ANNUAIRE.PutValue ('ANN_UTILISATEUR',V_PGI.User);
       //forme juridique. Celles ci sont renumérotée dans une combo à 3c
for i:=0 to FparamFjur.TobFjur.detail.count-1 do
   begin
   if (copy (StTot, 157,5) = FparamFjur.TobFjur.Detail[i].GetValue('Code')) then
      begin
      if FparamFjur.TobFjur.Detail[i].GetValue('Nouveau code')= '' 
         then break;
      Q2 := OpenSQL('SELECT JFJ_FORME,JFJ_FORMEGEN  FROM JUFORMEJUR WHERE JFJ_CODEDP = '
        + '"'+FparamFjur.TobFjur.Detail[i].GetValue('Nouveau code')+'"', TRUE,-1,'',true);
      if not Q2.eof then
      begin
        Form := Q2.FindField('JFJ_FORME').Asstring;
        T_ANNUAIRE.PutValue('ANN_FORME', Form);
        if Q2.FindField('JFJ_FORMEGEN').Asstring <> '' then
           T_ANNUAIRE.PutValue('ANN_FORMEGEN', Q2.FindField('JFJ_FORMEGEN').Asstring);
        end;
      Ferme (Q2);
      break;
      end;    
    end;     
If ctxScot  in V_PGI.PGIContexte then begin
   if ( copy(sttot,1011,2) <'01') or  ( copy(sttot,1011,2) >'12')  then i :=12
     else i:= StrtoInt (copy(sttot,1011,2));
   if (i < 1) or (i >12) then i:=12;
   T_ANNUAIRE.PutValue ('ANN_MOISCLOTURE', Format ('%2.2d',[i])) ;
   end;
ImportTob(T_ANNUAIRE,True);
T_ANNUAIRE.Free;
End;


procedure TFAsRecupTempo.InitArticle(TypeArt:string);
var  OB , OB_DETAIL : TOB;
begin
  // Création de l'enregistrement  article par défaut pour profil

  OB := TOB.Create('un article',nil,-1);
  OB_DETAIL := TOB.Create('ARTICLE',OB,-1);
  OB_DETAIL.PutValue('GA_TYPEARTICLE','PRE');
  OB_DETAIL.PutValue('GA_FAMILLETAXE1','NOR');
  OB_DETAIL.PutValue('GA_STATUTART','UNI');
  OB_DETAIL.PutValue('GA_CALCPRIXHT','AUC');
  OB_DETAIL.PutValue('GA_CALCPRIXTTC','AUC');
  OB_DETAIL.PutValue('GA_ACTIVITEREPRISE','F');
  OB_DETAIL.PutValue('GA_QUALIFUNITEVTE',GetParamSoc('SO_AFMESUREACTIVITE'));
  OB_DETAIL.PutValue('GA_QUALIFUNITEACT',GetParamSoc('SO_AFMESUREACTIVITE'));

If (TypeArt='PRE') then begin
if GetParamSoc ('SO_AFACOMPTE') <>'' then begin
  OB_DETAIL.PutValue('GA_ARTICLE',CodeArticleUnique2(GetParamSoc ('SO_AFACOMPTE'),''));
  OB_DETAIL.PutValue('GA_CODEARTICLE',GetParamSoc ('SO_AFACOMPTE'));
  OB_DETAIL.PutValue('GA_LIBELLE','Acompte forfaitaire');
  OB.InsertOrUpdateDB(True);
  end;
if GetParamSoc ('SO_AFPRESTATIONRES') <>'' then begin
  OB_DETAIL.PutValue('GA_ARTICLE',CodeArticleUnique2(GetParamSoc ('SO_AFPRESTATIONRES'),''));
  OB_DETAIL.PutValue('GA_CODEARTICLE',GetParamSoc ('SO_AFPRESTATIONRES'));
  OB_DETAIL.PutValue('GA_LIBELLE','Prestation defaut temps');
  OB.InsertOrUpdateDB(True);
  end;
  // mcd 19/06/02  création article si prestation globale
OB_DETAIL.PutValue('GA_ARTICLE',CodeArticleUnique2('PRGLOBAL',''));
OB_DETAIL.PutValue('GA_CODEARTICLE','PRGLOBAL');
OB_DETAIL.PutValue('GA_LIBELLE','Prestation facture globale');
OB.InsertOrUpdateDB(True);
end;
  // mcd 19/06/02  création article si frais globale
If (TypeArt='FRA') then begin
OB_DETAIL.PutValue('GA_ARTICLE',CodeArticleUnique2('FRGLOBAL',''));
OB_DETAIL.PutValue('GA_CODEARTICLE','FRGLOBAL');
OB_DETAIL.PutValue('GA_LIBELLE','Frais facture globale');
OB_DETAIL.PutValue('GA_TYPEARTICLE','FRA');
OB.InsertOrUpdateDB(True);
end;
  // mcd 19/06/02  création article si fourniture globale
If (TypeArt='FOU') then begin
OB_DETAIL.PutValue('GA_ARTICLE',CodeArticleUnique2('FOGLOBAL',''));
OB_DETAIL.PutValue('GA_CODEARTICLE','FOGLOBAL');
OB_DETAIL.PutValue('GA_LIBELLE','Fourniture facture globale');
OB_DETAIL.PutValue('GA_TYPEARTICLE','MAR');
OB.InsertOrUpdateDB(True);
end;
  OB.Free;


end;


procedure TFAsRecupTempo.RecupFour (StTot:string);
var T_CLIENT : TOB;
    champ, champ1 , ModeRempl :string;
    Lng_auxi,   jour, duree:integer;

Begin
     // récupération fournisseur
T_CLIENT := TOB.CREATE ('TIERS', NIL, -1);
T_CLIENT.initValeurs;
Lng_auxi:=VH^.Cpta[fbAux].Lg;


// il faut copier le compte sur la longueur en fichier
// fait avec le caractère de bourrage des paramètres
champ := StringOfChar (VH^.Cpta[fbAux].Cb, Lng_auxi);
champ1 := TRim (copy (StTot, 6, 10));
champ := champ1 + champ;
T_CLIENT.PutFixedStValue ('T_AUXILIAIRE', champ,1,lng_auxi, tcChaine, false);
T_CLIENT.PutFixedStValue ('T_TIERS', champ1,2,9, tcChaine, false);

    // on récupère le RIB
if (Trim(Copy (Sttot, 294,46)) <> '') then RecupRibFou (Copy (Sttot, 294,48), Copy (champ, 1 , Lng_Auxi), TRUE);
         // on récupère le contact = nom correspondant
If (Trim(Copy (Sttot, 350,40)) <>'') then
    begin
    RecupTiersContactFou (Sttot,champ,VH^.Cpta[fbAux].Lg);
    end;
T_CLIENT.PutValue ('T_NATUREAUXI','FOU');
T_CLIENT.PutFixedStValue ('T_LIBELLE', StTot, 16,35, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_ABREGE', StTot, 16,17, tcTrim, false);
T_CLIENT.PutValue ('T_EAN',  AffectTiers (Copy (Sttot,6,10)));

Lng_auxi:=VH^.Cpta[fbGene].Lg;
// il faut copier le compte collectif sur la longueur en fichier
// fait avec le caractère de bourrage des paramètres
champ := StringOfChar (VH^.Cpta[fbGene].Cb, Lng_auxi);
If (Trim(Copy (Sttot,412,10)) <>'') then champ1 :=  trim (Copy (StTot, 412,10))
    else champ1 :=VH^.DefautFou;
champ := champ1 + champ;
T_CLIENT.PutFixedStValue ('T_COLLECTIF', champ,1,lng_auxi, tcChaine, false);
T_CLIENT.PutFixedStValue ('T_PRENOM', StTot, 66,35, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_ADRESSE1', StTot, 101,35, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_ADRESSE2', StTot, 136,35, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_VILLE', StTot, 171,35, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_CODEPOSTAL', StTot, 206,5, tcTrim, false);
    // le code pays est systématiquement mis à FRA (= France PGI)
    // les code pays étranger ne sont pas traité dans la reprise.
    // ce sont des cas exceptionnels avec modif fiche après.
T_CLIENT.PutValue ('T_PAYS', 'FRA');
T_CLIENT.PutValue ('T_NATIONALITE', 'FRA');
T_CLIENT.PutValue ('T_DEVISE',V_PGI.DevisePivot);
T_CLIENT.PutValue ('T_LANGUE','FRA');
T_CLIENT.PutValue ('T_MULTIDEVISE','-');
                  //identification CEE
T_CLIENT.PutFixedStValue ('T_NIF', StTot, 393,18, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_TELEPHONE', StTot, 214,20, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_RVA', StTot, 274,20, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_FAX', StTot, 254,20, tcTrim, false);
T_CLIENT.PutFixedStValue ('T_TELEX', StTot, 234,20, tcTrim, false);
T_CLIENT.PutValue ('T_FACTURE',  AffectTiers (Copy (Sttot,6,10)));
T_CLIENT.PutValue ('T_DATECREATION',V_PGI.DateEntree );
T_CLIENT.PutValue ('T_DATEMODIF',V_PGI.DateEntree );
    // on renseigne monnaie de présentation client
If (Copy(StTot,411,1)='E') Then T_CLIENT.PutValue ('T_EURODEFAUT', 'X')
              Else T_CLIENT.PutValue ('T_EURODEFAUT', '-');

T_CLIENT.PutValue ('T_REGIMETVA', 'FRA');
T_CLIENT.PutValue('T_TVAENCAISSEMENT',GetParamSoc('SO_TVAENCAISSEMENT'));

T_CLIENT.PutValue ('T_SOUMISTPF', '-');
T_CLIENT.PutValue ('T_CONFIDENTIEL','-');
T_CLIENT.PutValue ('T_RELEVEFACTURE','-');
T_CLIENT.PutValue ('T_CREERPAR','IMP');
T_CLIENT.PutValue ('T_UTILISATEUR',V_PGI.User);
T_CLIENT.PutValue ('T_AVOIRRBT','-');
T_CLIENT.PutValue ('T_ISPAYEUR','-');
T_CLIENT.PutValue ('T_FACTUREHT','X');

  // traitement des modes de règlement

duree :=0;
if (Copy (Sttot, 344,3) <>'   ') then duree := StrtoInt (Copy (Sttot, 344,3));
jour:=0;
if (Copy (Sttot, 347,2)<>'  ') then jour := StrtoInt (Copy (Sttot, 347,2));
T_CLIENT.PutValue ('T_MODEREGLE', RecupModeReglFou (Copy (Sttot, 349,1), Copy (Sttot, 342,2),
    ModeRempl, duree, jour));
    // mcd 15/06/01 il faut créer la table tiers compl pour les tiers
REcupCLientCOmplFou(sttot, T_CLIENT.GetValue('T_AUXILIAIRE'),T_CLIENT.GetValue('T_TIERS'));
ImportTob(T_CLIENT,True);
T_CLIENT.Free;
End;

procedure TFAsRecupTempo.RecupArtFour (StTot:string);
var T_CATALOGUE,TobArt : TOB;
    CodeArt,Lib,Valeur  : string;
    Q : TQuery;
    PA : Double;
Begin
// recup si l'article existe...(récupération fourniture qui va se trouver dans table article et/ou table frais et port
  Q:=Nil;
  try
    Valeur := trim(FindEtReplace (Copy(stTot,6,15),'"',' ',True));
    Q:=OpenSQL('SELECT GA_ARTICLE,GA_LIBELLE,GA_FOURNPRINC FROM ARTICLE where GA_CODEARTICLE="'+ Valeur +'"',true,-1,'',true);
    if Not(Q.EOF) then
     begin
     CodeArt := Q.Fields[0].asString; Lib :=Q.Fields[1].asString;
     TobArt := tob.create('ARTICLE',Nil,-1);
     TobArt.SelectDB('',Q);
     end
    else Exit;
    T_CATALOGUE := TOB.CREATE ('CATALOGU', NIL, -1);
    T_CATALOGUE.initValeurs;

    T_CATALOGUE.PutFixedStValue ('GCA_REFERENCE', StTot, 30,15, tcTrim, false);
    if T_CATALOGUE.getValue('GCA_REFERENCE') = '' then
       T_CATALOGUE.PutValue ('GCA_REFERENCE', Valeur);
    T_CATALOGUE.PutFixedStValue ('GCA_TIERS', StTot, 21,9, tcTrim, false);
    T_CATALOGUE.PutValue('GCA_ARTICLE',CodeArt);
    T_CATALOGUE.PutValue('GCA_LIBELLE',Lib);
    T_CATALOGUE.PutValue('GCA_DATEREFERENCE',CurrentDate);
    T_CATALOGUE.PutValue('GCA_DATEREFERENCE',CurrentDate);
    PA := StrtoFloat (Copy (Sttot, 63,8));
    T_CATALOGUE.PutValue('GCA_PRIXBASE',PA);
    T_CATALOGUE.PutValue('GCA_CREERPAR','IMP');
    ImportTob(T_CATALOGUE,True);
      // maj du fournisseur principal sur l'article
    if TobArt.GetValue('GA_FOURNPRINC')='' then
       begin
       TobArt.PutValue('GA_FOURNPRINC',T_CATALOGUE.GetValue('GCA_TIERS'));
       TobArt.UpdateDB;
       end;
  finally
     Ferme(Q);
  end;
T_CATALOGUE.Free;
TobArt.Free;
End;

procedure TFAsRecupTempo.RecupTiersCOntactFou (StTot:string; champ:string;lng_auxi:integer);
var T_CONTACT: TOB;

Begin
T_CONTACT:= TOB.CREATE ('CONTACT', NIL, -1);
T_CONTACT.initValeurs;

T_CONTACT.PutFixedStValue ('C_AUXILIAIRE', champ, 1,lng_auxi, tcTrim, false);
T_CONTACT.PutFixedStValue ('C_NOM', StTot, 350,40, tcTrim, false);
T_CONTACT.PutValue ('C_NUMEROCONTACT',1);
T_CONTACT.PutValue ('C_NATUREAUXI','FOU');
T_CONTACT.PutValue ('C_TYPECONTACT','T');

ImportTob(T_CONTACT,True);
T_CONTACT.Free;
End;

function TFAsRecupTempo.RecupModeReglFou(FinMois, ModePaie, Cheque:string; Duree, Jour:integer ):string ;
var T_MODEREGL : TOB;
    Q : TQuery;
    max, Apartirde, Separepar,Arrondijour, Sduree, champ ,Ecartjour : string;
Begin
    { Cette fonction permet de rendre le mode de règlement associé
    aux modes gérés dans l'ancienne gamme.
    Si le mode n'existe pas, le crée.
    FinMois : O/N si calcul fin de mois ou pas
    ModePaie : CH, TR, LR ... Doit être en lien avec table MODEPAIE
    Durée : à X jours (30, 60 ...)
    Jour : le X (jour de l'échénace)  }

    // on accède d'abord à la table pour voir si le mode de règlement existe déjà
Sduree:= InttoStr (duree);
if (FinMois = 'O') then Apartirde := 'FIN' else Apartirde :='FAC';
case jour of
0 : Arrondijour := 'PAS';
1 : Arrondijour := 'DEB';
5 : Arrondijour := '05M';
10 : Arrondijour := '10M';
15 : Arrondijour := '15M';
20: Arrondijour := '20M';
25: Arrondijour := '25M';
30: Arrondijour := 'FIN';
31: Arrondijour := 'FIN';
else  Arrondijour := 'PAS';
end ;

champ :='1M';
Ecartjour :='';
SeparePar := champ;

Q:=OPENSQL('SELECT MR_MODEREGLE From MODEREGL WHERE MR_APARTIRDE = "' + Apartirde +
    '" AND MR_PLUSJOUR= ' + sduree + ' And MR_ARRONDIJOUR = "' + Arrondijour +
    '" And MR_NOMBREECHEANCE = 1 And MR_SEPAREPAR = "'+ Separepar +
    '" And MR_MP1 = "' + ModePaie + '"' ,true,-1,'',true);
If Not Q.EOF Then begin
    result :=  Q.Fields[0].asstring;
    end
Else begin
        // il faut créer le nouveau mode. on garde en memoire les n°
        // affecté aux modes de règlements. On regarde si existe avant
        // de le créer
    Ferme(Q);
    Max := Format ('%3.3d',[(NBMode+1)]);
    Q:=OpenSQL('SELECT MR_MODEREGLE FROM MODEREGL WHERE MR_MODEREGLE = "'+ max+ '"',TRUE,-1,'',true) ;
    while (Not Q.EOF) do begin
           Ferme(Q);
           Inc (Nbmode);
           Max := Format ('%3.3d',[(NBMode+1)]);
           Q:=OpenSQL('SELECT MR_MODEREGLE FROM MODEREGL WHERE MR_MODEREGLE = "'+ max +'"',TRUE) ;
           end ;
    T_MODEREGL := TOB.CREATE ('MODEREGL', NIL, -1);
    T_MODEREGL.InitValeurs;

    T_MODEREGL.PutValue ('MR_MODEREGLE', max);
    champ := Format ('%s à %d jours le %d, Fin mois %s', [ModePaie, duree, jour, Finmois]) ;
    T_MODEREGL.PutValue ('MR_LIBELLE', champ);
    T_MODEREGL.PutValue ('MR_ABREGE', Copy(champ,1,17));
    T_MODEREGL.PutValue ('MR_APARTIRDE', Apartirde);
    T_MODEREGL.PutValue ('MR_PLUSJOUR', duree);
    T_MODEREGL.PutValue ('MR_ARRONDIJOUR', Arrondijour);
    T_MODEREGL.PutValue ('MR_NOMBREECHEANCE', 1);
    T_MODEREGL.PutValue ('MR_SEPAREPAR', Separepar);
    T_MODEREGL.PutValue ('MR_ECARTJOURS', EcartJour);
    T_MODEREGL.PutValue ('MR_REMPLACEMIN', Cheque);
    T_MODEREGL.PutValue ('MR_MONTANTMIN', 0);
    ImportTob(T_MODEREGL,True) ;
    T_MODEREGL.Free;
    Result := max;
    end;
Ferme(Q);
End;


procedure TFAsRecupTempo.RecupRibFou (StTot, Clt:string; condi : boolean);
var T_RIB : TOB;
    QQ :TQuery;
    etb, guich, cpt, cle : string;
    max : integer;
Begin
// Alimentation de la table RIB. à parti client ou affaire

Etb := Copy (StTot,  26,5);
guich := Copy (StTot , 31,5);
cpt := Copy (StTot,  36,11);
cle := Copy (StTot,  48,2);
QQ:=OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="' + clt+ '" and R_ETABBQ="' + etb+
 '" and R_GUICHET="' + guich+ '" and R_NUMEROCOMPTE="' + cpt+ '" and R_CLERIB="' + cle+ '"' ,True,-1,'',true) ;
if  QQ.EOF
    then begin   // le RIB n'existe pas, on le crée
        Ferme(QQ) ;
        QQ:=OpenSQL('SELECT MAX(R_NUMERORIB) FROM RIB WHERE R_AUXILIAIRE="' + clt+ '"'  ,TRUE);
          if Not QQ.EOF then Max:=QQ.Fields[0].AsInteger+1 else Max:=1;
        T_RIB := TOB.CREATE ('RIB', NIL, -1);
        T_RIB.InitValeurs;
        T_RIB.PutValue ('R_AUXILIAIRE', Clt);
        T_RIB.PutValue ('R_NUMERORIB', Max);
        if (condi = TRUE) then T_RIB.PutValue ('R_PRINCIPAL', 'X')
            else T_RIB.PutValue ('R_PRINCIPAL', '-') ;
        T_RIB.PutValue ('R_ETABBQ', etb);
        T_RIB.PutValue ('R_GUICHET', guich);
        T_RIB.PutValue ('R_NUMEROCOMPTE',cpt);
        T_RIB.PutValue ('R_CLERIB',  cle);
        T_RIB.PutFixedStValue ('R_DOMICILIATION', StTot, 1,25, tcTrimr, false);
        T_RIB.PutValue ('R_DEVISE',V_PGI.DevisePivot);

        ImportTob(T_RIB,True);
        T_RIB.Free;
        end;
Ferme(QQ) ;

End;

// fct qui test existance Affaire, client,article, ressource et qui rend l'erreur
function TFAsRecupTempo.TestExistance (ZChamp,Art,TypArt,Ress,Aff, ZExer, Clt,Typ,Avenant:string; Var Part0,Part1, Part2, Part3,Affaire : STring;  Var Moisclot:Integer):Boolean ;
var   QQ : TQuery;
    Xchamp:string;
Begin
result:=True;
if (Aff <>'') then begin
    Affaire:= AffectAff(aff ,Zexer, clt, typ,Avenant,Part0,Part1, Part2, Part3, False, MoisClot);
    If Affaire ='' then begin
       Xchamp:=Zchamp + TraduitGa(' Affaire non trouvée');
       writeln (FichierLog, Xchamp);
       result:=False;
       exit;
       end;
   end;
   // mcd ajout test TypArt car, il faut vérifier couple article + typeart
if (Art <>'') then begin
  QQ := OpenSQL ('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_ARTICLE ="'+ Art+'" AND GA_TYPEARTICLE="'+ TYPART +'"',TRUE);
  if QQ.EOF then
     Begin
       Xchamp:=Zchamp + TraduitGa(' Article non trouvé');
       writeln (FichierLog, Xchamp);
       result:=False;
       Ferme (QQ);
       exit;
       end ;
  Ferme (QQ);
   end;
if (Ress <>'') then begin
  QQ := OpenSQL ('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_RESSOURCE ="'+ Ress+'"',TRUE);
  if QQ.EOF then
     Begin
       Xchamp:=Zchamp + TraduitGa(' Ressource non trouvé');
       writeln (FichierLog, Xchamp);
       result:=False;
       Ferme (QQ);
       exit;
       end ;
  Ferme (QQ);
  end;
end;

end.


