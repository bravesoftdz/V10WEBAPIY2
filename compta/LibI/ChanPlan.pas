// modif mbo 12.09.2005 - FQ 16608 = suite FQ 16286 - pb saisie et affichage avec 3 décimales sur exceptionnel
// modif mbo 30.09.2005 - FQ 15280 - autoriser l'exceptionnel sur les immos avec plan variable et derogatoire
// modif mbo 30.09.2005 - FQ 16757 - touche F6 pour récup du montant maximum
// BTY - 11/05 Ajout Calcul plan futur avec VNC et Gestion plan fiscal CRC 2002-10
//             en révision du plan
// BTY - 11/05 Ajout Calcul plan futur avec VNC en éléments exceptionnels
// TGA 29/11/2005 maj exceptionnel si sortie directe par bvalider
// BTY - 12/05 FQ 17106 Ne pas contrôler la durée fiscale / déjà amorti si l'immo avait déjà un plan fiscal
// BTY - 12/05 Afficher nelles zones Durée d'amortisssement et Valeur résiduelle
// BTY - 12/05 FQ 17174 Erreur ancienne : taux ECO pouvait être > taux FISCAL =>  à bloquer
// BTY - 12/05 FQ 17178 Erreur ancienne : modif duréeou méthode sans tabuler ne réactualise pas le taux
// BTY - 01/06 En annulation opération, OPECHANGEPLAN peut aussi concerner DPR
// BTY - 01/06 FQ 17373 Ouvrir le plan CRC2002-10 même à durée ECO inchangée
// MBO - 02/06 Impact de la saisie d'un antérieur dépréciation sur l'onglet exceptionnel
// BTY - 03/06 FQ 17446 Ouvrir Calcul plan futur avec la VNC si passage éco DEG -> LIN
// BTY - 03/06 FQ 17708 Champ 'Calcul plan futur' porte évènement parasite bAmortissementFiscalClick sur le OnClick
// MBO - 04/06 FQ 17923 prise en cpte nouvelle opération : modif date mise en service
// MBO - 05/06 FQ 18028 - prise en cpte de la bonne date de début d'amort en calcul durée déjà amortie (DureeValide)
// MBO - 05/06 FQ 17569 - modif des dates debut pour calcul des taux
// BTY - 06/06 FQ 18393 - En série, reprendre la date de l'opération
// BTY - 06/06 FQ 18394 - En série, reprendre le no de serie pour qu'en annulation on puisse annuler les autres opérations de la série
// MVG - 12/07/2006 pour correction conseils compilation
// BTY - 10/06 - Plan fiscal lié à la présence d'une prime d'équipement
// MBO - 10/06 - correction controles liés à l'exceptionnel appliqués en sortie de révision de plan
// BTY - 11/06 - Cocher et bloquer Plan fiscal si présence prime ou subvention
// BTY - 11/06 - FQ 19064 Date d'achat modifiée sur une immo au départ non amortissable
// BTY - 11/06 - FQ 19080 Durée minimale LIN = 1 mois, DEG = 36 mois
// BTY - 12/06 - FQ 19279 Appliquer FQ 18028 à une immo non amortissable: date début = date opération
// BTY - 04/07 - Nelle gestion fiscale, supprimer la notion de plan fiscal CRC200210
// BTY - 05/07 FQ 20256 En révision et annulation de révision, mettre le type de dérogatoire
// BTY - 05/07 FQ 20434 Enregistrement révision ou exceptionnel plante en web access sur le bloc-notes
// MBO - 11/06/2007 - Amélioration des messages en cas de saisie d'exceptionnel impossible
// MBO - 12/06/2007 - pas de révision du plan fiscal si immo remplaçante
// MBO - 06/07 - Composant remplaçant: afficher 'Remplac. composant 2E catégorie' dans la combo base fiscale
// BTY - 09/07 FQ 21502 Type base fiscale présente Rempl. Composant => remplacer items.delete par la propriété Plus qui filtre
// Va avec le paramétrage &#@ de la tablette TITYPEBASEFISC
// BTY - 09/07 FQ 21523 Voulez-vous enregistrer les modifs en sortant de la fiche immo sans modifs, mais après avoir annulé une révision de plan
unit ChanPlan;

interface

uses
  Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
{$IFDEF VER150}
  Variants,
{$ENDIF}
  Math,
  ExtCtrls,
  StdCtrls,
  Hctrls,
  Mask,
  DBCtrls,
  Spin,
  Db,
{$IFDEF EAGLCLIENT}
  uTOB,
{$ELSE}
  HDB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  HSysMenu,
  hmsgbox,
  HTB97,
  ComCtrls,
  HEnt1,
  HRichEdt,
  HRichOLE,
  HPanel,
  UiUtil,
  OpEnCour,
  ImPlan,
  ImOuPlan,
  ParamDat,
  ImEnt;


Type

  TChangePlanAmort = class(TForm)
    HPanel2: THPanel;
    Pages: TPageControl;
    PGeneral: TTabSheet;
    AmortEco: TGroupBox;
    tI_DUREEECO: THLabel;
    tI_TAUXECO: THLabel;
    Label1: TLabel;
    TauxEco: THNumEdit;
    A: TGroupBox;
    Label8: TLabel;
    tI_TAUXFISC: THLabel;
    tI_DUREEFISC: THLabel;
    TauxFiscal: THNumEdit;
    PDotationExcep: TTabSheet;
    HLabel3: THLabel;
    VNC: THNumEdit;
    DETAILDOTATION: TGroupBox;
    HLabel5: THLabel;
    GroupBox1: TGroupBox;
    bAjoutDotation: TRadioButton;
    bRemplaceDotation: TRadioButton;
    bRepriseDotation: TRadioButton;
    HPanel1: THPanel;
    HLabel1: THLabel;
    DATE_OPE: THCritMaskEdit;
    cMethodeEco: THValComboBox;
    DureeEco: TSpinEdit;
    cMethodeFiscale: THValComboBox;
    DureeFiscale: TSpinEdit;
    HM: THMsgBox;
    Label2: TLabel;
    Label3: TLabel;
    HPanel3: THPanel;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    bAmortissementFiscal: TCheckBox;
    Label5: TLabel;
    Label4: TLabel;
    HLabel7: THLabel;
    lBaseEco: TLabel;
    HLabel2: THLabel;
    lBaseFisc: TLabel;
    NOUVELLEDOT: TLabel;
    Label6: TLabel;
    DONTDOTEXC: THNumEdit;
    HPanel4: THPanel;
    HPanel5: THPanel;
    HLabel6: THLabel;
    DOTATIONEXC: THNumEdit;
    Label7: TLabel;
    DotationTotale: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    sDateAchat: THLabel;
    sCodeImmo: THLabel;
    HLabel11: THLabel;
    sLibelleImmo: THLabel;
    sValeurHT: THLabel;
    Label11: TLabel;
    HPanel6: THPanel;
    GroupBox2: TGroupBox;
    BLOC_NOTE: THRichEditOLE;
    bDureeRest: TCheckBox;
    {bPlanCRC: TCheckBox;
    stTypeBaseFisc: TLabel;}
    cTypeBaseFisc: THValComboBox;
    bDureeRestE: TCheckBox;
    stDureeDeja: TLabel;
    stValeurresiduelle: TLabel;
    Valeurresiduelle: THLabel;
    DureeDeja: THLabel;
    bDureeRestF: TCheckBox;           // révision plan
    bGestionFiscale: TCheckBox;       // révision plan
    procedure FormShow(Sender: TObject);
    procedure MajMontantDotation(Sender: TObject);
    procedure TypeDotationClick(Sender: TObject);
    procedure ModifPlanSurChangeDurMeth;// changement durée/Méthode
    procedure ModifPlanSurChangeExc;
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure AmortFiscExit(Sender: TObject);
    procedure AmortEcoExit(Sender: TObject);
    procedure DureeEcoExit (Sender: TObject); // BTY FQ 19095
    procedure bAmortissementFiscalClick(Sender: TObject);
    procedure DATE_OPEKeyPress(Sender: TObject; var Key: Char);
    procedure DOTATIONEXCKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure DATE_OPEExit(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    {procedure bPlanCRCClick(Sender: TObject); }
    procedure cTypeBaseFiscChange(Sender: TObject);
    procedure DureeRestExit(Sender: TObject);
    procedure DureeRestFExit(Sender: TObject);
  private
    { Déclarations privées }
    QPlan : TQuery;
    DotationExe : double;
    MaxAjout : double;
    MaxRemplace : double;
    MaxReprise : double;

    // ajout mbo 30.09.05 fq 15280
    VNCfinEx : double;   // VNC fin exercice en cours
    VRF      : double;    // VRF fin exercice en cours  ajout mbo 11.06.07
    BaseFisc : double;   // base amortissable fiscale
    DotExeFisc : double; // dotation fiscale pour exercice
    RepriseFisc : double;       // somme des antérieurs fiscaux - DotEco = somme antérieurs éco

    RevisionEco : double;
    RevisionFisc : double;
    InitVNC : double;
    DotEco : double;
    BaseEco : double;
    RepriseDep : double;    // mbo impact antérieur dépréciation
    RepriseEco : double;
    PlanCalcul : TPlanAmort;
    fValeurTheorique : double;
    fCumulEco : double;
    fCumulFisc : double;
    fOrdreSerie : integer;  // BTY FQ 18394
    fSBVPRI : double;
    fSBVMT : double;
    fNAM : boolean;
    fInalienabilite : integer;
    bFuturVNF : boolean;
    fRemplace : string;
    bDPI : boolean;
    QPlanCalcul : TQuery;
    fRepriseDR, fRepriseFEC : double;

    procedure VersLaFiche;
    procedure EnregistreChangePlan;
    procedure EnregChangePlan;
    function GereDetailDotation : TModalResult;
    procedure TraiteFinChangePlan;
    procedure EcritOrigImmo;
    procedure AffecteModifPlanCalcul(TypeAmort : integer);
    procedure AfficheOngletExceptionnel(AffPlan : TPlanAmort);
    procedure GereAmortissementTotal;
    function DureeValide (Methode : string; Duree : integer; bFiscal : boolean ) : boolean;
    //function DureeMinimum (Methode : string; Duree : integer ) : boolean; //XVI Conseil Compile...
    procedure AfficheDureeDejaAmortie;
    procedure AfficheVNCDebutExo;
    procedure RecupAncienPlanFiscal;
    procedure AfficheGestionFiscale(PlanCalcul:TPlanAmort; TypeAmort:integer);
    function  CalculAnterieursOK : boolean;
  public
    { Déclarations publiques }
    Plan : TPlanAmort;
    AncienPlan : TPlanAmort;
    CodeImmo : string;
    ElemExc : double;
    DateOpe : TDateTime;
    OrdreS : integer;
    DateAchat : TDateTime;
    TypeChangement : integer;
    MontantDot : double;
    TypeDot : string;
    TypeExc : string;
    //ajout mbo 30.09.05 fq 15280
    PlanFiscal : string;
    PlanEco : string;
    Derog : boolean;
    fDateSerie : TDateTime; // FQ 18393
    procedure EcritLogChangePlan;
  end;

// FQ 18393
//function ExecuteChangePlan(Lequel : string; TypeChangement : integer) : TModalResult;
function ExecuteChangePlan(var Lequel : string; TypeChangement : integer) : TModalResult;
procedure AnnuleChangementPlan(LogChPlan : TLogChPlan);
function OpeChangementPlan(TypeOpe : string) : boolean;

implementation

uses Outils;

{$R *.DFM}

{***********A.G.L.***********************************************
Auteur  ...... : CAY
Créé le ...... : 01/01/1950
Modifié le ... : 07/11/2005
Description .. : Révision du plan d'amortissement: typechangement=1
Suite ........ : Saisie d'un exceptionnel:         typechangement=2
Suite ........ : FQ 18393 Reprendre la date d'opération passée en entrée
Mots clefs ... :
*****************************************************************}
//function ExecuteChangePlan(Lequel : string; TypeChangement : integer) : TModalResult;
function ExecuteChangePlan(var Lequel : string; TypeChangement : integer) : TModalResult;
var ChangePlanAmort: TChangePlanAmort;
begin
// MVG 12/07/2006 pour corriger les conseils result:=mrCancel ;
// mbo - 30.09.2005 - FQ 15280 on autorise exceptionnel sur dérogatoire
//if (TypeChangement=2) and (existeSQL('SELECT I_METHODEFISC FROM IMMO WHERE I_IMMO="'+Lequel+'" AND I_METHODEFISC<>""')) then
//  PGIBox('Vous ne pouvez pas effectuer de dotations ou reprises exceptionnelles sur une immobilisation gérant l''amortissement dérogatoire','Changement de plan')
//else
  begin
  ChangePlanAmort:=TChangePlanAmort.Create(Application) ;
  try
    // FQ 18393
    //ChangePlanAmort.CodeImmo:=Lequel;
    ChangePlanAmort.CodeImmo:=ReadTokenSt(Lequel);
    ChangePlanAmort.fDateSerie := iDate1900;
    if Lequel <> '' then
      ChangePlanAmort.fDateSerie:= StrToDate(ReadTokenSt(Lequel));
    //
    // FQ 18394
    ChangePlanAmort.fOrdreSerie:= TrouveNumeroOrdreSerieLogSuivant;
    if Lequel <> '' then
      ChangePlanAmort.fOrdreSerie:= StrToInt(ReadTokenSt(Lequel));

    ChangePlanAmort.TypeChangement:=TypeChangement;
    ChangePlanAmort.ShowModal;
  finally
    result:=ChangePlanAmort.ModalResult;
    // FQ 18393
    if result = mrYes then
        begin
        Lequel := ChangePlanAmort.DATE_OPE.Text;
        // FQ 18394
        Lequel := Lequel + ';' + IntToStr(ChangePlanAmort.fOrdreSerie);
        end;
    ChangePlanAmort.free;
  end;
  end ;
end;

procedure AnnuleChangementPlan(LogChPlan : TLogChPlan);
var
  PlanActif,CodeToDel : string;
  QImmo : TQuery;
  Plan : TPlanAmort; // CA le 26/01/1999 pour recalcul des dates de dernières dotations
  eco,fisc : double;
  NumVersion : string;
begin
  QImmo:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+LogChPlan.LeQuel+'"', FALSE) ;
  QImmo.Edit ;
  CodeToDel:=QImmo.FindField('I_IMMO').AsString ;
  PlanActif:=QImmo.FindField('I_PLANACTIF').AsString ;
  ExecuteSQL('DELETE FROM IMMOAMOR WHERE IA_IMMO="'+CodeToDel+'" AND IA_NUMEROSEQ='+PlanActif) ;
  QImmo.FindField('I_MONTANTEXC').AsFloat:=LogChPlan.MontantExc;
  QImmo.FindField('I_TYPEEXC').AsString:=LogChPlan.TypeExc;
  QImmo.FindField('I_REVISIONECO').AsFloat:=LogChPlan.RevisionEco;
  QImmo.FindField('I_REVISIONFISCALE').AsFloat:=LogChPlan.RevisionFisc;
  QImmo.FindField('I_PLANACTIF').AsInteger:=LogChPlan.PlanActifAv;
  QImmo.FindField('I_DUREEECO').AsInteger:=LogChPlan.DureeEco;
  QImmo.FindField('I_DUREEFISC').AsInteger:=LogChPlan.DureeFisc;
  QImmo.FindField('I_METHODEECO').AsString:=LogChPlan.MethodeEco;
  QImmo.FindField('I_METHODEFISC').AsString:=LogChPlan.MethodeFisc;
  // CA - 19/08/2004 - Astuce pour changement plan NAM
  if LogChPlan.MethodeEco='NAM' then
  begin
    QImmo.FindField('I_DATEPIECEA').AsDateTime := LogChPlan.DatePieceA;
    QImmo.FindField('I_DATEAMORT').AsDateTime := LogChPlan.DateAmort;
    // BTY 11/06 Retrouver la durée d'inaliénabilité en cas de prime ou subvention
    QImmo.FindField('I_CORVRCEDDE').AsFloat := LogChPlan.DureePriSbv;
    // BTY 11/06 Supprimer les éventuels antérieurs sur le plan fiscal qui a pu être généré
    QImmo.FindField('I_REPRISEFISCAL').AsFloat:=LogChPlan.RepriseFisc;
  end;
  // BTY 11/05 Si plan fiscal CRC 2002-10 et/ou calcul plan futur sur VNC
  // Dans tous les cas, récupérer le précédent i_journala
  // car possible de réviser avec plan futur sur VNC (CD3) l'année suivant
  // une révision sans plan futur sur VNC (CDM)
  QImmo.FindField('I_JOURNALA').AsString:=LogChPlan.CodeEclat;

  NumVersion := LogChPlan.NumVersion;
  NumVersion := Copy (NumVersion, 1, POS('.',NumVersion)-1);

  {  if ((LogChPlan.TypeOpe='CD2') or (LogChPlan.TypeOpe='CD3')) then }
  if (LogChPlan.TypeOpe <> 'ELE') then
  begin
    // 09/07 Restituer l'ancienne valeur sinon va rester à sa nelle valeur
    // FQ 21523 Toujours restituer l'ancienne base fiscale
    // if (LogChPlan.MethodeFisc <> '') then
       QImmo.FindField('I_BASEFISC').AsFloat:=LogChPlan.BaseFiscAvMB;

    {if (LogChPlan.TypeOpe='CD2') then }
    // Présence plan fiscal CRC 2002-10 sans calcul plan futur sur VNC
    QImmo.FindField('I_REPRISEECO').AsFloat:=LogChPlan.RepriseEco;
    QImmo.FindField('I_REPRISEFISCAL').AsFloat:=LogChPlan.RepriseFisc;
    {end;}

    // 04/07 Gestion fiscale
    // 04/07 Plan futur avec la VNF si version 8 ou supérieure
    if (StrToInt ( Copy (NumVersion,1,1)) >= 8) then
    begin
      QImmo.FindField('I_NONDED').AsString:=LogChPlan.GestionFiscale;
      QImmo.FindField('I_FUTURVNFISC').AsString:=LogChPlan.CalculVNF;
    end;

    // Antérieurs dérogatoires
    QImmo.FindField('I_REPRISEDR').AsFloat := LogChPlan.RepriseDR;
    QImmo.FindField('I_REPRISEFEC').AsFloat := LogChPlan.RepriseFEC;
  end;


  //  mb fq 17569 eco := GetTaux(QImmo.FindField('I_METHODEECO').AsString, QImmo.FindField('I_DATEAMORT').AsDateTime,QImmo.FindField('I_DATEPIECEA').AsDateTime, QImmo.FindField('I_DUREEECO').AsInteger);
  // fisc := GetTaux(QImmo.FindField('I_METHODEFISC').AsString, QImmo.FindField('I_DATEAMORT').AsDateTime,QImmo.FindField('I_DATEPIECEA').AsDateTime, QImmo.FindField('I_DUREEFISC').AsInteger);
  eco := GetTaux(QImmo.FindField('I_METHODEECO').AsString, QImmo.FindField('I_DATEDEBECO').AsDateTime,QImmo.FindField('I_DATEDEBECO').AsDateTime, QImmo.FindField('I_DUREEECO').AsInteger);
  fisc := GetTaux(QImmo.FindField('I_METHODEFISC').AsString, QImmo.FindField('I_DATEDEBFIS').AsDateTime,QImmo.FindField('I_DATEDEBFIS').AsDateTime, QImmo.FindField('I_DUREEFISC').AsInteger);
  QImmo.FindField('I_TAUXECO').AsFloat:=eco;
  QImmo.FindField('I_TAUXFISC').AsFloat:=fisc;

  QImmo.FindField('I_TYPEDEROGLIA').AsString := TypeDerogatoire( nil, QImmo); // FQ 20256


  // Attention : ruse pour ne passer qu'un paramètre pour IL_TYPEOP
  // MajOpeEnCoursImmo ( QImmo, 'I_OPECHANGEPLAN','ELE" OR IL_TYPEOP="CDM" OR IL_TYPEOP="ELC','-');
  // BTY 11/05 Prendre en compte les cas CD2 et CD3
  // BTY 01/06 Décocher I_OPECHANGEPLAN si une seule opération parmi ELE/CDM/CD2/CD3/CDA/ELC/DPR/MMS
  // MBO 04/06 Ajout dans les opérations MMS fq 17923
  MajOpeEnCoursImmo ( QImmo,
                'I_OPECHANGEPLAN',
//                     'ELE" OR IL_TYPEOP="CDM" OR IL_TYPEOP="CD2" OR IL_TYPEOP="CD3" OR IL_TYPEOP="ELC',
               'ELE" OR IL_TYPEOP="MMS" OR IL_TYPEOP="CDM" OR IL_TYPEOP="CD2" OR IL_TYPEOP="CD3" OR IL_TYPEOP="ELC" OR IL_TYPEOP="CDA" OR IL_TYPEOP="DPR',
                     '-');
  Plan:=TPlanAmort.Create(true) ;// := CreePlan(True);
  try
    Plan.Recupere(LogChPlan.LeQuel,IntToStr(LogChPlan.PlanActifAv));
    QImmo.FindField('I_DATEDERMVTECO').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortEco);
    QImmo.FindField('I_DATEDERNMVTFISC').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortFisc);
  finally
    Plan.free ; //Detruit;
  end ;
  QImmo.Post ;
  Ferme(QImmo);
  // Ajout de l'ordre dans la requête pour ne pas détruire  toutes les opérations liées et enchaîner ainsi naturellement
  ExecuteSQL('DELETE FROM IMMOLOG WHERE IL_IMMO="'+CodeToDel+'" AND IL_PLANACTIFAP='+PlanActif+' AND IL_ORDRE='+IntToStr(LogChPlan.Ordre)) ;
end;

function OpeChangementPlan(TypeOpe : string) : boolean;
begin
  //result := (TypeOpe = 'ELE') or (TypeOpe = 'CDM') or (TypeOpe = 'ELC');
  // BTY 11/05 Nouveaux cas CD2 et CD3 (calcul plan futur sur VNC et plan fiscal CRC 2002-10)
  result := (TypeOpe = 'ELE') or (TypeOpe = 'CDM') or (TypeOpe = 'CD2') or (TypeOpe = 'CD3') or (TypeOpe = 'ELC');
end;


procedure TChangePlanAmort.FormShow(Sender: TObject);
var signe : integer; (*j : integer;*) masque : string ; //YCP 21/10/05
    Duree : string;
    // NbMois: Word;   BTY 01/06
    // NJour: integer;  BTY 01/06
begin
  inherited;
  QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+CodeImmo+'"', FALSE) ;

  // 04/07 Query qui reflète les données saisies en temps réel
  QPlanCalcul:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+CodeImmo+'"', FALSE) ;

  AncienPlan.Charge(QPlan);
  AncienPlan.Recupere(CodeImmo,QPlan.FindField('I_PLANACTIF').AsString);
  // BTY 11/05 Stocker les (antérieurs saisis + dotation de l'exo) ECO et FISCAL
  AncienPlan.GetCumulsDotExercice(VHImmo^.Encours.Deb, fCumulEco,fCumulFisc,True,True,true);
  Plan.copie(AncienPlan);
  PlanCalcul.copie(AncienPlan);
  sCodeImmo.Caption := QPlan.FindField('I_IMMO').AsString;
  sLibelleImmo.Caption := QPlan.FindField('I_LIBELLE').AsString;
  sDateAchat.Caption := QPlan.FindField('I_DATEPIECEA').AsString;
  sValeurHT.Caption := StrFMontant(QPlan.FindField('I_MONTANTHT').AsFloat,15,V_PGI.OkDecV,'',True);
  DateAchat := QPlan.FindField('I_DATEPIECEA').AsDateTime;
  if QPlan.FindField('I_TYPEEXC').AsString = 'RDO' then signe:=-1 else signe:=1;
  ElemExc := QPlan.FindField('I_MONTANTEXC').AsFloat * signe;
  BaseEco := QPlan.FindField('I_BASEECO').AsFloat;
  RepriseEco := QPlan.FindField('I_REPRISEECO').AsFloat;
  RepriseDep := QPlan.FindField('I_REPRISEDEP').AsFloat;  // impact antérieur dépréciation mbo 02.06

  // ajout mbo 30.09.05 - fq 15280
  BaseFisc := QPlan.FindField('I_BASEFISC').AsFloat;
  RepriseFisc := QPlan.FindField('I_REPRISEFISCAL').AsFloat;
  planFiscal := QPlan.FindField('I_METHODEFISC').AsString;
  PlanEco := QPlan.FindField('I_METHODEECO').AsString;
  if (PlanFiscal <> '') AND (PlanEco <> 'VAR') then derog:= true;

  // Test immo non amortissable
  fNAM := (PlanEco = 'NAM');
  // FQ 18393
  if fDateSerie <> iDate1900 then
     DATE_OPE.Text := DateToStr(fDateSerie);
  // BTY 10/06 Prime d'équipement
  fSBVPRI := QPlan.FindField('I_SBVPRI').AsFloat;
  fSBVMT := QPlan.FindField('I_SBVMT').AsFloat;
  // Durée d'inaliénabilité
  Duree := FloatToStr (QPlan.FindField('I_CORVRCEDDE').AsFloat);
  fInalienabilite := StrToInt (Duree);

  bFuturVNF := (QPlan.FindField ('I_FUTURVNFISC').AsString = '***');
  // modif mbo 12.06.07 fRemplace := QPlan.FindField ('I_REMPLACE').AsString;
  fRemplace := QPlan.FindField ('I_STRING1').AsString;
  bDPI := (QPlan.FindField ('I_DPI').AsString = 'X');
  fRepriseDR := QPlan.FindField ('I_REPRISEDR').AsFloat;
  fRepriseFEC := QPlan.FindField ('I_REPRISEFEC').AsFloat;

  bAjoutDotation.Checked := true;
  VersLaFiche;

  FocusControl(DATE_OPE);
  if (TypeChangement = 2) then HelpContext := 2111500
  else HelpContext := 2111400;
  PGeneral.TabVisible:=TypeChangement<>2;
  PDotationExcep.TabVisible:=not PGeneral.TabVisible ;
  if TypeChangement=1 then
  begin
    Caption := Caption + HM.Mess[5];
    Pages.ActivePage := PGeneral;
  end
  else
  begin
    Caption := Caption + HM.Mess[6];
    Pages.ActivePage := PDotationExcep;
  end;

  // FQ 16286 TGA 28/07/2005
  sValeurHt.displayformat := StrfMask(V_PGI.OkDecV,'', True);

  // BTY 11/05 Ajout choix durée restante et gestion plan fiscal CRC 200210
  // en révision de plan
  // BTY 11/05 Ajout choix durée restante en saisie exceptionnel
  // bDureeRest.Visible := ((TypeChangement=1) and ((PlanEco ='LIN') or (PlanEco ='VAR')) );
  if (TypeChangement=1) then
  begin
    bDureeRest.Visible := ((PlanEco ='LIN') or (PlanEco ='VAR'));
    bDureeRest.Enabled := bDureeRest.Visible;
    bDureeRest.Checked := ((bDureeRest.Enabled) and (QPlan.FindField('I_JOURNALA').AsString='***'));
  end else
  begin
    bDureeRestE.Visible := ((PlanEco ='LIN') or (PlanEco ='VAR'));
    bDureeRestE.Enabled := bDureeRestE.Visible;
    bDureeRestE.Checked := ((bDureeRestE.Enabled) and (QPlan.FindField('I_JOURNALA').AsString='***'));
  end;

  // BTY 11/06 CheckB visible si prime ou subvention
  {bPlanCRC.Visible := ((TypeChangement=1) and (PlanFiscal = '') and
                       ((PlanEco = 'LIN') or (PlanEco = 'DEG') or (fSBVPRI <> 0) or (fSBVMT <> 0)) );
  bPlanCRC.Enabled := False;
  // BTY 11/06 Check bloquée et cochée si prime ou subvention
  if (fSBVPRI <> 0) or (fSBVMT <> 0)  then
      bPlanCRC.Checked := True;
  stTypeBaseFisc.Visible := False;
  cTypeBaseFisc.Visible := False; }

  // BTY 12/05 Ajout champs durée déjà amortie et valeur résiduelle à fin N-1
  AfficheDureeDejaAmortie;
  AfficheVNCDebutExo;


  if (TypeChangement=2) then
  begin
    DotationTotale.caption:= StrFMontant(DotationExe,15,V_PGI.OkDecV,'',True);
    Nouvelledot.caption := StrFMontant(DotationExe,15,V_PGI.OkDecV,'',True);
  end;

  // FQ 16286 TGA 31/08/2005 mise en commentaire
  // lBaseEco.caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
  // lBasefisc.caption:= StrFMontant(0.00,15,V_PGI.OkDecV,'',True);

  // mbo - FQ 16286 08.09.2005-
  // mbo - fq 16608 suite 16286 - 12.09.05 - remise à plat des masques (pb 3 décimales)

  Masque:= StrfMask(V_PGI.OkDecV, '', True);

  Vnc.Masks.PositiveMask:= Masque;
  Vnc.Masks.NegativeMask := Masque;
  Vnc.Masks.ZeroMask := Masque;

  DotationExc.Masks.PositiveMask := Masque;
  DotationExc.Masks.NegativeMask := Masque;
  DotationExc.Masks.ZeroMask := Masque;

  DontDotexc.masks.PositiveMask := Masque;
  DontDotexc.masks.NegativeMask := '-' + Masque;
  DontDotexc.masks.ZeroMask := Masque

end;


procedure TChangePlanAmort.ModifPlanSurChangeExc;
var
  signe : integer;
begin
  signe := 1;
  if (bAjoutDotation.Checked) then
  begin
    TypeDot := 'DOT';
    Plan.SetMontantExc(Plan.MontantExc+DOTATIONEXC.Value);
    MontantDot := Valeur(DOTATIONEXC.Text) * signe;
    Plan.SetTypeExc('DOT');
  end
  else if bRepriseDotation.Checked then
  begin
    TypeDot := 'RDO';
//    signe := -1;
//    Plan.SetMontantExc(Plan.MontantExc-DOTATIONEXC.Value);
//    MontantDot := Valeur(DOTATIONEXC.Text) * signe;
    MontantDot := Valeur(DOTATIONEXC.Text);
    Plan.SetMontantExc(MontantDot);
    Plan.SetTypeExc('RDO');
  end
  else if bRemplaceDotation.Checked then
  begin
    TypeDot := 'SDO';
//    Plan.SetMontantExc(DOTATIONEXC.Value);
    MontantDot := (DOTATIONEXC.Value - Valeur(DotationTotale.Caption)) * signe;
    if MontantDot < 0 then  Plan.SetMontantExc ((-1)*MontantDot)
    else Plan.SetMontantExc (MontantDot);
    if (MontantDot < 0) then Plan.SetTypeExc('RDO')
    else Plan.SetTypeExc('DOT');
  end;
//  if (Plan.MontantExc < 0) and (Plan.TypeOpe<>'SDO') then Plan.SetTypeExc('RDO')
//  if (Plan.MontantExc < 0) then Plan.SetTypeExc('RDO')
//  else Plan.SetTypeExc(TypeDot);
//  else Plan.SetTypeExc('DOT');
  Plan.SetTypeOpe(TypeDot);
  RevisionEco := MontantDot;
  RevisionFisc := 0.00;
end;

procedure TChangePlanAmort.MajMontantDotation(Sender: TObject);
begin
  inherited;
  // mbo 12.09.05 - suite FQ 16608
  // pb en saisie dossier avec 3 décimales :

  //TGA 29/11/2005 (pb du à l'appel par onchange)
  DOTATIONEXC.Value:=Valeur(DOTATIONEXC.Text);

  GereDetailDotation;
end;

function TChangePlanAmort.GereDetailDotation : TModalResult;
var signe : integer; NelleDot, DontExc, DotExercice : double; Okok: boolean ;
  MontantMaxi : double;
begin
  Signe:=1; NelleDot:=0; DontExc:=0;
  // modif mbo 30.09.2005 - 15280 DotAnterieur:=DotEco;
  DotExercice :=DotationExe;  //ajout mbo fq 15280

Okok:=false ;

// Ajout du test sur l'onglet actif : mbo 26.10.06 pour éviter en révision de plan des messages incohérents
if Pages.ActivePage = PDotationExcep then
begin
  if (DOTATIONEXC.Value<0) then HM.execute(3,Caption,'')
  else
  begin
     {modif mbo 11.06.07 pour améliorer le message affiché en cas d'erreur
     if  bAjoutDotation.Checked and (Arrondi(DOTATIONEXC.Value,V_PGI.OkDecV)>MaxAjout) then HM.execute(11,Caption,FloatToStr(MaxAjout))
     else if bRepriseDotation.Checked and (Arrondi(DOTATIONEXC.Value,V_PGI.OkDecV)>MaxReprise) then HM.execute(11,Caption,FloatToStr(MaxReprise))
     else if bRemplaceDotation.Checked and (Arrondi(DOTATIONEXC.Value,V_PGI.OkDecV)>MaxRemplace) then HM.execute(11,Caption,FloatToStr(MaxRemplace))
     else okok := true;
     }
     if  bAjoutDotation.Checked then MontantMaxi := MaxAjout
     else if bRepriseDotation.Checked then MontantMaxi := MaxReprise
     else MontantMaxi := MaxRemplace;

     if MontantMaxi = 0 then
     begin
        if VNCFinex = 0 then
           HM.execute(19,Caption,'')
        else
        begin
           if (derog) and (VRF=0) then
              HM.execute(20,Caption,'');
        end;
     end else
     begin
        if Arrondi(DOTATIONEXC.Value,V_PGI.OkDecV)>MontantMaxi then
           HM.execute(11,Caption,FloatToStr(MontantMaxi))
        else
           okok := true;
     end;
  end;

  if Okok then
    begin
    if bAjoutDotation.Checked then
      begin
      NelleDot:=DotationExe+DOTATIONEXC.Value;
      DontExc:=DOTATIONEXC.Value+Plan.MontantExc;
      end
    else if bRepriseDotation.Checked then
      begin
      signe:=-1;
      NelleDot:=DotationExe-DOTATIONEXC.Value;
      DontExc:=-DOTATIONEXC.Value+Plan.MontantExc;
      end
    else if bRemplaceDotation.Checked then
      begin
      //DotAnterieur:=0.00;
      DotExercice:=0.00;
      NelleDot:=DOTATIONEXC.Value;
      DontExc:=DOTATIONEXC.Value-DotationExe+Plan.MontantExc;
      end;
    NOUVELLEDOT.Caption:=StrFMontant(NelleDot,15,V_PGI.OkDecV,'',True);
    DONTDOTEXC.Value:=DontExc;
    VNC.Value:=InitVNC-DotExercice-(DOTATIONEXC.Value*signe);
    result:=mrOk;
    end
  else

    begin
    // ajout mbo FQ 15280 - 30.09.2005
    // si erreur de saisie on réaffiche avec les montants d'origine
    VNC.Value := InitVNC-DotationExe;
    DOTATIONTOTALE.Caption := StrFMontant(DotationExe,15,V_PGI.OkDecV,'',True);;
    DONTDOTEXC.Value := 0.00;
    NOUVELLEDOT.Caption := StrFMontant(DotationExe,15,V_PGI.OkDecV,'',True);;

    // Fin ajout mbo
    FocusControl(DOTATIONEXC);
    result:=mrNone;
    end ;
end else
begin
  result := mrOk;
end;
end;


procedure TChangePlanAmort.TypeDotationClick(Sender: TObject);
begin
  inherited;
  if bAjoutDotation.Checked then DOTATIONEXC.Value:=0.00
  else if bRepriseDotation.Checked then DOTATIONEXC.Value:=0.00
  else if bRemplaceDotation.Checked then DOTATIONEXC.Value:=DotationExe;  //=DotEco ; modif mbo 15280
  GereDetailDotation;
end;


procedure TChangePlanAmort.BValiderClick(Sender: TObject);
begin
  ModalResult := mrNone;

  //TGA 29/11/2005
  DOTATIONEXC.Value:=Valeur(DOTATIONEXC.Text);

  TraiteFinChangePlan;
end;

procedure TChangePlanAmort.VersLaFiche;
begin
  if TypeChangement=1 then
  begin
    // BTY FQ 19095 Interdire de passer le mode ECO à Non amortissable
    // Pour une immo déjà non amortissable, laisser l'item pour pouvoir l'afficher en entrée
    // et bloquer la durée
    if not fNAM then
       cMethodeEco.Items.Delete(cMethodeEco.Items.IndexOf('Non amortissable'))
    else
       DureeEco.Enabled := False;

    cMethodeFiscale.Items.Delete(cMethodeFiscale.Items.IndexOf('Non amortissable'));

    // 06/07 Composant remplaçant => affichage TR2 dans la base fiscale sinon la retirer
    // FQ 21502 Items.Delete met du désordre dans les valeurs
    // Attention Plus réinitialise Itemindex à -1 => assigner Plus avant Itemindex
    if (fRemplace='') then
    begin
      //cTypeBaseFisc.Items.Delete(cTypeBaseFisc.Items.IndexOf(RechDom('TITYPEBASEFISC', 'TR2', FALSE)))
      cTypeBaseFisc.Plus := ' AND CO_CODE<>"TR2"';
    end else
    begin
      cTypeBaseFisc.ItemIndex := cTypeBaseFisc.Items.IndexOf(RechDom('TITYPEBASEFISC', 'TR2', FALSE));
      cTypeBaseFisc.Enabled := False;
    end;

    RecupAncienPlanFiscal;

    cMethodeEco.ItemIndex := cMethodeEco.Items.IndexOf(RechDom('TIMETHODEIMMO', AncienPlan.AmortEco.Methode, FALSE)) ;
    DureeEco.Value := AncienPlan.AmortEco.Duree;
    TauxEco.Value := AncienPlan.AmortEco.Taux*100;
    fValeurTheorique := Arrondi (AncienPlan.ValeurHT+AncienPlan.ValeurTvaRecuperable-AncienPlan.ValeurTvaRecuperee,V_PGI.OkDecV);

    bAmortissementFiscal.Checked := AncienPlan.Fiscal;
    bAmortissementFiscal.Enabled := (not AncienPlan.Fiscal);
    lBaseEco.Caption := StrFMontant(AncienPlan.AmortEco.Base,15,V_PGI.OkDecV,'',True);
    Pages.ActivePage := PGeneral;


    // 04/07 Plan fiscal automatiquement ajouté si prime ou subvention
    if (not AncienPlan.Fiscal) and ((fSBVPRI <> 0) or (fSBVMT <> 0))   then
    begin
       bAmortissementFiscal.Checked := True;
       bAmortissementFiscal.Enabled := False;
       A.Enabled := False;
       bDureeRestF.Visible := bDureeRest.Visible;
       bDureeRestF.Enabled := bDureeRest.Enabled;
       bDureeRestF.Checked := bDureeRest.Checked;
       bGestionFiscale.Enabled := False;
    end;
  end;
  AfficheOngletExceptionnel(AncienPlan);
end;

// Application de l'ancien plan fiscal dans le groupe A fiscal
procedure TChangePlanAmort.RecupAncienPlanFiscal;
begin
  A.Enabled:= AncienPlan.Fiscal;
  cMethodeFiscale.ItemIndex := cMethodeFiscale.Items.IndexOf(RechDom('TIMETHODEIMMO', AncienPlan.AmortFisc.Methode, FALSE)) ;
  DureeFiscale.Value := AncienPlan.AmortFisc.Duree;
  TauxFiscal.Value := AncienPlan.AmortFisc.Taux*100;
  lBaseFisc.Caption := StrFMontant(AncienPlan.AmortFisc.Base,15,V_PGI.OkDecV,'',True);

  bDureeRestF.Visible := ((PlanFiscal ='LIN') or (PlanFiscal ='VAR'));
  bDureeRestF.Enabled := bDureeRestF.Visible;
  bDureeRestF.Checked := ((bDureeRestF.Enabled) and (bFuturVNF));
  bGestionFiscale.Checked := AncienPlan.Gestionfiscale;
  bGestionFiscale.Enabled := False;

  if AncienPlan.Fiscal then
  begin
     // 06/07 Afficher TR2 si composant remplaçant
     if fRemplace = 'X' then
     begin
        cTypeBaseFisc.Itemindex:=cTypeBaseFisc.Items.indexof(RechDom('TITYPEBASEFISC','TR2',FALSE));
        cTypeBaseFisc.Enabled := False;
     end else
     begin
       if AncienPlan.AmortFisc.Base = AncienPlan.AmortEco.Base then
          cTypeBaseFisc.Itemindex:=cTypeBaseFisc.Items.indexof(RechDom('TITYPEBASEFISC','ECO',FALSE))
       else
          cTypeBaseFisc.Itemindex:=cTypeBaseFisc.Items.indexof(RechDom('TITYPEBASEFISC','THE',FALSE));
     end;
  end
  else
  begin
     cMethodeFiscale.Value := '';
     cTypeBaseFisc.Value := '';
  end;

  // Imoo est un composant remplaçant => bloquer le plan fiscal
  if (AncienPlan.Fiscal) and (fRemplace<> '') then
  begin
     A.Enabled := False;
     bDureeRestF.Enabled := False;
     bGestionFiscale.Enabled := False;
  end;
end;


procedure TChangePlanAmort.ModifPlanSurChangeDurMeth;
var  CasCD2 : boolean;
begin
    Plan.AmortEco.ModifieMethodeAmort(Plan.AmortEco.DateFinAmort,
                                       Plan.AmortEco.Revision,
                                       Valeur(lBaseEco.Caption),
                                       TauxEco.Value/100,
                                       DureeEco.Value,
                                       cMethodeEco.Value,
                                       Plan.ImmoCloture);
    Plan.AmortFisc.ModifieMethodeAmort(Plan.AmortFisc.DateFinAmort,
                                        Plan.AmortFisc.Revision,
                                        Valeur(lBaseFisc.Caption),
                                        TauxFiscal.Value/100,
                                        DureeFiscale.Value,
                                        cMethodeFiscale.Value,
                                        Plan.ImmoCloture);

    // 04/07 Test ajout d'un plan fiscal
    // <=> opération CD2 qui nécessitera le recalcul des antérieurs éco :
    // soit antérieurs saisis + antérieurs clôturés
    // soit antérieurs saisis recalculés + antérieurs clôturés
    {CasCD2 :=  ((not AncienPlan.Fiscal) and (bDureeRest.Enabled) and (not bDureeRest.Checked)
                and (cMethodeFiscale.Value = AncienPlan.AmortEco.methode)
                and (DureeFiscale.Value = AncienPlan.AmortEco.duree)
                and (Arrondi (TauxFiscal.Value, 2) = Arrondi(AncienPlan.AmortEco.taux * 100, 2))
                and (Arrondi (Valeur(lBaseFisc.Caption), V_PGI.OkDecV) =
                     Arrondi (AncienPlan.AmortEco.base, V_PGI.OkDecV))
                and (DureeEco.Value > AncienPlan.AmortEco.duree) ); }
    CasCD2 :=  ((not AncienPlan.Fiscal) and (cMethodeFiscale.Value <> '') );

    //Plan.SetTypeOpe('CDM');
    // BTY 11/05 Ajout Calcul plan futur sur VNC et Plan fiscal CRC 2002-10
    {if bDureeRest.Checked then Plan.SetTypeOpe('CD3')
    else if bPlanCRC.Checked then Plan.SetTypeOpe('CD2')
         else Plan.SetTypeOpe('CDM'); }
    // 04/07 Nelle gestion fiscale
    {if bDureeRest.Checked then Plan.SetTypeOpe('CD3')
    else if CasCD2 then Plan.SetTypeOpe('CD2')
         else Plan.SetTypeOpe('CDM');}
    if CasCD2 then Plan.SetTypeOpe('CD2')
    else if bDureeRest.Checked then Plan.SetTypeOpe('CD3')
         else Plan.SetTypeOpe('CDM');
end;



procedure TChangePlanAmort.TraiteFinChangePlan;
var Sender: Tobject;
    NumMess : integer;
    GestionFiscaleCochee : boolean;
begin

if GereDetailDotation=mrOk then
  begin

  // BTY 12/05 FQ 17178 Envoyer les OnExit au cas où on n'est pas sorti du champ courant
  if TypeChangement=1 then
     begin
     GestionFiscaleCochee := (bGestionFiscale.Enabled and bGestionFiscale.Checked);
     Sender := nil;
     AmortEcoExit(Sender);
     DureeEcoExit(Sender);
     AmortFiscExit(Sender);

     DureeRestExit(Sender);
     DureeRestFExit(Sender);

     // Coche Gestion fiscale cochée par l'utilisateur
     // La vérif théorique AfficheGestionFiscale recalcule et peut décocher systématiquement
     // => conserver ici la valeur si saisie par l'utilisateur
     if GestionFiscaleCochee and bGestionFiscale.Enabled then
        bGestionFiscale.Checked := True;
     end;

  // FQ 19095
  if fNAM and (cMethodeEco.Value='NAM') then
  begin
    HM.Execute(17, Caption, '');
    Pages.ActivePage := PGeneral; FocusControl(cMethodeEco);
  end else
  // FQ 19080
  if (not TraiteIntervaleDureeAmort(cMethodeEco.Value,0,DureeEco.Value,NumMess)) then
  begin
    HM.Execute(NumMess, Caption, '');
    Pages.ActivePage := PGeneral; FocusControl(DureeEco);
  end else if (bAmortissementFiscal.Checked) and
          (not TraiteIntervaleDureeAmort(cMethodeFiscale.Value,0,DureeFiscale.Value,NumMess)) then
  begin
    HM.Execute(NumMess, Caption, '');
    Pages.ActivePage := PGeneral; FocusControl(DureeFiscale);
  end else

  if DateAchat>StrToDate(DATE_OPE.Text) then
    begin
    HM.Execute(9, Caption, '');
    FocusControl(DATE_OPE);
    end
  else if (StrToDate(DATE_OPE.Text)<VHImmo^.Encours.Deb) or (StrToDate(DATE_OPE.Text)>VHImmo^.Encours.Fin) then
    begin
    HM.Execute (13,Caption,'');
    FocusControl (DATE_OPE);
    end
  else if ExisteSQL('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'+CodeImmo+'" AND IL_DATEOP>"'+USDATETIME(StrToDate(DATE_OPE.Text))+'"') then
    begin
    HM.Execute (10,Caption,'');
    FocusControl (DATE_OPE);
    end

  else if (TypeChangement=1) then
    begin
    if (DureeEco.Text='') then
       begin
       PGIInfo('Veuillez renseigner la durée économique.');
       FocusControl(DureeEco);
       end
    else if (bAmortissementFiscal.Checked) and (DureeFiscale.Text='') then
       begin
       PGIInfo('Veuillez renseigner la durée fiscale.');
       FocusControl(DureeFiscale);
       end
    else if (not DureeValide( cMethodeEco.Value, DureeEco.Value, False ))  then
       begin
       PGIInfo('La durée d''amortissement économique doit être supérieure à la durée déjà amortie.');
       FocusControl(DureeEco);
       end
    // BTY 12/05 FQ 17106 Ne pas contrôler la durée fiscale si le mode fiscal existait déjà
    // 04/07 Nelle gestion fiscale => toujours contrôler
    else if (bAmortissementFiscal.Checked) and
         (not DureeValide( cMethodeFiscale.Value, DureeFiscale.Value, True )) then
         //(and not AncienPlan.Fiscal) then
       begin
       PGIInfo('La durée d''amortissement fiscal doit être supérieure à la durée déjà amortie.');
       FocusControl(DureeFiscale);
       end
    // BTY 12/05 FQ 17174 Erreur ancienne : taux ECO pouvait être > taux FISCAL
    // 04/07 Nelle gestion fiscale => contrôler le taux si "Gestion fiscale" n'est pas cochée
    else if (bAmortissementFiscal.Checked) and  (TauxEco.Value > TauxFiscal.Value)
         and (not bGestionFiscale.Checked) then
       begin
       HM.Execute(16, Caption, '');
       FocusControl (DureeEco);
       end

    // test 2 plans identiques
    else if  (cMethodeFiscale.Value = cMethodeEco.Value)
         and (DureeFiscale.Value = DureeEco.Value)
         and (TauxFiscal.Value = TauxEco.Value)
         and (Valeur(lBaseFisc.Caption) = Valeur(lBaseEco.Caption))
         and (bDureeRestF.Checked = bDureeRest.Checked)
         and (CalculAnterieursOK)
         and (not bGestionFiscale.Checked) then
       begin
       HM.Execute(18, Caption, '');
       FocusControl (cMethodeEco);
       end
    else
       begin
       ModifPlanSurChangeDurMeth;// changement durée/Méthode
       EnregistreChangePlan;
       ModalResult:=mrYes ;
       end

    end
    else  // typechangement=2
       begin
       ModifPlanSurChangeExc;  // Elements exceptionnels
       EnregistreChangePlan;
       ModalResult:=mrYes ;
       end;

  end ;
end;

procedure TChangePlanAmort.EnregistreChangePlan;
begin
  ModalResult := mrNo;
  if Transactions(EnregChangePlan, 1)<>oeOK then HM.execute(1,Caption,'')
  else
  begin
    ModalResult := mrYes;
    VHImmo^.ChargeOBImmo := True;
    ImMarquerPublifi (True);   // CA le 06/10/2000
  end;
end;

procedure TChangePlanAmort.EnregChangePlan;
begin
  DateOpe := StrToDate(DATE_OPE.EditText); // CA le 20/01/99
  EcritOrigImmo;
  OrdreS :=-1; // pour forcer le calcul du Numero d'ordre en serie
  EcritLogChangePlan;
  if (TypeChangement=1) and (DOTATIONEXC.Value<>0) then
  begin
    AncienPlan.copie(Plan);     // Nouveau "Ancien Plan"
    TypeChangement:=2;          // simulation Elements exceptionnels
    ModifPlanSurChangeExc;      // on recupere l'element exceptionnel eventuel Revision
    Plan.AmortEco.ModifieMethodeAmort(Plan.AmortEco.DateFinAmort,RevisionEco,Valeur(lBaseEco.Caption),
                                TauxEco.Value/100,DureeEco.Value,cMethodeEco.Value,Plan.ImmoCloture);
    Plan.AmortFisc.ModifieMethodeAmort(Plan.AmortFisc.DateFinAmort,RevisionFisc,Valeur(lBaseFisc.Caption),
                                TauxFiscal.Value/100,DureeFiscale.Value,cMethodeFiscale.Value,Plan.ImmoCloture);
    EcritOrigImmo;
    EcritLogChangePlan;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 26/03/2004
Modifié le ... :   /  /
Description .. : Mise à jour d'une immobilisation suite à une revision du
Suite ........ : plan ou à une dotation exceptionnelle.
Suite ........ : FQ 13063 - Dans le cas d'une révision sur une
Suite ........ : immobilisation non amortissable, on effectue tous les calculs
Suite ........ : par rapport à la date d'opération.
Mots clefs ... :
*****************************************************************}
procedure TChangePlanAmort.EcritOrigImmo;
var Plan2 : TPlanAmort;
    wReCalcul : boolean;
    // mbo 11.06.07 pour messages à la compil wCumulEco, wCumulFiscal : double;
    TypeOpe : string;
begin
  QPlan.Edit;

  if TypeChangement=1 then
  begin
    { Mémorisation des dates d'acquisition et mise en service de l'immobilisation }
    if (QPlan.FindField('I_METHODEECO').AsString='NAM') then
    begin
      // FQ 19064
      { Pour faire les calculs, on se base sur la date d'opération }
      //QPlan.FindField('I_DATEPIECEA').AsDateTime := DateOpe;
      //QPlan.FindField('I_DATEAMORT').AsDateTime := DateOpe;
      if cMethodeEco.Value = 'DEG' then
         QPlan.FindField('I_DATEPIECEA').AsDateTime := DateOpe
      else
         QPlan.FindField('I_DATEAMORT').AsDateTime := DateOpe
    end;
    TypeOpe := Plan.TypeOpe;

    // Plan fiscal ajouté => retasser antérieurs éco et peut-être les recalculer
    if (Plan.TypeOpe = 'CD2') then
    begin
       Plan2:=TPlanAmort.Create(true);

       // Compatibilité avec version 7 FQ 17373 :
       // 06/07 Recalcul antérieurs saisis si plan fiscal ajouté = plan ECO,
       // et si durée éco augmentée à  méthode constante ou  méthode DEG->LIN à durée constante   // 06/07
       // et si pas de plan futur VNC
       wReCalcul := ((not AncienPlan.Fiscal) and (not bDureeRest.Checked)
                 and (cMethodeFiscale.Value = AncienPlan.AmortEco.methode)
                 and (DureeFiscale.Value = AncienPlan.AmortEco.duree)
                 and (Arrondi (TauxFiscal.Value, 2) = Arrondi(AncienPlan.AmortEco.taux * 100, 2))
                 and (Arrondi (Valeur(lBaseFisc.Caption), V_PGI.OkDecV) =
                      Arrondi (AncienPlan.AmortEco.base, V_PGI.OkDecV))
                 and ( ((DureeEco.Value > AncienPlan.AmortEco.duree)
                         and (cMethodeEco.Value = AncienPlan.AmortEco.methode))
                   or  ((DureeEco.Value = AncienPlan.AmortEco.duree)
                         and ((cMethodeEco.Value = 'LIN')
                         and (AncienPlan.AmortEco.methode = 'DEG')))) );

       if (wReCalcul) then
       begin
         // Copie du plan en cours pour recalculer les antérieurs ECO avec nelles méthode/durée/taux
         Plan2.Copie(Plan);
         Plan2.AmortEco.Reprise := 0;
         Plan2.CalculDateFinAmortReprise(Plan2.AmortEco);
         Plan2.CalculReprise(Plan2.AmortEco);
         QPlan.FindField('I_REPRISEECO').AsFloat :=Plan2.AmortEco.Reprise;
       end else
         // simple retassement
         QPlan.FindField('I_REPRISEECO').AsFloat := fCumulEco;

       // Recréer plan en cours
       Plan.Free ;
       Plan:=TPlanAmort.Create(true);
       Plan.TypeOpe:= TypeOpe;
       Plan2.Free ;
    end;

    QPlan.FindField('I_METHODEECO').AsString:=cMethodeEco.Value;
    QPlan.FindField('I_METHODEFISC').AsString:=cMethodeFiscale.Value;
    QPlan.FindField('I_DUREEECO').AsInteger:=DureeEco.Value;
    QPlan.FindField('I_DUREEFISC').AsInteger:=DureeFiscale.Value;
    QPlan.FindField('I_TAUXECO').AsFloat:=TauxEco.Value;
    QPlan.FindField('I_TAUXFISC').AsFloat:=TauxFiscal.Value;
    // à voir si CD2 Plan a été vidé et réalloué => ne contient plus rien
    QPlan.FindField('I_REVISIONECO').AsFloat:=Plan.AmortEco.Revision;
    QPlan.FindField('I_REVISIONFISCALE').AsFloat:=Plan.AmortFisc.Revision ;
    QPlan.FindField('I_BASEECO').AsFloat:=Valeur(lBaseEco.Caption);
    QPlan.FindField('I_BASEFISC').AsFloat:=Valeur(lBaseFisc.Caption);

    // BTY 11/05 Ajout plan futur sur VNC
    if bDureeRest.Checked then
       QPlan.FindField('I_JOURNALA').AsString := '***'
    else
       QPlan.FindField('I_JOURNALA').AsString := '';

    // 04/07 Nelle gestion fiscale : plan futur sur la VNF
    if bDureeRestF.Checked then
       QPlan.FindField('I_FUTURVNFISC').AsString := '***'
    else
       QPlan.FindField('I_FUTURVNFISC').AsString := '';

    // 04/07 Nelle gestion fiscale : Gestion fiscale
    if bGestionFiscale.Checked then
       QPlan.FindField('I_NONDED').AsString := 'X'
    else
       QPlan.FindField('I_NONDED').AsString := '-';

    // BTY FQ 19064
    if fNAM then
       if QPlan.FindField('I_METHODEECO').AsString = 'DEG' then
          begin
          QPlan.FindField('I_DATEDEBECO').AsDateTime := QPlan.FindField('I_DATEPIECEA').AsDateTime;
          QPlan.FindField('I_DATEDEBFIS').AsDateTime := QPlan.FindField('I_DATEPIECEA').AsDateTime;
          end
       else
          begin
          QPlan.FindField('I_DATEDEBECO').AsDateTime := QPlan.FindField('I_DATEAMORT').AsDateTime;
          QPlan.FindField('I_DATEDEBFIS').AsDateTime := QPlan.FindField('I_DATEAMORT').AsDateTime;
          end;


    // 04/07 Plan fiscal ajouté => ses antérieurs = anciens antérieurs ECO RETASSES
    if (Plan.TypeOpe = 'CD2') then
    begin
        QPlan.FindField('I_REPRISEFISCAL').AsFloat := fCumulEco;
        Plan.Charge(QPlan);
        // Recalculer antérieurs dérogatoires
        Plan.CalculReprise_Derog_Reint;
        QPlan.FindField('I_REPRISEDR').AsFloat := Plan.AmortDerog.Reprise;
        QPlan.FindField('I_REPRISEFEC').AsFloat := Plan.AmortReint.Reprise;
    end;

    // FQ 20256
    QPlan.FindField('I_TYPEDEROGLIA').AsString := TypeDerogatoire( nil, QPlan);
  end
  else if TypeChangement=2 then
  begin
    QPlan.FindField('I_TYPEEXC').AsString:=Plan.TypeExc ;
    QPlan.FindField('I_MONTANTEXC').AsFloat:=Plan.MontantExc ;

    // BTY 11/05 Ajout plan futur sur VNC
    if bDureeRestE.Checked then
       QPlan.FindField('I_JOURNALA').AsString := '***'
    else
       QPlan.FindField('I_JOURNALA').AsString := '';
  end;

  QPlan.Post;
  Plan.Calcul(QPlan,iDate1900);
  Plan.Sauve;
  QPlan.Edit;
  QPlan.FindField('I_DATEDERMVTECO').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortEco);
  QPlan.FindField('I_DATEDERNMVTFISC').AsDateTime := Plan.GetDateFinAmortEx(Plan.AmortFisc);
  QPlan.FindField('I_PLANACTIF').AsInteger := Plan.NumSeq;
  (* Tant pis on laisse les dates modifiées en non amortissable pour les calculs ultérieurs - CJ/CA - 19/08/2004 - FQ 13063
  { On met à jour la fiche avec les dates initiales }
  if bNam then
  begin
    QPlan.FindField('I_DATEPIECEA').AsDateTime := DatePieceA;
    QPlan.FindField('I_DATEAMORT').AsDateTime := DateAmort;
  end;
  *)
  CocheChampOperation (QPlan,'','I_OPECHANGEPLAN');
  QPlan.Post ;
end;

procedure TChangePlanAmort.EcritLogChangePlan;
var Ordre : integer;
    QueryW : TQuery;
begin
  Ordre := TrouveNumeroOrdreLogSuivant(Plan.CodeImmo);
  // ATTENTION : Penser à initialiser OrdreS=-1 lors du premier appel
  if OrdreS=-1 then OrdreS := TrouveNumeroOrdreSerieLogSuivant;
  QueryW:=OpenSQL('SELECT * FROM IMMOLOG WHERE IL_IMMO="'+Plan.CodeImmo+'"', FALSE) ;
  QueryW.Insert ;
  QueryW.FindField('IL_IMMO').AsString:=CodeImmo ;
  // BTY 11/05 Adapter le libellé si saisie exceptionnel avec calcul plan futur avec VNC
  //QueryW.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', Plan.TypeOpe, FALSE);
  if (TypeChangement=2) and (bDureeRestE.Checked) then
     QueryW.FindField('IL_LIBELLE').AsString:= TraduireMemoire('Eléments exc. & calcul plan futur avec VNC')
  else
     QueryW.FindField('IL_LIBELLE').AsString:=RechDom('TIOPEAMOR', Plan.TypeOpe, FALSE);
  QueryW.FindField('IL_TYPEMODIF').AsString:=AffecteCommentaireOperation('CHP'); //15/01/99 EPZ
  QueryW.FindField('IL_DATEOP').AsDateTime:= DateOpe;
  QueryW.FindField('IL_TYPEOP').AsString:=Plan.TypeOpe;
  QueryW.FindField('IL_ORDRE').AsInteger:=Ordre;
  QueryW.FindField('IL_ORDRESERIE').AsInteger:=fOrdreSerie; // FQ 18394 OrdreS;
  QueryW.FindField('IL_DUREEECO').AsInteger := AncienPlan.AmortEco.Duree;
  QueryW.FindField('IL_DUREEFISC').AsInteger := AncienPlan.AmortFisc.Duree;
  QueryW.FindField('IL_METHODEECO').AsString := AncienPlan.AmortEco.Methode;
  QueryW.FindField('IL_METHODEFISC').AsString := AncienPlan.AmortFisc.Methode;
  QueryW.FindField('IL_REVISIONECO').AsFloat := AncienPlan.AmortEco.Revision;
  QueryW.FindField('IL_REVISIONFISC').AsFloat := AncienPlan.AmortFisc.Revision;
  QueryW.FindField('IL_TYPEEXC').AsString := AncienPlan.TypeExc;
  QueryW.FindField('IL_MONTANTEXC').AsFloat := AncienPlan.MontantExc;
  QueryW.FindField('IL_TYPEDOT').AsString := TypeDot;
  QueryW.FindField('IL_MONTANTDOT').AsFloat := MontantDot;
  QueryW.FindField('IL_PLANACTIFAV').AsInteger := AncienPlan.NumSeq;
  QueryW.FindField('IL_PLANACTIFAP').AsInteger := Plan.NumSeq;
  // CA - 19/08/2004 - Astuce pour mémorise les dates en NAM
  if AncienPlan.AmortEco.Methode='NAM' then
    begin
    QueryW.FindField('IL_CODEMUTATION').AsString := FormatDateTime('ddmmyyyy',AncienPlan.DateAchat)+FormatDateTime('ddmmyyyy',AncienPlan.DateMiseEnService);
    QueryW.FindField('IL_CODECB').AsString := IntToStr(fInalienabilite);
    end;

//  FQ 20434 En web access : PutValue faux pour un query => findfield
//  {$IFDEF EAGLCLIENT}
//  QueryW.PutValue('IL_BLOCNOTE',BLOC_NOTE.LinesRTF.Text) ;
//  {$ELSE}
//  TBlobField(QueryW.FindField('IL_BLOCNOTE')).Assign(BLOC_NOTE.LinesRTF) ;
//  {$ENDIF}
  QueryW.FindField('IL_BLOCNOTE').AsString := RichToString (Bloc_note);

  // BTY 11/05 Si Plan futur sur VNC et/ou Plan fiscal CRC 2002-10
  // Dans tous les cas, conserver le i_journala de l'immo traitée
  // car possible de réviser sans plan futur sur VNC (CDM) l'année suivant
  // une révision avec plan futur sur VNC  (CD3)
  QueryW.FindField('IL_CODEECLAT').AsString := AncienPlan.Journala;

  if (TypeChangement=1) then
  begin
    {if ((Plan.TypeOpe = 'CD2') or (Plan.TypeOpe = 'CD3')) then
    begin
     QueryW.FindField('IL_BASEFISCAVMB').AsFloat := AncienPlan.AmortFisc.Base;
    end; }

    // 04/07 Nelle gestion fiscale, conserver l'ancienne base
    // FQ 21523 Toujours conserver l'ancienne base fiscale
    //if (cMethodeFiscale.Value <> '') then
        QueryW.FindField('IL_BASEFISCAVMB').AsFloat := AncienPlan.AmortFisc.Base;

    {if (Plan.TypeOpe = 'CD2') then }
    QueryW.FindField('IL_REPRISEECO').AsFloat := RepriseEco;
    QueryW.FindField('IL_REPRISEFISC').AsFloat := RepriseFisc;

    // 04/07 ancien Plan futur avec la VNF
    if AncienPlan.VncRestFisc then
       QueryW.FindField('IL_LIEUGEO').AsString := '***'
    else
       QueryW.FindField('IL_LIEUGEO').AsString := '';

    // 04/07 anciennne Gestion fiscale
    if AncienPlan.GestionFiscale then
       QueryW.FindField('IL_ETABLISSEMENT').AsString := 'X'
    else
       QueryW.FindField('IL_ETABLISSEMENT').AsString := '';

    // Antérieurs dérogatoires
    QueryW.FindField('IL_REVISIONDOTECO').AsFloat :=  fRepriseDR;
    QueryW.FindField('IL_REVISIONREPECO').AsFloat :=  fRepriseFEC;
  end;
  //end;

  QueryW.FindField('IL_VERSION').AsString := V_PGI.NumVersion;

  QueryW.Post;
  Ferme(QueryW) ;
end;


procedure TChangePlanAmort.BFermeClick(Sender: TObject);
var mr : integer;
begin
  mr := HM.execute(2,Caption,'');
  if mr=mrYes then TraiteFinChangePlan
  else ModalResult := mr;
  if ModalResult = mrCancel then ModalResult:= mrNone;
end;

// a priori cette procedure 'est jamais utilisée
// 04/07 SI SI
procedure TChangePlanAmort.AmortFiscExit(Sender: TObject);
{var
  NumMess : integer; //message compilateur} //XVI conseil compile
begin
  // FQ 19080 Déplacé dans le contrôle général du bouton Validation
  //if not TraiteIntervaleDureeAmort(cMethodeFiscale.Value,0,
  //                                      DureeFiscale.Value,NumMess) then
  //begin
  //  HM.Execute(NumMess, Caption, '');
  //  Pages.ActivePage := PGeneral; FocusControl(DureeFiscale); exit;
  //end;
  TauxFiscal.Value := GetTaux(cMethodeFiscale.Value,
                                QPlan.FindField('I_DATEAMORT').AsDateTime,
                                QPlan.FindField('I_DATEPIECEA').AsDateTime,
                                DureeFiscale.Value);

  bDureeRestF.Visible := ((cMethodeFiscale.Value = 'LIN') or (cMethodeFiscale.Value = 'VAR'));
  bDureeRestF.Enabled := (bDureeRestF.Visible);
  if (not bDureeRestF.Visible) then bDureeRestF.Checked := False;

  AffecteModifPlanCalcul(MethodeFisc);
end;

// a priori cette procedure n'est jamais utilisée SI !!
procedure TChangePlanAmort.AmortEcoExit(Sender: TObject);
var
  //NumMess : integer; message compilateur
  {OuvrirPlanCRC : boolean; }
  NAM : boolean;
begin
  inherited;
  // BTY 11/06 En non amortissable, mettre la durée à 0 sinon sera enregistrée telle quelle
  NAM := (cMethodeEco.Value = 'NAM');
  if NAM then
     begin
     DureeEco.Value := 0;
     DureeEco.Enabled := False;
    end
  else
     DureeEco.Enabled := True;

  if DureeEco.Text='' then DureeEco.Value := 1;  // 12; FQ 19095

  // FQ 19080 Déplacé dans le contrôle général du bouton Validation
  // if (DureeEco.Value <> 0) and
  //    (not TraiteIntervaleDureeAmort(cMethodeEco.Value,0,DureeEco.Value,NumMess)) then
  // begin
  //   HM.Execute(NumMess, Caption, '');
  //   Pages.ActivePage := PGeneral; FocusControl(DureeEco); exit;
  // end;

  TauxEco.Value := GetTaux(cMethodeEco.Value,
                                QPlan.FindField('I_DATEAMORT').AsDateTime,
                                QPlan.FindField('I_DATEPIECEA').AsDateTime,
                                DureeEco.Value);
  // BTY 03/06 FQ 17446
  bDureeRest.Visible := ((cMethodeEco.Value = 'LIN') or (cMethodeEco.Value = 'VAR'));
  bDureeRest.Enabled := (bDureeRest.Visible);
  if (not bDureeRest.Visible) then bDureeRest.Checked := False;
  //

  // BTY 11/05 Plan fiscal CRC 200210 ouvert si allongement de la durée
  // BTY 12/05 Cacher Plan fiscal CRC 200210 si durée diminuée après avoir été augmentée
  // BTY 01/06 FQ 17373 Ouvrir Plan CRC 2002-10 même à durée ECO inchangée
  // bPlanCRC.Enabled := ((bPlanCRC.Visible) and (DureeEco.Value > AncienPlan.AmortEco.Duree));
  // BTY 11/06 Plan fiscal bloqué si prime ou subvention
  {if ((DureeEco.Value = AncienPlan.AmortEco.Duree) and
      (cMethodeEco.Value = AncienPlan.AmortEco.Methode)) or
     (DureeEco.Value < AncienPlan.AmortEco.Duree) or
     (cMethodeEco.Value = 'NAM') or
     (fSBVPRI <> 0) or (fSBVMT <> 0)  then
     OuvrirPlanCRC := False
  else
     OuvrirPlanCRC := True;
  bPlanCRC.Enabled := ((bPlanCRC.Visible) and (OuvrirPlanCRC));
  // BTY 11/06 Plan fiscal peut être touché si pas de prime ni subvention
  if (fSBVPRI = 0) and (fSBVMT = 0)  then
      if (bPlanCRC.Visible) and (not bPlanCRC.Enabled) then bPlanCRC.Checked := False; }

  AffecteModifPlanCalcul(MethodeEco);
end;

// FQ 19095
procedure TChangePlanAmort.DureeEcoExit(Sender: TObject);
//var
  //mMess : integer; message compilateur
  {OuvrirPlanCRC : boolean; }
begin
  inherited;

  if DureeEco.Text='' then DureeEco.Value := 1;
 // if (DureeEco.Value = 0) then cMethodeEco.Value := 'NAM'; FQ 19095

  TauxEco.Value := GetTaux(cMethodeEco.Value,
                                QPlan.FindField('I_DATEAMORT').AsDateTime,
                                QPlan.FindField('I_DATEPIECEA').AsDateTime,
                                DureeEco.Value);

  {if ((DureeEco.Value = AncienPlan.AmortEco.Duree) and
      (cMethodeEco.Value = AncienPlan.AmortEco.Methode)) or
     (DureeEco.Value < AncienPlan.AmortEco.Duree) or
     (cMethodeEco.Value = 'NAM') or
     (fSBVPRI <> 0) or (fSBVMT <> 0)  then
     OuvrirPlanCRC := False
  else
     OuvrirPlanCRC := True;
  bPlanCRC.Enabled := ((bPlanCRC.Visible) and (OuvrirPlanCRC));

  // BTY 11/06 Plan fiscal peut être touché si pas de prime ni subvention
  if (fSBVPRI = 0) and (fSBVMT = 0)  then
      if (bPlanCRC.Visible) and (not bPlanCRC.Enabled) then bPlanCRC.Checked := False; }

  bDureeRest.Visible := ((cMethodeEco.Value = 'LIN') or (cMethodeEco.Value = 'VAR'));
  bDureeRest.Enabled := (bDureeRest.Visible);
  if (not bDureeRest.Visible) then bDureeRest.Checked := False;

  AffecteModifPlanCalcul(MethodeEco);
end;

procedure TChangePlanAmort.DureeRestExit(Sender: TObject);
begin
  if bDureeRest.Enabled then
     AffecteModifPlanCalcul(MethodeEco);
end;

procedure TChangePlanAmort.DureeRestFExit(Sender: TObject);
begin
  if bDureeRestF.Enabled then
     AffecteModifPlanCalcul(MethodeFisc);
end;


procedure TChangePlanAmort.bAmortissementFiscalClick(Sender: TObject);
begin
  //29/03/99 cette fonction ne sert pas :
  //Il n'est pas possible dans la 1ere version de
  // créer une méthode fiscale si pas init à la création de l'immo
  // 04/07 Cette fonction sert
  A.Enabled := bAmortissementFiscal.Checked;
  if not bAmortissementFiscal.Checked then
  begin
    cMethodeFiscale.Value := '';
    DureeFiscale.Value := DureeFiscale.MinValue;
    TauxFiscal.Value := 0;
    cTypeBaseFisc.Value := '';
    lBaseFisc.Caption := StrFMontant(0.00,15,V_PGI.OkDecV,'',True);
    bDureeRestF.Visible := False;
    bDureeRestF.Enabled := bDureeRestF.Visible;
    bDureeRestF.Checked := False;
    bGestionFiscale.Checked := False;
    bGestionFiscale.Enabled := False;
  end
  else
  begin
    if not AncienPlan.Fiscal then
    begin
       // Alimenter le groupe fiscal avec les données ECO
       // sauf si immo NAM avec prime eu subvention => forcer à LIN sur la durée d'inaliénabilité
       if cMethodeEco.Value ='NAM' then
       begin
         cMethodeFiscale.Value := 'LIN';
         DureeFiscale.Value := fInalienabilite;
         TauxFiscal.Value := GetTaux (cMethodeFiscale.Value,
                                     QPlan.FindField('I_DATEAMORT').AsDateTime,
                                     QPlan.FindField('I_DATEPIECEA').AsDateTime,
                                     DureeFiscale.Value);
         cTypeBaseFisc.Itemindex:=cTypeBaseFisc.Items.indexof(RechDom('TITYPEBASEFISC','ECO',FALSE));
         lBaseFisc.Caption := StrFMontant(AncienPlan.AmortEco.Base,15,V_PGI.OkDecV,'',True);
         bDureeRestF.Visible := False;
         bDureeRestF.Enabled := bDureeRestF.Visible;
         bDureeRest.Checked := False;
       end
       else
       begin
         // Alimenter le groupe fiscal avec les anciennes données ECO
         cMethodeFiscale.Value := AncienPlan.AmortEco.Methode;
         DureeFiscale.Value := AncienPlan.AmortEco.Duree;
         TauxFiscal.Value := AncienPlan.AmortEco.Taux*100;
         cTypeBaseFisc.Itemindex:=cTypeBaseFisc.Items.indexof(RechDom('TITYPEBASEFISC','THE',FALSE));
         lBaseFisc.Caption := StrFMontant(fValeurTheorique,15,V_PGI.OkDecV,'',True);
         bDureeRestF.Visible := ((PlanEco ='LIN') or (PlanEco ='VAR'));
         bDureeRestF.Enabled := bDureeRestF.Visible;
         bDureeRestF.Checked := ((bDureeRestF.Enabled) and (QPlan.FindField('I_JOURNALA').AsString='***'));
         bGestionFiscale.Checked := False;
         bGestionFiscale.Enabled := not (bDPI); // ouvert si immo sans DPI
         // détermine case Gestion fiscale
         AffecteModifPlanCalcul(MethodeFisc);
       end;
       // DPI ou prime ou subvention => pas de Gestion fiscale
       if (bDPI) or (fSBVPRI <> 0) or (fSBVMT <> 0)  then
       begin
          bGestionFiscale.Checked := False;
          bGestionFiscale.Enabled := False;
       end;

    end
    else RecupAncienPlanFiscal;

  end;
end;

procedure TChangePlanAmort.AffecteModifPlanCalcul(TypeAmort : integer);
begin
  if TypeAmort=MethodeEco then
  begin
    PlanCalcul.VncRestEco := bDureeRest.Checked;
    PlanCalcul.AmortEco.ModifieMethodeAmort(PlanCalcul.AmortEco.DateFinAmort,RevisionEco,PlanCalcul.AmortEco.Base,
                                      TauxEco.Value/100,DureeEco.Value,cMethodeEco.Value,PlanCalcul.ImmoCloture);
    PlanCalcul.CalculDotation(PlanCalcul.AmortEco,VHImmo^.Encours.Deb,false);
  end
  else if TypeAmort=MethodeFisc then
  begin
    PlanCalcul.Fiscal := True;
    PlanCalcul.CalculDateFinAmortissement(PlanCalcul.AmortFisc);

    // 04/07  PlanCalcul.AmortFisc.ModifieMethodeAmort(PlanCalcul.AmortFisc.DateFinAmort,RevisionFisc,PlanCalcul.AmortFisc.Base, etc.
    PlanCalcul.AmortFisc.ModifieMethodeAmort(PlanCalcul.AmortFisc.DateFinAmort,
                                             RevisionFisc,
                                             Valeur(lBaseFisc.Caption),
                                             TauxFiscal.Value/100,
                                             DureeFiscale.Value,
                                             cMethodeFiscale.Value,
                                             PlanCalcul.ImmoCloture);
    PlanCalcul.CalculDotation(PlanCalcul.AmortFisc,VHImmo^.Encours.Deb,false);
  end;
  AfficheOngletExceptionnel(PlanCalcul);
  AfficheGestionFiscale(PlanCalcul, TypeAmort);
end;

// 04/07 Valeur de la Check "Gestion Fiscale" en fonction des antérieurs théoriques
procedure TChangePlanAmort.AfficheGestionFiscale(PlanCalcul:TPlanAmort; TypeAmort:integer);
var AffPlan : TPlanAmort;
    CumulEcoTheorique : double ;
    CumulFiscTheorique : double ;

begin
 if TypeChangement=1 then
 begin
   AffPlan := TPlanAmort.Create(true) ;
   AffPlan.Copie(PlanCalcul);

   // antérieurs THEORIQUES éco et fiscal selon les données saisies
   AffPlan.CalculReprises;
   CumulEcoTheorique := AffPlan.AmortEco.Reprise;
   CumulFiscTheorique := AffPlan.AmortFisc.Reprise;

   if not AncienPlan.Fiscal then
   begin

     // DPI ou prime ou subvention => pas de Gestion fiscale
     if (not bDPI) and (fSBVPRI = 0) and (fSBVMT = 0)  then
     begin

       // mettre à jour la Query de travail avec les données saisies
       QPlanCalcul.Edit;
       if TypeAmort=MethodeEco then
       begin
         QPlanCalcul.FindField('I_METHODEECO').AsString  := cMethodeEco.Value;
         QPlanCalcul.FindField('I_DUREEECO').AsInteger  := DureeEco.Value;
         QPlanCalcul.FindField('I_TAUXECO').AsFloat  :=   TauxEco.Value;
         if bDureeRest.Checked then
            QPlanCalcul.FindField('I_JOURNALA').AsString := '***'
         else
            QPlanCalcul.FindField('I_JOURNALA').AsString := '';
       end
       else
       begin
         QPlanCalcul.FindField('I_METHODEFISC').AsString  := cMethodeFiscale.Value;
         QPlanCalcul.FindField('I_DUREEFISC').AsInteger  := DureeFiscale.Value;
         QPlanCalcul.FindField('I_TAUXFISC').AsFloat  :=   TauxFiscal.Value;
         QPlanCalcul.FindField('I_BASEFISC').AsFloat := Valeur(lBaseFisc.Caption);
         if bDureeRestF.Checked then
           QPlanCalcul.FindField('I_FUTURVNFISC').AsString := '***'
         else
           QPlanCalcul.FindField('I_FUTURVNFISC').AsString := '';
       end;

       // Antérieurs ECO > Fisc et <> 0 ou non conditionnent la valeur de "Gestion fiscale"
       if (AffPlan.Determine_Gestion_Fiscale(QPlanCalcul, CumulEcoTheorique, CumulFiscTheorique))=True then
       begin
         bGestionFiscale.Checked := True;
         bGestionFiscale.Enabled := False;
       end else
       begin
         bGestionFiscale.Checked := False;
         bGestionFiscale.Enabled := True;
       end;
     end;
   end; // AncienPlan.Fiscal

   AffPlan.Free ;
 end; // TypeChangement

end;

function  TChangePlanAmort.CalculAnterieursOK : boolean;
var Calcul : boolean;
    Plan2 : TPlanAmort;
    CumulEco, CumulFiscal,wRepriseEco : double;
begin

    // Test recalcul des antérieurs éco saisis I_RepriseEco :
    // ajout plan fiscal = plan éco et plan éco allongé sans calcul plan futur VNC
    Calcul :=  ((not AncienPlan.Fiscal) and (bDureeRest.Enabled) and (not bDureeRest.Checked)
                and (cMethodeFiscale.Value = AncienPlan.AmortEco.methode)
                and (DureeFiscale.Value = AncienPlan.AmortEco.duree)
                and (Arrondi (TauxFiscal.Value, 2) = Arrondi(AncienPlan.AmortEco.taux * 100, 2))
                and (Arrondi (Valeur(lBaseFisc.Caption), V_PGI.OkDecV) =
                     Arrondi (AncienPlan.AmortEco.base, V_PGI.OkDecV))
                and (DureeEco.Value > AncienPlan.AmortEco.duree) );

    if Calcul then
    begin
      Plan2:=TPlanAmort.Create(true);
      Plan2.Copie(Plan);
      Plan2.AmortEco.ModifieMethodeAmort(Plan2.AmortEco.DateFinAmort,
                                         Plan2.AmortEco.Revision,
                                         Valeur(lBaseEco.Caption),
                                         TauxEco.Value/100,
                                         DureeEco.Value,
                                         cMethodeEco.Value,
                                         Plan2.ImmoCloture);
      Plan2.AmortEco.Reprise := 0;
      Plan2.CalculDateFinAmortReprise(Plan2.AmortEco);
      Plan2.CalculReprise(Plan2.AmortEco);
      wRepriseEco := Plan2.AmortEco.Reprise;
      Plan2.GetCumulsDotExercice(VHImmo^.Encours.Deb, CumulEco,CumulFiscal,True,False,true);
      CumulEco := CumulEco + wRepriseEco;
      Plan2.Free ;
    end
    else
      CumulEco := fCumulEco;

    // Antérieurs fiscaux
    if (not AncienPlan.Fiscal) and (cMethodeFiscale.Value <> '') then
      CumulFiscal := fCumulEco
    else
      CumulFiscal := fCumulFisc;

    Result := (Arrondi (CumulEco, V_PGI.OkDecV) = Arrondi (CumulFiscal, V_PGI.OkDecV) );
end;



procedure TChangePlanAmort.AfficheOngletExceptionnel(AffPlan : TPlanAmort);
var CumulEco, CumulFisc, AnterieurEco, AnterieurFisc : double;
begin
  DotationExe := AffPlan.GetDotationExercice(VHImmo^.EnCours.Fin,AffPlan.AmortEco,false);
  // ajout mbo 30.09.05 - fq 15280
  DotExeFisc := AffPlan.GetDotationExercice(VHImmo^.EnCours.Fin,AffPlan.AmortFisc,false);
  AffPlan.GetCumulsDotExercice(VHImmo^.Encours.Deb,CumulEco, CumulFisc,false,false,true);
  DotEco:=CumulEco;      // cumul antérieurs économiques (sans la reprise)

  AnterieurEco  := CumulEco + RepriseEco + RepriseDep;  // mbo impact antérieur dépréciation
  AnterieurFisc := CumulFisc + RepriseFisc + RepriseDep;

  InitVNC := Arrondi (BaseEco-AnterieurEco, V_PGI.OkDecV );
  VNCfinEx:= InitVnc - Dotationexe;

  // ajout mbo - fq 15280
  if derog = false then
    begin
    MaxAjout := Arrondi ( InitVNC-DotationExe, V_PGI.OkDecV );
    MaxRemplace := InitVNC;
    MaxReprise := Arrondi (DotationExe, V_PGI.OkDecV );
    end
  else
    begin
    // ajout mbo 30.09.05 - fq 15280 - on borne à la vnc fiscale
    VRF := Arrondi ((BaseFisc-AnterieurFisc-DotExeFisc), V_PGI.OkDecV );
    MaxAjout := Arrondi ((BaseFisc-AnterieurFisc-DotExeFisc), V_PGI.OkDecV );
    if MaxAjout > Arrondi ( InitVNC-DotationExe, V_PGI.OkDecV ) then MaxAjout :=Arrondi ( InitVNC-DotationExe, V_PGI.OkDecV );
    MaxRemplace := Arrondi ((BaseFisc-AnterieurFisc), V_PGI.OkDecV );
    if MaxRemplace >Arrondi ( InitVNC, V_PGI.OkDecV ) then MaxRemplace :=Arrondi ( InitVNC, V_PGI.OkDecV );
    MaxReprise := Arrondi ( DotationExe, V_PGI.OkDecV );
    If MaxReprise > DotExeFisc then MaxReprise:=DotExeFisc;
    end;
  RevisionEco := 0.00;
  RevisionFisc := 0.00;
  VNC.Value:=Arrondi ( InitVNC-DotationExe, V_PGI.OkDecV );
  NOUVELLEDOT.Caption:=StrFMontant(DotationExe,15,V_PGI.OkDecV,'',True);
  DONTDOTEXC.Value   :=Valeur(FloatToStrF(AffPlan.MontantExc,ffFixed,20,V_PGI.OkDecV));
  DotationTotale.Caption:=StrFMontant(DotationExe,15,V_PGI.OkDecV,'',True);

end;

procedure TChangePlanAmort.DATE_OPEKeyPress(Sender: TObject;
  var Key: Char);
begin
  ParamDate(Self,Sender,Key);
end;

procedure TChangePlanAmort.DOTATIONEXCKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
begin
  inherited;
  // mbo - FQ 16757 - if Key=VK_F3 then GereAmortissementTotal;
  if Key=VK_F6 then GereAmortissementTotal;
end;

procedure TChangePlanAmort.GereAmortissementTotal;
var
  DotAnterieur,MaxDot,NelleDot : double;
  signe : integer;
begin
  signe := 1;MaxDot :=0;
  DotAnterieur := DotEco+RepriseEco+ RepriseDep;  // impact antérieur dépréciation mbo 02.06
  if bAjoutDotation.Checked then MaxDot:=MaxAjout
  else if bRemplaceDotation.Checked then begin DotAnterieur := 0.00; MaxDot:=MaxRemplace; end
  else if bRepriseDotation.Checked then begin signe:=-1; MaxDot:=MaxReprise; end;
  DOTATIONEXC.Value := MaxDot;
  NelleDot := DotationExe + DOTATIONEXC.Value;
  DONTDOTEXC.Value := DOTATIONEXC.Value + Plan.MontantExc;
  NOUVELLEDOT.Caption := StrFMontant(NelleDot,15,V_PGI.OkDecV,'',True);
  VNC.Value := InitVNC - DotAnterieur - (DOTATIONEXC.Value*signe);
end;

procedure TChangePlanAmort.DATE_OPEExit(Sender: TObject);
begin
if StrToDate(DATE_OPE.Text)<DateAchat then
  begin
  HM.Execute (9,Caption,'');
  FocusControl (DATE_OPE); exit;
  end
else if (StrToDate(DATE_OPE.Text)<VHImmo^.Encours.Deb) or (StrToDate(DATE_OPE.Text)>VHImmo^.Encours.Fin) then
  begin
  HM.Execute (13,Caption,'');
  FocusControl (DATE_OPE);
  end
else if ExisteSQL('SELECT IL_DATEOP FROM IMMOLOG WHERE IL_IMMO="'+CodeImmo+' AND IL_DATEOP>'+USDATETIME(StrToDate(DATE_OPE.Text))+'"') then
  begin
  HM.Execute (10,Caption,'');
  FocusControl (DATE_OPE);
  end;
end;

procedure TChangePlanAmort.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

procedure TChangePlanAmort.FormCreate(Sender: TObject);
begin
{$IFDEF SERIE1}
HelpContext:=511030 ;
{$ELSE}
//HelpContext:=2111100 ;
{$ENDIF}
Plan:=TPlanAmort.Create(true) ;
PlanCalcul:=TPlanAmort.Create(true) ;
AncienPlan:=TPlanAmort.Create(true) ;
end;

procedure TChangePlanAmort.FormDestroy(Sender: TObject);
begin
  Plan.Free ;
  PlanCalcul.free ;
  AncienPlan.free ;
  Ferme(QPlan);
  Ferme(QPlanCalcul);
end;

procedure TChangePlanAmort.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key=VK_F10 then
  begin
    key := 0;
    bValiderClick(nil);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 19/08/2004
Modifié le ... : 24/09/2004
Description .. : Contrôle de la validité de la durée d'amortissement saisie en
Suite ........ : fonction de la durée d'amortissement déjà pratiquée.
Suite ........ : - 24/09/2004 - CA - FQ 14645 : Ajout de la méthode en
Suite ........ : paramètre pour ne pas faire de traitement sur une immo non
Suite ........ : amortissable
Mots clefs ... :
*****************************************************************}
function TChangePlanAmort.DureeValide(Methode : string; Duree: integer;  bFiscal: boolean): boolean;
var PremMois, PremAnnee,  NbMois : Word;
begin
  if Methode='NAM' then Result := True
  else
  begin
    // fq 18028 mbo 02.05.06 c'est la méthode qui compte et non le plan pour déterminer la date de début d'amort
    //if bFiscal then NombreMois(QPlan.FindField('I_DATEAMORT').AsDateTime,StrToDate(Date_Ope.Text), PremMois, PremAnnee,  NbMois)
    // else NombreMois(QPlan.FindField('I_DATEPIECEA').AsDateTime,StrToDate(Date_Ope.Text), PremMois, PremAnnee,  NbMois);

    if Methode = 'DEG' then NombreMois(QPlan.FindField('I_DATEPIECEA').AsDateTime,StrToDate(Date_Ope.Text), PremMois, PremAnnee,  NbMois)
    else NombreMois(QPlan.FindField('I_DATEAMORT').AsDateTime,StrToDate(Date_Ope.Text), PremMois, PremAnnee,  NbMois);

    // FQ 18028 Result := ((NbMois-1) < Duree); pour gérer le cas où durée déjà amortie = nvelle durée
    Result := ((NbMois-1) <= Duree);
  end;
end;

// FQ 19080
{function TChangePlanAmort.DureeMinimum(Methode : string; Duree: integer): boolean;
//var PremMois, PremAnnee,  NbMois : Word; //XVI Conseil Compile...
begin
  result:=FALSE ; //XVI Conseil conmpile....
  if Methode='NAM' then Result := True
  else
  begin
    if Methode = 'DEG' then
       result := (Duree > 35)
    else if Methode = 'LIN' then
       result := (Duree > 0);
  end;
end;} //XVI Conseil Compile...


{***********A.G.L.***********************************************
Auteur  ...... : BTY
Créé le ...... : 09/11/2005
Modifié le ... : 09/11/2005
Description .. : Cocher/Décocher la check Gestion d'un plan fiscal CRC 2002-10
Mots clefs ... :
*****************************************************************}

{procedure TChangePlanAmort.bPlanCRCClick(Sender: TObject);
begin

  stTypeBaseFisc.Visible:= (bPlanCRC.Checked);
  cTypeBaseFisc.Visible:= (bPlanCRC.Checked);
  cMethodeFiscale.Enabled := not (bPlanCRC.Checked);
  DureeFiscale.Enabled := not (bPlanCRC.Checked);
  TauxFiscal.Enabled := not (bPlanCRC.Checked);
  bAmortissementFiscal.Checked := (bPlanCRC.Checked);
  if bPlanCRC.Checked then
     begin
     // Alimenter le groupe fiscal avec les données ECO
     // BTY 11/06 sauf si immo NAM avec prime eu subvention => forcer à LIN sur la durée d'inaliénabilité
     if cMethodeEco.Value ='NAM' then
        begin
        cMethodeFiscale.Value := 'LIN';
        DureeFiscale.Value := fInalienabilite;
        TauxFiscal.Value := GetTaux (cMethodeFiscale.Value,
                                     QPlan.FindField('I_DATEAMORT').AsDateTime,
                                     QPlan.FindField('I_DATEPIECEA').AsDateTime,
                                     DureeFiscale.Value);
        end
     else
        begin
        cMethodeFiscale.Value := AncienPlan.AmortEco.Methode;
        DureeFiscale.Value := AncienPlan.AmortEco.Duree;
        TauxFiscal.Value := AncienPlan.AmortEco.Taux*100;
        end;
     cTypeBaseFisc.Itemindex:=cTypeBaseFisc.Items.indexof(RechDom('TITYPEAMORFISC','THE',FALSE));
     lBaseFisc.Caption := StrFMontant(fValeurTheorique,15,V_PGI.OkDecV,'',True);
     end
  else
     begin
     // RAZ du groupe fiscal
     cTypeBaseFisc.Value := '';
     cMethodeFiscale.Value := '';
     DureeFiscale.Value := AncienPlan.AmortFisc.Duree;
     TauxFiscal.Value := AncienPlan.AmortFisc.Taux*100;
     lBaseFisc.Caption := StrFMontant(AncienPlan.AmortFisc.Base,15,V_PGI.OkDecV,'',True);
     end;

end; }

{***********A.G.L.***********************************************
Auteur  ...... : BTY
Créé le ...... : 10/11/2005
Modifié le ... : 04/04/2007
Description .. : En cas de gestion d'un plan fiscal CRC 2002-10 :
Suite ........ : choix de la base fiscale
Suite ........ : Affichage base fiscale en fonction du choix de cette combo
Mots clefs ... : 
*****************************************************************}

procedure TChangePlanAmort.cTypeBaseFiscChange(Sender: TObject);
begin
  if   cTypeBaseFisc.Value='ECO' then lBaseFisc.Caption := StrFMontant (AncienPlan.AmortEco.Base,15,V_PGI.OkDecV,'',True)
  else lBaseFisc.Caption := StrFMontant (fValeurTheorique,15,V_PGI.OkDecV,'',True);

  AffecteModifPlanCalcul(MethodeFisc);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BTY
Créé le ...... : 05/12/2005
Modifié le ... :   /  /
Description .. : Affichage standard de la durée déjà amortie en mois et en
Suite ........ : jours
Mots clefs ... :
*****************************************************************}
procedure TChangePlanAmort.AfficheDureeDejaAmortie;
var
  PremMois, PremAnnee, NbMois: Word;
  NbJour: integer;
  DateImmo: TDateTime;

begin
    NOMBREMOIS(StrToDate(QPlan.FindField('I_DATEAMORT').AsString), VHImmo^.Encours.Deb,
               PremMois, PremAnnee, NbMois);
    //si dégressif : date d'achat
    //sinon        : date de mise en service
    if (AncienPlan.AmortEco.Methode = 'DEG') then
         DateImmo := StrToDate(QPlan.FindField('I_DATEPIECEA').AsString)
    else DateImmo := StrToDate(QPlan.FindField('I_DATEAMORT').AsString);

    if ( VarIsNull(DateImmo) or VarIsNull(VHImmo^.Encours.Deb-1)
    or (AncienPlan.AmortEco.Methode = 'NAM')
    or ((VHImmo^.Encours.Deb-1) < DateImmo) ) then
      begin
      DureeDeja.Caption := '';
      end
    else
      begin
      if ((NbMois - 1) > QPlan.FindField('I_DUREEECO').AsInteger) then
         NbJour := QPlan.FindField('I_DUREEECO').AsInteger * 30
      else
         NbJour := NombreJour360 (DateImmo, VHImmo^.Encours.Deb-1);
      if NbJour < 0 then NbJour := 0;
      NOMBREMOIS(DateImmo, VHImmo^.Encours.Deb, PremMois, PremAnnee, NbMois);
      if ((NbMois - 1) > QPlan.FindField('I_DUREEECO').AsInteger) then
          NbMois := QPlan.FindField('I_DUREEECO').AsInteger + 1;
      DureeDeja.Caption := Format(HM.Mess[14], [NbMois - 1, NbJour]);
      stDureeDeja.Caption := stDureeDeja.Caption + Format(HM.Mess[15],[DateToStr(VHImmo^.Encours.Deb-1)]);
      end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BTY
Créé le ...... : 05/12/2005
Modifié le ... :   /  /
Description .. : Affichage VNC début d'exercice (Base Eco - antérieurs -
Suite ........ : antérieurs saisis)
Suite......... : 13/02/2006 - modif mbo : impact antérieur dépréciation sur le calcul de la vnc
Mots clefs ... :
*****************************************************************}
procedure TChangePlanAmort.AfficheVNCDebutExo;
var CumulEco,CumulFisc,AnterieurEco : double;
begin
  // Récup antérieurs calculés jusqu'à fin N-1
  AncienPlan.GetCumulsDotExercice(VHImmo^.Encours.Deb,CumulEco, CumulFisc, False, False, True);
  AnterieurEco  := CumulEco + RepriseEco + RepriseDep; // impact antérieur dépréciation mbo 02.06
  ValeurResiduelle.Caption :=StrFMontant(Arrondi (BaseEco-AnterieurEco, V_PGI.OkDecV ), 15, V_PGI.OkDecV, '', True);
end;



end.



