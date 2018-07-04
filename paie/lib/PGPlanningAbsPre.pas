{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 01/01/2002
Modifié le ... :   /  /
Description .. : Edition du planning des absences
Mots clefs ... : PAIE;CP
*****************************************************************   }
{
PT1   29/04/2002 SB      On edite les abences des mvts dont la date debut concerne la periode saisie
PT2-1 05/12/2002 SB V591 Intégration des compteurs Abs en cours en fn des champs de la table
PT2-2 05/12/2002 SB V591 Optimisation requête & ajout calcul mouvements en cours d'intégration
PT3   08/10/2003 SB V_42 Econges Spéc. CEGID Gestion des abences
PT4   02/12/2003 SB V_42 Ajout absences intervalle supérieur dates planning
PT5   03/05/2004 SB V_50 FQ 11280 Modif champ TYPERTT => TYPEABS
PT6   15/02/2005 SB V_60 FQ 11313 Création absence sur click droit planning
PT7   18/04/2005 SB V_60 FQ 12175 Erreur SQL Oracle, refonte construction requete
PT8   25/07/2005 SB V_60 FQ 12399 Ajout gestion assistante
PT9-1 25/07/2005 SB V_65 FQ 12398 Ajout Filtre des absences
PT9-2 25/07/2005 SB V_65 FQ 12398 Ajout Filtre des absences
PT10  25/07/2005 SB V_65 Ajout planning du jour
PT11  26/07/2005 SB V_65 Accès à la validation pour les respons. hierarchiques
PT12  19/09/2005 SB V_65 FQ 11313 Suite création new planning, déportation de procédure dans pgplanningoutils
PT13  26/09/2005 SB V_65 Ajout de l'aide en ligne
PT14  05/12/2005 SB V_65 FQ 12737 Ajout sélection du niveau hierarchique
PT15-1 28/02/2006 SB V_65 Refonte traitement pour mise en place gestion des couleurs, gestion planning
PT15-2 28/02/2006 SB V_65 FQ 12926 Optimisation + Refonte tri du planning par nom resp., puis nom sal.
PT16   19/06/2006 SB V_65 FQ 13231 Retrait des mvt absences annulées
PT17  27/06/2006  SB V_65 Améliorations et optimisations de l'export des compteurs et ajout de nouveaux filtres d'absence
PT17-2 27/06/2006 SB V_65 ajout critère total nb jour absence
PT18   04/07/2006 SB V_65 Nouveaux paramètres d'execution pour Suivi des absences
PT19   08/09/2006 SB V_65 Reprise PT15-1
PT20   28/06/2007 NA V_80 Ajout gestion des présences
}
unit PGPlanningAbsPre;

interface



uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  {$IFDEF ABSETEMPS}
  SaisiePersoPlanning,
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
  MaineAgl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ENDIF}
  HCtrls, StdCtrls, HTB97, ExtCtrls, hplanning, UTOB,  HmsgBox,
  HEnt1, HeureUtil,HPanel,UiUtil,PgPlanningOutils, Utobdebug;

 

type

  TPlanningAbsPre = class(TForm)
    PPanel: TPanel;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    TRECAP: TToolbarButton97;
    Bperso: TToolbarButton97;
    TExcel: TToolbarButton97;
    BExporter: TToolbarButton97;
    PRECAP: TPanel;
    GBAbsence: TGroupBox;
    Valide: THLabel;
    Attente: THLabel;
    EnCoursCP: THLabel;
    NbValidePri: THLabel;
    NbAttentePri: THLabel;
    NbValideRTT: THLabel;
    NbAttenteRTT: THLabel;
    EnCoursRtt: THLabel;
    GBRecap: TGroupBox;
    LBN1: THLabel;
    AcquisN1: THLabel;
    PrisN1: THLabel;
    RestantN1: THLabel;
    LBN: THLabel;
    AcquisN: THLabel;
    PrisN: THLabel;
    RestantN: THLabel;
    LBRTT: THLabel;
    AcquisRtt: THLabel;
    PrisRTT: THLabel;
    RestantRTT: THLabel;
    LBAcquis: THLabel;
    LBPris: THLabel;
    LBRestant: THLabel;
    BtnDetail: TToolbarButton97;
    TCREERPRESENCE: TToolbarButton97;
    TCREERABSENCE: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);  { PT15 }
    procedure BFermeClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure TRECAPCLICK(Sender: TObject);
    procedure TCREERPRESENCECLICK(Sender: TObject);
    procedure TCREERABSENCECLICK(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BpersoClick(Sender: TObject);
    //Selection Item
    procedure HAvertirApplication(Sender: TObject; FromItem, ToItem: TOB; Actions: THPlanningAction);
    procedure BExporterClick(Sender: TObject);
    procedure TExcelClick(Sender: TObject);
    //DoubleClick sur ressource
    procedure DoubleCLickSpecPlanning(ACol, ARow: INTEGER; TypeCellule: TPlanningTypeCellule; T: TOB = nil);
    //DoubleClick sur itemDoubl
    procedure DoubleCLickPlanning(Sender: TObject);
    procedure BtnDetailClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure LoadDesignPlanning; { PT18  }
    procedure InitialiseOngletRecap; { PT18 }

  private
    { Déclarations privées }
    H: THPlanning;
    TypUtilisat, StWhTobEtats, StWhTobItems, StWhTobPres, StWhTobRessource, StUtilisat, StAbs: string;    { PT7 }
    StFiltreAbs,StTypeConge,StNivHierar, GbMode,GbTypeSoldeAbs  : string;  { PT9-2 PT14 PT17 PT18 }
    FontColor, FontName, FontStyle, ColorBackground: string;
    FontSize: Integer;
    TobRessources, TobItems, TobEtats, Tob_Recapitulatif, TobAbsDelete, TobMotifAbs {PT2-1,Tob_Encours}: TOB;
    InitDateDebAbs, InitDateFinAbs, DateDebAbs, DateFinAbs: TDateTime;
    Rupture: TNiveauRupture;
    MultiNiveau, OnShow: Boolean;
    ChampLibOrg : TChampLibOrg; { PT17 }
//    ChampsOrg: array[1..4] of string;  PT17
//    Champslibre: array[1..4] of string; PT17
    GbJourAbs,GbSoldeAbs : Double; { PT17-2 PT18 }
    GbTDerItem : Tob; { PT18  }
    NbrColonnesEntete : Integer;
    function  ChargeTobPlaning: boolean;
    function  LoadAbsencePeriode(DD, DF: TDateTime; OnLoad: Boolean): Boolean;
    procedure WMAppStartup(var msg: TMessage); message SBWM_APPSTARTUP;
    procedure RefrehRecapPlanning(sal: string; TItem: Tob);
    procedure ChargeTob_Recapitulatif;
    { DEB PT6 }
    procedure InitItemAbsence(Sender: TObject; var Item: TOB; var Cancel: boolean);
    procedure CreateItemAbsence(Sender: TObject; Item: TOB; var Cancel: boolean);
    procedure CreateItemPresence(Sender: TObject; Item: TOB; var Cancel: boolean); // pt20
    procedure ModifyItemAbsence(Sender: TObject; Item: TOB; var Cancel: boolean);
    Function  LanceFicheAbsenceModify(TEnr : Tob ) : Boolean;
    Procedure InitItemAbsenceModify (TEnr : Tob ; Var Ret : String);
    function  AccesUtilisat(Sal : String) : Boolean;
    Procedure DeleteItemAbsence(Sender: TObject; Item: TOB; var Cancel: boolean);
    function  SuppressionAbsence(Item : Tob; AvecRecap : Boolean): Boolean;
    procedure EtireItemAbsence(Sender: TObject; Source, Destination: TOB; Option: THPlanningOptionLink; var Cancel: boolean);
    Procedure PgAppliqCritereTotalJourAbs; { PT17-2 }
    Procedure PgAppliqCritereSoldeAbs; 	{ PT18 }

    { FIN PT6 }
    //    procedure IsOkItemAbsence
  public
    { Déclarations publiques }
  end;
procedure PGPlanningAbsPres(DateDebAbs, DateFinAbs: TDateTime; TypUtilisat, StWhTobEtats, StWhTobItems, StWhTobPres, StWhTobRessource, StUtilisat : string; NiveauRupt: TNiveauRupture;
  MultiNiveau: boolean = False; FiltreAbs : string = '' ; NivHierar :  string = '' ;TypeConge : string = ''; JourAbs : double = 0; SoldeAbs : double = 0; TypeSoldeAbs : string = ''; Mode : string = ''); { PT17 }

var
  PlanningAbsPre: TPlanningAbsPre;

implementation

uses  EntPaie,PgOutilsEagl ;

{$R *.DFM}

procedure PGPlanningAbsPres(DateDebAbs, DateFinAbs: TDateTime; TypUtilisat, StWhTobEtats, StWhTobItems, StWhTobPres, StWhTobRessource, StUtilisat : string; NiveauRupt: TNiveauRupture;
  MultiNiveau: boolean = False; FiltreAbs : string = '' ;  NivHierar : string = ''; TypeConge : string = ''; JourAbs : double = 0; SoldeAbs : double = 0; TypeSoldeAbs : string = ''; Mode : string = ''); { PT17 }
Var
  PP: THPanel;
begin    { DEB PT10 }
  PlanningAbsPre := nil;
  PP := FindInsidePanel();
  if Assigned(PP) then
  begin
    PlanningAbsPre := TPlanningAbsPre.Create(Application);
    InitInside(PlanningAbsPre, PP);
    PlanningAbsPre.DateDebAbs := DateDebAbs;
    PlanningAbsPre.DateFinAbs := DateFinAbs;
    PlanningAbsPre.TypUtilisat := TypUtilisat;
    PlanningAbsPre.StWhTobEtats := StWhTobEtats;
    PlanningAbsPre.StWhTobItems := StWhTobItems;
    PlanningAbsPre.StWhTobPres := StWhTobPres;
    PlanningAbsPre.StWhTobRessource := StWhTobRessource;
    PlanningAbsPre.StUtilisat := StUtilisat;
    PlanningAbsPre.StFiltreAbs := FiltreAbs;
    PlanningAbsPre.StTypeConge := TypeConge; { PT17 }
    PlanningAbsPre.StNivHierar := NivHierar; { PT14 }
    PlanningAbsPre.Rupture := NiveauRupt;
    PlanningAbsPre.MultiNiveau := MultiNiveau;
    PlanningAbsPre.GbJourAbs :=  JourAbs ; { PT17-2 }
    PlanningAbsPre.GbSoldeAbs :=  SoldeAbs ;   { PT18 }
    PlanningAbsPre.GbTypeSoldeAbs := TypeSoldeAbs ;  { PT18 }
    PlanningAbsPre.GbMode :=  Mode ;  { PT18 }
    PlanningAbsPre.Show;
  end
  else  { FIN PT10 }
  begin
    PlanningAbsPre := TPlanningAbsPre.Create(Application);
  try
    PlanningAbsPre.DateDebAbs := DateDebAbs;
    PlanningAbsPre.DateFinAbs := DateFinAbs;
    PlanningAbsPre.TypUtilisat := TypUtilisat;
    PlanningAbsPre.StWhTobEtats := StWhTobEtats;
    PlanningAbsPre.StWhTobItems := StWhTobItems;
    PlanningAbsPre.StWhTobPres := StWhTobPres;
    PlanningAbsPre.StWhTobRessource := StWhTobRessource;
    PlanningAbsPre.StUtilisat := StUtilisat;
    PlanningAbsPre.StFiltreAbs := FiltreAbs; { PT9-2 }
    PlanningAbsPre.StTypeConge := TypeConge; { PT17 }
    PlanningAbsPre.StNivHierar := NivHierar;  { PT14 }
    PlanningAbsPre.Rupture := NiveauRupt;
    PlanningAbsPre.MultiNiveau := MultiNiveau;
    PlanningAbsPre.GbJourAbs :=  JourAbs ; { PT17-2 }
    PlanningAbsPre.GbSoldeAbs :=  SoldeAbs ;   { PT18 }
    PlanningAbsPre.GbTypeSoldeAbs := TypeSoldeAbs ;  { PT18 }
    PlanningAbsPre.GbMode :=  Mode ;  { PT18 }
    PlanningAbsPre.ShowModal;
  finally
    if PlanningAbsPre <> nil then PlanningAbsPre.Free;
  end;
  end;
end;

procedure TPlanningAbsPre.WMAppStartup(var msg: TMessage);
begin
  CLose;
end;

procedure TPlanningAbsPre.FormShow(Sender: TObject);
var
  ListeColonnesEntete, TempColonne : String;
begin
  OnShow := True;
  FontColor := 'ClBlack';
  FontStyle := 'G';
  FontName := 'Times New Roman';
  FontSize := 10;
  ColorBackground := 'clWhite';

  InitDateDebAbs := PlanningAbsPre.DateDebAbs;
  InitDateFinAbs := PlanningAbsPre.DateFinAbs;

  InitialiseChampLibOrg(ChampLibOrg); { PT17 }

  if H <> nil then H.Free;
  H := THPlanning.Create(Application);
  H.Parent := PPanel;
  H.activate := False;
  { Ressources }
  H.ResChampID := 'PSA_SALARIE';
  { Etats }
  H.EtatChampCode := 'MOTIFETAT';             { PT15-1 }
  H.EtatChampLibelle := 'LIBMOTIFETAT';       { PT15-1 }
  H.EtatChampBackGroundColor := 'BACKGROUNDCOLOR';
  H.EtatChampFontColor := 'FONDCOLOR';
  H.EtatChampFontStyle := 'FONTSTYLE';
  H.EtatChampFontSize := 'FONTSIZE';
  H.EtatChampFontName := 'FONTNAME';
  { Items }
  H.ChampLineID := 'PCN_SALARIE';
  H.ChampdateDebut := 'PCN_DATEDEBUTABS';
  H.ChampDateFin := 'PCN_DATEFINABS';
  H.ChampEtat := 'MOTIFVALIDRESP';  { PT15-1 }
  H.ChampLibelle := 'ABREGE';
  H.ChampHint := 'PCN_LIBELLE';

  H.IntervalDebut := PlanningAbsPre.DateDebAbs;
  H.IntervalFin := PlanningAbsPre.DateFinAbs;

  { Liste des champs d'entete }
  H.TokenFieldColEntete := '';
  { Liste des champs d'entete }
  H.TokenFieldColFixed := ListeChampRupt(';',Rupture) + 'PSA_SALARIE;NOM_SAL;NOM_SERVICE';  { PT15-2 } { PT17 }
//  H.TokenAlignColFixed := 'C';
  MiseEnFormeColonneFixeRupt(H,Rupture,3); { PT12 }

  ListeColonnesEntete := H.TokenFieldColFixed;
  TempColonne := READTOKENST(ListeColonnesEntete);
  NbrColonnesEntete := 0;
  While TempColonne <> '' do
  begin
    Inc(NbrColonnesEntete);
    TempColonne := READTOKENST(ListeColonnesEntete);
  end;

  { DEB PT18 }
  if GBMode = 'SUIVI' then
  Caption := 'Suivi des absences/présences du ' + DateToStr(H.IntervalDebut) + ' au ' + DateToStr(H.IntervalFin)
  else
  Caption := 'Planning des absences/présences du ' + DateToStr(H.IntervalDebut) + ' au ' + DateToStr(H.IntervalFin);
{ FIN PT18 }


   {Création des trois TOB }
  if not ChargeTobPlaning then PostMessage(Handle, SBWM_APPSTARTUP, 0, 0)
  else
  begin
    InitialiseOngletRecap;  { PT18 }
    LoadDesignPlanning;   { PT18 }
    LoadConfigPlanning(H);  { PT12 }
    H.activate := True;
  end;
end;


function TPlanningAbsPre.ChargeTobPlaning: boolean;
var
  Q: TQuery;
  T, T1, T2, T3, TobCalend, Tob_Sal, Tob_Motifs, Tob_ValidResp : Tob;  { PT15-1 }
  St, StSal, StCond : string;
  NomChamp: array[1..3] of string;
  ValChamp: array[1..3] of variant;
  DateDebAbs, DateFinAbs: TdateTime;
  TJ: integer;
  DateCalculer, StRupture, StOrg, StRes, StDate : string;    { PT7 }
  I, j : Integer;
begin
  result := False;
  if TobItems <> nil then TobItems.Free;
  if TobEtats <> nil then TobEtats.Free;
  if TobRessources <> nil then TobRessources.Free;
  //if TypUtilisat<>'SAL' then
//  if Tob_Recapitulatif<>nil then Tob_Recapitulatif.free;
  TobRessources := TOB.Create('Les ressources', nil, -1);
  Tob_Sal := TOB.Create('Les salariés', nil, -1);
  TobItems := TOB.Create('Les items', nil, -1);
  TobEtats := TOB.Create('Les états', nil, -1);
  Tob_Motifs := TOB.Create('Les motifs', nil, -1);             { PT15-1 }
  Tob_ValidResp := TOB.Create('etat validation', nil, -1);     { PT15-1 }

  //if TypUtilisat<>'SAL' then
  Q := nil;

  { Chargement des ressources }
  StSal := PlanningAbsPre.StWhTobRessource;
  if Length(StSal) < 10 then StSal := '' else StSal := ' AND ' + Copy(StSal, 6, Length(StSal)); { PT15-2 }
  { DEB PT7 }
  StRes := '';  StAbs := '';
  StOrg := RendCritereSalarie(MultiNiveau,TypUtilisat,StUtilisat,StNivHierar); { PT15-2 }
  if StOrg <> '' then
    Begin
    StAbs := 'AND PCN_SALARIE IN '+ StOrg +') ';
    StRes := 'AND PSE_SALARIE IN '+ StOrg +') ';
    End;
  { DEB PT15-2 }
  StRupture := ListeChampRupt(',',Rupture);
  if PlanningAbsPre.DateDebAbs = PlanningAbsPre.DateFinAbs then
       StDate := 'AND ABSSAL.PCN_DATEDEBUTABS<="' + UsDateTime(PlanningAbsPre.DateDebAbs) + '" '+
                 'AND ABSSAL.PCN_DATEFINABS>="' + UsDateTime(PlanningAbsPre.DateFinAbs) + '" '
  else StDate := 'AND ((ABSSAL.PCN_DATEDEBUTABS>="' + UsDateTime(PlanningAbsPre.DateDebAbs) + '" '+
                 'AND ABSSAL.PCN_DATEDEBUTABS<="' + UsDateTime(PlanningAbsPre.DateFinAbs) + '") ' +
                 'OR (ABSSAL.PCN_DATEFINABS>="' + UsDateTime(PlanningAbsPre.DateDebAbs) + '" '+
                 'AND ABSSAL.PCN_DATEFINABS<="' + UsDateTime(PlanningAbsPre.DateFinAbs) + '") ' +
                 'OR (ABSSAL.PCN_DATEDEBUTABS<="' + UsDateTime(PlanningAbsPre.DateDebAbs) + '" '+
                 'AND ABSSAL.PCN_DATEFINABS>="' + UsDateTime(PlanningAbsPre.DateFinAbs) + '")) ';
  { DEB PT9-2 }
  If  (StFiltreAbs='SA1') OR (StFiltreAbs='SA2') then StCond := StTypeConge  { PT17 }
  else StCond := '';
  if (StFiltreAbs='ABS') OR ((StFiltreAbs='SA1') AND (StCond <> '')) then { PT17 }
    StRes := 'AND PSE_SALARIE IN (SELECT ABSSAL.PCN_SALARIE '+
    'FROM ABSENCESALARIE ABSSAL,MOTIFABSENCE '+
    'WHERE ABSSAL.PCN_TYPECONGE=PMA_MOTIFABSENCE  '+
    'AND ##PMA_PREDEFINI## (ABSSAL.PCN_TYPEMVT="ABS" OR ABSSAL.PCN_TYPECONGE="PRI") '+
    'AND ABSSAL.PCN_ETATPOSTPAIE <> "NAN" '+ { PT16 }
    'AND ABSSAL.PCN_MVTDUPLIQUE<>"X" AND PMA_MOTIFEAGL="X"  AND PMA_EDITPLANABS="X" '+
    StDate +' '+ StCond +') '+ StRes  { PT17 }
    else
    if (StFiltreAbs='PRE') OR ((StFiltreAbs='SA2') AND (StCond <> '')) then  { PT17 }
       StRes := 'AND PSE_SALARIE NOT IN (SELECT ABSSAL.PCN_SALARIE '+
      'FROM ABSENCESALARIE ABSSAL,MOTIFABSENCE  '+
      'WHERE ABSSAL.PCN_TYPECONGE=PMA_MOTIFABSENCE  '+
      'AND ##PMA_PREDEFINI## (ABSSAL.PCN_TYPEMVT="ABS" OR ABSSAL.PCN_TYPECONGE="PRI") '+
      'AND ABSSAL.PCN_ETATPOSTPAIE <> "NAN" '+ { PT16 }
      'AND ABSSAL.PCN_MVTDUPLIQUE<>"X" AND PMA_MOTIFEAGL="X" AND PMA_EDITPLANABS="X" '+
      StDate +' '+ StCond +') ' + StRes;  { PT17 }
  { FIN PT9-2 }
  { FIN PT15-2 }
  StRes := 'SELECT DISTINCT PSE_SALARIE ,PSE_CODESERVICE, ' + StRupture +
    ' PSA_SALARIE,SUBSTRING(PSA_LIBELLE,1,36)||" "||PSA_PRENOM NOM_SAL ' +
    'FROM DEPORTSAL,SALARIES ' +
    'WHERE PSE_SALARIE=PSA_SALARIE ' + StSal + StRes ;
    //'ORDER BY ' + StRupture + 'PSE_CODESERVICE,NOM_SAL,PSA_SALARIE';   { PT15-2 }
  Q := OpenSql(StRes, True);
  Tob_Sal.LoadDetailDB('SAL_ARIES', '', '', Q, False);
  Ferme(Q);


  MiseEnFormeTobRupt('RES', Tob_Sal,TobRessources,Rupture);   { PT12 }

  {StRes := ' AND PCN_SALARIE IN (SELECT DISTINCT PSE_SALARIE FROM DEPORTSAL ';
  if Pos('PSA_', StOrg) > 0 then StRes := StRes + 'LEFT JOIN SALARIES ON PSE_SALARIE=PSA_SALARIE ' + StOrg + ')'
  else StRes := StRes + StOrg + ')';   }
  { FIN PT7 }
  { Chargement des états }
  { DEB PT15-1 }
  St := 'SELECT * FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## PMA_MOTIFEAGL="X" AND PMA_EDITPLANABS="X" ';
  Q := OpenSql(St, True);
  if not Q.Eof then Tob_Motifs.LoadDetailDB('Les motifs', '', '', Q, False);
  Ferme(Q);

  St := 'SELECT CO_TYPE,CO_CODE,CO_LIBELLE FROM COMMUN WHERE CO_TYPE="PAE"';
  Q := OpenSql(St, True);
  if not Q.eof then Tob_ValidResp.LoadDetailDB('etat validation', '', '', Q, False);
  Ferme(Q);

  { Boucle sur les motifs }
  For j:=0 to Tob_Motifs.detail.count-1 do
     Begin
     T1 := Tob_Motifs.Detail[j];
     { Boucle sur l'état de la validation }
     For i := 0 to Tob_ValidResp.Detail.COunt - 1 do
       begin
       T2 := Tob_ValidResp.Detail[i];
       T := TOB.Create('Les états', TobEtats, -1);
       T.AddChampSupValeur('MOTIFETAT',T1.GetValue('PMA_MOTIFABSENCE')+T2.GetValue('CO_CODE'));
       T.AddChampSupValeur('LIBMOTIFETAT',T2.GetValue('CO_LIBELLE'));
       if T2.GetValue('CO_CODE')='VAL' then T.AddChampSupValeur('BACKGROUNDCOLOR',T1.GetValue('PMA_PGCOLORVAL'))
       else if T2.GetValue('CO_CODE')='ATT' then T.AddChampSupValeur('BACKGROUNDCOLOR',T1.GetValue('PMA_PGCOLORATT'))
       else if T2.GetValue('CO_CODE')='NAN' then T.AddChampSupValeur('BACKGROUNDCOLOR',T1.GetValue('PMA_PGCOLORANN'))
       else if T2.GetValue('CO_CODE')='REF' then T.AddChampSupValeur('BACKGROUNDCOLOR',T1.GetValue('PMA_PGCOLORREF'))
       else T.AddChampSupValeur('BACKGROUNDCOLOR',T1.GetValue('PMA_PGCOLORACTIF'));
      
       T.AddChampSupValeur('FONDCOLOR', FontColor);
       T.AddChampSupValeur('FONTNAME', FontName);
       T.AddChampSupValeur('FONTSTYLE', FontStyle);
       T.AddChampSupValeur('FONTSIZE', FontSize);
       end;
     end;
  FreeAndNil(Tob_ValidResp);
  FreeAndNil(Tob_Motifs);
  { FIN PT15-1 }   
  { Chargement des items }
  LoadAbsencePeriode(PlanningAbsPre.DateDebAbs, PlanningAbsPre.DateFinAbs, True);

  //Chargement du récapitulatif
  //ChargeTob_Recapitulatif; { PT17 }


   { Suppression des salariés si critère total jour d'absence non rempli }
   PgAppliqCritereTotalJourAbs; { PT17-2 }
   PgAppliqCritereSoldeAbs;     { PT18 }

  {PT2-1 Mise en commentaire : on utilise désormais les champs de la table puis supprimé }

  if (H.Interval = PiDemiJour) then
  begin
    { Chargement des salaries}
    StSal := PlanningAbsPre.StWhTobRessource;
    Q := OpenSql('SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,psa_standcalend,PSA_CALENDRIER,' +
      'etb_standcalend FROM SALARIES LEFT join etabcompl on psa_etablissement=etb_etablissement ' +
      'LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE ' + StSal, True);
    Tob_Sal.LoadDetailDB('SAL_ARIES', '', '', Q, False);
    Ferme(Q);
    { Chargement des calendriers }
    Q := OpenSql('SELECT * FROM CALENDRIER', True);
    TobCalend := Tob.create('CAl_endrier', nil, -1);
    TobCalend.LoadDetailDB('CAL_ENDRIER', '', '', Q, False);
    Ferme(Q);
    T1 := Tob_Sal.FindFirst([''], [''], False);
    while T1 <> nil do
    begin
      if (T1.GetValue('PSA_STANDCALEND') = 'PER') then
      begin
        NomChamp[1] := 'ACA_STANDCALEN';
        NomChamp[2] := 'ACA_SALARIE';
        NomChamp[3] := 'ACA_JOUR';
        ValChamp[1] := T1.GetValue('PSA_CALENDRIER');
        ValChamp[2] := T1.GetValue('PSA_SALARIE');
      end
      else
        if (T1.GetValue('PSA_STANDCALEND') = 'ETB') then
      begin
        NomChamp[1] := 'ACA_STANDCALEN';
        NomChamp[2] := 'ACA_JOUR';
        NomChamp[3] := '';
        ValChamp[1] := T1.GetValue('ETB_STANDCALEND');
        ValChamp[3] := '';
      end
      else
        if (T1.GetValue('PSA_STANDCALEND') = 'ETS') then
      begin
        NomChamp[1] := 'ACA_STANDCALEN';
        NomChamp[2] := 'ACA_JOUR';
        NomChamp[3] := '';
        ValChamp[1] := T1.GetValue('PSA_CALENDRIER');
        ValChamp[3] := '';
      end;
      T2 := Tobitems.FindFirst(['PCN_SALARIE'], [T1.getValue('PSA_SALARIE')], False);
      while T2 <> nil do
      begin
        DateDebAbs := T2.GetValue('PCN_DATEDEBUTABS');
        TJ := DayOfWeek(DateDebAbs);
        if TJ = 1 then TJ := 7 else TJ := TJ - 1;
        if T1.GetValue('PSA_STANDCALEND') = 'PER' then ValChamp[3] := TJ else ValChamp[2] := TJ;
        T3 := TobCalend.FindFirst(NomChamp, ValChamp, False);
        if T3 <> nil then
        begin
          if T2.GetValue('PCN_DEBUTDJ') = 'MAT' then
            DateCalculer := DateToStr(DateDebAbs) + ' ' + FloatToStrTime(8, 'HH:MM:SS') //T3.GetValue('ACA_HEUREDEB1')
          else
            DateCalculer := DateToStr(DateDebAbs) + ' ' + FloatToStrTime(14, 'HH:MM:SS'); //T3.GetValue('ACA_HEUREDEB2')
          T2.PutValue('PCN_DATEDEBUTABS', StrToDateTime(DateCalculer));
        end;
        DateFinAbs := T2.GetValue('PCN_DATEFINABS');
        TJ := DayOfWeek(DateFinAbs);
        if TJ = 1 then TJ := 7 else TJ := TJ - 1;
        if T1.GetValue('PSA_STANDCALEND') = 'PER' then ValChamp[3] := TJ else ValChamp[2] := TJ;
        T3 := TobCalend.FindFirst(NomChamp, ValChamp, False);
        if T3 <> nil then
        begin
          if T2.GetValue('PCN_FINDJ') = 'MAT' then
            DateCalculer := DateToStr(DateFinAbs) + ' ' + FloatToStrTime(12, 'HH:MM:SS') //T3.GetValue('ACA_HEUREFIN1')
          else
            DateCalculer := DateToStr(DateFinAbs) + ' ' + FloatToStrTime(18, 'HH:MM:SS'); //T3.GetValue('ACA_HEUREFIN2')
          T2.PutValue('PCN_DATEFINABS', StrToDateTime(DateCalculer));
        end;
        T2 := Tobitems.FindNext(['PCN_SALARIE'], [T1.getValue('PSA_SALARIE')], False);
      end;
      T1 := Tob_Sal.FindNext([''], [''], False);
    end;
    TobCalend.free;
  end; //FIN (H.Interval=PiDemiJour)

  (*
    TProgress.Value:=TProgress.Value+1;
    TProgress.SubText:='Affectation du planning..';
  *)
//  tobdebug(TobRessources);
//  tobdebug(TobEtats);
// tobdebug(TobItems);

  H.TobRes := TobRessources;
  H.TobEtats := TobEtats;
  H.TobItems := TobItems;
  (*
    TProgress.free;
  *)

  { DEB PT6 }
  if Assigned(TobAbsDelete) then FreeAndNil(TobAbsDelete);
  TobAbsDelete := TOB.Create('ABSENCESALARIE', nil, -1);
  { FIN PT6 }

  if Tob_Sal <> nil then Tob_Sal.free;
  result := (TobRessources.detail.count > 0);
  If not (TobRessources.detail.count > 0) then { PT17 }
    Begin
    PgiInfo('Aucun salarié ne remplit les critères sélectionnés.',Caption); { PT17 }
    if IsInside(Self) then THPanel(parent).CloseInside;
    End;
			{ DEB PT18 }
  if GbMode = 'SUIVI' then
     Begin
     BExporterClick(nil);
     FreeAndNil(TobRessources);
     result := False;
     End; { FIN PT18 }            
end;

procedure TPlanningAbsPre.DoubleCLickPlanning(Sender: TObject);
var
  T: TOB;
begin
  T := H.GetCurItem;
  if Assigned(T) then LanceFicheAbsenceModify(T); { PT6 }
end;

procedure TPlanningAbsPre.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;   { PT15-2 }                     
end;


procedure TPlanningAbsPre.BFermeClick(Sender: TObject);
begin
  PlanningAbsPre.Close;
  if IsInside(Self) then THPanel(parent).CloseInside;
end;

procedure TPlanningAbsPre.BFirstClick(Sender: TObject);
var
  YY, MM, JJ: WORD;
  DD: TDateTime;
begin
  if H = nil then exit;
  DD := H.IntervalDebut;
  DecodeDate(H.IntervalDebut, YY, MM, JJ);
  if (MM < 6) or ((MM = 6) and (JJ = 1)) then
  begin
    YY := YY - 1;
    DD := DD - 1;
  end;
  H.Activate := False;
  DateDebAbs := EncodeDate(YY, 6, 1);
  LoadAbsencePeriode(DateDebAbs, DD, False);
  H.IntervalDebut := DateDebAbs;
  H.DateOfStart := DateDebAbs;
  Caption := 'Planning des absences du ' + DateToStr(H.IntervalDebut) + ' au ' + DateToStr(H.IntervalFin);
  H.Activate := True;
end;

procedure TPlanningAbsPre.BLastClick(Sender: TObject);
var
  YY, MM, JJ: WORD;
  DF: TDateTime;
begin
  if H = nil then exit;
  DF := H.IntervalFin + 1;
  DecodeDate(H.IntervalFin, YY, MM, JJ);
  if (MM > 5) or ((MM = 5) and (JJ = 31)) then YY := YY + 1;
  H.Activate := False;
  DateFinAbs := EncodeDate(YY, 5, 31);
  LoadAbsencePeriode(DF, DateFinAbs, False);
  H.IntervalFin := DateFinAbs;
  H.DateOfStart := DF;
  Caption := 'Planning des absences du ' + DateToStr(H.IntervalDebut) + ' au ' + DateToStr(H.IntervalFin);
  H.Activate := True;
end;

procedure TPlanningAbsPre.BNextClick(Sender: TObject);
var
  DF: TDateTime;
begin
  if H = nil then exit;
  H.Activate := False;
  DF := H.IntervalFin + 1;
  DateFinAbs := FindeMois(PlusMois(H.IntervalFin, 1));
  LoadAbsencePeriode(DF, DateFinAbs, False);
  H.IntervalFin := DateFinAbs;
  H.DateOfStart := DF;
  Caption := 'Planning des absences du ' + DateToStr(H.IntervalDebut) + ' au ' + DateToStr(H.IntervalFin);
  H.Activate := True;
end;

procedure TPlanningAbsPre.BPrevClick(Sender: TObject);
var
  DD: TdateTime;
begin
  if H = nil then exit;
  H.Activate := False;
  DD := H.IntervalDebut - 1;
  DateDebAbs := DebutdeMois(PlusMois(H.IntervalDebut, -1));
  LoadAbsencePeriode(DateDebAbs, DD, False);
  H.IntervalDebut := DateDebAbs;
  H.DateOfStart := DateDebAbs;
  Caption := 'Planning des absences du ' + DateToStr(H.IntervalDebut) + ' au ' + DateToStr(H.IntervalFin);
  H.Activate := True;
end;

procedure TPlanningAbsPre.TRECAPCLICK(Sender: TObject);
begin
  AgllanceFiche('PAY', 'MULRECAPSAL', '', '', TypUtilisat);
end;

procedure TPlanningAbsPre.BImprimerClick(Sender: TObject);
var
  S: string;
begin
  H.TypeEtat := 'E';
  H.NatureEtat := 'PCG';
  H.CodeEtat := 'PL1';
  S := 'DATEDEB=' + FormatDateTime('dd/mm/yyyy', H.IntervalDebut) + '^E' + IntToStr(Byte(otDate));
  S := S + '`DATEFIN=' + FormatDateTime('dd/mm/yyyy', H.IntervalFin) + '^E' + IntToStr(Byte(otDate));
  H.CriteresToPrint := S;
  H.Print;
end;

function TPlanningAbsPre.LoadAbsencePeriode(DD, DF: TDateTime; OnLoad: Boolean): Boolean;
var
  StWhere, St, StDate, stPres, StAbsence: string;      // Pt20
  Q: TQuery;
  TobAbsence, T1, T2, Texist: Tob;
  i: integer;
begin
  { Chargement des items }
  //  result := False;

  if StFiltreAbs='PRE' then exit;    { PT15-2 }

  FreeAndNil(TobAbsence);  { PT6 }
  StWhere := PlanningAbsPre.StWhTobItems;
  if Pos('WHERE', StWhere) > 0 then StWhere := 'AND ' + Copy(StWhere, Pos('WHERE', StWhere) + 5, Length(StWhere));
  // Deb Pt20
  StPres := PlanningAbsPre.StWhTobPres;
  if Pos('WHERE', StPres) > 0 then StPres :=  Copy(StPres, Pos('WHERE', StPres) + 5, Length(StPres));
  if stwhere = '' then Stpres := 'AND ' + StPres else   Stpres := 'OR ' +  Stpres;

  // fin Pt20

  {  if StFiltreAbs='ABS' then
    StPer :=

  else

  if StFiltreAbs='PRE' else}

  //PT2-2 optimisation requête SQL
  { DEB PT15-2 }
  if DD = DF then StDate := 'AND PCN_DATEDEBUTABS<="' + UsDateTime(DD) + '" AND PCN_DATEFINABS>="' + UsDateTime(DF) + '" '
  else StDate := 'AND ((PCN_DATEDEBUTABS>="' + UsDateTime(DD) + '" AND PCN_DATEDEBUTABS<="' + UsDateTime(DF) + '") ' + //PT1 Ajout condition
    'OR (PCN_DATEFINABS>="' + UsDateTime(DD) + '" AND PCN_DATEFINABS<="' + UsDateTime(DF) + '") ' +
    'OR (PCN_DATEDEBUTABS<="' + UsDateTime(DD) + '" AND PCN_DATEFINABS>="' + UsDateTime(DF) + '")) '; //PT4 { PT7 }
  { FIN PT15-2 }
  St := 'SELECT PCN_ORDRE,PCN_SALARIE,PCN_TYPEMVT,PCN_TYPECONGE,PCN_ETABLISSEMENT,' +
    'PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_SENSABS,PCN_JOURS,PCN_HEURES,PCN_LIBELLE,' +
    'PCN_APAYES,PCN_CODETAPE,PCN_DATEVALIDITE,PCN_ETATPOSTPAIE,PCN_EXPORTOK,' +
    'PCN_VALIDRESP,PCN_VALIDSALARIE,PCN_DEBUTDJ,PCN_FINDJ,PCN_MVTDUPLIQUE,' +
    'PMA_CALENDCIVIL,PMA_MOTIFEAGL,PMA_TYPEABS ' + { PT5 }
    'FROM ABSENCESALARIE ' +
    ' LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PCN_TYPECONGE=PMA_MOTIFABSENCE  ' +
    'WHERE (PCN_TYPEMVT="ABS" OR PCN_TYPEMVT = "PRE" OR PCN_TYPECONGE="PRI") AND PCN_MVTDUPLIQUE<>"X" ' +
    'AND PCN_ETATPOSTPAIE <> "NAN" '+  { PT16 }
    'AND PMA_MOTIFEAGL="X" AND PMA_EDITPLANABS="X" ' +StDate + StWhere + Stpres + StAbs;   { PT15-1  PT20}
  Q := Opensql(St, True);
  if not Q.eof then
  begin
    TobAbsence := Tob.Create('ABSENCE_SALARIE', nil, -1);
    TobAbsence.LoadDetailDB('ABSENCE_SALARIE', '', '', Q, False);
  end
  else
  begin
    PGIInfo('Aucun mouvement d''absence et de présence à éditer sur le planning du ' + DateToStr(DD) + ' au ' + DateToStr(DF) + '!', Caption);
    //Ferme(Q);
    //Exit;
  end;
  Ferme(Q);

  if Assigned(TobAbsence) then   { PT6 }
         for i := 0 to TobAbsence.detail.count - 1 do
         begin
         T1 := TobAbsence.detail[i];
         T1.AddChampSup('ABREGE', False);
         // pt20
         if T1.Getvalue('PCN_TYPEMVT') = 'PRE' then
         T1.PutValue('ABREGE', RechDom('PGMOTIFPRESENCE', T1.GetValue('PCN_TYPECONGE'), False))
         else
         T1.PutValue('ABREGE', RechDom('PGMOTIFABSENCE', T1.GetValue('PCN_TYPECONGE'), False));

         T1.AddChampSupValeur('MOTIFVALIDRESP',T1.GetValue('PCN_TYPECONGE')+T1.GetValue('PCN_VALIDRESP')); { PT15-1 }
         if TobItems <> nil then
           Texist := tobItems.FindFirst(['PCN_SALARIE', 'PCN_TYPEMVT', 'PCN_ORDRE'], [T1.GetValue('PCN_SALARIE'), T1.GetValue('PCN_TYPEMVT'), T1.GetValue('PCN_ORDRE')], False)
         else Texist := nil;
         if Texist = nil then
           begin
           T2 := Tob.Create('ABSENCE_SALARIE', tobItems, -1);
           T2.Dupliquer(T1, True, True, True);
           end;
         end;
  H.TobItems := TobItems;
  if Assigned(TobAbsence) then TobAbsence.free;
  { PT6 Chargement des motifs absence }
  TobMotifAbs := Tob.Create('MOTIF_ABSENCE',nil,-1);
  TobMotifAbs.LoadDetailDB('MOTIFABSENCE','','',nil,False);
  result := True;
end;

procedure TPlanningAbsPre.BpersoClick(Sender: TObject);
{$IFDEF ABSETEMPS}
var
  eColorStart, eColorEnd: TColor;
  {$ENDIF}
begin
  {$IFDEF ABSETEMPS}
  eCongePersoPlanning('ConfigPlanningConges', H, eColorStart, eColorEnd);
  {$ENDIF}
  H.Refresh;
end;


procedure TPlanningAbsPre.HAvertirApplication(Sender: TObject; FromItem,
  ToItem: TOB; Actions: THPlanningAction);
var
  Sal: string;
begin
  //if TypUtilisat='SAL' then exit;
  {if Assigned(FromItem) then
     Begin
     FreeAndNil(GbTDerItem);
     GbTDerItem.Dupliquer(FromItem,True,True);
     End;                                     }

  if Tob_Recapitulatif = nil then exit;
  if Actions = paclickLeft then
  begin
    if FromItem <> nil then
    begin
      Sal := FromItem.GetValue('PCN_SALARIE');
      RefrehRecapPlanning(sal, FromItem);
      //H.Raffraichir;
    end;
  end;
end;

procedure TPlanningAbsPre.BExporterClick(Sender: TObject);
begin
  if not assigned(Tob_Recapitulatif) then ChargeTob_Recapitulatif; { PT17 }
  PgGenereTobExport(TobRessources,TobItems,Tob_Recapitulatif,H.TokenFieldColFixed,caption,GbMode); { PT17 } { PT18 }
end;


procedure TPlanningAbsPre.TExcelClick(Sender: TObject);
begin
  with TSaveDialog.Create(Self) do
  begin
    Filter := 'Microsoft Excel (*.xls)|*.XLS';
    if Execute then H.ExportToExcel(True, FileName);
    Free;
  end;
end;

procedure TPlanningAbsPre.DoubleCLickSpecPlanning(ACol, ARow: INTEGER;
  TypeCellule: TPlanningTypeCellule; T: TOB);
begin
  case TypeCellule of
    ptcRessource: if T <> nil then
        RefrehRecapPlanning(T.GetValue('PSA_SALARIE'), nil);
  end;
end;

procedure TPlanningAbsPre.RefrehRecapPlanning(sal: string; TItem: Tob);
var
  T: Tob;
  //ValPri,AttPri,ValRtt,AttRtt : double;
begin
  {PT2-1 Mise en commentaire : on utilise désormais les champs de la table
  ValPri:=Tob_Encours.Somme('JOURS',['PCN_SALARIE','PCN_TYPECONGE','PCN_VALIDRESP'],[Sal,'PRI','VAL'],False);
  AttPri:=Tob_Encours.Somme('JOURS',['PCN_SALARIE','PCN_TYPECONGE','PCN_VALIDRESP'],[Sal,'PRI','ATT'],False);
  ValRtt:=Tob_Encours.Somme('JOURS',['PCN_SALARIE','PCN_TYPECONGE','PCN_VALIDRESP'],[Sal,'RTT','VAL'],False);
  ValRtt:=ValRtt+Tob_Encours.Somme('JOURS',['PCN_SALARIE','PCN_TYPECONGE','PCN_VALIDRESP'],[Sal,'RTS','VAL'],False);
  AttRtt:=Tob_Encours.Somme('JOURS',['PCN_SALARIE','PCN_TYPECONGE','PCN_VALIDRESP'],[Sal,'RTT','ATT'],False);
  AttRtt:=AttRtt+Tob_Encours.Somme('JOURS',['PCN_SALARIE','PCN_TYPECONGE','PCN_VALIDRESP'],[Sal,'RTS','ATT'],False);
  }
  if OnShow then
  begin
    if not assigned(Tob_Recapitulatif) then ChargeTob_Recapitulatif; { PT17 }
    OnShow := False;
    BtnDetail.Down := True;
    PRecap.visible := True;
  end;
  if Tob_Recapitulatif <> nil then
    T := Tob_Recapitulatif.FindFirst(['PRS_SALARIE'], [Sal], False)
  else T := nil;
  if T <> nil then
  begin
    GBRecap.Caption := 'Récapitulatif des absences du salarié ' + T.GetValue('PRS_SALARIE') + ' ' + T.GetValue('PRS_LIBELLE') + ' ' + T.GetValue('PRS_PRENOM') + ' au ' +
      DateToStr(VH_Paie.PGEcabDateIntegration);
    if TItem <> nil then
      GBAbsence.Caption := FloatToStr(TItem.GetValue('PCN_JOURS')) + ' jour(s) de ' + RechDom('PGMOTIFABSENCE', TItem.GetValue('PCN_TYPECONGE'), False)
    else
      GBAbsence.Caption := '';
    NbValidePri.Caption := FloatToStr(T.GetValue('PRS_PRIVALIDE')); //DEB PT2-1
    NbAttentePri.Caption := FloatToStr(T.GetValue('PRS_PRIATTENTE'));
    NbValideRTT.Caption := FloatToStr(T.GetValue('PRS_RTTVALIDE'));
    NbAttenteRTT.Caption := FloatToStr(T.GetValue('PRS_RTTATTENTE')); //FIN PT2-1
    AcquisN1.caption := FloatToStr(T.GetValue('PRS_ACQUISN1'));
    PrisN1.caption := FloatToStr(T.GetValue('PRS_PRISN1'));
    RestantN1.caption := FloatToStr(T.GetValue('PRS_RESTN1'));
    AcquisN.caption := FloatToStr(T.GetValue('PRS_ACQUISN'));
    PrisN.caption := FloatToStr(T.GetValue('PRS_PRISN'));
    RestantN.caption := FloatToStr(T.GetValue('PRS_RESTN'));
    AcquisRtt.caption := FloatToStr(T.GetValue('PRS_CUMRTTACQUIS'));
    PrisRTT.caption := FloatToStr(T.GetValue('PRS_CUMRTTPRIS'));
    RestantRTT.caption := FloatToStr(T.GetValue('PRS_CUMRTTREST')); //PT2-1
  end;
end;

procedure TPlanningAbsPre.ChargeTob_Recapitulatif;
var
  StSal, Join, st: string;
  Q: TQuery;
begin
  if Tob_Recapitulatif <> nil then
  begin
    Tob_Recapitulatif.free;
    Tob_Recapitulatif := nil;
  end;
  StSal := PlanningAbsPre.StWhTobRessource;
  if StSal <> '' then
  begin
    Join := '';
    if Pos('PSA_', StSal) > 0 then Join := 'LEFT JOIN SALARIES ON PSA_SALARIE=PRS_SALARIE ';
    if Pos('PSE_', StSal) > 0 then Join := Join + 'LEFT JOIN DEPORTSAL ON PSE_SALARIE=PRS_SALARIE';
    StSal := Join + ' ' + StSal;
  end;
  { DEB PT15-2 }
  st := RendCritereSalarie(MultiNiveau,TypUtilisat,StUtilisat,StNivHierar);
  if (St <>'') and (StSal<>'') then StSal := StSal + 'AND PRS_SALARIE IN '+St+')'
  else
    if (St <>'') and (StSal='') then StSal := StSal + 'WHERE PRS_SALARIE IN '+St+')';
  { FIN PT15-2 }


  st := 'SELECT PRS_SALARIE,PRS_LIBELLE,PRS_PRENOM,PRS_DATEMODIF,PRS_ACQUISN,PRS_ACQUISN1,' +
    'PRS_PRISN,PRS_PRISN1,PRS_CUMRTTACQUIS,PRS_CUMRTTPRIS,PRS_RESTN,PRS_RESTN1, ' +
    'PRS_CUMRTTREST,PRS_PRIATTENTE,PRS_PRIVALIDE,PRS_RTTATTENTE,PRS_RTTVALIDE ' + //PT2-1
  'FROM RECAPSALARIES ' + StSal;
  Q := Opensql(St, True);
  if not Q.eof then
  begin
    Tob_Recapitulatif := TOB.Create('Le récapitulatif', nil, -1);
    Tob_Recapitulatif.LoadDetailDB('RECAP_SALARIE', '', '', Q, False);
  end
  else
  begin
    GBRecap.visible := False;
    GBAbsence.Visible := False; //PT2-1
  end;
  Ferme(Q);
end;

procedure TPlanningAbsPre.BtnDetailClick(Sender: TObject);
begin
  if not assigned(Tob_Recapitulatif) then ChargeTob_Recapitulatif; { PT17 }
  PRecap.visible := BtnDetail.Down;
end;
{ DEB PT6 }
procedure TPlanningAbsPre.InitItemAbsence(Sender: TObject; var Item: TOB;
  var Cancel: boolean);
begin
  Cancel := True;
  Item := Tob.Create('ABSENCESALARIE', TobItems, -1);
  Item.AddChampSupValeur('CREATION', 'TEMP', False);
  Cancel := false;
end;


procedure TPlanningAbsPre.CreateItemAbsence(Sender: TObject; Item: TOB;
  var Cancel: boolean);
var
  Sal, Lib, DD, DF, Ret : string;
begin
  Cancel := True;

  Sal := Item.GetString('PCN_SALARIE');
  If not AccesUtilisat(Sal) then exit;
  lib := RechDom('PGSALARIE',Sal,False);
  DD  := Item.GetString('PCN_DATEDEBUTABS');
  DF  := Item.GetString('PCN_DATEFINABS');
  Ret := AglLanceFiche('PAY', 'EABSENCE', 'PCN_SALARIE=' + Sal, Sal, Sal + '!' + lib + ';E;;ACTION=CREATION;;' + PlanningAbsPre.TypUtilisat+';'+DD+';'+DF);
  If Ret<>'' then
     Begin
     H.Activate := False;
     InitItemAbsenceModify(Item,Ret);
     ChargeTob_Recapitulatif;
     RefrehRecapPlanning(sal, Item);
     //tobdebug(item);
    // pt20  if Item <> nil then H.InvalidateItem(Item);
     H.Activate := true;
     Cancel := false;
     End;
end;
{ FIN PT6 }

{ DEB PT6 }
procedure TPlanningAbsPre.ModifyItemAbsence(Sender: TObject; Item: TOB;
  var Cancel: boolean);
begin
  Cancel := LanceFicheAbsenceModify (Item);
end;

function TPlanningAbsPre.LanceFicheAbsenceModify(TEnr: Tob): Boolean;
Var
  Sal, TypeMvt, Ordre, Etab, ValidResp, exportok, StAction, ret, DD, DF: string;
begin
    Result := True ;
    if not Assigned(TEnr) then exit;
    Sal := TEnr.GetValue('PCN_SALARIE');
    If not AccesUtilisat(Sal) then exit;
    TypeMvt        := TEnr.GetValue('PCN_TYPEMVT');
    Ordre          := IntToStr(TEnr.GetValue('PCN_ORDRE'));
    Etab           := TEnr.GetValue('PCN_ETABLISSEMENT');
    ValidResp      := TEnr.GetValue('PCN_VALIDRESP');
    exportok       := TEnr.GetValue('PCN_EXPORTOK');
    DD             := TEnr.GetString('PCN_DATEDEBUTABS');
    DF             := TEnr.GetString('PCN_DATEFINABS');

    // pt20   (si type de mvt <> 'presence" )
  if Typemvt <> 'PRE' then
  begin

    if PlanningAbsPre.TypUtilisat = 'SAL' then
    begin
      if ValidResp = 'ATT' then StAction := 'ACTION=MODIFICATION' else StAction := 'ACTION=CONSULTATION';
      ret := AglLanceFiche('PAY', 'EABSENCE', 'PCN_SALARIE=' + sal, TYPEMVT + ';' + sal + ';' + Ordre, sal + ';E;' + Etab + ';' + StAction + ';;' + PlanningAbsPre.TypUtilisat+';'+DD+';'+DF)
    end
    else
    begin
      if ExportOk = '-' then StAction := 'ACTION=MODIFICATION' else StAction := 'ACTION=CONSULTATION';
      if IfMotifabsenceSaisissable(TEnr.GetValue('PCN_TYPECONGE'), PlanningAbsPre.TypUtilisat) then StAction := 'ACTION=MODIFICATION'; //PT3
      if (PlanningAbsPre.TypUtilisat = 'RESP') or (PlanningAbsPre.TypUtilisat = 'ADM') or (PlanningAbsPre.TypUtilisat = '') then
        ret := AglLanceFiche('PAY', 'EABSENCE', 'PCN_SALARIE=' + sal, TYPEMVT + ';' + sal + ';' + Ordre, sal + ';E;' + Etab + ';' + StAction + ';;' + PlanningAbsPre.TypUtilisat+';'+DD+';'+DF)
    end;
  end  // pt20 : appel fiche présence
  else
     ret := AglLanceFiche('PAY', 'MVTPRESENCE', 'PCN_SALARIE=' + sal, TYPEMVT + ';' + sal + ';' + Ordre, sal + ';P;' + Etab + ';' + StAction + ';;' + PlanningAbsPre.TypUtilisat+';'+DD+';'+DF);

    if (ret <> '') and (PlanningAbsPre.TypUtilisat <> 'SAL') then
    begin
      Result := False;
      if Pos('SUPPR',Ret)>0 then
        Begin
     //   H.Activate := False;
        H.DeleteItem(TEnr, False);
        ChargeTob_Recapitulatif;
        RefrehRecapPlanning(TEnr.GetString('PCN_SALARIE'), nil);
       // Tenr.ChangeParent(nil,-1);
       // FreeAndNil(Tenr);
       // H.Activate := True;
        End
      else
        Begin
        H.Activate := False;
        InitItemAbsenceModify(TEnr,Ret);
        ChargeTob_Recapitulatif;
        RefrehRecapPlanning(sal, TEnr);
        if TEnr <> nil then H.InvalidateItem(TEnr);
        H.Activate := True;

        End;
    end;
end;
{ FIN PT6 }

{ DEB PT6 }
procedure TPlanningAbsPre.InitItemAbsenceModify(TEnr: Tob ; Var Ret : String);
Var T_MotifAbs : Tob;
begin
  if Not Assigned(TEnr) then exit;
  if Ret = '' then exit;
  TEnr.PutValue('PCN_ORDRE'         ,StrToInt(ReadTokenSt(Ret)));
  TEnr.PutValue('PCN_TYPECONGE'     ,ReadTokenSt(Ret));
  // pt20
  if TEnr.GetString('PCN_TYPECONGE')='PRI' then
  TEnr.PutValue('PCN_TYPEMVT'       ,'CPA');
  // pt20
  //else
  // TEnr.PutValue('PCN_TYPEMVT'       ,'ABS');
  TEnr.PutValue('PCN_DATEDEBUTABS'  ,StrToDate(ReadTokenSt(Ret)));
  TEnr.PutValue('PCN_DATEFINABS'    ,StrToDate(ReadTokenSt(Ret)));
  TEnr.PutValue('PCN_DEBUTDJ'       ,ReadTokenSt(Ret));
  TEnr.PutValue('PCN_FINDJ'         ,ReadTokenSt(Ret));
  TEnr.PutValue('PCN_JOURS'         ,Valeur(ReadTokenSt(Ret)));
  TEnr.PutValue('PCN_HEURES'        ,Valeur(ReadTokenSt(Ret)));
  TEnr.PutValue('PCN_LIBELLE'       ,ReadTokenSt(Ret));
 // pt20 TEnr.AddChampSupValeur('ABREGE'   ,RechDom('PGMOTIFABSENCE', TEnr.GetValue('PCN_TYPECONGE'), False));
  TEnr.PutValue('PCN_LIBCOMPL1'     ,ReadTokenSt(Ret));
  TEnr.PutValue('PCN_LIBCOMPL2'     ,ReadTokenSt(Ret));
  TEnr.PutValue('PCN_VALIDRESP'     ,ReadTokenSt(Ret));
  TEnr.PutValue('PCN_EXPORTOK'      ,ReadTokenSt(Ret));

  TEnr.AddChampSupValeur('MOTIFVALIDRESP',TEnr.GetValue('PCN_TYPECONGE')+TEnr.GetValue('PCN_VALIDRESP')); { PT19 }
  if Assigned(TobMotifAbs) then
    Begin
      T_MotifAbs:=TobMotifAbs.FindFirst(['PMA_MOTIFABSENCE'],[TEnr.GetString('PCN_TYPECONGE')],False);
      if Assigned(T_MotifAbs) then
      begin
          // Deb pt20   Recherche du type de mvt selon le motif et récupère l'abrégé dans la bonne tablette
        if TEnr.GetString('PCN_TYPECONGE')<> 'PRI' then Tenr.PutValue('PCN_TYPEMVT', T_MotifAbs.GetValue('PMA_TYPEMOTIF'));
        if Tenr.getvalue('PCN_TYPEMVT') = 'PRE' then
        TEnr.AddChampSupValeur('ABREGE'   ,RechDom('PGMOTIFPRESENCE', TEnr.GetValue('PCN_TYPECONGE'), False))
        else
        TEnr.AddChampSupValeur('ABREGE'   ,RechDom('PGMOTIFABSENCE', TEnr.GetValue('PCN_TYPECONGE'), False));
        //  Fin Pt20

        TEnr.AddChampSupValeur('PMA_TYPEABS',T_MotifAbs.GetValue('PMA_TYPEABS'));
      End;
    End;



end;
{ FIN PT6 }

{ DEB PT6 }
function TPlanningAbsPre.AccesUtilisat(Sal : String): Boolean;
begin
  Result := True;
  if ((PlanningAbsPre.TypUtilisat = 'SAL') OR (PlanningAbsPre.TypUtilisat = 'ASS')) and (Sal <> StUtilisat) then { PT8 }
    begin
    PgiBox('Vous ne pouvez consulter les absences de vos collaborateurs.', Caption);
    Result := False;
    Exit;
    end;
{  else   PT11 Mise en commentaire, autorise la saisie modification
    if (PlanningAbsPre.TypUtilisat = 'RESP') then
    begin
      Q := OpenSql('SELECT PSE_RESPONSABS FROM DEPORTSAL WHERE PSE_SALARIE="' + Sal + '"', True);
      if not Q.eof then
        if Q.FindField('PSE_RESPONSABS').AsString <> StUtilisat then
        begin
          PgiBox('Vous ne pouvez consulter les absences des collaborateurs de ' + Rechdom('PGSALARIE', Q.FindField('PSE_RESPONSABS').AsString, False) + ' !',
            'Planning des absences');
          Ferme(Q);
          Result := False;
          Exit;
        end;
      Ferme(Q);
    end;        }

end;
{ FIN PT6 }

{ DEB PT6 }
procedure TPlanningAbsPre.DeleteItemAbsence(Sender: TObject; Item: TOB;
  var Cancel: boolean);
Var Sal : String;
begin
  Cancel := True;
  Sal := Item.GetString('PCN_SALARIE');
  If not AccesUtilisat(Sal) then exit;

  if Item.GetString('PCN_EXPORTOK')<>'-' then
     Begin
     PgiBox('Vous ne pouvez supprimer que les absences en attente d''intégration!',caption);
     Exit;
     End;
  if PlanningAbsPre.TypUtilisat = 'RESP' then
     Begin
     If not IfMotifabsenceSaisissable(Item.GetString('PCN_TYPECONGE'),PlanningAbsPre.TypUtilisat) then
       Begin
       PgiBox('Vous ne pouvez supprimer une absence non saisissable par le responsable!',caption);
       Exit;
       End;
     If (Item.GetString('PCN_VALIDRESP') = 'REF') OR (Item.GetString('PCN_VALIDRESP') = 'NAN') then
       Begin
       PgiBox('Vous ne pouvez supprimer une absence refusée ou annulée!',caption);
       Exit;
       End;
     if (Item.GetString('PCN_VALIDSALARIE') = 'SAL') then
       Begin
       PgiBox('Vous ne pouvez supprimer une absence saisie par le salarié!',caption);
       Exit;
       End;
     End;

   if PgiAsk('Confirmez-vous la suppression de l''absence : '+Item.GetValue('PCN_LIBELLE')+'?',Caption) = MrNo then Exit;

   Cancel:=SuppressionAbsence(Item,True);
end;
{ FIN PT6 }

{ DEB PT6 }
function TPlanningAbsPre.SuppressionAbsence(Item : Tob; AvecRecap : Boolean): Boolean;
var
   StClause : String;
   EJoursPris: double;
begin
   Result := True;
   EJoursPris := Item.GetValue('PCN_JOURS');
   StClause   := '';
   if AvecRecap then
     if (Item.GetString('PCN_TYPECONGE') = 'PRI') then
        StClause := 'SET PRS_PRIATTENTE=PRS_PRIATTENTE-' + StrfPoint(EJoursPris)
     else
       if RechDom('PGMOTIFABSENCERTT',Item.GetString('PCN_TYPECONGE'),False) <> '' then
          StClause := 'SET PRS_RTTATTENTE=PRS_RTTATTENTE-' + StrfPoint(EJoursPris);
   try
      BeginTrans;
      if (StClause<>'') and (AvecRecap) then
         ExecuteSql('UPDATE RECAPSALARIES '+StClause+
                    ' WHERE PRS_SALARIE="' + Item.GetString('PCN_SALARIE') + '"');
      ExecuteSql('DELETE ABSENCESALARIE '+
                 'WHERE PCN_SALARIE="'+Item.GetString('PCN_SALARIE')+'" '+
                 'AND PCN_TYPEMVT="'+Item.GetString('PCN_TYPEMVT')+'" '+
                 'AND PCN_ORDRE='+IntToStr(Item.GetValue('PCN_ORDRE')));
      CommitTrans;
      Result := False;
      if Assigned(TobAbsDelete) then Item.Changeparent(TobAbsDelete, -1);
      ChargeTob_Recapitulatif;
      RefrehRecapPlanning(Item.GetString('PCN_SALARIE'), nil);
//      if Assigned(Item) then H.InvalidateItem(Item);
   except
      Rollback;
      PGIBox('Suppression annulée : Echec lors de la suppression de l''absence.', caption);
   end;
end;
{ FIN PT6 }

{ DEB PT6 }
procedure TPlanningAbsPre.EtireItemAbsence(Sender: TObject; Source,
  Destination : TOB; Option: THPlanningOptionLink; var Cancel: boolean);
begin
    Cancel := True;
    //polLier, polDelete, polCreate, polExtend, polReduce, polMove, polCopy,
    //polLiaisonEffective, polSupprimeLiaison
end;
{ FIN PT6 }

{ DEB PT13 }
procedure TPlanningAbsPre.HelpBtnClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;
{ FIN PT13 }

{ DEB PT15-2 }
procedure TPlanningAbsPre.FormDestroy(Sender: TObject);
begin
  FreeAndNil(TobRessources);
  FreeAndNil(TobItems);
  FreeAndNil(TobEtats);
  FreeAndNil(TobAbsDelete);  { PT6 }
  FreeAndNil(TobMotifAbs);   { PT6 }
  FreeAndNil(Tob_Recapitulatif);
  H.free;
end;
{ FIN PT15-2 }
{ DEB PT17-2 }
procedure TPlanningAbsPre.PgAppliqCritereTotalJourAbs;
var NbJ : double;
    Sal : String;
    i   : integer;
    T1, T2 : Tob;
begin
   If (GbJourAbs > 0) And Assigned(TobRessources) and assigned(Tobitems) then
     Begin
     Sal := '';
     For i:= 0 To Tobitems.detail.count-1 do
        Begin
        T1 := Tobitems.detail[i];
        if Sal = T1.getvalue('PCN_SALARIE') then begin  continue; end;
        Sal := T1.getvalue('PCN_SALARIE');
        NbJ := Tobitems.Somme('PCN_JOURS',['PCN_SALARIE'],[Sal],False);
        if NbJ >= GbJourAbs then
            Begin
            T2 := TobRessources.FindFirst(['PSA_SALARIE'],[Sal],False);
            if Assigned(T2) then T2.free;
            End;
        End;
     End;
end;
{ FIN PT17-2 }
{ DEB PT18 }
procedure TPlanningAbsPre.PgAppliqCritereSoldeAbs;
Var
  Tob_Export, TSld, T1, T2 : Tob;
  Sal :  String;
  i   : Integer;
  NbJ : Double;
begin
   If (GbSoldeAbs > 0) And Assigned(TobRessources) then
     Begin
     if not assigned(Tob_Recapitulatif) then ChargeTob_Recapitulatif;
     Tob_Export := PgGenereTobExport(TobRessources,TobItems,Tob_Recapitulatif,H.TokenFieldColFixed,caption,GbMode,True);
     Sal := '';
     For i:= 0 To Tob_Recapitulatif.detail.count-1 do
        Begin
        T1 := Tob_Recapitulatif.detail[i];
        //if Sal = T1.getvalue('PRS_SALARIE') then begin  continue; end;
        Sal := T1.getvalue('PRS_SALARIE');
        TSld := Tob_Export.FindFirst(['PSA_SALARIE'],[Sal],False);
        NbJ := 0;
        If Assigned(TSld) then
           Begin
           if (GbTypeSoldeAbs = 'SLDCP') AND (isnumeric(TSld.GetValue('CPSOLDE'))) then
              NbJ := TSld.GetValue('CPSOLDE')
           else if (GbTypeSoldeAbs = 'SLDRTT') and (isnumeric(TSld.GetValue('RTTSOLDE'))) then
              NbJ := TSld.GetValue('RTTSOLDE')
           else
              Begin
              if (isnumeric(TSld.GetValue('CPSOLDE'))) then NbJ := TSld.GetValue('CPSOLDE');
              if (isnumeric(TSld.GetValue('RTTSOLDE'))) then  NbJ := NbJ + TSld.GetValue('RTTSOLDE');
              end;
           if NbJ <= GbSoldeAbs then
             Begin
             T2 := TobRessources.FindFirst(['PSA_SALARIE'],[Sal],False);
             if Assigned(T2) then T2.free;
             End;
           End;  
        End;
     End;                                             
FreeAndNil(Tob_Export);
end;
{ FIN PT18 }
{ DEB PT18 }
procedure TPlanningAbsPre.LoadDesignPlanning;
var
   OkOk : Boolean;
begin

  BFirst.enabled := (PlanningAbsPre.StFiltreAbs = ''); { PT9-2 }
  BPrev.enabled  := (PlanningAbsPre.StFiltreAbs = ''); { PT9-2 }
  BNext.enabled  := (PlanningAbsPre.StFiltreAbs = ''); { PT9-2 }
  BLast.enabled  := (PlanningAbsPre.StFiltreAbs = ''); { PT9-2 }

  H.Align := alClient;
  H.TextAlign := taLeftJustify;

    { Les évènements }
  H.OnDblClick := DoubleCLickPlanning;
  H.OnDblClickSpec := DoubleCLickSpecPlanning;
  H.OnAvertirApplication := HAvertirApplication;

  H.Personnalisation := False;
  H.Legende := True;

  if TypUtilisat = 'SAL' then TRecap.enabled := False;

  H.Interval := piJour; // ;piDemiJour; //
  H.CumulInterval := pciSemaine;


  H.MultiLine := False;
  H.JourFerieCalcule         := False;
  H.GestionJoursFeriesActive := True;
  H.ActiveSaturday := True;
  H.ActiveSunday := True;
  H.ColorJoursFeries := StringToColor('12189695');
  H.ColorOfSaturday := StringToColor('12189695');
  H.ColorOfSunday := StringToColor('12189695');
  H.ColorBackground := StringToColor(ColorBackground);
  H.ColorSelection := clTeal; //StringToColor(FontColor);
  H.ActiveLigneDate := True;
  H.ActiveLigneGroupeDate := True;
  if H.IntervalDebut = H.IntervalFin then H.ColSizeData := 250;   { PT10 }

   { Taille des colonnes fixes }
  H.JourneeDebut := StrToTime('07:00');
  H.JourneeFin := StrToTime('19:00');
  H.DateFormat := 'dd mm'; //hh:mm
  H.DebutAM := StrToTime('08:00');
  H.FinAM := StrToTime('12:00');
  H.DebutAP := StrToTime('14:00');
  H.FinAP := StrToTime('18:00');
  { DEB PT6 }
 //pt20  H.DisplayOptionCreation       := ((TypUtilisat<>'SAL') AND (TypUtilisat<>'ASS')); { PT8 }
  H.DisplayOptionCreation       := false; // pt20
  H.DisplayOptionModification   := ((TypUtilisat<>'SAL') AND (TypUtilisat<>'ASS')); { PT8 }
  H.DisplayOptionSuppression    := ((TypUtilisat<>'SAL') AND (TypUtilisat<>'ASS')); { PT8 }
  OkOk := False; //(TypUtilisat<>'SAL');
  H.DisplayOptionEtirer         := OkOk;
  H.DisplayOptionReduire        := OkOk;
  H.MoveHorizontal              := OkOk;
  H.DisplayOptionCopie          := False;
  H.DisplayOptionDeplacement    := False;
  H.DisplayOptionSuppressionLiaison := False;
  H.DisplayOptionLiaison := False;
  H.DisplayOptionLier := False;
  { FIN PT6 }
  if (TypUtilisat='SAL') or (TypUtilisat='ASS') then H.Autorisation := [];  { PT8 }
  H.ColorSelection := clNavy;

  { Les évènements item }
  H.OnInitItem   := InitItemAbsence;
  H.OnCreateItem := CreateItemAbsence;
  H.OnModifyItem := ModifyItemAbsence;
  H.OnDeleteItem := DeleteItemAbsence;
  H.OnLink       := EtireItemAbsence;
  H.MouseAlready := True;



end;
{ FIN PT18 }
{ DEB PT18 }
procedure TPlanningAbsPre.InitialiseOngletRecap;
begin
  AcquisN1.caption := '';
  PrisN1.caption := '';
  RestantN1.caption := '';
  AcquisN.caption := '';
  PrisN.caption := '';
  RestantN.caption := '';
  AcquisRtt.caption := '';
  PrisRTT.caption := '';
  RestantRTT.caption := '';
  NbValidePri.caption := '';
  NbAttentePri.caption := '';
  NbValideRtt.caption := '';
  NbAttenteRtt.caption := '';

end;
{ FIN PT18 }

//pt20
procedure TPlanningAbsPre.CreateItemPresence(Sender: TObject; Item: TOB;
  var Cancel: boolean);
var
  Sal, Lib, DD, DF, Ret : string;
begin
  Cancel := True;

  Sal := Item.GetString('PCN_SALARIE');
  If not AccesUtilisat(Sal) then exit;
  lib := RechDom('PGSALARIE',Sal,False);
  DD  := Item.GetString('PCN_DATEDEBUTABS');
  DF  := Item.GetString('PCN_DATEFINABS');

  Ret := AglLanceFiche('PAY', 'MVTPRESENCE', 'PCN_SALARIE=' + Sal, Sal, Sal + '!' + lib + ';P;;ACTION=CREATION;;' + PlanningAbsPre.TypUtilisat+';'+DD+';'+DF);
  If Ret<>'' then
     Begin
     H.Activate := False;
     InitItemAbsenceModify(Item,Ret);
     ChargeTob_Recapitulatif;
     RefrehRecapPlanning(sal, Item);
    // tobdebug(item);
    // pt20  if Item <> nil then H.InvalidateItem(Item);
     H.Activate := true;
     Cancel := false;
     End;
end;
{ FIN PT6 }
procedure TPlanningAbsPre.Tcreerabsenceclick(Sender: TObject);
var
T, tobRessource : Tob;
cancel : boolean;
 Zone : TZoneSelected;
 datedeb, datefin : Tdatetime;
 salarie : string;
begin
  Zone := H.ZoneSelected;
  tobRessource := H.GetRessourceOfRow(H.Row);
  salarie := tobressource.getvalue('PSA_SALARIE');
  datedeb := zone.DebutSelection - NbrColonnesEntete + 1;
  datefin := zone.FinSelection - NbrColonnesEntete +1 ;

  inititemabsence(sender, T, cancel);
  T.putvalue('PCN_SALARIE', salarie);
  T.putvalue('PCN_DATEDEBUTABS', datedeb);
  T.putvalue('PCN_DATEFINABS', datefin);

  Createitemabsence(Sender, T, cancel);
end;

procedure TPlanningAbsPre.Tcreerpresenceclick(Sender: TObject);
var
T, tobRessource : Tob;
cancel : boolean;
 Zone : TZoneSelected;
 datedeb, datefin : Tdatetime;
 salarie : string;
begin
  Zone := H.ZoneSelected;
  tobRessource := H.GetRessourceOfRow(H.Row);
  salarie := tobressource.getvalue('PSA_SALARIE');
  datedeb := zone.DebutSelection - NbrColonnesEntete + 1;
  datefin := zone.FinSelection - NbrColonnesEntete +1 ;

  inititemabsence(sender, T, cancel);
  T.putvalue('PCN_SALARIE', salarie);
  T.putvalue('PCN_DATEDEBUTABS', datedeb);
  T.putvalue('PCN_DATEFINABS', datefin);

  Createitempresence(Sender, T, cancel);
end;

end.

