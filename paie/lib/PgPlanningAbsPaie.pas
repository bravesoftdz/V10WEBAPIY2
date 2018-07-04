{***********UNITE*************************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 07/11/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{
PT1 07/11/2005 SB V_65 FQ 12625 Epurement des salariés chargés en cas de filtre absence
PT2 23/01/2006 SB V_65 Simplification de la requête SQL
PT3-1 28/02/2006 SB V_65 Refonte traitement pour mise en place gestion des couleurs, gestion planning
PT3-2 28/02/2006 SB V_65 FQ 12926 Optimisation + Refonte tri du planning par nom resp., puis nom sal.
PT4   19/06/2006 SB V_65 FQ 13231 Retrait des mvt absences annulées
PT5   07/07/2006 SB V_65 FQ 13346 Retrait de la condition PSA_CONGESPAYES
PT6   21/07/2006 SB V_65 FQ 13394 Correction anomalie en fermeture
PT7  07/09/2006 SB V_70 FQ 13394 Reprise PT6
}
unit PgPlanningAbsPaie;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, HTB97, ExtCtrls,Hpanel,UiUtil,HPlanning,PgPlanningOutils,Utob,
  {$IFDEF EAGLCLIENT}
  MaineAgl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ENDIF}
  HeureUtil,Hmsgbox,Hent1;

type
  TPlanningAbsPaie = class(TForm)
    PPanel: TPanel;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BImprimer: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    TABSENCE: TToolbarButton97;
    Bperso: TToolbarButton97;
    TExcel: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);  { PT6 }
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BtnDetailClick(Sender: TObject);
    procedure TExcelClick(Sender: TObject);
    procedure BpersoClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure TABSENCEClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
  private
    { Déclarations privées }
    DateDebAbs, DateFinAbs: TDateTime;
    TypUtilisat, StWhTobEtats, StWhTobItems, StWhTobRessource,StFiltreAbs : string;
    Rupture: TNiveauRupture;
    THPlanningAbs : THPlanning;
    FontColor, FontName, FontStyle, ColorBackground: string;
    FontSize : Integer;
    ChampsOrg: array[1..4] of string;
    Champslibre: array[1..4] of string;
    TobRessources, GblTobItems, GblTobEtats, TobAbsDelete, TobMotifAbs: Tob;  { PT7 }
    Function ChargeTobPlanningAbsPaie : Boolean;
    function LoadAbsencePeriode(DD, DF: TDateTime; OnLoad: Boolean): Boolean;
    //DoubleClick sur ressource
    procedure DoubleCLickSpecPlanning(ACol, ARow: INTEGER; TypeCellule: TPlanningTypeCellule; T: TOB = nil);
    //DoubleClick sur item
    procedure DoubleCLickPlanning(Sender: TObject);
    procedure HAvertirApplication(Sender: TObject; FromItem, ToItem: TOB; Actions: THPlanningAction);
    Function  LanceFicheAbsenceModify(TEnr : Tob ) : Boolean;

    procedure InitItemAbsence(Sender: TObject; var Item: TOB; var Cancel: boolean);
    procedure CreateItemAbsence(Sender: TObject; Item: TOB; var Cancel: boolean);
    procedure ModifyItemAbsence(Sender: TObject; Item: TOB; var Cancel: boolean);
    Procedure InitItemAbsenceModify (TEnr : Tob ; Var Ret : String);
    Procedure DeleteItemAbsence(Sender: TObject; Item: TOB; var Cancel: boolean);
    function  SuppressionAbsence(Item : Tob; AvecRecap : Boolean): Boolean;
  public
    { Déclarations publiques }

  end;

procedure PGPlanningAbsencePaie(DateDebAbs, DateFinAbs: TDateTime; TypUtilisat, StWhTobEtats, StWhTobItems, StWhTobRessource : string; NiveauRupt: TNiveauRupture;
  FiltreAbs : string = '');

(** PT7 Mise en commentaire
var
  PlanningAbsPaie: TPlanningAbsPaie;
**)

implementation

uses EntPaie
   {$IFDEF ABSETEMPS}
   ,SaisiePersoPlanning
   {$ENDIF}
   , DB;

{$R *.DFM}

procedure PGPlanningAbsencePaie(DateDebAbs, DateFinAbs: TDateTime; TypUtilisat, StWhTobEtats, StWhTobItems, StWhTobRessource : string; NiveauRupt: TNiveauRupture;
  FiltreAbs : string = '');
Var
  PP: THPanel;
  PlanningAbsPaie: TPlanningAbsPaie; { PT7 }
begin    { DEB PT10 }
  // PlanningAbsPaie := nil; { PT7 }
  try
  PP := FindInsidePanel();
  if Assigned(PP) then
  begin
    PlanningAbsPaie := TPlanningAbsPaie.Create(Application);
    InitInside(PlanningAbsPaie, PP);
    PlanningAbsPaie.DateDebAbs := DateDebAbs;
    PlanningAbsPaie.DateFinAbs := DateFinAbs;
    PlanningAbsPaie.TypUtilisat := TypUtilisat;
    PlanningAbsPaie.StWhTobEtats := StWhTobEtats;
    PlanningAbsPaie.StWhTobItems := StWhTobItems;
    PlanningAbsPaie.StWhTobRessource := StWhTobRessource;
    PlanningAbsPaie.StFiltreAbs := FiltreAbs;
    PlanningAbsPaie.Rupture := NiveauRupt;
    PlanningAbsPaie.Show;
  end
  else  { FIN PT10 }
  begin
    PlanningAbsPaie := TPlanningAbsPaie.Create(Application);
  try
    PlanningAbsPaie.DateDebAbs := DateDebAbs;
    PlanningAbsPaie.DateFinAbs := DateFinAbs;
    PlanningAbsPaie.TypUtilisat := TypUtilisat;
    PlanningAbsPaie.StWhTobEtats := StWhTobEtats;
    PlanningAbsPaie.StWhTobItems := StWhTobItems;
    PlanningAbsPaie.StWhTobRessource := StWhTobRessource;
    PlanningAbsPaie.StFiltreAbs := FiltreAbs; { PT9-2 }
    PlanningAbsPaie.Rupture := NiveauRupt;
    PlanningAbsPaie.ShowModal;
      finally
        if Assigned(PlanningAbsPaie) then  FreeAndNil(PlanningAbsPaie); { PT7 }
      end;
    end;
  except  { PT7 }
    on E: Exception do
      ShowMessage('Plantage : ' + E.Message);
  end;
end;

procedure TPlanningAbsPaie.FormShow(Sender: TObject);
var
   OkOk : Boolean;
begin
  //***OnShow := True;
  FontColor := 'ClBlack';
  FontStyle := 'G';
  FontName := 'Times New Roman';
  FontSize := 10;
  ColorBackground := 'clWhite';

  ChampsOrg[1] := VH_Paie.PGLibelleOrgStat1;
  ChampsOrg[2] := VH_Paie.PGLibelleOrgStat2;
  ChampsOrg[3] := VH_Paie.PGLibelleOrgStat3;
  ChampsOrg[4] := VH_Paie.PGLibelleOrgStat4;
  Champslibre[1] := VH_Paie.PgLibCombo1;
  Champslibre[2] := VH_Paie.PgLibCombo2;
  Champslibre[3] := VH_Paie.PgLibCombo3;
  Champslibre[4] := VH_Paie.PgLibCombo4;
  BFirst.enabled := (StFiltreAbs = '');
  BPrev.enabled  := (StFiltreAbs = '');
  BNext.enabled  := (StFiltreAbs = '');
  BLast.enabled  := (StFiltreAbs = '');

  (** PT7 
  if THPlanningAbs <> nil then THPlanningAbs.Free;
  **)
  THPlanningAbs := THplanning.Create(Self);  { PT7 }
  THPlanningAbs.Parent := PPanel;
  THPlanningAbs.Align := alClient;
  THPlanningAbs.activate := False;
  { Les évènements }
  THPlanningAbs.OnDblClick := DoubleCLickPlanning;
  THPlanningAbs.OnDblClickSpec := DoubleCLickSpecPlanning;
  THPlanningAbs.OnAvertirApplication := HAvertirApplication;
  THPlanningAbs.Personnalisation := False;
  THPlanningAbs.Legende := True;
  { Ressources }
  THPlanningAbs.ResChampID := 'PSA_SALARIE';
  { Etats }
  THPlanningAbs.EtatChampCode := 'MOTIFETAT';       { PT3-1 }
  THPlanningAbs.EtatChampLibelle := 'LIBMOTIFETAT'; { PT3-1 }
  THPlanningAbs.EtatChampBackGroundColor := 'BACKGROUNDCOLOR';
  THPlanningAbs.EtatChampFontColor := 'FONDCOLOR';
  THPlanningAbs.EtatChampFontStyle := 'FONTSTYLE';
  THPlanningAbs.EtatChampFontSize := 'FONTSIZE';
  THPlanningAbs.EtatChampFontName := 'FONTNAME';
  THPlanningAbs.TextAlign := taLeftJustify;
  { Items }
  THPlanningAbs.ChampLineID := 'PCN_SALARIE';
  THPlanningAbs.ChampdateDebut := 'PCN_DATEDEBUTABS';
  THPlanningAbs.ChampDateFin := 'PCN_DATEFINABS';
  THPlanningAbs.ChampEtat := 'MOTIFVALIDRESP';  { PT3-1 }
  THPlanningAbs.ChampLibelle := 'ABREGE';
  THPlanningAbs.ChampHint := 'PCN_LIBELLE';
  {Personalisation}
  THPlanningAbs.Interval := piJour; // ;piDemiJour; //
  THPlanningAbs.CumulInterval := pciSemaine;
  THPlanningAbs.IntervalDebut := DateDebAbs;
  THPlanningAbs.IntervalFin := DateFinAbs;
  THPlanningAbs.MultiLine := False;
  THPlanningAbs.GestionJoursFeriesActive := True;
  THPlanningAbs.ActiveSaturday := True;
  THPlanningAbs.ActiveSunday := True;
  THPlanningAbs.ColorJoursFeries := StringToColor('12189695');
  THPlanningAbs.ColorOfSaturday := StringToColor('12189695');
  THPlanningAbs.ColorOfSunday := StringToColor('12189695');
  THPlanningAbs.ColorBackground := StringToColor(ColorBackground);
  THPlanningAbs.ColorSelection := clTeal; //StringToColor(FontColor);
  THPlanningAbs.ActiveLigneDate := True;
  THPlanningAbs.ActiveLigneGroupeDate := True;
  if THPlanningAbs.IntervalDebut = THPlanningAbs.IntervalFin then THPlanningAbs.ColSizeData := 250;   { PT10 }
  { Alignement des colonnes fixes }

  { Liste des champs d'entete }
  THPlanningAbs.TokenFieldColEntete := '';
  { Liste des champs d'entete }
  THPlanningAbs.TokenFieldColFixed := ListeChampRupt(';',Rupture) + 'PSA_SALARIE;NOM_SAL';
  MiseEnFormeColonneFixeRupt(THPlanningAbs,Rupture);

  { Taille des colonnes fixes }
  THPlanningAbs.JourneeDebut := StrToTime('07:00');
  THPlanningAbs.JourneeFin := StrToTime('19:00');
  THPlanningAbs.DateFormat := 'dd mm'; //hh:mm
  THPlanningAbs.DebutAM := StrToTime('08:00');
  THPlanningAbs.FinAM := StrToTime('12:00');
  THPlanningAbs.DebutAP := StrToTime('14:00');
  THPlanningAbs.FinAP := StrToTime('18:00');

  THPlanningAbs.DisplayOptionCreation       := ((TypUtilisat<>'SAL') AND (TypUtilisat<>'ASS'));
  THPlanningAbs.DisplayOptionModification   := ((TypUtilisat<>'SAL') AND (TypUtilisat<>'ASS'));
  THPlanningAbs.DisplayOptionSuppression    := ((TypUtilisat<>'SAL') AND (TypUtilisat<>'ASS'));
  OkOk := False; //(TypUtilisat<>'SAL');
  THPlanningAbs.DisplayOptionEtirer         := OkOk;
  THPlanningAbs.DisplayOptionReduire        := OkOk;
  THPlanningAbs.MoveHorizontal              := OkOk;
  THPlanningAbs.DisplayOptionCopie          := False;
  THPlanningAbs.DisplayOptionDeplacement    := False;
  THPlanningAbs.DisplayOptionSuppressionLiaison := False;
  THPlanningAbs.DisplayOptionLiaison := False;
  THPlanningAbs.DisplayOptionLier := False;

  if (TypUtilisat='SAL') or (TypUtilisat='ASS') then THPlanningAbs.Autorisation := [];
  THPlanningAbs.ColorSelection := clNavy;

  { Les évènements item }
  THPlanningAbs.OnInitItem   := InitItemAbsence;
  THPlanningAbs.OnCreateItem := CreateItemAbsence;
  THPlanningAbs.OnModifyItem := ModifyItemAbsence;
  THPlanningAbs.OnDeleteItem := DeleteItemAbsence;
  THPlanningAbs.MouseAlready := True;


  {Création des trois TOB }
  if not ChargeTobPlanningAbsPaie then PostMessage(Handle, SBWM_APPSTARTUP, 0, 0)
  else
  begin
    LoadConfigPlanning(THPlanningAbs);
    Caption := 'Planning des absences du ' + DateToStr(THPlanningAbs.IntervalDebut) + ' au ' + DateToStr(THPlanningAbs.IntervalFin);
    THPlanningAbs.activate := True;
  end;
end;

function TPlanningAbsPaie.ChargeTobPlanningAbsPaie: Boolean;
Var
Tob_Sal,TobCalend, T, T1,T2,T3,Tob_Motifs, Tob_ValidResp : Tob;  { PT3-1 }
Q : TQuery;
St,StSal,StRes,StRupture,DateCalculer,StDate : String;
i,j,TJ : integer;
NomChamp: array[1..3] of string;
ValChamp: array[1..3] of variant;
begin
  result := False;
  if GblTobItems <> nil then   GblTobItems.Free;  { PT7 }
  if GblTobEtats <> nil then   GblTobEtats.Free;  { PT7 }
  if TobRessources <> nil then TobRessources.Free;

  TobRessources := TOB.Create('Les ressources', nil, -1);
  Tob_Sal := TOB.Create('Les salariés', nil, -1);
  GblTobItems := TOB.Create('Les items', nil, -1);   { PT7 }
  GblTobEtats := TOB.Create('Les états', nil, -1);  { PT7 }
  Tob_Motifs := TOB.Create('Les motifs', nil, -1);           { PT3-1 }
  Tob_ValidResp := TOB.Create('etat validation', nil, -1);   { PT3-1 }

  Q := nil;
  { Chargement des ressources }
  StRes := StWhTobRessource;  { PT7 }
  if Length(StRes) < 10 then
    StRes := ' WHERE PSA_LIBELLE<>"" '; { PT5 PSA_CONGESPAYES="X" }
  // else StRes := StRes + 'AND PSA_CONGESPAYES="X" '; PT5

  StRupture := ListeChampRupt(',',Rupture);
  { DEB PT3-2 }
  if DateDebAbs = DateFinAbs then  { PT7 }
    StDate := 'AND ABSSAL.PCN_DATEDEBUTABS<="' + UsDateTime(DateDebAbs) + '" ' +
      'AND ABSSAL.PCN_DATEFINABS>="' + UsDateTime(DateFinAbs) + '" '
  else
    StDate := 'AND ((ABSSAL.PCN_DATEDEBUTABS>="' + UsDateTime(DateDebAbs) + '" ' +
      'AND ABSSAL.PCN_DATEDEBUTABS<="' + UsDateTime(DateFinAbs) + '") ' +
      'OR (ABSSAL.PCN_DATEFINABS>="' + UsDateTime(DateDebAbs) + '" ' +
      'AND ABSSAL.PCN_DATEFINABS<="' + UsDateTime(DateFinAbs) + '") ' +
      'OR (ABSSAL.PCN_DATEDEBUTABS<="' + UsDateTime(DateDebAbs) + '" ' +
      'AND ABSSAL.PCN_DATEFINABS>="' + UsDateTime(DateFinAbs) + '")) ';

  if StFiltreAbs='ABS' then
    StRes := StRes +' AND PSA_SALARIE IN (SELECT ABSSAL.PCN_SALARIE '+
    'FROM ABSENCESALARIE ABSSAL,MOTIFABSENCE '+  { PT2 }
    'WHERE ABSSAL.PCN_TYPECONGE=PMA_MOTIFABSENCE  '+
    'AND ##PMA_PREDEFINI## (ABSSAL.PCN_TYPEMVT="ABS" OR ABSSAL.PCN_TYPECONGE="PRI") '+
    'AND ABSSAL.PCN_ETATPOSTPAIE <> "NAN" '+ { PT4 }
    'AND ABSSAL.PCN_MVTDUPLIQUE<>"X" AND PMA_EDITPLANPAIE="X" '+ StDate +' ) '
    else
    if StFiltreAbs='PRE' then
       StRes := StRes + ' AND PSA_SALARIE NOT IN (SELECT ABSSAL.PCN_SALARIE '+
      'FROM ABSENCESALARIE ABSSAL,MOTIFABSENCE '+  { PT2 }
      'WHERE ABSSAL.PCN_TYPECONGE=PMA_MOTIFABSENCE  '+
      'AND ##PMA_PREDEFINI## (ABSSAL.PCN_TYPEMVT="ABS" OR ABSSAL.PCN_TYPECONGE="PRI") '+
      'AND ABSSAL.PCN_ETATPOSTPAIE <> "NAN" '+  { PT4 }
      'AND ABSSAL.PCN_MVTDUPLIQUE<>"X" AND PMA_EDITPLANPAIE="X" '+ StDate +') ' ;

  StRes := 'SELECT ' + StRupture +' PSA_SALARIE,SUBSTRING(PSA_LIBELLE,1,36)||" "||PSA_PRENOM NOM_SAL ' +
    'FROM SALARIES ' + StRes ;
  //  'ORDER BY ' + StRupture + 'NOM_SAL,PSA_SALARIE';
  { FIN PT3-2 }

  Q := OpenSql(StRes, True);
  Tob_Sal.LoadDetailDB('SAL_ARIES', '', '', Q, False);
  Ferme(Q);

{ Chargement des états }
  { DEB PT3-1 }
  St := 'SELECT * FROM MOTIFABSENCE WHERE ##PMA_PREDEFINI## PMA_EDITPLANPAIE="X" ';
  Q := OpenSql(St, True);
  if not Q.Eof then Tob_Motifs.LoadDetailDB('Les motifs', '', '', Q, False);
  Ferme(Q);

  St := 'SELECT CO_TYPE,CO_CODE,CO_LIBELLE FROM COMMUN WHERE CO_TYPE="PET"';
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
      T := TOB.Create('Les états', GblTobEtats, -1);  { PT7 }
       T.AddChampSupValeur('MOTIFETAT',T1.GetValue('PMA_MOTIFABSENCE')+T2.GetValue('CO_CODE'));
       T.AddChampSupValeur('LIBMOTIFETAT',T2.GetValue('CO_LIBELLE'));
       if T2.GetValue('CO_CODE')='...' then T.AddChampSupValeur('BACKGROUNDCOLOR',T1.GetValue('PMA_PGCOLORACTIF'))
       else T.AddChampSupValeur('BACKGROUNDCOLOR',T1.GetValue('PMA_PGCOLORPAY'));
       T.AddChampSupValeur('FONDCOLOR', FontColor);
       T.AddChampSupValeur('FONTNAME', FontName);
       T.AddChampSupValeur('FONTSTYLE', FontStyle);
       T.AddChampSupValeur('FONTSIZE', FontSize);
       end;
     end;
  FreeAndNil(Tob_ValidResp);
  FreeAndNil(Tob_Motifs);
  { FIN PT3-1 } 

  { Chargement des items }
  LoadAbsencePeriode(DateDebAbs, DateFinAbs, True); { PT3-2 }  { PT7 }

  if (StFiltreAbs = 'ABS') then
    EpureTobRessource(Tob_Sal, GblTobItems); { PT1 }  { PT7 }

  MiseEnFormeTobRupt('RES', Tob_Sal,TobRessources,Rupture);          { PT1 Dpct du code }


  //Chargement du récapitulatif
//  ChargeTob_Recapitulatif;


  if (THPlanningAbs.Interval = PiDemiJour) then
  begin
    { Chargement des salaries}
    StSal := StWhTobRessource;  { PT7 }
    Q := OpenSql('SELECT PSA_SALARIE,PSA_LIBELLE,PSA_PRENOM,psa_standcalend,PSA_CALENDRIER,' +
      'etb_standcalend FROM SALARIES LEFT join etabcompl on psa_etablissement=etb_etablissement '+ StSal, True);
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
      T2 := GblTobitems.FindFirst(['PCN_SALARIE'], [T1.getValue('PSA_SALARIE')], False);  { PT7 }
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
            DateCalculer := DateToStr(DateDebAbs) + ' ' + FloatToStrTime(8, 'HH:MM:SS')
          else
            DateCalculer := DateToStr(DateDebAbs) + ' ' + FloatToStrTime(14, 'HH:MM:SS');
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
            DateCalculer := DateToStr(DateFinAbs) + ' ' + FloatToStrTime(12, 'HH:MM:SS')
          else
            DateCalculer := DateToStr(DateFinAbs) + ' ' + FloatToStrTime(18, 'HH:MM:SS');
          T2.PutValue('PCN_DATEFINABS', StrToDateTime(DateCalculer));
        end;
        T2 := GblTobItems.FindNext(['PCN_SALARIE'], [T1.getValue('PSA_SALARIE')], False);  { PT7 }
      end;
      T1 := Tob_Sal.FindNext([''], [''], False);
    end;
    TobCalend.free;
  end;


  THPlanningAbs.TobRes := TobRessources;
  THPlanningAbs.TobEtats := GblTobEtats;  { PT7 }
  THPlanningAbs.TobItems := GblTobItems;  { PT7 }



  if Assigned(TobAbsDelete) then FreeAndNil(TobAbsDelete);
  TobAbsDelete := TOB.Create('ABSENCESALARIE', nil, -1);


  if Tob_Sal <> nil then Tob_Sal.free;
  result := True;
end;





function TPlanningAbsPaie.LoadAbsencePeriode(DD, DF: TDateTime;
  OnLoad: Boolean): Boolean;
var
  StWhere, St, StDate : string;
  Q: TQuery;
  TobAbsence, T1, T2, Texist: Tob;
  i: integer;
begin
  { Chargement des items }
//  result := False;
  FreeAndNil(TobAbsence);
  if StFiltreAbs='PRE' then exit;   { PT3-2 }

  StWhere := StWhTobItems;
  if Pos('WHERE', StWhere) > 0 then StWhere := 'AND ' + Copy(StWhere, Pos('WHERE', StWhere) + 5, Length(StWhere));
  { DEB PT3-2 }
  if DD = DF then StDate := 'AND PCN_DATEDEBUTABS<="' + UsDateTime(DD) + '" AND PCN_DATEFINABS>="' + UsDateTime(DF) + '" '
  else StDate := 'AND ((PCN_DATEDEBUTABS>="' + UsDateTime(DD) + '" AND PCN_DATEDEBUTABS<="' + UsDateTime(DF) + '") ' + //PT1 Ajout condition
    'OR (PCN_DATEFINABS>="' + UsDateTime(DD) + '" AND PCN_DATEFINABS<="' + UsDateTime(DF) + '") ' +
    'OR (PCN_DATEDEBUTABS<="' + UsDateTime(DD) + '" AND PCN_DATEFINABS>="' + UsDateTime(DF) + '")) '; //PT4 { PT7 }

  St := 'SELECT PCN_ORDRE,PCN_SALARIE,PCN_TYPEMVT,PCN_TYPECONGE,PCN_ETABLISSEMENT,' +
    'PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_SENSABS,PCN_JOURS,PCN_HEURES,PCN_LIBELLE,' +
    'PCN_APAYES,PCN_CODETAPE,PCN_DATEVALIDITE,PCN_ETATPOSTPAIE,PCN_EXPORTOK,' +
    'PCN_VALIDRESP,PCN_VALIDSALARIE,PCN_DEBUTDJ,PCN_FINDJ,PCN_MVTDUPLIQUE,' +
    'PMA_CALENDCIVIL,PMA_MOTIFEAGL,PMA_TYPEABS ' +
    'FROM ABSENCESALARIE ' +
    ' LEFT JOIN MOTIFABSENCE ON ##PMA_PREDEFINI## PCN_TYPECONGE=PMA_MOTIFABSENCE  ' +
    'WHERE (PCN_TYPEMVT="ABS" OR PCN_TYPECONGE="PRI") AND PCN_MVTDUPLIQUE<>"X" '+
    'AND PCN_ETATPOSTPAIE <> "NAN" AND PMA_EDITPLANPAIE="X" ' +  { PT4 }
    StDate + StWhere ; //PT4 { PT7 }
  { FIN PT3-2 }  
  Q := Opensql(St, True);
  if not Q.eof then
  begin
    TobAbsence := Tob.Create('ABSENCE_SALARIE', nil, -1);
    TobAbsence.LoadDetailDB('ABSENCE_SALARIE', '', '', Q, False);
  end
  else
  begin
    PGIInfo('Aucun mouvement d''absence à éditer sur le planning du ' + DateToStr(DD) + ' au ' + DateToStr(DF) + '!', Caption);
  end;
  Ferme(Q);
  if Assigned(TobAbsence) then
         for i := 0 to TobAbsence.detail.count - 1 do
         begin
         T1 := TobAbsence.detail[i];
         T1.AddChampSup('ABREGE', False);
         T1.PutValue('ABREGE', RechDom('PGMOTIFABSENCE', T1.GetValue('PCN_TYPECONGE'), False));
         T1.AddChampSupValeur('MOTIFVALIDRESP',T1.GetValue('PCN_TYPECONGE')+T1.GetValue('PCN_CODETAPE')); { PT3-1 }
      if GblTobItems <> nil then { PT7 }
        Texist := GblTobItems.FindFirst(['PCN_SALARIE', 'PCN_TYPEMVT', 'PCN_ORDRE'], [T1.GetValue('PCN_SALARIE'), T1.GetValue('PCN_TYPEMVT'),
          T1.GetValue('PCN_ORDRE')], False)
      else
        Texist := nil;
         if Texist = nil then
           begin
        T2 := Tob.Create('ABSENCE_SALARIE', GblTobItems, -1);
           T2.Dupliquer(T1, True, True, True);
           end;
         end;
  THPlanningAbs.TobItems := GblTobItems;
  if Assigned(TobAbsence) then TobAbsence.free;

  TobMotifAbs := Tob.Create('MOTIF_ABSENCE',nil,-1);
  TobMotifAbs.LoadDetailDB('MOTIFABSENCE','','',nil,False);
  result := True;
end;


procedure TPlanningAbsPaie.FormClose(Sender: TObject;  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TPlanningAbsPaie.BFermeClick(Sender: TObject);
begin
Close;
end;

{ DEB PT6 }
procedure TPlanningAbsPaie.FormDestroy(Sender: TObject);
begin
  if Assigned(THPlanningAbs) then
    FreeAndNil(THPlanningAbs);
  FreeAndNil(TobRessources);
  FreeAndNil(GblTobItems);
  FreeAndNil(GblTobEtats);
  FreeAndNil(TobAbsDelete);
  FreeAndNil(TobMotifAbs);
end;
{ FIN PT6 }


procedure TPlanningAbsPaie.BFirstClick(Sender: TObject);
var
  YY, MM, JJ: WORD;
  DD: TDateTime;
begin
  if THPlanningAbs = nil then exit;
  DD := THPlanningAbs.IntervalDebut;
  DecodeDate(THPlanningAbs.IntervalDebut, YY, MM, JJ);
  if (MM < 6) or ((MM = 6) and (JJ = 1)) then
  begin
    YY := YY - 1;
    DD := DD - 1;
  end;
  THPlanningAbs.Activate := False;
  DateDebAbs := EncodeDate(YY, 6, 1);
  LoadAbsencePeriode(DateDebAbs, DD, False);
  THPlanningAbs.IntervalDebut := DateDebAbs;
  THPlanningAbs.DateOfStart := DateDebAbs;
  Caption := 'Planning des absences du ' + DateToStr(THPlanningAbs.IntervalDebut) + ' au ' + DateToStr(THPlanningAbs.IntervalFin);
  THPlanningAbs.Activate := True;
end;



procedure TPlanningAbsPaie.BPrevClick(Sender: TObject);
var
  DD: TdateTime;
begin
  if THPlanningAbs = nil then exit;
  THPlanningAbs.Activate := False;
  DD := THPlanningAbs.IntervalDebut - 1;
  DateDebAbs := DebutdeMois(PlusMois(THPlanningAbs.IntervalDebut, -1));
  LoadAbsencePeriode(DateDebAbs, DD, False);
  THPlanningAbs.IntervalDebut := DateDebAbs;
  THPlanningAbs.DateOfStart := DateDebAbs;
  Caption := 'Planning des absences du ' + DateToStr(THPlanningAbs.IntervalDebut) + ' au ' + DateToStr(THPlanningAbs.IntervalFin);
  THPlanningAbs.Activate := True;
end;

procedure TPlanningAbsPaie.BNextClick(Sender: TObject);
var
  DF: TDateTime;
begin
  if THPlanningAbs = nil then exit;
  THPlanningAbs.Activate := False;
  DF := THPlanningAbs.IntervalFin + 1;
  DateFinAbs := FindeMois(PlusMois(THPlanningAbs.IntervalFin, 1));
  LoadAbsencePeriode(DF, DateFinAbs, False);
  THPlanningAbs.IntervalFin := DateFinAbs;
  THPlanningAbs.DateOfStart := DF;
  Caption := 'Planning des absences du ' + DateToStr(THPlanningAbs.IntervalDebut) + ' au ' + DateToStr(THPlanningAbs.IntervalFin);
  THPlanningAbs.Activate := True;
end;

procedure TPlanningAbsPaie.BLastClick(Sender: TObject);
var
  YY, MM, JJ: WORD;
  DF: TDateTime;
begin
  if THPlanningAbs = nil then exit;
  DF := THPlanningAbs.IntervalFin + 1;
  DecodeDate(THPlanningAbs.IntervalFin, YY, MM, JJ);
  if (MM > 5) or ((MM = 5) and (JJ = 31)) then YY := YY + 1;
  THPlanningAbs.Activate := False;
  DateFinAbs := EncodeDate(YY, 5, 31);
  LoadAbsencePeriode(DF, DateFinAbs, False);
  THPlanningAbs.IntervalFin := DateFinAbs;
  THPlanningAbs.DateOfStart := DF;
  Caption := 'Planning des absences du ' + DateToStr(THPlanningAbs.IntervalDebut) + ' au ' + DateToStr(THPlanningAbs.IntervalFin);
  THPlanningAbs.Activate := True;
end;

procedure TPlanningAbsPaie.BtnDetailClick(Sender: TObject);
begin
  //PRecap.visible := BtnDetail.Down;
end;

procedure TPlanningAbsPaie.TExcelClick(Sender: TObject);
begin
  with TSaveDialog.Create(Self) do
  begin
    Filter := 'Microsoft Excel (*.xls)|*.XLS';
    if Execute then THPlanningAbs.ExportToExcel(True, FileName);
    Free;
  end;
end;

procedure TPlanningAbsPaie.BpersoClick(Sender: TObject);
{$IFDEF ABSETEMPS}
var
  eColorStart, eColorEnd: TColor;
  {$ENDIF}
begin
  {$IFDEF ABSETEMPS}
  eCongePersoPlanning('ConfigPlanningConges', THPlanningAbs, eColorStart, eColorEnd);
  {$ENDIF}  
  THPlanningAbs.Refresh;
end;

procedure TPlanningAbsPaie.BImprimerClick(Sender: TObject);
var
  S: string;
begin
  THPlanningAbs.TypeEtat := 'E';
  THPlanningAbs.NatureEtat := 'PCG';
  THPlanningAbs.CodeEtat := 'PL1';
  S := 'DATEDEB=' + FormatDateTime('dd/mm/yyyy', THPlanningAbs.IntervalDebut) + '^E' + IntToStr(Byte(otDate));
  S := S + '`DATEFIN=' + FormatDateTime('dd/mm/yyyy', THPlanningAbs.IntervalFin) + '^E' + IntToStr(Byte(otDate));
  THPlanningAbs.CriteresToPrint := S;
  THPlanningAbs.Print;
end;

procedure TPlanningAbsPaie.DoubleCLickPlanning(Sender: TObject);
 Var T: TOB;
begin
  T := THPlanningAbs.GetCurItem;
  if Assigned(T) then LanceFicheAbsenceModify(T);
end;

procedure TPlanningAbsPaie.DoubleCLickSpecPlanning(ACol, ARow: INTEGER;
  TypeCellule: TPlanningTypeCellule; T: TOB);
begin
   case TypeCellule of
    ptcRessource: if T <> nil then 
         AglLanceFiche ('PAY','SALARIE_CP', '', T.GetValue('PSA_SALARIE') ,'');
  end;

end;

procedure TPlanningAbsPaie.HAvertirApplication(Sender: TObject; FromItem,
  ToItem: TOB; Actions: THPlanningAction);
var
  Sal: string;
begin
  if Actions = paclickLeft then
  begin
    if FromItem <> nil then
    begin
      Sal := FromItem.GetValue('PCN_SALARIE');
//***      RefrehRecapPlanning(sal, FromItem);
    end;
  end;
end;

function TPlanningAbsPaie.LanceFicheAbsenceModify(TEnr: Tob): Boolean;
Var
  Sal, TypeMvt, Ordre, Etab, ret, DD, DF: string;
begin
    Result := True ;
    if not Assigned(TEnr) then exit;
    Sal := TEnr.GetValue('PCN_SALARIE');
    TypeMvt        := TEnr.GetValue('PCN_TYPEMVT');
    Ordre          := IntToStr(TEnr.GetValue('PCN_ORDRE'));
    Etab           := TEnr.GetValue('PCN_ETABLISSEMENT');
    DD             := TEnr.GetString('PCN_DATEDEBUTABS');
    DF             := TEnr.GetString('PCN_DATEFINABS');
    Ret := AglLanceFiche ('PAY','MVTABSENCE','',TypeMvt+';'+Sal+';'+Ordre,Sal +';A;'+Etab+';ACTION=MODIFICATION;;;'+DD+';'+DF);

   if (ret <> '') then
    begin
      Result := False;
      if Pos('SUPPR',Ret) > 0 then
        Begin
        THPlanningAbs.DeleteItem(TEnr, False);
        //***ChargeTob_Recapitulatif;
        //RefrehRecapPlanning(TEnr.GetString('PCN_SALARIE'), nil);
        End
      else
      Begin
        InitItemAbsenceModify(TEnr,Ret);
        //***ChargeTob_Recapitulatif;
        //RefrehRecapPlanning(sal, TEnr);
        if TEnr <> nil then THPlanningAbs.InvalidateItem(TEnr);
      End;
    end;
end;

procedure TPlanningAbsPaie.CreateItemAbsence(Sender: TObject; Item: TOB;
  var Cancel: boolean);
var
  Sal, Etab, Lib, DD, DF, Ret : string;
begin
  Cancel := True;
  Sal := Item.GetString('PCN_SALARIE');
  Etab:= RechDom('PGSALARIEETAB',Sal,False);
  lib := RechDom('PGSALARIE',Sal,False);
  DD  := Item.GetString('PCN_DATEDEBUTABS');
  DF  := Item.GetString('PCN_DATEFINABS');
  Ret := AglLanceFiche('PAY','MVTABSENCE', '',Sal,Sal+';A;'+Etab+';ACTION=CREATION;;;'+DD+';'+DF);
  If Ret<>'' then
     Begin
     InitItemAbsenceModify(Item,Ret);
     //***ChargeTob_Recapitulatif;
     //RefrehRecapPlanning(sal, Item);
     if Item <> nil then THPlanningAbs.InvalidateItem(Item);
     Cancel := false;
     End;   
end;

procedure TPlanningAbsPaie.DeleteItemAbsence(Sender: TObject; Item: TOB;
  var Cancel: boolean);
Var Sal : String;
begin
  Cancel := True;
  Sal := Item.GetString('PCN_SALARIE');

  if Item.GetString('PCN_CODETAPE')='P' then
     Begin
     PgiBox('Vous ne pouvez supprimer une absence payée!',caption);
     Exit;
     End
  else
  if Item.GetString('PCN_CODETAPE')='C' then
     Begin
     PgiBox('Vous ne pouvez supprimer une absence clôturée!',caption);
     Exit;
     End
  else
  if Item.GetString('PCN_CODETAPE')='S' then
     Begin
     PgiBox('Vous ne pouvez supprimer une absence soldée!',caption);
     Exit;
     End;

  if PgiAsk('Confirmez-vous la suppression de l''absence : '+Item.GetValue('PCN_LIBELLE')+'?',Caption) = MrNo then Exit;

   Cancel:=SuppressionAbsence(Item,True);
end;


procedure TPlanningAbsPaie.InitItemAbsence(Sender: TObject; var Item: TOB;
  var Cancel: boolean);
begin
  Cancel := True;
  Item := Tob.Create('ABSENCESALARIE', GblTobItems, -1); { PT7 }
  Item.AddChampSupValeur('CREATION', 'TEMP', False);
  Cancel := false;
end;

procedure TPlanningAbsPaie.ModifyItemAbsence(Sender: TObject; Item: TOB;
  var Cancel: boolean);
begin
  Cancel := LanceFicheAbsenceModify (Item);
end;

function TPlanningAbsPaie.SuppressionAbsence(Item: Tob;
  AvecRecap: Boolean): Boolean;
begin
   Result := True;
   try
      BeginTrans;
      ExecuteSql('DELETE ABSENCESALARIE '+
                 'WHERE PCN_SALARIE="'+Item.GetString('PCN_SALARIE')+'" '+
                 'AND PCN_TYPEMVT="'+Item.GetString('PCN_TYPEMVT')+'" '+
                 'AND PCN_ORDRE='+IntToStr(Item.GetValue('PCN_ORDRE')));
      CommitTrans;
      Result := False;
      if Assigned(TobAbsDelete) then Item.Changeparent(TobAbsDelete, -1);
   except
      Rollback;
      PGIBox('Suppression annulée : Echec lors de la suppression de l''absence.', caption);
   end;
end;

procedure TPlanningAbsPaie.InitItemAbsenceModify(TEnr: Tob;
  var Ret: String);
Var T_MotifAbs : Tob;
begin
  if Not Assigned(TEnr) then exit;
  if Ret = '' then exit;
  TEnr.PutValue('PCN_ORDRE'         ,StrToInt(ReadTokenSt(Ret)));
  TEnr.PutValue('PCN_TYPECONGE'     ,ReadTokenSt(Ret));
  if TEnr.GetString('PCN_TYPECONGE')='PRI' then
  TEnr.PutValue('PCN_TYPEMVT'       ,'CPA')
  else
  TEnr.PutValue('PCN_TYPEMVT'       ,'ABS');
  TEnr.PutValue('PCN_DATEDEBUTABS'  ,StrToDate(ReadTokenSt(Ret)));
  TEnr.PutValue('PCN_DATEFINABS'    ,StrToDate(ReadTokenSt(Ret)));
  TEnr.PutValue('PCN_DEBUTDJ'       ,ReadTokenSt(Ret));
  TEnr.PutValue('PCN_FINDJ'         ,ReadTokenSt(Ret));
  TEnr.PutValue('PCN_JOURS'         ,Valeur(ReadTokenSt(Ret)));
  TEnr.PutValue('PCN_HEURES'        ,Valeur(ReadTokenSt(Ret)));
  TEnr.PutValue('PCN_LIBELLE'       ,ReadTokenSt(Ret));
  TEnr.AddChampSupValeur('ABREGE'   ,RechDom('PGMOTIFABSENCE', TEnr.GetValue('PCN_TYPECONGE'), False));
  TEnr.PutValue('PCN_LIBCOMPL1'     ,ReadTokenSt(Ret));
  TEnr.PutValue('PCN_LIBCOMPL2'     ,ReadTokenSt(Ret));
  TEnr.PutValue('PCN_VALIDRESP'     ,ReadTokenSt(Ret));
  TEnr.PutValue('PCN_EXPORTOK'      ,ReadTokenSt(Ret));
  TEnr.PutValue('PCN_CODETAPE'      ,ReadTokenSt(Ret));
  if Assigned(TobMotifAbs) then
    Begin
    T_MotifAbs:=TobMotifAbs.FindFirst(['PMA_MOTIFABSENCE'],[TEnr.GetString('PCN_TYPECONGE')],False);
    if Assigned(T_MotifAbs) then
      TEnr.AddChampSupValeur('PMA_TYPEABS',T_MotifAbs.GetValue('PMA_TYPEABS'));
    End;
end;

procedure TPlanningAbsPaie.TABSENCEClick(Sender: TObject);
begin
  AglLanceFiche('PAY', 'ABSENCE_MUL', '', '', ';MENU;;ABS');
end;

procedure TPlanningAbsPaie.HelpBtnClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;



end.




