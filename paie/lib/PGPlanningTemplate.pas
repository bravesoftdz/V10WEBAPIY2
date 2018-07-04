{***********UNITE*************************************************
Auteur  ...... : GGU
Créé le ...... : 20/06/2007
Modifié le ... :   /  /
Description .. : TOM visualisation d'un planning : Fiche ancêtre
Mots clefs ... : PLANNING
*****************************************************************

}
unit PGPlanningTemplate;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
{$IFDEF EAGLCLIENT}
  UtileAGL,eMul,MaineAgl,HStatus,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,
{$ENDIF}
  HCtrls, StdCtrls, HTB97, ExtCtrls, hplanning, UTOB, Dialogs,
  HEnt1, HeureUtil ,PgPlanningOutils, ComCtrls, Grids, Menus;

Type
  TPlanningEvent      = procedure (Planning : THPlanning; Sender : TObject) of Object;
  TPlanningPopupEvent = procedure (Item: TOB; ZCode: Integer; var ReDraw: Boolean; IdRessource : String; CurentDate : TDateTime; ZoneSelectionnee : TZoneSelected) of Object;
  TPlanningTobEvent   = procedure (Planning : THPlanning; Tob : TOB) of Object;
  TPlanningItemsProcedure = procedure (Deb, Fin : TDateTime) of Object;

type
  TFormPlanning = class(TForm)
    PPanel: TPanel;
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BImprimer: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    TRECAP: TToolbarButton97;
    Bperso: TToolbarButton97;
    TExcel: TToolbarButton97;
    BExporter: TToolbarButton97;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    BtnDetail: TToolbarButton97;
    PRECAP: TPanel;
    PCPlannings: TPageControl;
    BtnHighlight: TToolbarButton97;
    PopupHighlight: THPopupMenu;
    HSplitter1: THSplitter;
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure TExcelClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BtnDetailClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure PCPlanningsChange(Sender: TObject);
    procedure PCPlanningsChanging(Sender: TObject; var AllowChange: Boolean);
    procedure BNavigateClick(Sender: TObject);
  private
    FIntervalDebut, FIntervalFin : TDateTime;
    //Evènements et procedures à gérer par les classes filles
    FOnCreatePlanning    : TPlanningEvent;
    FOnChargeTobPlanning : TNotifyEvent;
    FOnPopup             : TPlanningPopupEvent;
    FOnAskRefreshDetails : TPlanningTobEvent;
    FOnBeforeShow        : TNotifyEvent;
    FLoadItemsProcedure, FLoadEtatsProcedure, FLoadRessourcesProcedure: TPlanningItemsProcedure ;  //  , FDeleteItemsProcedure
    procedure CreatePlanning(PlanningInterval : THPlanningInterval; PlanningCumulInterval : THPlanningCumulInterval; PlanningParent : TWinControl; PlanningModeOutlook : Boolean = False);
    function ChargeTobPlanning : Boolean;
    procedure ClickPlanning(Sender: TObject);
    procedure SetAllPlanIntervalDebut(const Value: TDateTime);
    procedure SetAllPlanIntervalFin(const Value: TDateTime);
    procedure HAvertirApplication(Sender: TObject; FromItem, ToItem: TOB;
      Actions: THPlanningAction);
    procedure OptionPopup(Item: TOB; ZCode: Integer; var ReDraw: Boolean);

  protected
    Plannings : Array of THPlanning ;
    TobRessources, TobEtats, TobItems: Tob;
    FontName, FontStyle : string;
    FontColor, DefaultColorBackground : TColor;
    FontSize : Integer;
    ListIntervales : TPlanningIntervalList;
    AskRefreshDetailsForAllClick : Boolean;
    procedure ShowPlanning;
    procedure DelItems(Deb, Fin: TDateTime; ChampsDate : String; ChampsDateFin : String = '');
    procedure CreateDetailsLabel(var LeLabel: THLabel; Parent: TWinControl;
      cLeft, cTop, cWidth, cHeight: Integer; cCaption: String; LibelleMode : Boolean = True;
      cVisible : Boolean = True);
    procedure CreateDetailsGroupBox(var LeGroupBox: THGroupBox;
      cLeft, cTop, cWidth, cHeight: Integer; cCaption: String; cAlign: TAlign = alClient;
      cVisible: Boolean = True);
    // Variables qui permettent de gérer les dates de début et de fin de l'ensemble des plannings de manière global
    property AllPlanIntervalDebut : TDateTime read FIntervalDebut write SetAllPlanIntervalDebut;
    property AllPlanIntervalFin   : TDateTime read FIntervalFin   write SetAllPlanIntervalFin;
  public
    //Evènements à gérer par les classes filles
    property OnCreatePlanning     : TPlanningEvent          read FOnCreatePlanning     Write FOnCreatePlanning;
    property OnChargeTobPlanning  : TNotifyEvent            read FOnChargeTobPlanning  Write FOnChargeTobPlanning;
    property OnPopup              : TPlanningPopupEvent     read FOnPopup              Write FOnPopup;
    property OnAskRefreshDetails  : TPlanningTobEvent       read FOnAskRefreshDetails  Write FOnAskRefreshDetails;
    property OnBeforeShow         : TNotifyEvent            read FOnBeforeShow         Write FOnBeforeShow;
    property LoadItemsProcedure   : TPlanningItemsProcedure read FLoadItemsProcedure   Write FLoadItemsProcedure;
    property LoadEtatsProcedure   : TPlanningItemsProcedure read FLoadEtatsProcedure   Write FLoadEtatsProcedure;
    property LoadRessourcesProcedure: TPlanningItemsProcedure read FLoadRessourcesProcedure Write FLoadRessourcesProcedure;
    procedure Highlight( FiltreNomChamps : Array of String; FiltreValeurChamps : Array of Variant; Activate : Boolean = True);
  end;

var
  FormPlanning: TFormPlanning;

implementation

uses DateUtils, Math, HMsgBox, ed_tools;

{$R *.dfm}

procedure TFormPlanning.HAvertirApplication(Sender: TObject; FromItem,
  ToItem: TOB; Actions: THPlanningAction);
begin
  if Actions = paclickLeft then
  begin
    if  ((not AskRefreshDetailsForAllClick) and (FromItem <> Nil)) or (AskRefreshDetailsForAllClick) then
      if Assigned(OnAskRefreshDetails) then
        OnAskRefreshDetails(Plannings[PCPlannings.activePageIndex],FromItem);
  end;
end;

procedure TFormPlanning.FormShow(Sender: TObject);
begin
  { Activation du planning affiché }
  Plannings[PCPlannings.activePageIndex].activate := True;
end;


function TFormPlanning.ChargeTobPlanning: Boolean;
Var
  IndexPlanning, IndexEtat : Integer;
  tempTob : Tob;
begin
  try
    if Assigned(TobItems) then FreeAndNil(TobItems);
    if Assigned(TobEtats) then FreeAndNil(TobEtats);
    if Assigned(TobRessources) then FreeAndNil(TobRessources);
    TobRessources := TOB.Create('Les ressources', nil, -1);
    TobItems := TOB.Create('Les items', nil, -1);
    TobEtats := TOB.Create('Les états', nil, -1);
    { Chargement des ressources }
    if Assigned(LoadRessourcesProcedure) then LoadRessourcesProcedure(AllPlanIntervalDebut, AllPlanIntervalFin);
    { Chargement des Etats }
    if Assigned(LoadEtatsProcedure) then LoadEtatsProcedure(AllPlanIntervalDebut, AllPlanIntervalFin);
    { Chargement des Items }
    if Assigned(LoadItemsProcedure) then LoadItemsProcedure(AllPlanIntervalDebut, AllPlanIntervalFin);

    If (TobRessources.Detail.Count = 0) Or (TobEtats.Detail.Count = 0) Then
    Begin
        Result := False;
        PGIBox ('La sélection ne renvoie aucun résultat.');
        Exit;
    End;

    for IndexPlanning := 0 to Length(Plannings) -1 do
    begin
      Plannings[IndexPlanning].TobRes   := TobRessources;
      Plannings[IndexPlanning].TobEtats := TobEtats;
      Plannings[IndexPlanning].TobItems := TobItems;
    end;
    result := True;
    if Assigned(OnChargeTobPlanning) then OnChargeTobPlanning(Self);
    { Ajout d'un champs dans la tob des etats et sauvegarde de la couleur }
    TobEtats.detail[0].AddChampSup('SAVECOLOR', True);
    for indexEtat := 0 to TobEtats.FillesCount(0)-1 do
    begin
      tempTob := TobEtats.Detail[indexEtat];
      tempTob.PutValue('SAVECOLOR',tempTob.GetString(Plannings[0].EtatChampBackGroundColor));
    end;
  except
    result := False;
  end;
end;

procedure TFormPlanning.FormDestroy(Sender: TObject);
var
  IndexPlanning : Integer;
begin
  for IndexPlanning := 0 to Length(Plannings) -1 do
  begin
    if Assigned(Plannings[IndexPlanning]) then FreeAndNil(Plannings[IndexPlanning]);
  end;
  FreeAndNil(TobRessources);
  FreeAndNil(TobItems);
  FreeAndNil(TobEtats);
end;

procedure TFormPlanning.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TFormPlanning.TExcelClick(Sender: TObject);
begin
  with TSaveDialog.Create(Self) do
  begin
    Filter := 'Microsoft Excel (*.xls)|*.XLS';
    if Execute then Plannings[PCPlannings.activePageindex].ExportToExcel(True, FileName);
    Free;
  end;
end;

procedure TFormPlanning.BImprimerClick(Sender: TObject);
var
  S: string;
begin
  With Plannings[PCPlannings.activePageindex] do
  begin
    TypeEtat := 'E';
    NatureEtat := 'PCG';
    CodeEtat := 'PL1';
    S := 'DATEDEB=' + FormatDateTime('dd/mm/yyyy', IntervalDebut) + '^E' + IntToStr(Byte(otDate));
    S := S + '`DATEFIN=' + FormatDateTime('dd/mm/yyyy', IntervalFin) + '^E' + IntToStr(Byte(otDate));
    CriteresToPrint := S;
    Print;
  end;
end;

procedure TFormPlanning.BFermeClick(Sender: TObject);
begin
  Close;
end;

procedure TFormPlanning.BtnDetailClick(Sender: TObject);
begin
  PRECAP.Visible := BtnDetail.Down;
  HSplitter1.Visible := BtnDetail.Down;
end;

procedure TFormPlanning.HelpBtnClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFormPlanning.CreatePlanning(PlanningInterval : THPlanningInterval; PlanningCumulInterval : THPlanningCumulInterval; PlanningParent : TWinControl; PlanningModeOutlook : Boolean = False);
var
  indexPlanning : Integer;
  stCaption, PlanningDateFormat : String;
  TempTabSheet : TTabSheet;
begin
  { Si le parents est un pageControl, on crée un nouvel onglet et on se met dedans }
  if PlanningParent is TPageControl then
  begin
    TempTabSheet := TTabSheet.Create(PlanningParent);
    TempTabSheet.PageControl := (PlanningParent as TPageControl);
    PlanningParent := TempTabSheet;
  end;
  SetLength(Plannings,Length(Plannings)+1);
  indexPlanning := Length(Plannings)-1;
  Plannings[indexPlanning] := THplanning.Create(Self);
  Plannings[indexPlanning].Parent := PlanningParent;
  Plannings[indexPlanning].Align := alClient;
  Plannings[indexPlanning].Visible := True;
  if PlanningModeOutlook = True then
  begin
    stCaption := 'Outlook';
    PlanningDateFormat := 'dd/mm';
  end else if PlanningInterval = piDemiHeure then
  begin
    stCaption := 'Demi-heure';
    PlanningDateFormat := 'hh:mm';
  end else if PlanningInterval = piHeure then
  begin
    stCaption := 'Heure';
    PlanningDateFormat := 'hh:mm';
  end else if PlanningInterval = piDemiJour then
  begin
    stCaption := 'Demi-journée';
    PlanningDateFormat := 'dd/mm tt';
  end else if PlanningCumulInterval = pciSemaine then
  begin
    stCaption := 'Semaine';
    PlanningDateFormat := 'dd/mm';
    PCPlannings.ActivePageIndex := indexPlanning;
  end else if PlanningCumulInterval = pciMois then
  begin
    stCaption := 'Mois';
    PlanningDateFormat := 'dd/mm';
  end else if PlanningCumulInterval = pciTrimestre then
  begin
    stCaption := 'Trimestre';
    PlanningDateFormat := 'dd/mm';
  end else if PlanningCumulInterval = pciSemestre then
  begin
    stCaption := 'Semestre';
    PlanningDateFormat := 'dd/mm';
  end else if PlanningCumulInterval = pciAnnee then
  begin
    stCaption := 'Année';
    PlanningDateFormat := 'dd/mm';
  end else begin
    stCaption := 'Jour';
    PlanningDateFormat := 'dd/mm';
  end;

  if PlanningParent is TTabSheet then
    (PlanningParent as TTabSheet).Caption := stCaption
  else if PlanningParent is TPanel then
    (PlanningParent as TPanel).Caption :=  stCaption;

  with Plannings[indexPlanning] do
  begin
    Align := alClient;
    activate := False;
    { Les évènements }
    OnDblClick := ClickPlanning;
    OnAvertirApplication := HAvertirApplication;
    Personnalisation := False;
    Legende := True;
    EtatChampFontColor := 'FONTCOLOR';
    EtatChampFontStyle := 'FONTSTYLE';
    EtatChampFontSize := 'FONTSIZE';
    EtatChampFontName := 'FONTNAME';
    TextAlign := taLeftJustify;
    { Personalisation }
    IntervalDebut := AllPlanIntervalDebut;
    IntervalFin   := AllPlanIntervalFin;
    MultiLine := False;
    GestionJoursFeriesActive := True;
    ActiveSaturday := True;
    ActiveSunday   := True;
    ColorBackground  := DefaultColorBackground;
    ColorJoursFeries := StringToColor('12189695');
    ColorOfSaturday  := StringToColor('12189695');
    ColorOfSunday    := StringToColor('12189695');
    ColorSelection := clNavy;
    ActiveLigneDate := True;
    ActiveLigneGroupeDate := True;
    if IntervalDebut = IntervalFin then ColSizeData := 250;   { PT10 }
    { Liste des champs d'entete }
    TokenFieldColEntete := '';
    { Taille des colonnes fixes }
    JourneeDebut := StrToTime('00:00');
    JourneeFin   := StrToTime('23:00');
    DebutAM      := StrToTime('00:00');
    FinAM        := StrToTime('12:00');
    DebutAP      := StrToTime('12:00');
    FinAP        := StrToTime('23:00');
    DisplayOptionCreation           := False;
    DisplayOptionModification       := False;
    DisplayOptionSuppression        := False;
    DisplayOptionEtirer             := False;
    DisplayOptionReduire            := False;
    MoveHorizontal                  := False;
    DisplayOptionCopie              := False;
    DisplayOptionDeplacement        := False;
    DisplayOptionSuppressionLiaison := False;
    DisplayOptionLiaison            := False;
    DisplayOptionLier               := False;
    { Les évènements item }
//    OnInitItem   := InitItemAbsence;
//    OnCreateItem := CreateItemAbsence;
//    OnModifyItem := ModifyItemAbsence;
//    OnDeleteItem := DeleteItemAbsence;
    MouseAlready := True;
    Interval := PlanningInterval;
    CumulInterval := PlanningCumulInterval;
    DateFormat := PlanningDateFormat;
    VoirToutesLesHeures := True;

    OnOptionPopup := OptionPopup;

    if PlanningModeOutlook then
    begin
      ModeOutlook := True;
      IntervalOutlook := pioHeure;
      VisionOutlook := voWeek;
      ReductionItemOutlook := True;
    end;
  end;
  if Assigned(OnCreatePlanning) then OnCreatePlanning(Plannings[indexPlanning],self);
end;

procedure TFormPlanning.PCPlanningsChange(Sender: TObject);
begin
  Plannings[PCPlannings.ActivePageindex].Activate := True;
end;

procedure TFormPlanning.PCPlanningsChanging(Sender: TObject;
  var AllowChange: Boolean);
begin
  Plannings[PCPlannings.ActivePageindex].Activate := False;
end;

procedure TFormPlanning.BNavigateClick(Sender: TObject);
var
  DD, DF: TDateTime;
  PlanningActif : THPlanning;
  Diff : Extended;
begin
  PlanningActif := Plannings[PCPlannings.ActivePageIndex];
  if PlanningActif = nil then exit;
  PlanningActif.Activate := False;
  DD := AllPlanIntervalDebut;
  DF := AllPlanIntervalFin;
  Diff := (DF-DD);
  if Sender = BFirst then  //On décale d'une période d'affichage (début et fin)
  begin
    TobItems.ClearDetail;
    AllPlanIntervalDebut := DD - Diff;
    AllPlanIntervalFin := DF - Diff;
    If Assigned(LoadItemsProcedure) then LoadItemsProcedure(DD - Diff, DF - Diff);
  end else if Sender = BPrev then  //on décale la fin d'1 semaine
  begin
    AllPlanIntervalDebut := DD - 7;
    If Assigned(LoadItemsProcedure) then LoadItemsProcedure(DD - 7, DD-1);
  end else if Sender = BNext then  //On décale le début d'une semaine
  begin
    AllPlanIntervalFin := DF + 7;
    If Assigned(LoadItemsProcedure) then LoadItemsProcedure(DF+1, DF+7);
  end else if Sender = BLast then //On décale d'une période d'affichage (début et fin)
  begin
    TobItems.ClearDetail;
    AllPlanIntervalDebut := DD + Diff;
    AllPlanIntervalFin := DF + Diff;
    If Assigned(LoadItemsProcedure) then LoadItemsProcedure(DD + Diff, DF + Diff);
  end;
  PlanningActif.Activate := True;
end;                                   

procedure TFormPlanning.SetAllPlanIntervalDebut(
  const Value: TDateTime);
var
  indexPlanning : Integer;
begin
  FIntervalDebut := Value;
  For indexPlanning := 0 to Length(Plannings) -1 do
  begin
    Plannings[indexPlanning].IntervalDebut := FIntervalDebut;
  end;
end;

procedure TFormPlanning.SetAllPlanIntervalFin(
  const Value: TDateTime);
var
  indexPlanning : Integer;
begin
  FIntervalFin := Value;
  For indexPlanning := 0 to Length(Plannings) -1 do
  begin
    Plannings[indexPlanning].IntervalFin := FIntervalFin;
  end;
end;

procedure TFormPlanning.ClickPlanning(Sender: TObject);
Var
  T: TOB;
begin
  T := Plannings[PCPlannings.ActivePageIndex].GetCurItem;
  if  ((not AskRefreshDetailsForAllClick) and (T <> Nil)) or (AskRefreshDetailsForAllClick) then
    if Assigned(OnAskRefreshDetails) then
      OnAskRefreshDetails(Plannings[PCPlannings.activePageIndex],T);
end;

//procedure TFormPlanning.GereOptionPopup(Item: TOB; ZCode: Integer;
//  var ReDraw: Boolean);
//Var
//  CurentDate : TDateTime;
//  TempTob : Tob;
//  Cycle : String;
//begin
//  if (ZCode = 99) then  { Création d'une exception }
//  begin
//    Cycle := '';
//    { On cherche la date et le cycle courants }
//    if not Assigned (Item) then
//    begin
//      CurentDate := Plannings[PCPlannings.ActivePageIndex].GetDateOfCol(Plannings[PCPlannings.ActivePageIndex].Col);
//      TempTob := nil;
//      { En mode outlook }
//      if Plannings[PCPlannings.ActivePageIndex].ModeOutlook then
//      begin
//        if TobRessources.FillesCount(0) > 1 then
//          {Si on a plusieurs ressources, on ne sais pas laquelle choisir }
//          PGIInfo('Vous devez selectionner le cycle sur lequel l''exception doit être appliquée.')
//        else
//          {Si on a 1 seule ressource }
//          TempTob := TobRessources.Detail[0];
//      end else
//        TempTob := Plannings[PCPlannings.ActivePageIndex].GetRessourceOfRow(Plannings[PCPlannings.ActivePageIndex].Row);
//      if Assigned(TempTob) then
//        Cycle := TempTob.GetString('CYCLE');
//    end else begin
//      CurentDate := Item.GetDateTime('JOUR');
//      Cycle := Item.GetString('CYCLE_OU_MODELECYCLE');
//    end;
//    if (Cycle <> '') and (CurentDate <> iDate1900) then
//    begin
//      if GestionPresence.DateOfException(CurentDate, Cycle, TypePlanning) then
//        AGLLanceFiche('PAY', 'EXCEPTCYCLE','',TypePlanning+';'+Cycle+';'+DateToStr(CurentDate),'ACTION=MODIFICATION')
//      else
//        AGLLanceFiche('PAY', 'EXCEPTCYCLE','','','ACTION=CREATION;PYA_TYPECYCLE='+TypePlanning+';PYA_CYCLE='+Cycle+';PYA_DATEEXCEPTION='+DateToStr(CurentDate));
//      Plannings[PCPlannings.ActivePageIndex].Activate := False;
//      GestionPresence.ReloadException;
//      DelItemsPeriode (CurentDate, CurentDate, TobItems);
//      LoadItemsPeriode(CurentDate, CurentDate, TobItems);
//      Plannings[PCPlannings.ActivePageIndex].Activate := True;
//    end;
//  end;
//end;

procedure TFormPlanning.DelItems(Deb, Fin: TDateTime; ChampsDate : String; ChampsDateFin : String = '');
var
  TempTob : Tob;
  indexItem : Integer;
  Date, DateFin : TDateTime;
  DDeb, DFin : TDateTime;
begin
  if length(Plannings) > 0 then
  begin
    DDeb := DateOf(Deb);
    DFin := DateOf(Fin);
    InitMoveProgressForm(nil, TraduireMemoire('Suppression des évènements'), TraduireMemoire('Veuillez patienter...'), TobItems.FillesCount(0), False, True);
    for indexItem := TobItems.FillesCount(0) -1 downto 0 do
    begin
      TempTob := TobItems.Detail[indexItem];
      DateFin := idate1900;
      Date := DateOf(TempTob.GetDateTime(ChampsDate));
      if ChampsDateFin <> '' then DateFin := DateOf(TempTob.GetDateTime(ChampsDateFin));
      if  ((Date >= DDeb) and (Date <= DFin))
      or  ((ChampsDateFin <> '') and (   ((DateFin >= DDeb) and (DateFin <= DFin))
                                      or ((Date    <= DDeb) and (DateFin >= DFin)) )) then
      begin
        MoveCurProgressForm('');
        TempTob.ChangeParent(nil, -1);
        FreeAndNil(TempTob);
      end;
    end;
    FiniMoveProgressForm;
  end;
end;

procedure TFormPlanning.CreateDetailsLabel(var LeLabel : THLabel; Parent : TWinControl; cLeft, cTop, cWidth, cHeight : Integer; cCaption : String; LibelleMode : Boolean = True; cVisible : Boolean = True);
begin
  LeLabel := THLabel.Create(FormPlanning);
  LeLabel.Parent := Parent;
  with LeLabel do
  begin
    Left    := cLeft;
    Top     := cTop;
    Width   := cWidth;
    Height  := cHeight;
    Caption := cCaption;
    if LibelleMode then
    begin
      Font.Charset := DEFAULT_CHARSET;
      Font.Color   := clTeal;
      Font.Height  := -11;
      Font.Name    := 'MS Sans Serif';
      Font.Style   := [];
      ParentFont   := False;
    end;
    Visible := cVisible;
  end;
end;

procedure TFormPlanning.CreateDetailsGroupBox(var LeGroupBox: THGroupBox;
  cLeft, cTop, cWidth, cHeight: Integer; cCaption: String; cAlign: TAlign = alClient;
  cVisible: Boolean = True);
begin
  LeGroupBox := THGroupBox.Create(FormPlanning);
  LeGroupBox.Parent := PRECAP;
  with LeGroupBox do
  begin
    Left    := cLeft;
    Top     := cTop;
    Width   := cWidth;
    Height  := cHeight;
    Caption := cCaption;
    Align   := cAlign;
    ParentBiDiMode := False;
    BiDiMode := bdRightToLeft;
    TabOrder := 0;
    Visible := cVisible;
  end;
end;

procedure TFormPlanning.OptionPopup(Item: TOB; ZCode: Integer;
  var ReDraw: Boolean);
Var
  CurrentDate : TDateTime;
  TempTob : Tob;
  IDRessource : String;
  ZoneSel : TZoneSelected;
  NbrColonnesEntete : Integer;
begin
  if Assigned(OnPopup) then
  begin
    IDRessource := '';
    { On cherche la date et le cycle courants }
    CurrentDate := Plannings[PCPlannings.ActivePageIndex].GetDateOfCol(Plannings[PCPlannings.ActivePageIndex].Col);
    if not Assigned (Item) then
    begin
      TempTob := nil;
      { En mode outlook }
      if Plannings[PCPlannings.ActivePageIndex].ModeOutlook then
      begin
        if TobRessources.FillesCount(0) > 1 then
          {Si on a plusieurs ressources, on ne sais pas laquelle choisir }
          PGIInfo('Vous devez selectionner l''objet sur lequel le traitement doit être appliquée.')
        else
          {Si on a 1 seule ressource }
          TempTob := TobRessources.Detail[0];
      end else
        TempTob := Plannings[PCPlannings.ActivePageIndex].GetRessourceOfRow(Plannings[PCPlannings.ActivePageIndex].Row);
      if Assigned(TempTob) then
        IDRessource := TempTob.GetString(Plannings[PCPlannings.ActivePageIndex].ResChampID);
    end else begin
      IDRessource := Item.GetString(Plannings[PCPlannings.ActivePageIndex].ChampLineID);
    end;
    ZoneSel := Plannings[PCPlannings.ActivePageIndex].ZoneSelected;
    { Bidouillage pour corriger un bogue de l'AGL sur les dates de la zone sélectionnée : }
    NbrColonnesEntete := Plannings[PCPlannings.ActivePageIndex].FixedCols;
    ZoneSel.DebutSelection := ZoneSel.DebutSelection - NbrColonnesEntete + 1;
    ZoneSel.FinSelection   := ZoneSel.FinSelection   - NbrColonnesEntete + 1;
    { Fin de la bidouille }
    OnPopup(Item, ZCode, ReDraw, IDRessource, CurrentDate, ZoneSel);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 11/07/2007
Modifié le ... :   /  /    
Description .. : Crée les onglets, charge les données et affiche le planning
Suite ........ : si les ressources et les états ne sont pas vides
Mots clefs ... :
*****************************************************************}
procedure TFormPlanning.ShowPlanning;
begin
  FontColor := clBlack;
  FontStyle := 'G';
  FontName := 'Times New Roman';
  DefaultColorBackground := clWhite;
  FontSize := 10;
  
  { On crée un planning pour chaque interval choisit }
  if (piaHeure in ListIntervales) then
    CreatePlanning(piHeure, pciJour, PCPlannings);
  if (piaDemiJournee in ListIntervales) then
    CreatePlanning(piDemiJour, pciSemaine, PCPlannings);
  if (piaSemaine in ListIntervales) then
    CreatePlanning(piJour, pciSemaine, PCPlannings);
  if (piaMois in ListIntervales) then
    CreatePlanning(piJour, pciMois, PCPlannings);
  if (piaOutlook in ListIntervales) then
    CreatePlanning(piHeure, pciJour, PCPlannings, True);
  if (piaTrimestre in ListIntervales) then
    CreatePlanning(piSemaine, pciTrimestre, PCPlannings);
  if (piaSemestre in ListIntervales) then
    CreatePlanning(piMois, pciSemestre, PCPlannings);

  { Création et chargement des trois TOB }
  if not ChargeTobPlanning then
    PostMessage(Handle, SBWM_APPSTARTUP, 0, 0)
  Else begin
    if Assigned(OnBeforeShow) then OnBeforeShow(self);
    ShowModal;
  end;
end;

procedure TFormPlanning.Highlight(FiltreNomChamps: array of String;
  FiltreValeurChamps: array of Variant; Activate : Boolean = True);
var
  indexEtat : Integer;
  tempTob: Tob;
  stColorField : String;
begin
  Plannings[PCPlannings.ActivePageindex].Activate := False;
  stColorField := Plannings[PCPlannings.ActivePageindex].EtatChampBackGroundColor;
  if Activate then
  begin
    tempTob := TobEtats.FindFirst(FiltreNomChamps, FiltreValeurChamps, False);
    While Assigned(tempTob) do
    begin
      tempTob.PutValue(stColorField,'1');
      tempTob := TobEtats.FindNext(FiltreNomChamps, FiltreValeurChamps, False);
    end;
  end;
  for indexEtat := 0 to TobEtats.FillesCount(0)-1 do
  begin
    tempTob := TobEtats.Detail[indexEtat];
    if Activate then
    begin
      if (tempTob.GetString(stColorField) = '1') then
        tempTob.PutValue(stColorField, tempTob.GetString('SAVECOLOR'))
      else
        tempTob.PutValue(stColorField, clInactiveCaptionText);
    end else begin
      tempTob.PutValue(stColorField, tempTob.GetString('SAVECOLOR'));
    end;
  end;
  Plannings[PCPlannings.ActivePageindex].RefreshEtats;
  Plannings[PCPlannings.ActivePageindex].Activate := True;
end;

end.
