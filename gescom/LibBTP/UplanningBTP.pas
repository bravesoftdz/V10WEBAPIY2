unit UplanningBTP;

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
  StdCtrls,
  Mask,
  Hctrls,
  ExtCtrls,
  HTB97,
  HPanel,
  HEnt1,
  Lookup,
  HMsgBox,
  AglInit,
{$IFDEF EAGLCLIENT}
  MaineAGL,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  Fe_main,
{$ENDIF}
  uTOB,
  HSysMenu, TntExtCtrls, TntStdCtrls, Grids, TntGrids,DateUtils,UdateUtils,
  ImgList, HImgList,Paramsoc, Menus, Spin,Math, ComCtrls;

type
  TModegestion = (tmgModif,tmgConsult);
  TModeGestionPC = (TmPcPlanCharge,TmPcPlanningCha,TmcPlanCharfnc);
  TmodeAffichage = (TmAHeure,TmaJours);

  ThePos = class
    Acol : Integer;
    Arow : integer;
  end;

  TheMergeItem = class (TObject)
    Debut : Integer;
    Fin : Integer;
    Numero : Integer;
  end;

  TlistMerge = class (TList)
  private
    function GetItems(Indice : integer): TheMergeItem;
    procedure SetItems(Indice : integer; const Value: TheMergeItem);
    function Add(AObject: TheMergeItem): Integer;
  public
    property Items [Indice : integer] : TheMergeItem read GetItems write SetItems;
    function Find ( Numero : integer) : TheMergeItem;
    procedure clear; override;
    destructor destroy; override;
  end;

  TheColonne = class (TObject)
    Nom : string;
    Indice : Integer;
    JourSemaine : integer;
    // pour la gestion des dates
    TheDate : TDateTime;
    Libelle : string;
    semaine : Integer;
    Mois : Integer;
  end;

  TListColonnes = class (Tlist)
  private
    function GetItems(Indice : integer): TheColonne;
    procedure SetItems(Indice : integer; const Value: TheColonne);
    function Add(AObject: TheColonne): Integer;
  public
    property Items [Indice : integer] : TheColonne read GetItems write SetItems;
    function Find (NomColonne : string) : TheColonne;
    function FindIndice (TheDate : TDateTime) : Integer;
    procedure clear; override;
    destructor destroy; override;
  end;


  TFplanningBTP = class(TForm)
    PCENTRE: TPanel;
    PHAUT: TPanel;
    PBAS: TPanel;
    PACTIONS: TPanel;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    GMois: THGrid;
    TLIBAFFAIRE: THLabel;
    TLIBELLE: THLabel;
    TLIBHRSPREV: THLabel;
    TLIBRESTEAPLANNIF: THLabel;
    GTHeures: THGrid;
    TLIBTOTALHRS: THLabel;
    TLibEffectif: THLabel;
    PCENTER: TPanel;
    G2: THGrid;
    G1: THGrid;
    GSem: THGrid;
    GDate: THGrid;
    GTEffectifs: THGrid;
    BRefresh: TToolbarButton97;
    BOuvrirFermer: TToolbarButton97;
    ImgListPlusMoins: THImageList;
    POPACTIONS: TPopupMenu;
    PlannifCh: TMenuItem;
    BMODIF: TToolbarButton97;
    BANNULE: TToolbarButton97;
    LNBRRESS: TLabel;
    NBRRESS: THSpinEdit;
    LNBHRS: TLabel;
    NBHRS: THNumEdit;
    BPARAMS: TToolbarButton97;
    TTPARAMS: TToolWindow97;
    DateD: THCritMaskEdit;
    DateD_: THCritMaskEdit;
    HLabel1: THLabel;
    HLabel2: THLabel;
    BREFRESHP: TToolbarButton97;
    TBVIEWARBO: TToolbarButton97;
    TTARBO: TToolWindow97;
    TTVARBO: TTreeView;
    CAffHeures: TCheckBox;
    GroupBox1: TGroupBox;
    RBHeures: TRadioButton;
    RbJours: TRadioButton;
    GFONCTION: THGrid;
    LPOTENTIEL: THLabel;
    LAFFECTE: THLabel;
    LRESTE: THLabel;
    LFONCTION: THLabel;
    BImprimer: TToolbarButton97;
    N1: TMenuItem;
    AffectRessources: TMenuItem;
    procedure BFermeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure GSemDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure BOuvrirFermerClick(Sender: TObject);
    procedure BRefreshClick(Sender: TObject);
    procedure PlannifChClick(Sender: TObject);
    procedure BMODIFClick(Sender: TObject);
    procedure BANNULEClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure NBRRESSChange(Sender: TObject);
    procedure NBHRSExit(Sender: TObject);
    procedure BREFRESHPClick(Sender: TObject);
    procedure BPARAMSClick(Sender: TObject);
    procedure TTPARAMSClose(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure TBVIEWARBOClick(Sender: TObject);
    procedure TTVARBOClick(Sender: TObject);
    procedure TTARBOClose(Sender: TObject);
    procedure CAffHeuresClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure AffectRessourcesClick(Sender: TObject);
  private
    DetailModif : boolean;
    ModeGestionGrid : TModeGestion;
    DureeJournee : double;
    ModeColor : Boolean;
    LastPos : ThePos;
    JourFermeture : array[0..6] of Integer;    // Liste des jours de fermeture
    NbJours : Integer;
    TobAffiches : TOb;
    TOBAffects,TOBAffectsDel : TOB;
    TOBTOTJours,TOBTOTFONC : TOB;
    LesColsG1,LesColsG2,LesColsGTH : string;
    LesColonnesG1,lesColonnesG2 : TListColonnes;
    MergedSem : TlistMerge;
    MergedMois : TlistMerge;
    { Déclarations privées }
    procedure LoadLesTobs;
    procedure PrepareAffichage;
    procedure CreateTobs;
    procedure FreeTobs;
    //
    function ConstitueDatesCha (Datedebut : TdateTime; detaille : boolean=false) : string;
    function GetColCounts(LaChaine: string): Integer;
    procedure ReajusteEcran;
    function GetWidth(Grid: Thgrid): integer;
    procedure PositionneTexteEntete (TheControl : THlabel; TheChamps : string;Lescolonnes : TListColonnes; TheGrid : THGrid);
    procedure SetDates(GDate: Thgrid);
    procedure SeteventsGrid (Enabled : Boolean);
    //
    procedure ConstitueGrilleG1;
    procedure G1TopLeftChanged (Sender: TObject);
    procedure G1PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure G1RowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G1GetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas;AState: TGridDrawState);
    //
    procedure ConstitueGrilleG2;
    function ConstitueStringG2: string;
    procedure G2TopLeftChanged (sender : TObject);
    procedure G2RowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G2GetCellCanvas (ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure G2PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure G2Cellenter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G2MouseDown(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);
    //
    procedure GTheuresDrawCell(Sender: TObject; ACol, ARow: Integer;Rect: TRect; State: TGridDrawState);
    procedure GTEffectifsDrawCell(Sender: TObject; ACol, ARow: Integer;Rect: TRect; State: TGridDrawState);
    procedure GMoisDrawCell(Sender: TObject; ACol, ARow: Integer;Rect: TRect; State: TGridDrawState);
    procedure GDateDrawCell(Sender: TObject; ACol, ARow: Integer;Rect: TRect; State: TGridDrawState);
    //
    procedure GFONCTIONDrawCell (Sender: TObject; ACol, ARow: Integer;Rect: TRect; State: TGridDrawState);
    procedure GFONCTIONCellenter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    //
    procedure ConstitueGrilleTH;
    procedure DefinilesSemaines;
    procedure DefinilesMois;
    function MergingCell(Sender: TObject; Coldeb,ColFin, aCol, ARow: Integer; State: TGridDrawState): boolean;
    procedure DrawButton(Grid: TStringGrid; Arect: TRect; Fin: boolean);
    procedure OwnerDrawText(Grid: TStringGrid;State: TGridDrawState;Acol : Integer;  rect: Trect; Texte: String; UFormat: Cardinal; TheColor : Tcolor=0);
    //
    function ConstitueCumJour: string;
    procedure AddOneTotal;
    procedure Colorise(Colonne: Integer; Ajout: Boolean);
    procedure ClearCanvas(Grid: TStringGrid; Acol: Integer; rect: Trect);
    procedure CalculEffectifs;
    function GetTOBLigne(TOBDet: TOB; Arow: integer): TOB;
    procedure BouttonVisible(Arow: integer; Mode: boolean);
    procedure Calculreste(TOBLL: TOB);
    function ConstitueDatesNat(Datedebut: TdateTime; Detaille : boolean=false): string;
    procedure SetFermeture(var JF: array of integer);
    function IsDayClose(LeItem: TheColonne; JF: array of integer): boolean;
    procedure ReinitDatas;
    procedure Raye(Grid: THgrid; Canvas: TCanvas; ACol, ARow: integer);
    procedure SetActionMenu (Arow : Integer);
    function ConstitueRequete (Chantier,NatureP : string; Detaille : boolean) : Widestring;
    procedure ConstitueTOBPlannChantier;
    procedure AddOneDETAIL(TOBL: TOB);
    procedure PositionneChamps(TOBL, TT: TOB);
    procedure AddChampsPlanningCha(TOBL: TOB; Level,NumPhase,NumOrdre: integer; LibellePhase: string;Bcolor,Color : Tcolor);
    function ConstitueConsoDetPlann(Datedebut: TdateTime): WideString;
    function GetLibChantier(CodeChantier: string): string;
    procedure AffecteHrsRess(ColDeb, ARow, ColFin : integer);
    procedure AjouteAffectation(TOBL,TOBP: TOB; Acol: integer);
    procedure SupprimeAffectation(TOBL, TOBP, TOBI: TOB; Acol: integer);
    procedure CumuleTOB(TOBL,TOBP: TOB; Acol: Integer; NbHrs: Double;sens: string='');
    function FindLaDateinTOB(LesDate: TOB; LaDate: TDateTime): TOB;
    procedure AnnuleSaisie;
    procedure CumuleTOBLigne(TA, TOBC: TOB; Sens: string = '');
    procedure ValidelaSaisie;
    procedure PositionneToutValide;
    procedure RefreshChantier(TT: TOB);
    function SetListeChantiers (Prefixe : string = 'GL') : Widestring;
    function SetListeTypeRessources: string;
    procedure RBmodeAffichageClick(Sender: TObject);
    function GetDureeSt(DD: Double): string;
    procedure  ConstitueCumulFonction (Fonction : string);
    function ConstitueInfoFonction(Fonction: string): string;
    function ConstitueCumFonctionJour(Fonction: string): WideString;
    procedure AddOneTotalFonction;
    procedure ShowThisHint(var HintStr: String; var CanShow: Boolean;
      var HintInfo: THintInfo);
    function GetDureeDD(DD: Double): double;
    function ConstituePlannifieCh(Datedebut: TdateTime): string;
    procedure AjoutePlannifieChantier (TOBAdd : TOB = nil; ChantierF : string = '';FonctionF : string='');
    procedure ReloadPlannifChantier(Chantier: string);
    procedure G2HauteurLigne(Index: integer);

  public
    { Déclarations publiques }
    ModeGestion : TModeGestionPC;
    Chantier : String;
    DateDepart : TDateTime;
    DateFin : TDateTime;
    ListeChantiers : TOB;
    ModeAffichage : TmodeAffichage;
    Etablissement : string;
  end;
Const
  Shortdn : array [1..7] of string = ('Lun.','Mar.','Mer.','Jeu.','Ven.','Sam','Dim.');
  Shortmn : array [1..12] of string = ('Jan.','Fev.','Mar.','Avril','Mai','Juin','Juil.','Aout','Sept.','Oct.','Nov.','Dec.');
  TXT_MARG: TPoint = (x: 2; y: 2);

procedure OpenPlanningBtp ( TheModeGestion : TModeGestionPC; TheListChantier : TOB = nil ; Chantier : string = '' ; Etablissement : string = ''; Datedepart: TDateTime = 0;DateFin : TDateTime=0; ModeAffichage:TmodeAffichage=TmaJours);

implementation
uses Variants,UAFO_Ferie,CalcOLEGenericBTP,Ulog, DB,BTPUtil,Entgc,UPrintScreen,BTPLANNING
;

{$R *.dfm}

procedure OpenPlanningBtp ( TheModeGestion : TModeGestionPC; TheListChantier : TOB = nil; Chantier : string = '' ; Etablissement : string = ''; Datedepart: TDateTime = 0;DateFin : TDateTime=0; ModeAffichage:TmodeAffichage=TmaJours);
var XX : TFplanningBTP;
begin
  if (not VH_GC.SeriaPlanCharge) and (not V_PGI.VersionDemo) then
  begin
    PGIInfo('Ce module n''est pas sérialisé. Il ne peut être utilisé.','Plan de charge chantiers');
    exit;
  end;
  XX := TFplanningBTP.Create(Application);
  XX.Modegestion := TheModeGestion;
  XX.Chantier := Chantier;
  XX.DateDepart := DateDepart;
  XX.DateFin := DateFin;
  XX.ListeChantiers := TheListChantier;
  XX.ModeAffichage := ModeAffichage;
  XX.Etablissement  := Etablissement;
  if TheListChantier <> nil then XX.Etablissement := TheListChantier.GetString('ETABLISSEMENT');
  try
    XX.ShowModal;
  finally
    XX.Free;
  end;
end;

procedure TFplanningBTP.BFermeClick(Sender: TObject);
begin
  close;
end;

procedure TFplanningBTP.CreateTobs;
begin
  TobAffiches := TOB.create ('LES CHANTIERS',nil,-1);
  TOBAffects := TOB.Create ('LES AFFECTATIONS JOUR',nil,-1);
  TOBAffectsDel := TOB.Create ('LES AFFECTATIONS SUPP',nil,-1);
  TOBTOTJours := TOB.Create('LES TOTAUX',nil,-1);
  TOBTOTFONC := TOB.Create('LES TOTAUX FONC',nil,-1);
end;


Function TFplanningBTP.IsDayClose(LeItem: TheColonne; JF:array of integer):boolean;
var QuelJour : integer;
    laDate : Tdatetime;
begin
  Result:=False;
  QuelJour := LeItem.JourSemaine;
  LaDate := LeItem.TheDate;
  If (JF[QuelJour-1]=-1) then BEGIN Result:=True; Exit; END;
  if isJourFerie (LaDate) then BEGIN Result:=True; Exit; END;
end;


Procedure TFplanningBTP.SetFermeture (Var JF:array of integer);
Var i:integer;
begin
For i:=0 To 5 do JF[i]:=GetParamSoc('SO_JOURFERMETURE'+IntToStr(i+1));
JF[6]:=GetParamSoc('SO_JOURFERMETURE'); // dimanche
end;


procedure TFplanningBTP.FormCreate(Sender: TObject);
begin
  self.WindowState := wsMaximized;
  ModeGestionGrid := TmgConsult;
  DureeJournee := UdateUtils.DureeJour / 60;
  LastPos := ThePos.Create;
  SetFermeture(JourFermeture);
  CreateTobs;
  //
  LesColonnesG1 := TListColonnes.Create;
  LesColonnesG2 := TListColonnes.create;
  MergedSem := TlistMerge.create;
  MergedMois := TlistMerge.create;
  Self.ShowHint := true;
  Application.OnShowHint := ShowThisHint;
end;

procedure TFplanningBTP.FreeTobs;
begin
  TobAffiches.free;
  TOBAffects.free;
  TOBAffectsDel.free;
  TOBTOTJours.free;
  TOBTOTFONC.free;
end;

function TFplanningBTP.ConstitueCumJour: string;
var II : Integer;
    DDay,FDay : tdatetime;
    SQl : string;
begin
  //
  DDay := DateDepart;
  //
  For II := 1 to NBjours do
  begin
    DDay := StrToDateTime(DateToStr(DDay)+' 00:00:00');
    FDay := StrToDateTime(DateToStr(DDay)+' 23:59:59');
    //
    SQL := '(SELECT SUM(BAT_NBHRS) '+
           'FROM BAFFECTCHANT '+
           'WHERE '+
           'BAT_DATE >= "'+USDATETIME(DDay)+'" AND '+
           'BAT_DATE <= "'+USDATETIME(FDay)+'" ';
    if Chantier <> '' then
    begin
      SQL := SQL + ' AND BAT_AFFAIRE="'+Chantier+'"';
    end;
    SQL := SQL + ') ';

    if II = 1 then Result := SQL+' AS TJ'+InttoStr(II)
              else Result := Result + ','+SQL+' AS TJ'+InttoStr(II);
    DDay := IncDay(DDay,1);
  end;
end;


function TFplanningBTP.ConstitueCumFonctionJour (Fonction : string) : WideString;
var II : Integer;
    DDay,FDay : tdatetime;
    SQl : string;
begin
  //
  DDay := DateDepart;
  //
  For II := 1 to NBjours do
  begin
    DDay := StrToDateTime(DateToStr(DDay)+' 00:00:00');
    FDay := StrToDateTime(DateToStr(DDay)+' 23:59:59');
    //
    SQL := '(SELECT SUM(BAT_NBHRS) '+
           'FROM BAFFECTCHANT '+
           'WHERE '+
           'BAT_FONCTION="'+Fonction+'" AND '+
           'BAT_DATE >= "'+USDATETIME(DDay)+'" AND '+
           'BAT_DATE <= "'+USDATETIME(FDay)+'" ';
    SQL := SQL + ') ';

    if II = 1 then Result := SQL+' AS TJ'+InttoStr(II)
              else Result := Result + ','+SQL+' AS TJ'+InttoStr(II);
    DDay := IncDay(DDay,1);
  end;
end;

procedure TFplanningBTP.AddOneTotal ;
var II : Integer;
    TOBTT : TOB;
begin
  TOBTT := TOB.Create ('UN TOT',TOBTOTJours,-1);
  For II := 1 to NBjours do
  begin
    TOBTT.AddChampSupValeur('TJ'+InttoStr(II),0);
  end;
end;

procedure TFplanningBTP.AddOneTotalFonction ;
var II : Integer;
    TOBTT : TOB;
begin
  TOBTT := TOB.Create ('UN TOT',TOBTOTFONC,-1);
  For II := 1 to NBjours do
  begin
    TOBTT.AddChampSupValeur('PJ'+InttoStr(II),0); // plannifiable
    TOBTT.AddChampSupValeur('CJ'+InttoStr(II),0); // consommé (plannifié)
    TOBTT.AddChampSupValeur('RJ'+InttoStr(II),0); // Reste
    TOBTT.AddChampSupValeur('ABS'+InttoStr(II),0); // Absence
    TOBTT.AddChampSupValeur('GRC'+InttoStr(II),0); // GRC
    TOBTT.AddChampSupValeur('INT'+InttoStr(II),0);
    TOBTT.AddChampSupValeur('MAT'+InttoStr(II),0);
    TOBTT.AddChampSupValeur('CHA'+InttoStr(II),0);
  end;
end;

procedure TFplanningBTP.AddOneDETAIL (TOBL : TOB) ;
var II : Integer;
begin
  For II := 1 to NBjours do
  begin
    TOBL.AddChampSup('D'+InttoStr(II),false);
    TOBL.AddChampSup('PLA'+InttoStr(II),false);
  end;
end;

procedure TFplanningBTP.CalculEffectifs;
var II : Integer;
    Zone : string;
    result : Double;
begin
  if TOBTOTJours.detail.count = 0 then exit;
  For II := 1 to NBjours do
  begin
    Zone := 'TJ'+InttoStr(II);
    result := arrondi(TOBTOTJours.detail[0].Getdouble(Zone) /  DureeJournee,0);
    if result = 0 then GTEffectifs.Cells[II-1,0] := ''
                  else GTEffectifs.Cells[II-1,0] := FloattoStr(result);
  end;
end;

procedure TFplanningBTP.Calculreste (TOBLL : TOB);
begin
  TOBLL.SetDouble('RESTE',TOBLL.GetDouble('BUDGET')-TOBLL.GetDouble('DEJAPLANN'));
end;


procedure TFplanningBTP.LoadLesTobs;
var QQ : TQuery;
    SQL : string;
    II : Integer;
begin
  if Modegestion = TmPcPlanCharge then
  begin
    SQl := ConstitueRequete('','',false);
    QQ := OpenSQL(SQL,True,-1,'',true);
    if Not QQ.eof then
    begin
      TobAffiches.LoadDetailDB('UN DETAIL','','',QQ,false);
      for II := 0 to TOBAffiches.detail.count -1 do
      begin
    	  if (VarIsNull(TOBAffiches.detail[II].GetValue('DEJAPLANN')) or
           (VarAsType(TOBAffiches.detail[II].GetValue('DEJAPLANN'), varString) = #0 )) then
        begin
          TOBAffiches.detail[II].SetDouble('DEJAPLANN',0);
        end;
        Calculreste (TOBAffiches.detail[II]);
      end;
    end;
    Ferme(QQ);

    (*
    AddLignes(100);
    RemplitGrid;
    *)

  end else if Modegestion = TmcPlanCharfnc then
  begin
    SQl := ConstitueRequete('','',false);
    QQ := OpenSQL(SQL,True,-1,'',true);
    if Not QQ.eof then
    begin
      TobAffiches.LoadDetailDB('UN DETAIL','','',QQ,false);
      for II := 0 to TOBAffiches.detail.count -1 do
      begin
    	  if (VarIsNull(TOBAffiches.detail[II].GetValue('DEJAPLANN')) or
           (VarAsType(TOBAffiches.detail[II].GetValue('DEJAPLANN'), varString) = #0 )) then
        begin
          TOBAffiches.detail[II].SetDouble('DEJAPLANN',0);
        end;
        Calculreste (TOBAffiches.detail[II]);
      end;
    end;
    Ferme(QQ);
  end else if (Modegestion = TmPcPlanningCha) then
  begin
    ConstitueTOBPlannChantier;
  end;
  //
  SQL := 'SELECT '+ConstitueCumJour;
  QQ := OpenSQL(SQL,True,-1,'',true);
  if Not QQ.eof then
  begin
    TOBTOTJours.LoadDetailDB('UN TOTAL','','',QQ,false);
  end else
  begin
    AddOneTotal;
  end;

  AjoutePlannifieChantier;
end;

function TFplanningBTP.ConstitueStringG2 : string;
var II : Integer;
begin
  Result := '';
  //
  For II := 1 to NBjours do
  begin
    if II = 1 then Result := 'D'+InttoStr(II)
    else Result := Result + ';D'+InttoStr(II);
  end;
  Result := Result + ';';
end;


procedure TFplanningBTP.SeteventsGrid (Enabled : Boolean);
begin
  if Enabled then
  begin
    G1.OnDblClick  := BOuvrirFermerClick;
    G1.OnTopLeftChanged := G1TopLeftChanged;
    G1.OnRowEnter := G1RowEnter;
    G1.PostDrawCell :=  G1PostDrawCell;
    G1.GetCellCanvas := G1GetCellCanvas;
    //
    G2.OnTopLeftChanged := G2TopLeftChanged;
    G2.GetCellCanvas := G2GetCellCanvas;
    G2.OnCellEnter := G2Cellenter;
    G2.onRowEnter := G2RowEnter;
    G2.OnMouseDown :=  G2MouseDown;
    G2.PostDrawCell := G2PostDrawCell;
    //
    GSem.OnDrawCell := GSemDrawCell;
    GMois.OnDrawCell := GMoisDrawCell;
    GDate.OnDrawCell := GDateDrawCell;
    GTHeures.OnDrawCell := GTheuresDrawCell;
    GTEffectifs.OnDrawCell := GTEffectifsDrawCell;
    GFONCTION.onDrawCell := GFONCTIONDrawCell;
    GFONCTION.OnCellEnter := GFONCTIONCellenter;
    RBHeures.OnClick := RBmodeAffichageClick;
    RbJours.OnClick :=  RBmodeAffichageClick;

  end else
  begin
    G1.OnDblClick := nil;
    G1.OnTopLeftChanged := nil;
    G1.OnRowEnter := nil;
    G1.GetCellCanvas := nil;
    //
    G2.OnTopLeftChanged := nil;
    G2.GetCellCanvas := nil;
    G2.OnCellEnter := nil;
    G2.onRowEnter := nil;
    G2.OnMouseDown := nil;
    G2.PostDrawCell := nil;
    //
    GSem.OnDrawCell := nil;
    GMois.OnDrawCell := nil;
    GDate.OnDrawCell := nil;
    GTHeures.OnDrawCell := nil;
    GTEffectifs.OnDrawCell := nil;
    GFONCTION.onDrawCell := nil;
    G1.PostDrawCell := nil;
    RBHeures.OnClick := nil;
    RbJours.OnClick :=  nil;
  end;
end;

procedure TFplanningBTP.PrepareAffichage;
var II : integer;
begin
  //
  SeteventsGrid (false);
  //
  G1.RowCount := TobAffiches.detail.count;
  G2.RowCount := TobAffiches.detail.count;
  //
  ConstitueGrilleTH;
  CalculEffectifs;
  AddOneTotalFonction;
  //
  LesColsG2 := ConstitueStringG2;
  ConstitueGrilleG2;
  DefinilesSemaines;
  DefinilesMois;
  //
  ConstitueGrilleG1;
  ReajusteEcran;
  SeteventsGrid (True);
  G2.Invalidate;
  G2.col := 0;
  GDate.Invalidate;
  GSem.Invalidate;
  GMois.Invalidate;
//  G1RowEnter(self,G1.row,cancel,false); 
end;

procedure TFplanningBTP.FormDestroy(Sender: TObject);
begin
  FreeTobs;
  LesColonnesG1.free;
  LesColonnesG2.free;
  MergedSem.Free;
  MergedMois.free;
  LastPos.Free;
  Application.OnShowHint := nil;
end;

function TFplanningBTP.ConstituePlannifieCh (Datedebut: TdateTime): string;
var II : Integer;
    DDay : tdatetime;
begin
  //
  DDay := Datedebut;
  //
  For II := 1 to NBjours do
  begin
    DDay := StrToDateTime(DateToStr(DDay)+' 00:00:00');
    //
    if II = 1 then Result := '0 AS PLA'+InttoStr(II)
              else Result := Result + ',0 AS PLA'+InttoStr(II);
    //
    DDay := IncDay(DDay,1);
  end;
end;

function TFplanningBTP.ConstitueDatesCha(Datedebut: TdateTime;detaille : boolean=false): string;
var II : Integer;
    DDay,FDay : tdatetime;
    SQl : string;
begin
  //
  DDay := Datedebut;
  //
  For II := 1 to NBjours do
  begin
    DDay := StrToDateTime(DateToStr(DDay)+' 00:00:00');
    FDay := StrToDateTime(DateToStr(DDay)+' 23:59:59');
    //
    SQL := '(SELECT SUM(BAT_NBHRS) '+
           'FROM BAFFECTCHANT '+
           'WHERE '+
           'BAT_DATE >= "'+USDATETIME(DDay)+'" AND '+
           'BAT_DATE <= "'+USDATETIME(FDay)+'" AND '+
           'BAT_AFFAIRE=GL_AFFAIRE';
    if Detaille then
    begin
      Sql := Sql + ' AND BAT_FONCTION=FONCTION';
    end;
    Sql := Sql +')';
    if II = 1 then Result := SQL+' AS D'+InttoStr(II)
              else Result := Result + ','+SQL+' AS D'+InttoStr(II);
    //
    DDay := IncDay(DDay,1);
  end;
end;


function TFplanningBTP.ConstitueConsoDetPlann(Datedebut: TdateTime): Widestring;
var II : Integer;
    DDay,FDay : tdatetime;
    SQl : string;
begin
  //
  DDay := Datedebut;
  //
  For II := 1 to NBjours do
  begin
    DDay := StrToDateTime(DateToStr(DDay)+' 00:00:00');
    FDay := StrToDateTime(DateToStr(DDay)+' 23:59:59');
    //
    SQL := '('+
           'SELECT SUM(BAT_NBHRS) '+
           'FROM BAFFECTCHANT '+
           'WHERE '+
           'BAT_DATE >= "'+USDATETIME(DDay)+'" AND '+
           'BAT_DATE <= "'+USDATETIME(FDay)+'" AND '+
           'BAT_AFFAIRE=B1.BAT_AFFAIRE AND BAT_FONCTION=B1.BAT_FONCTION AND BAT_NUMORDRE=B1.BAT_NUMORDRE'+
            ')';
    if II = 1 then Result := SQL+' AS D'+InttoStr(II)
              else Result := Result + ','+SQL+' AS D'+InttoStr(II);
    DDay := IncDay(DDay,1);
  end;
end;


function TFplanningBTP.ConstitueDatesNat(Datedebut: TdateTime; Detaille : boolean=false): string;
var II : Integer;
    DDay,FDay : tdatetime;
    SQl : string;
begin
  //
  DDay := Datedebut;
  //
  For II := 1 to NBjours do
  begin
    DDay := StrToDateTime(DateToStr(DDay)+' 00:00:00');
    FDay := StrToDateTime(DateToStr(DDay)+' 23:59:59');
    //
    SQL := '(SELECT SUM(BAT_NBHRS) '+
           'FROM BAFFECTCHANT '+
           'WHERE '+
           'BAT_DATE >= "'+USDATETIME(DDay)+'" AND '+
           'BAT_DATE <= "'+USDATETIME(FDay)+'" AND '+
           'BAT_FONCTION=FONCTION';
    if Detaille then
    begin
      Sql := Sql + ' AND BAT_AFFAIRE=GL_AFFAIRE';
    end;
    Sql := Sql +')';
    if II = 1 then Result := SQL+' AS D'+InttoStr(II)
              else Result := Result + ','+SQL+' AS D'+InttoStr(II);
    DDay := IncDay(DDay,1);
  end;
end;

function TFplanningBTP.GetLibChantier (CodeChantier : string) : string;
var QQ : TQuery;
begin
  Result := '';
  QQ := OpenSQL('SELECT AFF_LIBELLE FROM AFFAIRE WHERE AFF_AFFAIRE="'+CodeChantier+'"',True,1,'',true);
  if not QQ.eof then Result := QQ.fields[0].AsString;
  ferme (QQ);
end;

procedure TFplanningBTP.FormShow(Sender: TObject);
var cancel : boolean;
begin
  if ModeAffichage = TmAHeure then RBHeures.Checked := true else RbJours.checked := true;
  if ModeAffichage = TmAHeure then
  begin
    TLIBTOTALHRS.Caption := 'Total des heures';
    TLIBHRSPREV.Caption := 'Heures Prévues';
  end else
  begin
    TLIBHRSPREV.Caption := 'Jours Prévus';
    TLIBTOTALHRS.Caption := 'Total des jours';
  end;
  TLIBTOTALHRS.left := GTHeures.Left - TLIBTOTALHRS.Width - 5;

  if (Modegestion = TmcPlanCharfnc) then Caption := 'Plan de charge / Fonction';
  if (Modegestion = TmPcPlanCharge) then Caption := 'Plan de charge / Chantier';
  if (Modegestion = TmPcPlanningCha) then
  begin
    Caption := 'Planning du chantier ('+BTPCodeAffaireAffiche (chantier)+') '+GetLibChantier(Chantier);
  end;
  //
  NBRRESS.Value := 1;
  NBHrs.value := ARRONDI(NBRRESS.value * DureeJournee, V_PGI.OkDecQ);
  //
  LastPos.Acol := -1;
  LastPos.Arow := -1;
  if DateDepart = 0 then
  begin
    DateDepart :=  Now;
    DateFin := IncDay (DateDepart,20);
  end;
  //
  DateDepart := StrToDate(DateToStr(DateDepart));
  DateFin := StrToDateTime(DateToStr(DateFin)+' 23:59:59');
  //
  DATED.Text := DateToStr(DateDepart);
  DATED_.Text := DateToStr(DateFin);
  //
  NbJours :=  DaysBetween (DateDepart,DateFin)+1;
  //
  ReinitDatas;
  LoadLesTobs;
  if TobAffiches.detail.count = 0 then
  begin
    PGIInfo ('Aucune donnée existante pour la sélection');
  end;
  if TobAffiches.detail.count > 0 then
  begin
    PrepareAffichage;
    G1RowEnter(Self,0,cancel,false);
    GFONCTION.visible := false;
    BValider.Visible := false;
    LNBRRESS.visible := false;
    NBRRESS.visible := false;
    LNBHRS.visible := false;
    NBHRS.visible := false;

    if (Modegestion = TmcPlanCharfnc) or (Modegestion = TmPcPlanCharge) then
    begin
      BMODIF.visible := false;
    end;
    if (Modegestion = TmPcPlanningCha) then
    begin
      BMODIF.VISIBLE := true;
      TBVIEWARBO.visible := true;
    end;
  end else
  begin
    PostMessage(Self.Handle,WM_CLOSE,0,0);
  end;
end;

function TFplanningBTP.GetColCounts (LaChaine : string) : Integer;
var SS : string;
    TheC : string;
begin
  Result := 0;
  TheC := LaChaine;
  repeat
    SS := READTOKENST(TheC);
    if SS <> '' then Inc(Result);
  until SS ='';
end;

procedure TFplanningBTP.G2HauteurLigne (Index : integer);
var II : integer;
    TT : TOB;
    RR : integer;
begin
  TT := TobAffiches.detail[Index]; if TT = nil then exit;
  RR := 18;
  For II := 1 to NBjours do
  begin
    if TT.GetDouble('PLA'+InttoStr(II)) <> 0 then
    begin
      RR := 28;
      break;
    end;
  end;
  G2.RowHeights [Index] := RR;
  G1.RowHeights [Index] := G2.RowHeights [Index];
end;

procedure TFplanningBTP.ConstitueGrilleG2;
var st,FF      : string;
    i       : Integer;
    UneDate : TheColonne;
    DDay : TDateTime;
begin
  FF:='#';
  if V_PGI.OkDecQ>0 then
  begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecQ-1 do
    begin
      FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;
  //
//  G2.DefaultRowHeight := 18;
  //
  G2.ColCount := NbJours;  
  GDate.ColCount := G2.ColCount;
  GSem.ColCount := GDate.ColCount;
  GMois.ColCount := GDate.ColCount;
  GTHeures.ColCount := GDate.ColCount;
  GTEffectifs.ColCount := GDate.ColCount;
  GFONCTION.colcount := GDate.colCount;
  //
  St := LesColsG2;
  DDay := DateDepart;
  //
  for i := 0 to G2.ColCount - 1 do
  begin
//    G2.ColWidths[i] := 45;
    G2.ColWidths[i] := 45;
    G2.ColAligns[i] := taCenter;
    G2.ColFormats[i] := FF+';-'+FF+'; ;';
    //
    UneDate := TheColonne.Create;
    UneDate.Nom := 'D'+InttoStr(I+1);
    UneDate.Libelle := Shortdn[DayOfTheWeek(DDay)]+' '+ IntToStr(DayOf(DDay));
    UneDate.TheDate := DDay;
    UneDate.semaine := NumSemaine(DDay);
    UneDate.Mois := NumMois(DDay);
    UneDate.JourSemaine := DayOfTheWeek(DDay);
    lesColonnesG2.Add(UneDate);
    //
    GDate.Cells[I,0] := UneDate.Libelle;
    GSem.Cells[I,0] := 'Sem. '+InttoStr(UneDate.semaine);
    Gmois.Cells[I,0] := Shortmn[ UneDate.Mois ];
    //
    DDay := IncDay(DDay,1);
  end;
  //
  TobAffiches.PutGridDetail(G2,false,false,LesColsG2);
end;

procedure TFplanningBTP.ConstitueGrilleG1;
var st,FF      : string;
    Nam     : String;
    i      : Integer;
    UneColonne : TheColonne;
begin
  if Modegestion = TmPcPlanCharge then
  begin
    LesColsG1 := 'INDIC;AFFAIRE;LIBELLE;BUDGET;RESTE;';
  end else if Modegestion = TmcPlanCharfnc then
  begin
    LesColsG1 := 'INDIC;LIBNATURE;BUDGET;RESTE;';
  end else if Modegestion = TmPcPlanningCha then
  begin
    LesColsG1 := 'INDIC;LIBNATURE;BUDGET;RESTE;';
  end;
  FF:='#';
  if V_PGI.OkDecQ>0 then
  begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecQ-1 do
    begin
      FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;
  //
  G1.DefaultRowHeight := 18;
  //
  G1.ColCount := GetColCounts(LesColsG1);
  //
  St := LesColsG1;
  //
  for i := 0 to G1.ColCount - 1 do
  begin
    Nam := ReadTokenSt(St);
    UneColonne := TheColonne.Create;

    if Nam = 'INDIC' then
    begin
      G1.ColWidths[i] := 20;
      G1.ColAligns[i] := taCenter;
      //
      UneColonne.Nom := 'INDIC';
      UneColonne.Indice := i;
      LesColonnesG1.Add(UneColonne);
    end else if Nam = 'AFFAIRE' then
    begin
      G1.ColWidths[i] := 100;
      G1.ColAligns[i] := taLeftJustify;
      //
      UneColonne.Nom := 'AFFAIRE';
      UneColonne.Indice := i;
      LesColonnesG1.Add(UneColonne);
    end else if Nam = 'LIBELLE' then
    begin
      G1.ColWidths[i] := 270;
      G1.ColAligns[i] := taLeftJustify;
      //
      UneColonne.Nom := 'LIBELLE';
      UneColonne.Indice := i;
      LesColonnesG1.Add(UneColonne);
    end else if Nam = 'LIBNATURE' then
    begin
      G1.ColWidths[i] := 320;
      G1.ColAligns[i] := taLeftJustify;
      //
      UneColonne.Nom := 'LIBNATURE';
      UneColonne.Indice := i;
      LesColonnesG1.Add(UneColonne);
    end else if Nam = 'BUDGET' then
    begin
      G1.ColWidths[i] := 80;
      G1.ColAligns[i] := taRightJustify;
      G1.ColFormats[i] := FF+';-'+FF+'; ;';
      //
      UneColonne.Nom := 'BUDGET';
      UneColonne.Indice := i;
      LesColonnesG1.Add(UneColonne);
    end else if Nam = 'RESTE' then
    begin
      G1.ColWidths[i] := 80;
      G1.ColAligns[i] := taRightJustify;
      G1.ColFormats[i] := FF+';-'+FF+'; ;';
      //
      UneColonne.Nom := 'RESTE';
      UneColonne.Indice := i;
      LesColonnesG1.Add(UneColonne);
    end;

  end;
  //
  TobAffiches.PutGridDetail(G1,false,false,LesColsG1);
end;


function TFplanningBTP.GetWidth (Grid : Thgrid) : integer;
var ii : Integer;
begin
  Result := 0;
  for II := 0 to Grid.ColCount - 1 do
  begin
    result := result + Grid.ColWidths[II] + Grid.GridLineWidth;
  end;
end;

procedure TFplanningBTP.SetDates (GDate : Thgrid);
var I : Integer;
begin
  for I := 0 to lesColonnesG2.count -1 do
  begin
  end;
end;


procedure TFplanningBTP.ReajusteEcran;
var II : integer;
begin
  G1.BeginUpdate;
  G2.BeginUpdate;
  GDate.BeginUpdate;
  GTHeures.BeginUpdate;
  GTEffectifs.BeginUpdate;
  GSem.BeginUpdate;
  Gmois.BeginUpdate;
  if TobAffiches.detail.count = 0 then exit;
  // positionnement des grilles -- l'une par rapport a l'autre
  G1.top := PCenter.top;
  G1.left := PCenter.left;
  G1.Width := GetWidth(G1) + 10;
  G2.Left := G1.left+G1.Width;
  G2.Width := PCENTER.Width - G2.Left;
  G2.height := PCenter.ClientHeight;
  G1.height := G2.ClientHeight+ (2 * G1.GridLineWidth);
  //
  for II := 0 to TOBAffiches.detail.count -1 do
  begin
    G2HauteurLigne (II);
  end;
  //
  if Modegestion = TmPcPlanCharge then
  begin
    PositionneTexteEntete (TLIBAFFAIRE,'AFFAIRE',LesColonnesG1,G1);
    PositionneTexteEntete (TLIBELLE,'LIBELLE',LesColonnesG1,G1);
    PositionneTexteEntete (TLIBHRSPREV,'BUDGET',LesColonnesG1,G1);
    PositionneTexteEntete (TLIBRESTEAPLANNIF,'RESTE',LesColonnesG1,G1);
  end else
  begin
    TLIBAFFAIRE.Visible := false;
    PositionneTexteEntete (TLIBELLE,'LIBNATURE',LesColonnesG1,G1);
    PositionneTexteEntete (TLIBHRSPREV,'BUDGET',LesColonnesG1,G1);
    PositionneTexteEntete (TLIBRESTEAPLANNIF,'RESTE',LesColonnesG1,G1);
  end;
  //
  GDate.Enabled := false;
  GDate.CacheEdit;
  GDate.Left := G2.Left+(2 * G2.GridLineWidth);
  GDate.Width := G2.Width;
  For II := 0 to G2.ColCount -1 do
  begin
   GDate.ColWidths[II]  := G2.ColWidths [II];
  end;
  SetDates (GDate);
  //
  GTHeures.Left := GDate.Left;
  GTHeures.Width := GDate.Width;
  For II := 0 to G2.ColCount -1 do
  begin
   GTHeures.ColWidths[II]  := G2.ColWidths [II];
  end;
  TLIBTOTALHRS.left := GTHeures.Left - TLIBTOTALHRS.Width - 5;
  //
  GTEffectifs.Left := GDate.Left;
  GTEffectifs.Width := GDate.Width;
  //
  GFONCTION.left := GDate.left;
  GFONCTION.width := GDate.width;
  //
  LPOTENTIEL.left := GFONCTION.Left - LPOTENTIEL.Width - 5;
  LAFFECTE.left := GFONCTION.Left - LAFFECTE.Width - 5;
  LRESTE.left := GFONCTION.Left - LRESTE.Width - 5;

  For II := 0 to G2.ColCount -1 do
  begin
   GTEffectifs.ColWidths[II]  := G2.ColWidths [II];
   GFONCTION.ColWidths[II]  := G2.ColWidths [II];
  end;
  TLibEffectif.left := GTHeures.Left - TLibEffectif.Width - 5;
  //
  GSem.Left := GDate.Left;
  GSem.Width := GDate.Width;
  For II := 0 to G2.ColCount -1 do
  begin
    GSem.ColWidths[II]  := G2.ColWidths [II];
  end;
  //
  GMois.Left := GDate.Left;
  GMois.Width := GDate.Width;
  For II := 0 to G2.ColCount -1 do
  begin
    GMois.ColWidths[II]  := G2.ColWidths [II];
  end;
  //
  G1.EndUpdate;
  G2.EndUpdate;
  GDate.EndUpdate;
  GTHeures.EndUpdate;
  GTEffectifs.EndUpdate;
  GSem.EndUpdate;
  Gmois.EndUpdate;
end;

procedure TFplanningBTP.PositionneTexteEntete (TheControl : THlabel; TheChamps : string;Lescolonnes : TListColonnes; TheGrid : THGrid);
var OO : TheColonne;
    Larg,PosG : Integer;
begin
  if TheControl = nil then Exit;
  OO := Lescolonnes.Find(TheChamps);
  if OO <> nil then
  begin
    Larg := TheGrid.CellRect(OO.Indice,0).Right - TheGrid.CellRect(OO.Indice,0).left;
    PosG := TheGrid.CellRect(OO.Indice,0).Left + ((larg - TheControl.Width) div 2) ;
    TheControl.Left := PosG;
  end;
end;

{ TListColonnes }

function TListColonnes.Add(AObject: TheColonne): Integer;
begin
  result := Inherited Add(AObject);
end;

procedure TListColonnes.clear;
var indice : integer;
begin
  for Indice := 0 to Count -1 do
  begin
    TheColonne(Items [Indice]).free;
  end;
  inherited;
end;

destructor TListColonnes.destroy;
begin
  clear;
  inherited;
end;

function TListColonnes.Find(NomColonne: string): TheColonne;
var II : Integer;
    OO : TheColonne;
begin
  Result := nil;
  for II := 0 to Count -1 do
  begin
    OO := Items [II];
    if OO.Nom = NomColonne then
    begin
      result := OO;
      break;
    end;
  end;
end;

function TListColonnes.FindIndice(TheDate: TDateTime): Integer;
var II : Integer;
    OO : TheColonne;
begin
  result := -1; 
  for II := 0 to Count -1 do
  begin
    OO := Items[II];
    if OO.TheDate = TheDate then
    begin
      Result := II;
      Break;
    end;
  end;
end;

function TListColonnes.GetItems(Indice: integer): TheColonne;
begin
  result := TheColonne (Inherited Items[Indice]);
end;

procedure TListColonnes.SetItems(Indice: integer; const Value: TheColonne);
begin
  Inherited Items[Indice]:= Value;
end;

procedure TFplanningBTP.G1TopLeftChanged(Sender: TObject);
begin
  SeteventsGrid (False);
  G2.TopRow  := G1.TopRow;
  SeteventsGrid (true);
  G2.Invalidate;
end;

procedure TFplanningBTP.G2TopLeftChanged(sender: TObject);
begin
  SeteventsGrid(false);
  G1.TopRow  := G2.TopRow;
  GDate.LeftCol  := G2.leftCol;
  GSem.LeftCol  := G2.leftCol;
  GMois.LeftCol  := G2.leftCol;
  GTHeures.LeftCol := G2.leftCol;
  GTEffectifs.LeftCol := G2.leftCol;
  GFONCTION.leftCol := G2.leftCol;
  SeteventsGrid (true);
  G2.invalidate;
  G1.Invalidate;
  GDate.invalidate;
  GSem.Invalidate;
  Gmois.Invalidate;
  GTHeures.Invalidate;
  GTEffectifs.Invalidate;
  GFONCTION.invalidate;
end;

procedure TFplanningBTP.G1RowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var TT : TOB;
    BAffichePotentiel : boolean;
begin
  BAffichePotentiel := false;
  TT := GetTOBLigne(TobAffiches,Ou);
  SeteventsGrid (False);
  G2.Row := G1.row;
  GFONCTION.Visible := false;
  LPOTENTIEL.visible := false;
  LAFFECTE.visible := false;
  LRESTE.visible := false;
  LFONCTION.visible := false;
  if Modegestion <> TmPcPlanningCha then
  begin
    BouttonVisible (Ou,false);
    BouttonVisible (G1.row,True);
    if Modegestion = TmPcPlanCharge then
    begin
      if (TT.GetInteger('LEVEL') > 0) and (TT.GetString('FONCTION')<>'---') then BAffichePotentiel := true;
    end else if (Modegestion = TmcPlanCharfnc) and (TT.GetString('FONCTION')<>'---') then BAffichePotentiel := true;

  end else
  begin
    if (TT.GetString('FONCTION')<>'---') and (TT.GetString('FONCTION')<>'') then BAffichePotentiel := true;
  end;
  SetActionMenu (G1.Row);
  if BAffichePotentiel then
  begin
    ConstitueCumulFonction (TT.GetString('FONCTION'));
    LPOTENTIEL.visible := true;
    LAFFECTE.visible := true;
    LRESTE.visible := true;
    GFONCTION.Visible := True;
    GFONCTION.invalidate;
    LFONCTION.visible := true; 
    LFONCTION.Caption := trim(TT.getString('LIBFONCTION'))
  end;
  SeteventsGrid (true);
end;

function TFplanningBTP.GetTOBLigne (TOBDet : TOB; Arow : integer) : TOB;
begin
  result := nil;
  if Arow >= TOBdet.detail.count then exit;
  result := TOBdet.detail[Arow];
end;

procedure TFplanningBTP.BouttonVisible(Arow: integer; Mode : boolean);
var TOBL: TOB;
  Arect: Trect;
begin
  BOuvrirFermer.Opaque := false;
  BOuvrirFermer.Parent := G1;
  if Mode then
  begin
    if Arow > TOBAffiches.detail.count then exit;
    TOBL := GetTOBLigne(TobAffiches, Arow); // récupération TOB ligne
    if (TOBL <> nil)  then
    begin
      ARect := G1.CellRect(0, Arow);
      with BOuvrirFermer do
      begin
        if (TOBL.GetValue('LEVEL')=0) then
        begin
          Top := Arect.top - G1.GridLineWidth;
          Left := Arect.Left;
          Width := Arect.Right - Arect.Left;
          Height := Arect.Bottom - Arect.Top;
          Parent := G1;
          if TOBL.GetValue('OPENED')=0 then
          begin
            BOuvrirFermer.ImageIndex := 0;
          end else
          begin
            BOuvrirFermer.ImageIndex := 1;
          end;
          Visible := true;
        end;
      end;
    end;
  end else
  begin
    BOuvrirFermer.visible := false;
  end;
end;


procedure TFplanningBTP.G2RowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
  SeteventsGrid (False);
  G1.Row := G2.row;
  BouttonVisible (Ou,false);
  BouttonVisible (G1.row,True);
  SeteventsGrid (true);
  G1RowEnter(Self,G1.row,cancel,false);
end;

(*
procedure TFplanningBTP.DrawButton(Grid : TStringGrid;rect : TRect;Fin : boolean);
var
  OldBrushStyle : TBrushStyle ;
  OldPenStyle   : TPenStyle;
begin
  OldBrushStyle :=  Grid.Canvas.Brush.Style ;
  OldPenStyle   :=   Grid.Canvas.Pen.Style ;
  Grid.Canvas.Brush.Style := bsClear;
  Grid.Canvas.FillRect(rect);
  Frame3D(Grid.Canvas,Rect,clWhite,clBtnShadow,1);
//  Frame3D(Grid.Canvas,Rect,clGray,clgray,1);
  Grid.Canvas.Brush.Style := OldBrushStyle;
  Grid.Canvas.Pen.Style   := OldPenStyle ;
end;
*)

procedure TFplanningBTP.DrawButton(Grid : TStringGrid;Arect : TRect;Fin : boolean);

  procedure TireLigne(GS : TstringGrid; DebX, DebY, FinX, FinY: integer);
  begin
    GS.Canvas.MoveTo(DebX, DebY);
    GS.Canvas.LineTo(FinX, FinY);
  end;

var
  OldBrushStyle : TBrushStyle ;
  OldPenStyle   : TPenStyle;
begin
  OldBrushStyle :=  Grid.Canvas.Brush.Style ;
  OldPenStyle   :=   Grid.Canvas.Pen.Style ;
  Grid.Canvas.Brush.Style := bsClear;
  Grid.Canvas.FillRect(Arect);
  // ligne du haut
  TireLigne(Grid,ARect.left, ARect.Top, ARect.Right, ARect.top);
  // ligne du bas
  TireLigne(Grid,ARect.Left, ARect.Bottom, ARect.Right, ARect.Bottom);
  // trait gauche
  TireLigne(Grid,ARect.Left, ARect.top, ARect.Left, ARect.Bottom);
  // trait droit
  TireLigne(Grid,ARect.right, ARect.Top, ARect.Right, ARect.bottom);

  Grid.Canvas.Brush.Style := OldBrushStyle;
  Grid.Canvas.Pen.Style   := OldPenStyle ;
end;


Procedure TFplanningBTP.OwnerDrawText(Grid : TStringGrid;State: TGridDrawState;Acol : Integer; rect : Trect;Texte : String;UFormat : Cardinal;TheColor  : Tcolor=0 );
var
  OldBrushStyle : TBrushStyle ;
  OldPenStyle   : TPenStyle;
begin
  Rect.Left := Rect.left + TXT_MARG.x;
  OldBrushStyle :=  Grid.Canvas.Brush.Style ;
  OldPenStyle   :=  Grid.Canvas.Pen.Style ;

  if (gdSelected in state) then
  begin
    Grid.Canvas.font.Color := clHighlightText;
  end else
  begin
    IF theColor<> 0 then grid.canvas.font.color := TheColor
                    else Grid.Canvas.font.Color := clWindowText;
  end;
  Grid.Canvas.Brush.Style := bsClear;
  Grid.Canvas.Pen.Style := psClear;
  DrawText(Grid.Canvas.Handle,PChar(Texte), -1, Rect ,UFormat);
  Grid.Canvas.Brush.Style := OldBrushStyle;
  Grid.Canvas.Pen.Style := OldPenStyle ;

end;

Procedure TFplanningBTP.ClearCanvas(Grid : TStringGrid;Acol : Integer; rect : Trect);
var
  OldBrushStyle : TBrushStyle ;
  OldPenStyle   : TPenStyle;
begin
  Rect.Left := Rect.left + TXT_MARG.x;
  OldBrushStyle :=  Grid.Canvas.Brush.Style ;
  OldPenStyle   :=  Grid.Canvas.Pen.Style ;
  Grid.Canvas.Brush.Style := bsClear;
  Grid.Canvas.Pen.Style := psClear;
  Grid.Canvas.Brush.Style := OldBrushStyle;
  Grid.Canvas.Pen.Style := OldPenStyle ;
end;

Function TFplanningBTP.MergingCell(Sender: TObject;Coldeb,ColFin : integer;aCol,ARow : Integer;State: TGridDrawState) : boolean;
var
  txtRect: TRect;
  str : String;
  Grid : TStringGrid;
  OldRect : Trect;
  Rect: TRect;
  LastColVisi : Integer;
begin
  Grid := (Sender as TStringGrid);
  Result  := False;
  str := Grid.Cells[ColDeb, 0];
  if (Acol < ColDeb) or (Acol > ColFin) then exit;
  LastColVisi := Grid.leftcol + Grid.VisibleColCount;
  if Grid.LeftCol > Coldeb then Coldeb := Grid.LeftCol;
  if LastColVisi < Colfin then Colfin := LastColVisi;
//  if Grid.VisibleColCount > Coldeb then Coldeb := LeftCol;
  Rect.Left   := Grid.CellRect(ColDeb,0).Left;
  Rect.Top    := Grid.CellRect(ColDeb,0).Top;
  Rect.Right  := Grid.CellRect(ColFin,0).Right;
  Rect.Bottom := Grid.CellRect(ColFin,0).Bottom;
  Grid.Canvas.FillRect(Rect);

  DrawButton(Grid,Rect,true);

  str := Grid.Cells[ColDeb, 0];
  if Grid.VisibleColCount < colFin Then
  begin
    if Grid.Canvas.TextWidth(str) > (Rect.Right-Rect.Left) Then
    begin
      txtRect := OldRect;
    end;
  end;
  OwnerDrawText(Grid,State,Acol,Rect,str,DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL);
  Result := true;
end;

{ TlistMerge }

function TlistMerge.Add(AObject: TheMergeItem): Integer;
begin
  result := Inherited Add(AObject);
end;

procedure TlistMerge.clear;
var indice : integer;
begin
  for Indice := 0 to Count -1 do
  begin
    TheMergeItem(Items [Indice]).free;
  end;
  inherited;
end;

destructor TlistMerge.destroy;
begin
  clear;
  inherited;
end;

function TlistMerge.Find(Numero: integer): TheMergeItem;
var II : Integer;
    OO : TheMergeItem;
begin
  Result := nil;
  for II := 0 to Count -1 do
  begin
    OO := Items [II];
    if (Numero <= OO.Fin) and (Numero >= OO.Debut) then
    begin
      result := OO;
      break;
    end;
  end;
end;

function TlistMerge.GetItems(Indice: integer): TheMergeItem;
begin
  result := TheMergeItem (Inherited Items[Indice]);
end;

procedure TlistMerge.SetItems(Indice: integer;
  const Value: TheMergeItem);
begin
  Inherited Items[Indice]:= Value;
end;

procedure TFplanningBTP.DefinilesMois;
var II : Integer;
    MoisCourant : Integer;
    MergeItem,LastMois : TheMergeItem;
    OneColonne : TheColonne;
begin
  MoisCourant := -1;
  LastMois := nil;
  for II := 0 to lesColonnesG2.Count -1 do
  begin
    OneColonne := lesColonnesG2.Items[II];
    if OneColonne.Mois <> MoisCourant then
    begin
      if LastMois <> nil then
      begin
        LastMois.Fin := II-1;
      end;
      MergeItem := TheMergeItem.Create;
      MergeItem.Debut := II;
      MergeItem.Fin := -1;
      MergeItem.Numero := OneColonne.Mois;
      MergedMois.Add(MergeItem);
      MoisCourant := OneColonne.Mois;
      //
      LastMois := MergeItem;
    end;
  end;
  LastMois.Fin := lesColonnesG2.count -1;
end;

procedure TFplanningBTP.DefinilesSemaines;
var II : Integer;
    SemCourante : Integer;
    MergeItem,LastSemaine : TheMergeItem;
    OneColonne : TheColonne;
begin
  SemCourante := -1;
  LastSemaine := nil;
  for II := 0 to lesColonnesG2.Count -1 do
  begin
    OneColonne := lesColonnesG2.Items[II];
    if OneColonne.semaine <> SemCourante then
    begin
      if LastSemaine <> nil then
      begin
        LastSemaine.Fin := II-1;
      end;
      MergeItem := TheMergeItem.Create;
      MergeItem.Debut := II;
      MergeItem.Fin := -1;
      MergeItem.Numero := OneColonne.semaine;
      MergedSem.Add(MergeItem);
      SemCourante := OneColonne.semaine;
      //
      LastSemaine := MergeItem;
    end;
  end;
  LastSemaine.Fin := lesColonnesG2.count -1;
end;

procedure TFplanningBTP.GSemDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
var grid : THGrid;
    II : Integer;
    OO :  TheMergeItem;
begin
  grid := (Sender as THgrid);
  for II := 0 to MergedSem.Count -1 do
  begin
    OO := MergedSem.Items [II];
    MergingCell(Grid ,OO.debut,OO.Fin,ACol,ARow,State);
  end;
end;

procedure TFplanningBTP.GMoisDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
var grid : THGrid;
    II : Integer;
    OO :  TheMergeItem;
begin
  grid := (Sender as THgrid);
  for II := 0 to MergedMois.Count -1 do
  begin
    OO := MergedMois.Items [II];
    MergingCell(Grid ,OO.debut,OO.Fin,ACol,ARow,State);
  end;
end;

procedure TFplanningBTP.GDateDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect; State: TGridDrawState);
var grid : THGrid;
begin
  grid := (Sender as THgrid);
  Grid.Canvas.FillRect(Rect);
  DrawButton(Grid,Rect,true);
  if IsDayClose (lesColonnesG2.items[Acol],JourFermeture) and not(gdSelected in state) then
  begin
//    Grid.Canvas.Brush.Color := clAqua;
    Grid.Canvas.Brush.Color := $f4f2a5;
    Grid.Canvas.FillRect(rect);
  end else if (gdSelected in state)  then
  begin
    Grid.Canvas.Brush.Color := clHighlight ;
    Grid.Canvas.FillRect(rect);
  end;
  OwnerDrawText(Grid,State,Acol,Rect,grid.cells[Acol,Arow],DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL);
end;

procedure TFplanningBTP.GTEffectifsDrawCell(Sender: TObject; ACol,ARow: Integer; Rect: TRect; State: TGridDrawState);
var grid : THGrid;
  stcell : string;
begin
  grid := (Sender as THgrid);
  ClearCanvas (grid,ACol,Rect);
  Grid.Canvas.FillRect(Rect);
  DrawButton(Grid,Rect,true);
  if IsDayClose (lesColonnesG2.items[Acol],JourFermeture) and not(gdSelected in state) then
  begin
//    Grid.Canvas.Brush.Color := clAqua;
    Grid.Canvas.Brush.Color := $f4f2a5;
    Grid.Canvas.FillRect(rect);
  end else if (gdSelected in state)  then
  begin
    Grid.Canvas.Brush.Color := clHighlight ;
    Grid.Canvas.FillRect(rect);
  end;
  DrawButton(Grid,Rect,true);
  if IsDayClose (lesColonnesG2.items[Acol],JourFermeture)  then stCell := ''
                                                           else StCell := grid.cells[Acol,Arow];
  OwnerDrawText(Grid,State,Acol,Rect,StCell,DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL);
end;

procedure TFplanningBTP.GTheuresDrawCell(Sender: TObject; ACol,ARow: Integer; Rect: TRect; State: TGridDrawState);
var grid : THGrid;
    TT : TOB;
    DD : double;
    St,TheText : string;
begin
  grid := (Sender as THgrid);
  if TOBTOTJours.detail.count = 0 then exit;
  TT := TOBTOTJours.detail[0];
  St := 'TJ'+InttoStr(Acol+1);
  DD := TT.getDouble(St);
  TheText := GetDureeSt (DD); if theText = '0' then TheText := '';

  Grid.Canvas.FillRect(Rect);
  ClearCanvas (grid,ACol,Rect);
  DrawButton(Grid,Rect,true);
  if IsDayClose (lesColonnesG2.items[Acol],JourFermeture) and not(gdSelected in state) then
  begin
//    Grid.Canvas.Brush.Color := clAqua;
    Grid.Canvas.Brush.Color := $f4f2a5;
    Grid.Canvas.FillRect(rect);
  end else if (gdSelected in state)  then
  begin
    Grid.Canvas.Brush.Color := clHighlight ;
    Grid.Canvas.FillRect(rect);
  end;
  DrawButton(Grid,Rect,true);
  OwnerDrawText(Grid,State,Acol,Rect,TheText,DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL);
//  OwnerDrawText(Grid,State,Acol,Rect,grid.cells[Acol,Arow],DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL);
end;


procedure TFplanningBTP.G1GetCellCanvas(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var TOBL : TOB;
begin
  if (Modegestion = TmPcPlanCharge) then
  begin
    TOBL := GetTOBLigne(TobAffiches,Arow);
    if ((LesColonnesG1.items[Acol].Nom = 'LIBELLE') or (LesColonnesG1.items[Acol].Nom = 'AFFAIRE'))and
        (TOBL.GetInteger('LEVEL')=0)  and not (gdSelected in Astate) then
    begin
      Canvas.Brush.Color := clLtGray ;
      Canvas.font.Size := 10;
      Canvas.font.Style := [fsbold];
    end;
    if ((LesColonnesG1.items[Acol].Nom = 'LIBELLE') or (LesColonnesG1.items[Acol].Nom = 'AFFAIRE'))and
        (TOBL.GetInteger('LEVEL')=0)  and (gdSelected in Astate) then
    begin
      Canvas.font.Size := 10;
      Canvas.font.Style := [fsbold];
    end;
    if (LesColonnesG1.items[Acol].Nom <> 'INDIC') and (TOBL.GetInteger('LEVEL')>0) and not (gdSelected in Astate) then
    begin
      Canvas.Brush.Color := TOBL.GetInteger('BCOLOR');
      Canvas.Font.Color := TOBL.GetInteger('COLOR');
    end;
    if (LesColonnesG1.items[Acol].Nom = 'RESTE')  then
    begin
      if TOBL.GetDouble('RESTE') < 0 then
      begin
        Canvas.Font.Color := clred;
        Canvas.font.Style := [fsbold];
        if not (gdSelected in Astate) then Canvas.Brush.Color := clWindow;
      end;
    end;
  end;
  if (Modegestion = TmcPlanCharfnc) then
  begin
    TOBL := GetTOBLigne(TobAffiches,Arow);
    if (LesColonnesG1.items[Acol].Nom = 'LIBNATURE') and (TOBL.GetInteger('LEVEL')=0) and not (gdSelected in Astate)  then
    begin
//      Canvas.Brush.Color := clLtGray ;
      Canvas.Brush.Color := TOBL.GetInteger('BCOLOR') ;
      Canvas.Font.Color := TOBL.GetInteger('COLOR');
      Canvas.font.Size := 10;
      Canvas.font.Style := [fsbold];
    end;
    if (LesColonnesG1.items[Acol].Nom = 'LIBNATURE') and (TOBL.GetInteger('LEVEL')=0) and (gdSelected in Astate)  then
    begin
      Canvas.font.Size := 10;
      Canvas.font.Style := [fsbold];
    end;
    if (LesColonnesG1.items[Acol].Nom = 'LIBNATURE') and (TOBL.GetInteger('LEVEL')>0) and not (gdSelected in Astate)  then
    begin
      Canvas.Brush.Color := clInfoBk;
    end;
    if (LesColonnesG1.items[Acol].Nom = 'RESTE') and not (gdSelected in Astate) then
    begin
      if TOBL.GetDouble('RESTE') < 0 then
      begin
        Canvas.Font.Color := clred;
        Canvas.font.Style := [fsbold];
        if not (gdSelected in Astate) then Canvas.Brush.Color := clWindow;
      end;
    end;
  end;
  if (Modegestion = TmPcPlanningCha)then
  begin
    TOBL := GetTOBLigne(TobAffiches,Arow);
    if (LesColonnesG1.items[Acol].Nom = 'LIBNATURE') and (TOBL.GetInteger('LEVEL')=0) and not (gdSelected in Astate) then
    begin
      Canvas.Brush.Color := clLtGray ;
      Canvas.font.Size := 10;
      Canvas.font.Style := [fsbold];
    end;
    if (LesColonnesG1.items[Acol].Nom = 'LIBNATURE') and (TOBL.GetInteger('LEVEL')=0) and (gdSelected in Astate) then
    begin
      Canvas.font.Size := 10;
      Canvas.font.Style := [fsbold];
    end;
    if (LesColonnesG1.items[Acol].Nom = 'LIBNATURE') and (TOBL.GetInteger('LEVEL')>0) and not (gdSelected in Astate) then
    begin
      Canvas.Brush.Color := TOBL.GetInteger('BCOLOR') ;
      Canvas.Font.Color := TOBL.GetInteger('COLOR');
//      Canvas.Brush.Color := clInfoBk;
    end;
    if (LesColonnesG1.items[Acol].Nom = 'RESTE') and not (gdSelected in Astate) then
    begin
      if TOBL.GetDouble('RESTE') < 0 then
      begin
        Canvas.Font.Color := clred;
        Canvas.font.Style := [fsbold];
        if not (gdSelected in Astate) then Canvas.Brush.Color := clWindow;
      end;
    end;
  end;
end;

procedure TFplanningBTP.G2GetCellCanvas(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var TT : TOB;
begin
  TT := GetTOBLigne(TobAffiches,Arow);

  if (IsDayClose (lesColonnesG2.items[Acol],JourFermeture)) and not (gdSelected in Astate) then
  begin
//    Canvas.Brush.Color := clAqua;
    Canvas.Brush.Color := $f4f2a5;
  end;

  if (Trim(G2.cells[Acol,Arow]) <> '') and not (gdSelected in Astate) then
  begin
    if TT.GetInteger('LEVEL') = 0 then
    begin
      Canvas.Brush.Color := clLtGray;
      Canvas.Font.Color := clwindow;
    end else
    begin
//      Canvas.Brush.Color := clInfoBk;
      if TT.GetInteger('BCOLOR') = 0 then Canvas.Brush.Color := clInfoBk
                                     else Canvas.Brush.Color := TT.GetInteger('BCOLOR');
      if TT.GetInteger('COLOR') <> 0 then Canvas.Font.Color := TT.GetInteger('COLOR');
    end;
  end;
end;

procedure TFplanningBTP.Colorise (Colonne : Integer; Ajout : Boolean);
begin
  ModeColor := Ajout;
  //
  GTHeures.Col := Colonne;
  GTheures.InvalidateCell (Colonne,0);
  GTEffectifs.Col := Colonne;
  GTEffectifs.InvalidateCell (Colonne,0) ;
  GDate.Col := Colonne;
  GDate.InvalidateCell (Colonne,0) ;
  GFONCTION.col := Colonne;
  GFONCTION.InvalidateCell (Colonne,0) ;
  //
end;

procedure TFplanningBTP.G2Cellenter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if (Acol <> G2.col) and (Acol <> 0)  then
  begin
    Colorise(acol,false);
  end;
  Colorise(G2.col,True);
end;

procedure TFplanningBTP.CumuleTOB (TOBL,TOBP: TOB; Acol : Integer; NbHrs : Double ; sens : string ='');
var TT : TOB;
    ZZ,CZ,DZ : string;
begin
  ZZ := 'TJ'+InttoStr(Acol+1);
  CZ := 'D'+InttoStr(Acol+1);
  DZ := 'DEJAPLANN';
  if TOBTOTJours.detail.count > 0 then
  begin
    TT := TOBTOTJours.detail[0];
    if Sens = '' then
    begin
      TT.SetDouble(ZZ,TT.GetDouble(ZZ)+NbHrs);
      TOBL.SetDouble(CZ,TOBL.GetDouble(CZ)+NbHrs);
      TOBL.SetDouble(DZ,TOBL.GetDouble(DZ)+NbHrs);
      if TOBP <> nil then
      begin
        TOBP.SetDouble(CZ,TOBP.GetDouble(CZ)+NbHrs);
        TOBP.SetDouble(DZ,TOBP.GetDouble(DZ)+NbHrs);
      end;
    end else
    begin
      TT.SetDouble(ZZ,TT.GetDouble(ZZ)-NbHrs);
      TOBL.SetDouble(CZ,TOBL.GetDouble(CZ)-NbHrs);
      TOBL.SetDouble(DZ,TOBL.GetDouble(DZ)-NbHrs);
      if TOBP <> nil then
      begin
        TOBP.SetDouble(CZ,TOBP.GetDouble(CZ)-NbHrs);
        TOBP.SetDouble(DZ,TOBP.GetDouble(DZ)-NbHrs);
      end;
    end;
    TOBL.SetDouble('RESTE',TOBL.GetDouble('BUDGET')-TOBL.GetDouble('DEJAPLANN'));
    if TOBP <> nil then TOBP.SetDouble('RESTE',TOBP.GetDouble('BUDGET')-TOBP.GetDouble('DEJAPLANN'));
  end;
end;

procedure TFplanningBTP.SupprimeAffectation (TOBL,TOBP,TOBI : TOB; Acol:integer);
begin
  CumuleTOB(TOBL,TOBP,Acol,TOBI.GetDouble('BAT_NBHRS'),'-');
  if TOBI.GetInteger('DBPRESENT')= 0 then TOBI.Free
                                     else TOBI.ChangeParent(TOBL.detail[1],-1);
end;

procedure TFplanningBTP.AjouteAffectation (TOBL,TOBP: TOB; Acol: Integer);
var TOBI : TOB;
    TheDate : TDateTime;
begin
  if TOBL.detail.count = 0 then exit;
  TheDate := TheColonne(lesColonnesG2.Items [Acol]).TheDate;
  TOBI := TOB.Create ('BAFFECTCHANT',TOBL.detail[0],-1);
  TOBI.AddChampSupValeur('DBPRESENT',0);
  TOBI.SetString('BAT_AFFAIRE',Chantier);
  TOBI.SetInteger('BAT_NUMORDRE',TOBL.GetInteger('NUMORDRE'));
  TOBI.SetString('BAT_FONCTION',TOBL.GetString('FONCTION'));
  TOBI.SetString('BAT_LIBELLE',TOBL.GetString('LIBNATURE'));
  TOBI.SetBoolean('BAT_PLANNIFIE',false);
  TOBI.SetDateTime('BAT_DATE',TheDate);
  TOBI.SetDouble('BAT_NBHRS',Arrondi(VALEUR(NBHRS.text),V_PGI.okdecQ));
  CumuleTOB(TOBL,TOBP,Acol,TOBI.GetDouble('BAT_NBHRS'));
end;

function TFplanningBTP.FindLaDateinTOB (LesDate : TOB; LaDate : TDateTime) : TOB;
var II : Integer;
begin
  Result := nil;
  IF LesDate.Detail.Count = 0 then exit;
  for II := 0 to LesDate.Detail.Count -1 do
  begin
    if LesDate.detail[II].GetDateTime('BAT_DATE')=LaDate then
    begin
      Result := LesDate.detail[II];
      break;
    end;
  end;
end;

procedure TFplanningBTP.AffecteHrsRess (ColDeb,ARow,ColFin : integer);
var II,IRow : Integer;
    TOBL,TOBI,TOBP : TOB;
    TheDate : TDateTime;
    OneColonne : TheColonne;
begin
  TOBL := GetTOBLigne(TobAffiches,ARow); if TOBL = nil then Exit;
  if TOBL.detail.count = 0 then exit;
  TOBP := TOB(TOBL.Data);
  if TOBP <> nil then IRow := TOBP.GetIndex
                 else Irow := -1;
  for II := ColDeb to ColFin do
  begin
    OneColonne := lesColonnesG2.Items[II];
    TheDate := OneColonne.TheDate;
    if IsDayClose (OneColonne,JourFermeture) then continue;
    TOBI := FindLaDateinTOB (TOBL.Detail[0],TheDate);
    if TOBI<> nil then
    begin
      if TOBI.GetBoolean('BAT_PLANNIFIE') then continue;
      SupprimeAffectation (TOBL,TOBP,TOBI,II);
    end else
    begin
      AjouteAffectation (TOBL,TOBP,II);
    end;
    TOBL.PutLigneGrid (G1,ARow,false,false,LesColsG1);
    TOBL.PutLigneGrid(G2,Arow,false,false,LesColsG2);
    if TOBTOTJours.detail.count > 0 then TOBTOTJours.detail[0].PutLigneGrid(GTHeures,0,false,false,LesColsGTH);
    G1.invalidate;
    G2.invalidate;
    GtHeures.invalidate;
    CalculEffectifs;
  end;
  if TOBP <> nil then
  begin
    TOBP.PutLigneGrid (G1,IRow,false,false,LesColsG1);
    TOBP.PutLigneGrid (G2,IRow,false,false,LesColsG2);
  end;
end;

procedure TFplanningBTP.G2MouseDown(Sender: TObject; Button: TMouseButton;Shift: TShiftState; X, Y: Integer);

  function IsLigneAffectable (Arow : integer) : boolean;
  var TOBL : TOB;
  begin
    Result := false;
    TOBL := GetTOBLigne(TobAffiches,ARow);
    if TOBL.detail.count = 0 then exit;
    Result := True;
  end;

var Acol,Arow : Integer;
begin
  if ModeGestionGrid = tmgConsult then Exit;
  G2.MouseToCell(X,Y,Acol,Arow);
  if not IsLigneAffectable (Arow) then Exit;
  //
  DetailModif := true;
  //
  if (ssShift in Shift) and ( Button=mbleft) then
  begin
    if (LastPos.Arow = Arow) and (LastPos.Acol <> Acol) then
    begin
      AffecteHrsRess(LastPos.Acol,Arow,Acol);
    end;
  end else if (ssleft in Shift) and ( Button=mbleft) then
  begin
    AffecteHrsRess(Acol,Arow,Acol);
    LastPos.acol := Acol+1;
    LastPos.Arow := Arow;
  end;
end;

procedure TFplanningBTP.ConstitueGrilleTH;

  function ConstitueTotHeures : string;
  var II : Integer;
  begin
    For II := 1 to NBjours do
    begin
      Result := Result + 'TJ'+InttoStr(II)+';';
    end;
  end;
  
var st,FF      : string;
    i      : Integer;
begin
  LesColsGTH := ConstitueTotHeures;
  FF:='#';
  if V_PGI.OkDecQ>0 then
  begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecQ-1 do
    begin
      FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;
  //
  GTHeures.DefaultRowHeight := 18;
  //
  GTHeures.ColCount := NbJours;
  //
  St := LesColsGTH;
  //
  for i := 0 to GTHeures.ColCount - 1 do
  begin
    GTHeures.ColWidths[i] := 20;
    GTHeures.ColAligns[i] := taCenter;
    GTHeures.ColFormats[i] := FF+';-'+FF+';;';
  end;
  //
  GTHeures.RowCount := 1;
  TOBTOTJours.PutGridDetail(GTHeures,false,false,LesColsGTH);
end;

procedure TFplanningBTP.BOuvrirFermerClick(Sender: TObject);

  procedure AjouteTOB (TOBAffiches,TOBAA: TOB ; PosInsere : integer);
  var IDeb : integer;
      TOBII : TOB;
  begin
    Ideb := PosInsere+1;
    repeat
      TOBII := TOBAA.detail[0];
      Calculreste (TOBII);
      TOBII.ChangeParent(TOBAffiches,Ideb);
      G1.InsertRow(Ideb);
      TOBII.PutLigneGrid (G1,Ideb,false,false,LesColsG1);
      G2.InsertRow(Ideb);
      G2HauteurLigne (Ideb);
      TOBII.PutLigneGrid (G2,Ideb,false,false,LesColsG2);
      inc(Ideb);
    until TOBAA.detail.count =0;
  end;

  procedure AjouteDetail (TOBL: TOB ;position : integer);
  var Chantier,Nature : string;
    SQL : WideString;
    QQ : TQuery;
    TOBAdd : TOB;
    cancel : boolean;
  begin
    Chantier := '';
    Nature := '';
    TOBAdd := TOB.Create ('LES DETAIL',nil,-1);
    if Modegestion = TmPcPlanCharge then
    begin
      Chantier := TOBL.getString('AFFAIRE');
      SQl := ConstitueRequete(chantier,'',true);
      QQ := OpenSQL(SQL,True,-1,'',true);
      if not QQ.eof then
      begin
        TOBADD.LoadDetailDB('UN DETAIL','','',QQ,false);
      end;
      ferme (QQ);
    end else if Modegestion = TmcPlanCharfnc then
    begin
      Nature := TOBL.getString('FONCTION');
      SQl := ConstitueRequete('',Nature,true);
      QQ := OpenSQL(SQL,True,-1,'',true);
      if not QQ.eof then
      begin
        TOBADD.LoadDetailDB('UN DETAIL','','',QQ,false);
      end;
      ferme (QQ);

    end;
    AjoutePlannifieChantier (TOBAdd,Chantier, Nature);
    //
    if TOBADD.detail.count > 0 then
    begin
      AjouteTOB (TOBAffiches,TOBADD,Position);
    end;
    //
    TOBL.SetInteger('OPENED',1);
    G1RowEnter(Self,position,cancel,false);
    TOBAdd.free;
   end;

  procedure RetireDetail (TOBL: TOB;Position : integer);
  var II : Integer;
      Chantier,Nature : string;
      TOBS : TOB;
      cancel : boolean;
  begin
    Chantier := TOBL.getString('AFFAIRE');
    Nature := TOBL.getString('FONCTION');
    //
    II := Position +1;
    repeat
      if II >= TOBAffiches.detail.count then break; 
      TOBS := GetTOBLigne(TobAffiches,II);
      if (Modegestion = TmPcPlanCharge) and (TOBS.GetString('AFFAIRE')<>Chantier) then break;
      if (Modegestion = TmcPlanCharfnc) and (TOBS.GetString('FONCTION')<>Nature) then break;
      TOBS.Free;
      G1.DeleteRow(II);
      G2.DeleteRow(II);
    until ( 1=2);
    TOBL.SetInteger('OPENED',0);
    G1RowEnter(Self,position,cancel,false);
    for II := 0 to TOBAffiches.detail.count -1 do
    begin
      G2HauteurLigne (II);
    end;
  end;


var Position : integer;
    TOBl : TOB;
begin
  Position := G1.Row;
  TOBL := GetTOBLigne(TobAffiches,Position);
  if (Modegestion = TmPcPlanCharge) or (Modegestion =TmcPlanCharfnc) then
  begin
    if (TOBL.GetInteger('LEVEL') = 0) then
    begin
      if TOBL.GetInteger('OPENED')= 0 then
      begin
        AjouteDetail (TOBL,Position);
      end else
      begin
        RetireDetail (TOBL,Position);
      end;
    end;
  end;
end;

procedure TFplanningBTP.G1PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var Arect : Trect;
    TT : Tob;
    SS,TheText : string;
    DD : double;
    TheColor : Tcolor;
begin
  TT := GetTOBLigne(TobAffiches,Arow);
  if TT = nil then exit;
  SS := BTPCodeAffaireAffiche (TT.GetString('AFFAIRE'));
  ARect := G1.CellRect(ACol, ARow);
  if (Modegestion = TmPcPlanCharge) then
  begin
    if (LesColonnesG1.items[Acol].Nom = 'AFFAIRE') and (TT.GetInteger('LEVEL')=0) then
    begin
      G1.Canvas.FillRect(ARect);
      G1.Canvas.Brush.Style := bsSolid;
      G1.Canvas.TextOut (Arect.left+1,Arect.Top +2,SS);
    end;
    if (LesColonnesG1.items[Acol].Nom = 'AFFAIRE') and (TT.GetInteger('LEVEL')>0) then
    begin
      G1.Canvas.FillRect(ARect);
      exit;
    end;
    if (LesColonnesG1.items[Acol].Nom = 'LIBELLE') and (TT.GetInteger('LEVEL')>0) then
    begin
      G1.Canvas.FillRect(ARect);
      G1.Canvas.Brush.Style := bsSolid;
      if (gdSelected in Astate) then
      begin
//        G1.Canvas.Font.color := clBlack;
      end;
      G1.Canvas.TextOut (Arect.left+1,Arect.Top +2,'     '+G1.Cells[Acol,Arow]);
    end;
    if (LesColonnesG1.items[Acol].Nom = 'BUDGET') then
    begin
      G1.Canvas.FillRect(ARect);
      DD := TT.getDouble('BUDGET');
      TheText := GetDureeSt (DD);
      OwnerDrawText(G1,AState,Acol,ARect,TheText,DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL);
    end;
    if (LesColonnesG1.items[Acol].Nom = 'RESTE') then
    begin
      G1.Canvas.FillRect(ARect);
      DD := TT.getDouble('RESTE'); if DD < 0 then TheColor := Clred else TheColor := 0;
      TheText := GetDureeSt (DD);
      OwnerDrawText(G1,AState,Acol,ARect,TheText,DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL,TheColor);
      (*
      if TT.GetDouble('RESTE') < 0 then
      begin
        Raye(G1,G1.Canvas,ACol,ARow);
      end;
      *)
    end;
  end;
  //
  if (Modegestion = TmcPlanCharfnc)  then
  begin
    if (LesColonnesG1.items[Acol].Nom = 'LIBNATURE') and (TT.GetInteger('LEVEL')>0) then
    begin
      G1.Canvas.FillRect(ARect);
      G1.Canvas.Brush.Style := bsSolid;
      if (gdSelected in Astate) then
      begin
        G1.Canvas.Font.color := clBlack;
      end;
      G1.Canvas.TextOut (Arect.left+1,Arect.Top +2,'     ('+SS+') '+G1.Cells[Acol,Arow]);
    end;
    if (LesColonnesG1.items[Acol].Nom = 'BUDGET') then
    begin
      G1.Canvas.FillRect(ARect);
      DD := TT.getDouble('BUDGET');
      TheText := GetDureeSt (DD);
      OwnerDrawText(G1,AState,Acol,ARect,TheText,DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL);
    end;
    if (LesColonnesG1.items[Acol].Nom = 'RESTE') then
    begin
      G1.Canvas.FillRect(ARect);
      DD := TT.getDouble('RESTE'); if DD < 0 then TheColor := Clred else TheColor := 0;
      TheText := GetDureeSt (DD);
      OwnerDrawText(G1,AState,Acol,ARect,TheText,DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL,TheColor);
      (*
      if TT.GetDouble('RESTE') < 0 then
      begin
        Raye(G1,G1.Canvas,ACol,ARow);
      end;
      *)
    end;
  end;
  if (Modegestion = TmPcPlanningCha)  then
  begin
    if (LesColonnesG1.items[Acol].Nom = 'LIBNATURE') and (TT.GetInteger('LEVEL')>0) then
    begin
      G1.Canvas.FillRect(ARect);
      G1.Canvas.Brush.Style := bsSolid;
      (*
      if (gdSelected in Astate) then
      begin
        G1.Canvas.Font.color := clBlack;
      end;
      *)
      G1.Canvas.TextOut (Arect.left+1,Arect.Top +3,'      '+G1.Cells[Acol,Arow]);
    end;
    if (LesColonnesG1.items[Acol].Nom = 'BUDGET') then
    begin
      G1.Canvas.FillRect(ARect);
      DD := TT.getDouble('BUDGET');
      TheText := GetDureeSt (DD);
      OwnerDrawText(G1,AState,Acol,ARect,TheText,DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL);
    end;
    if (LesColonnesG1.items[Acol].Nom = 'RESTE') then
    begin
      G1.Canvas.FillRect(ARect);
      DD := TT.getDouble('RESTE'); if DD < 0 then TheColor := Clred else TheColor := 0;
      TheText := GetDureeSt (DD);
      OwnerDrawText(G1,AState,Acol,ARect,TheText,DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL,TheColor);
      (*
      if TT.GetDouble('RESTE') < 0 then
      begin
        Raye(G1,G1.Canvas,ACol,ARow);
      end;
      *)
    end;
  end;
end;

function TFplanningBTP.GetDureeSt (DD : Double) : string;
var ZZ : double;
    Sens : integer;
    HDM : double;
    DELTAHRS : double;
begin
  result := '';
  HDM := DureeJOurnee / 2;
  if DD > 0 then Sens := 1 else Sens := -1;
  if DD = 0 then exit;
  if RBHeures.checked then
  begin
    result := Format('%5.2f',[DD]);
  end else
  begin
    ZZ := Arrondi(DD/DureeJournee,0)*Dureejournee;
    DELTAHRS := DD - ZZ;
    if (DD < DureeJournee) then ZZ := 0.5   // durre inférieure a 1 journée
    else if ((DD - ZZ) < 0)  then
      ZZ := Arrondi(DD/DureeJournee,0)-0.5
    else if ((DD - ZZ) > 0) and ((DELTAHRS) <= HDM) then
      ZZ := Arrondi(DD/DureeJournee,0)+0.5
    else if ((DD - ZZ) > 0) and ((DELTAHRS) > HDM) then
      ZZ := Arrondi(DD/DureeJournee,0)+1
    else
      ZZ := Arrondi(DD/DureeJournee,0);
    ZZ := ZZ * Sens;
    result := Format('%5.1f',[ZZ]);
  end;
end;

function TFplanningBTP.GetDureeDD (DD : Double) : double;
var ZZ : double;
begin
  result := 0;
  if DD = 0 then exit;
  if RBHeures.checked then
  begin
    result := DD;
  end else
  begin
    ZZ := Arrondi(DD/DureeJournee,0)*Dureejournee;
    if (DD < DureeJournee) then ZZ := 0.5
    else if ((DD - ZZ) < 0)  then ZZ := Arrondi(DD/DureeJournee,0)-0.5
    else if ((DD - ZZ) > 0) and ((DD - ZZ) <= 0.5) then ZZ := Arrondi(DD/DureeJournee,0)+0.5
    else if ((DD - ZZ) > 0) and ((DD - ZZ) >0.5) then ZZ := Arrondi(DD/DureeJournee,0)+1
    else ZZ := Arrondi(DD/DureeJournee,0);
    result := ZZ;
  end;
end;


procedure TFplanningBTP.G2PostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var TT,TOBI,TOBL : Tob;
    Arect : Trect;
    TheText,St : string;
    DD : Double;
    TheDate : Tdatetime;
    OneColonne : TheColonne;
begin
  ARect := G2.CellRect(ACol, ARow);
  //
  //
  TT := GetTOBLigne(TobAffiches,Arow);
  if TT = nil then exit;
  //
  if (Modegestion = TmPcPlanningCha) and (gdSelected in Astate) and (ModeGestionGrid=tmgModif) then
  begin
    if TT.detail.count = 0 then Raye(G2,G2.Canvas,ACol,ARow);
  end;
  G2.Canvas.FillRect(ARect);
  if CAffHeures.checked then
  begin
    G2.Canvas.Brush.Style := bsSolid;
    if GetDureeSt (TT.GetDouble('PLA'+InttoStr(Acol+1))) <> '' then
    begin
      TheText := trim(GetDureeSt (TT.GetDouble('D'+InttoStr(Acol+1))));
      OwnerDrawText(G2,AState,Acol,ARect,TheText,DT_SINGLELINE or DT_CENTER or DT_TOP or DT_EDITCONTROL);

      TheText := '('+Trim(GetDureeSt (TT.GetDouble('PLA'+InttoStr(Acol+1))))+')';
      OwnerDrawText(G2,AState,Acol,ARect,TheText,DT_SINGLELINE or DT_CENTER or DT_BOTTOM or DT_EDITCONTROL);
    end else
    begin
      TheText := trim(GetDureeSt (TT.GetDouble('D'+InttoStr(Acol+1))));
      OwnerDrawText(G2,AState,Acol,ARect,TheText,DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL);
    end;
  end;
  //
  if TT.detail.count = 0 then exit;
  TOBL := TOB(TT.data);
  if TOBL = nil then exit;
  OneColonne := lesColonnesG2.Items[Acol];
  TheDate := OneColonne.TheDate;
  TOBI := FindLaDateinTOB (TT.Detail[0],TheDate);
  if TOBI<> nil then
  begin
    if TOBI.getBoolean('BAT_PLANNIFIE') then Raye(G2,G2.Canvas,ACol,ARow);
  end;
end;


procedure TFplanningBTP.ReinitDatas;
begin
  TobAffiches.ClearDetail;
  TOBAffects.ClearDetail;
  TOBAffectsDel.ClearDetail;
  TOBTOTJours.ClearDetail;
  TOBTOTFONC.ClearDetail;
  //
  LesColonnesG1.clear;
  lesColonnesG2.clear;
  MergedSem.clear;
  MergedMois.clear;
end;


procedure TFplanningBTP.BRefreshClick(Sender: TObject);
var cancel : boolean;
begin
//
  IF DetailModif then
  begin
    if PGIAsk('ATTENTION : les données ont été modifiés.#13#10 Etes-vous sur de vouloir annuler ?')=mryes then
    begin
      DetailModif := false;
    end else Exit;
  end;
//
  ReinitDatas;
  LoadlesTobs;
  PrepareAffichage;
  G1RowEnter(Self,G1.row,cancel,false);
end;

procedure TFplanningBTP.Raye(Grid: THgrid; Canvas: TCanvas; ACol, ARow: integer);
var  Arect: Trect;
begin
  Arect := Grid.CellRect(Acol, Arow);
  //
  Canvas.Brush.Color := clActiveCaption;
  Canvas.Brush.Style := bsBDiagonal;
  Canvas.Pen.Color := clActiveCaption;
  Canvas.Pen.Mode := pmCopy;
  Canvas.Pen.Style := psClear;
  Canvas.Pen.Width := 1;
  Canvas.Rectangle(Arect.Left, Arect.Top, Arect.Right + 1, Arect.Bottom + 1);
end;

procedure TFplanningBTP.SetActionMenu (Arow: Integer);
var TT : TOB;
begin
  if Arow > TobAffiches.detail.count then exit;
  TT := GetTOBLigne(TobAffiches,Arow);
  if TT = nil then exit;
  G1.PopupMenu := nil;
  if (Modegestion = TmPcPlanCharge) then
  begin
    if TT.GetInteger('LEVEL')=0 then G1.PopupMenu := POPACTIONS;
  end else if (Modegestion = TmcPlanCharfnc) then
  begin
    if TT.GetInteger('LEVEL')=1then G1.PopupMenu := POPACTIONS;
  end else if (Modegestion = TmPcPlanningCha) then
  begin
    PlannifCh.Visible := false;
    N1.visible := false;
    G1.PopupMenu := POPACTIONS;
  end;
end;

procedure TFplanningBTP.PlannifChClick(Sender: TObject);
var TT : TOB;
    chantier : string;
begin
  TT := GetTOBLigne(TobAffiches,G1.row); if TT = nil then Exit;
  chantier := TT.GetString('AFFAIRE');
  OpenPlanningBtp (TmPcPlanningCha,nil,Chantier,Etablissement,DateDepart,DateFin,ModeAffichage);
  RefreshChantier(TT);
end;

procedure TFplanningBTP.RefreshChantier(TT : TOB);
var SQl : String;
    QQ : Tquery;
    Chantier,Fonction : string;
    OkOk,Detail : Boolean;
    II,Ind0,Open : Integer;
begin
  detail := false;
  if Modegestion = TmPcPlanCharge then Chantier := TT.getString('AFFAIRE');
  if Modegestion = TmcPlanCharfnc then
  begin
    Chantier := TT.getString('AFFAIRE');
    Fonction := TT.getString('FONCTION');
    detail := true;
  end;
  //
  Open := TT.GetInteger('OPENED');
  Ind0 := TT.GetIndex;
  OkOk := false;
  SQl := ConstitueRequete(Chantier,Fonction,detail);
  QQ := OpenSQL(SQL,True,-1,'',true);
  if not QQ.eof then
  begin
    TT.InitValeurs;
    TT.SelectDB('',QQ);
    TT.SetInteger('OPENED',Open);
    //
    Calculreste (TT);
    Okok := True;
    AjoutePlannifieChantier (TT,Chantier,Fonction);
    TT.PutLigneGrid(G1,Ind0,False,False,LesColsG1);
    TT.PutLigneGrid(G2,Ind0,False,False,LesColsG2);
  end;
  ferme (QQ);
  if OKOK then
  begin
    if Ind0 < TobAffiches.Detail.count then
    begin
      for II := Ind0+1 to TobAffiches.Detail.count -1 do
      begin
        if TobAffiches.detail[II].GetString('AFFAIRE')<> Chantier then break;
        if Modegestion = TmPcPlanCharge then Fonction := TobAffiches.detail[II].getString('FONCTION') else
        if Modegestion = TmcPlanCharfnc then Chantier := TobAffiches.detail[II].getString('AFFAIRE');

        SQl := ConstitueRequete(Chantier,Fonction,true);
        QQ := OpenSQL(SQL,True,-1,'',true);
        if Not QQ.eof then
        begin
          TOBAffiches.detail[II].InitValeurs;
          TobAffiches.detail[II].SelectDB('',QQ);
          Calculreste (TobAffiches.detail[II]);
          AjoutePlannifieChantier (TOBAffiches.detail[II],Chantier,Fonction);
          TobAffiches.detail[II].PutLigneGrid(G1,II,False,False,LesColsG1);
          TobAffiches.detail[II].PutLigneGrid(G2,II,False,False,LesColsG2);
        end;
        ferme (QQ);
      end;
    end;
  end;
  for II := 0 to TOBAffiches.detail.count -1 do
  begin
    G2HauteurLigne (II);
  end;
  //
  TOBTOTJours.ClearDetail;
  SQL := 'SELECT '+ConstitueCumJour;
  QQ := OpenSQL(SQL,True,-1,'',true);
  if Not QQ.eof then
  begin
    TOBTOTJours.LoadDetailDB('UN TOTAL','','',QQ,false);
  end else
  begin
    AddOneTotal;
  end;
  TOBTOTJours.PutGridDetail(GTHeures,false,false,LesColsGTH);
  CalculEffectifs;
  G1.invalidate;
  G2.invalidate;
end;

procedure TFplanningBTP.BMODIFClick(Sender: TObject);
begin
  DetailModif := false;
  BMODIF.Visible := false;
  BVALIDER.visible := true;
  BABANDON.visible := false;
  BANNULE.Visible := true;
  LNBRRESS.visible := true;
  NBRRESS.visible := true;
  LNBHRS.visible := true;
  NBHRS.visible := true;
  ModeGestionGrid := tmgModif;
  Brefresh.hint := 'Annuler les modifications';
end;

procedure TFplanningBTP.CumuleTOBLigne(TA,TOBC : TOB; Sens : string = '');
var TOBP : TOB;
    II : Integer;
begin
  TOBP := TOB(TA.Data);
  if TOBC.GetDouble('BAT_NBHRS') <> 0 then
  begin
    II := lesColonnesG2.FindIndice(TOBC.GetDateTime('BAT_DATE'));
    if II <> -1 then
    begin
      CumuleTOB (TA,TOBP,II,TOBC.GetDouble('BAT_NBHRS'),sens);
    end;
  end;
end;


procedure TFplanningBTP.AnnuleSaisie;
var II,JJ,IP : Integer;
    TA : TOB;
    TB : TOB;
    TC : TOB;
    TP : TOB;
begin
  IP := -1;
  for II := 0 to TobAffiches.Detail.count -1 do
  begin
    if TobAffiches.Detail[II].Detail.count = 0 then continue;
    TA := TobAffiches.Detail[II];
    TP := TOB(TA.Data);
    if TP <> nil then IP := TP.GetIndex;
    TB := TA.Detail[1]; // les elements supprimes
    if TB.Detail.count > 0 then
    begin
      JJ := 0;
      repeat
        TC := TB.detail[JJ];
        CumuleTOBLigne(TA,TC );
        TC.ChangeParent(TA.detail[0],-1);
      until JJ >= TB.Detail.count;
    end;

    TB := TA.Detail[0];
    if TB.Detail.count > 0 then
    begin
      JJ := 0;
      repeat
        TC := TB.detail[JJ];
        if TC.GetInteger('DBPRESENT')=0 then
        begin
          // Suppression des ajouts dans la saisie.
          CumuleTOBLigne(TA,TC,'-');
          TC.Free;
        end else inc(JJ);
      until JJ >= TB.Detail.count;
    end;
    
    TA.PutLigneGrid (G1,II,false,false,LesColsG1);
    TA.PutLigneGrid(G2,II,false,false,LesColsG2);
    if IP <> 0 then
    begin
      TP.PutLigneGrid (G1,IP,false,false,LesColsG1);
      TP.PutLigneGrid(G2,IP,false,false,LesColsG2);
    end;
    if TOBTOTJours.detail.count > 0 then TOBTOTJours.detail[0].PutLigneGrid(GTHeures,0,false,false,LesColsGTH);
    G1.invalidate;
    G2.invalidate;
    GtHeures.invalidate;
  end;
  CalculEffectifs;
end;

procedure TFplanningBTP.BANNULEClick(Sender: TObject);
begin
  IF DetailModif then
  begin
    if PGIAsk('ATTENTION : les données ont été modifiés.#13#10 Etes-vous sur de vouloir annuler ?')=mryes then
    begin
      AnnuleSaisie;
      DetailModif := false;
    end else Exit;
  end;
  ModeGestionGrid := tmgConsult;
  BMODIF.Visible := true;
  BVALIDER.visible := false;
  BABANDON.visible := true;
  BANNULE.Visible := false;
  LNBRRESS.visible := false;
  NBRRESS.visible := false;
  LNBHRS.visible := false;
  NBHRS.visible := false;
  Brefresh.hint := 'Raffraichir les données';

end;

procedure TFplanningBTP.ValidelaSaisie;

  function InsereDB (TT : TOB) : boolean;
  var II : Integer;
  begin
    Result := false;
    if not GetNumCompteur ('BZZ',iDate1900,II) then
    begin
      Exit;
    end;
    TT.SetInteger('BAT_IDAFFECT',II);
    TT.SetString('BAT_AFFAIRE',chantier);
    TT.SetAllModifie(true);
    Result := TT.InsertDB(nil);
  end;

var II,JJ : Integer;
    TA,TS,TP : TOB;
begin
  for II := 0 to TobAffiches.Detail.Count -1 do
  begin
    TA:= TobAffiches.Detail[II];
    if TA.detail.count = 0 then Continue;
    // 1ere etape suppression des eliminés
    TS :=TA.detail[1];
    for JJ := 0 to TS.detail.count -1 do
    begin
      if TS.Detail[JJ].GetInteger('DBPRESENT')=1 then
      begin
        if not TS.Detail[JJ].DeleteDB then
        begin
          V_PGI.IOError := oeUnknown;
          Exit;
        end;
      end;
    end;
    // Creation ou mise a jour des plannifications
    TP :=TA.detail[0];
    for JJ := 0 to TP.detail.count -1 do
    begin
      if TP.Detail[JJ].GetInteger('DBPRESENT')=1 then
      begin
        if not TP.Detail[JJ].UpdateDB then
        begin
          V_PGI.IOError := oeUnknown;
          Exit;
        end;
      end else
      begin
        if Not InsereDB(TP.Detail[JJ]) then
        begin
          V_PGI.IOError := oeUnknown;
          Exit;
        end;
      end;
    end;
  end;
end;

procedure TFplanningBTP.PositionneToutValide;
var II,JJ : Integer;
    TD : TOB;
begin
  for II := 0 to TOBAffiches.detail.count -1 do
  begin
    TD := TOBAffiches.detail[II];
    if TD.Detail.count = 0 then continue;
    TD.detail[1].ClearDetail;
    for JJ := 0 to TD.detail[0].Detail.count -1 do
    begin
      TD.detail[0].detail[JJ].SetInteger('DBPRESENT',1);
    end;
  end;
  DetailModif := false;
end;

procedure TFplanningBTP.BValiderClick(Sender: TObject);
var XX : TIOErr;
    cancel : boolean;
begin
  if DetailModif then
  begin
    if PGIAsk('Vous aller validez les modifications apportées à la planification de ce chantier#13#10 Confirmez-vous ?') = mryes then
    begin
      XX := TRANSACTIONS(ValidelaSaisie,0);
      if XX <> OeOk then
      begin
        PGIError('ATTENTION : la validation n''a pu être effectuée');
        Exit;
      end else
      begin
        PositionneToutValide;
      end;
    end;
  end;
  ModeGestionGrid := tmgConsult;
  BMODIF.Visible := true;
  BVALIDER.visible := false;
  BABANDON.visible := true;
  BANNULE.Visible := false;
  LNBRRESS.visible := false;
  NBRRESS.visible := false;
  LNBHRS.visible := false;
  NBHRS.visible := false;
  Brefresh.hint := 'Raffraichir les données';
  G1RowEnter (self,G1.row,cancel,false);;
end;

function TFplanningBTP.SetListeChantiers (Prefixe : string = 'GL') : Widestring;
var ii : integer;
begin
  result := '';
  if (ListeChantiers.detail.count = 0) and (ListeChantiers.getString('TOUS')='-') then exit;
  if ListeChantiers.getString('TOUS')='X' then
  begin
    result := result + ' '+prefixe+'_AFFAIRE IN ( SELECT AFF_AFFAIRE FROM AFFAIRE '+ListeChantiers.getString('CRITERES')+') ';
  end else
  begin
    for II := 0 to ListeChantiers.detail.count -1 do
    begin
      if II = 0 then result := result + ' '+prefixe+'_AFFAIRE IN ('
                else result := result + ',';
      result := result + '"'+ListeChantiers.detail[ii].getString('CHANTIER')+'"';
    end;
    result := result + ') ';
  end;
end;

function TFplanningBTP.SetListeTypeRessources : string;

  function ConstituelisteTR (TheListe : string ) : string;
  var SA,SX : String;
  begin
    result := '';
    SA := TheListe;
    repeat
      SX := READTOKENST(SA);
      if SX <> '' then
      begin
        if result = '' then result := result + '"'+SX+'"'
                       else result := result + ',"'+SX+'"';
      end;
    until SX = '';
  end;

begin
  if (ListeChantiers.getString('TYPERESSOURCES')<>'') and (ListeChantiers.getString('TYPERESSOURCES')<>'<<Tous>>') then
  begin
    result := result + ' AND BNP_TYPERESSOURCE IN ('+ConstituelisteTR(ListeChantiers.getString('TYPERESSOURCES'))+') ';
  end;
end;

function TFplanningBTP.ConstitueRequete (Chantier,NatureP : string; Detaille : boolean) : Widestring;
begin
  if Modegestion = TmPcPlanCharge then
  begin
    if not Detaille then
    begin
      result := 'SELECT GL_AFFAIRE AS AFFAIRE,"" AS FONCTION,SUM(QTEBUDGET) AS BUDGET,'+
                '(SELECT SUM(BAT_NBHRS) FROM '+
                'BAFFECTCHANT WHERE BAT_AFFAIRE=GL_AFFAIRE) AS DEJAPLANN,'+
                '0 AS RESTE,AFF_LIBELLE AS LIBELLE,"" AS LIBFONCTION,'+
                ConstitueDatesCha (DateDepart) +','+
                ConstituePlannifieCh (DateDepart) +','+
                '0 AS OPENED,0 AS LEVEL," " AS INDIC,0 AS COLOR,0 AS BCOLOR '+
                'FROM BTPLANCHARGE ';
      if Chantier <> '' then result:= result + 'WHERE GL_AFFAIRE="'+chantier+'" '
                        else result := result + 'WHERE '+SetListeChantiers;
      result := result + SetListeTypeRessources;
      result := result+'GROUP BY GL_AFFAIRE,AFF_LIBELLE '+
                'ORDER BY GL_AFFAIRE';
    end else
    begin
      result := 'SELECT GL_AFFAIRE AS AFFAIRE,FONCTION,SUM(QTEBUDGET) AS BUDGET,'+
               '(SELECT SUM(BAT_NBHRS) FROM '+
               'BAFFECTCHANT WHERE BAT_AFFAIRE=GL_AFFAIRE AND BAT_FONCTION=FONCTION) AS DEJAPLANN, '+
               '0 AS RESTE,LIBELLE,LIBELLE AS LIBFONCTION,'+
                ConstitueDatesCha (DateDepart,true) +','+
                ConstituePlannifieCh (DateDepart) +','+
               '0 AS OPENED,1 AS LEVEL," " AS INDIC,COLOR,BCOLOR '+
               'FROM BTPLANCHARGE '+
               'WHERE GL_AFFAIRE="'+chantier+'" ';
      if natureP <> '' then result := result + ' AND FONCTION="'+NatureP+'" ';
      result := result + SetListeTypeRessources;
      result := result+ 'GROUP BY GL_AFFAIRE,FONCTION,LIBELLE,COLOR,BCOLOR '+
               'ORDER BY GL_AFFAIRE,FONCTION';
    end;
  end else if Modegestion = TmcPlanCharfnc then
  begin
    if not Detaille then
    begin
      result := 'SELECT FONCTION,"" AS AFFAIRE, SUM(QTEBUDGET) AS BUDGET,'+
                 '(SELECT SUM(BAT_NBHRS) FROM '+
                 'BAFFECTCHANT WHERE BAT_FONCTION=FONCTION) AS DEJAPLANN,'+
                 '0 AS RESTE,LIBELLE AS LIBNATURE,LIBELLE AS LIBFONCTION,'+
                 ConstitueDatesNat (DateDepart) +','+
                  ConstituePlannifieCh (DateDepart) +','+
                 '0 AS OPENED,0 AS LEVEL," " AS INDIC,COLOR,BCOLOR '+
                 'FROM BTPLANCHARGE '+
                 'WHERE ';
      result := result + SetListeChantiers;
      if natureP <> '' then result := result + 'AND FONCTION="'+NatureP+'" ';
      result := result + SetListeTypeRessources;
      result := result + 'GROUP BY FONCTION,LIBELLE,COLOR,BCOLOR '+
                 'ORDER BY FONCTION';
    end else
    begin
      result := 'SELECT FONCTION AS FONCTION,GL_AFFAIRE AS AFFAIRE,SUM(QTEBUDGET) AS BUDGET,'+
               '(SELECT SUM(BAT_NBHRS) FROM '+
               'BAFFECTCHANT WHERE BAT_FONCTION=FONCTION AND BAT_AFFAIRE=GL_AFFAIRE) AS DEJAPLANN, '+
               '0 AS RESTE,AFF_LIBELLE AS LIBNATURE,LIBELLE AS LIBFONCTION,'+
               ConstitueDatesNat (DateDepart,true) +','+
               ConstituePlannifieCh (DateDepart) +','+
               '0 AS OPENED,1 AS LEVEL," " AS INDIC,0 AS COLOR,0 AS BCOLOR '+
               'FROM BTPLANCHARGE '+
               'WHERE FONCTION="'+NatureP+'" ';
      if Chantier <> '' then  result := result + ' AND GL_AFFAIRE="'+chantier+'" '
                        else  result := result + ' AND '+SetListeChantiers+' ';
      result := result+'GROUP BY FONCTION,GL_AFFAIRE,AFF_LIBELLE,LIBELLE '+
               'ORDER BY FONCTION,GL_AFFAIRE';
    end;
  end;
end;

procedure TFplanningBTP.AddChampsPlanningCha(TOBL: TOB; Level,NumPhase,NumOrdre: integer; LibellePhase: string;Bcolor,Color : Tcolor);
begin
  TOBL.AddChampSupValeur('AFFAIRE','');
  TOBL.AddChampSupValeur('NUMPHASE',0);
  TOBL.AddChampSupValeur('NUMORDRE',0);
  TOBL.AddChampSupValeur('FONCTION','');
  TOBL.AddChampSupValeur('LIBFONCTION','');
  TOBL.AddChampSupValeur('LIBNATURE','');
  TOBL.AddChampSupValeur('BUDGET',0);
  TOBL.AddChampSupValeur('DEJAPLANN',0);
  TOBL.AddChampSupValeur('RESTE',0);
  TOBL.AddChampSupValeur('OPENED',0);
  TOBL.AddChampSupValeur('LEVEL',Level);
  TOBL.AddChampSupValeur('INDIC','');
  TOBL.AddChampSupValeur('BCOLOR',Bcolor);
  TOBL.AddChampSupValeur('COLOR',color);
  if NumOrdre <> 0 then
  begin
    TOBL.SetInteger('NUMORDRE',NumOrdre);
    TOBL.SetInteger('NUMPHASE',NumPhase);
  end;
  TOBL.SetString('LIBNATURE',LibellePhase);
  TOBL.SetString('LIBFONCTION',LibellePhase);
  AddOneDETAIL (TOBL);
end;

procedure TFplanningBTP.PositionneChamps (TOBL,TT : TOB);
begin
  TOBL.SetString('FONCTION',TT.GetString('FONCTION'));
end;

procedure TFplanningBTP.ConstitueTOBPlannChantier;

  procedure AddTOBFilles(TOBL : TOB);
  begin
    TOB.Create ('LE DETAIL',TOBL,-1);
    TOB.Create ('LES SUP',TOBL,-1);
  end;

  function FonctionExist (Fonction: string;NumPhase : integer) : TOB;
  var II : Integer;
  begin
    result := nil;
    for II := 0 to TobAffiches.detail.count -1 do
    begin
      if TobAffiches.detail[II].GetInteger('NUMPHASE') <  NumPhase then continue;
      if TobAffiches.detail[II].GetInteger('NUMPHASE') >  NumPhase then break;
      if TobAffiches.detail[II].GetString('FONCTION') = Fonction then
      begin
        result := TobAffiches.detail[II];
      end;
    end;
  end;

  function TrouveLigne (TOBSource,TOBA : TOB) : TOB;
  var II : integer;
      TT : TOB;
  begin
    result := nil;
    for II := 0 To TOBSource.detail.count -1 do
    begin
      TT := TOBSource.detail[II];
      if (TT.GetInteger('NUMORDRE')=TOBA.GetInteger('BAT_NUMORDRE')) and
         (TT.GetString('FONCTION') = TOBA.GetString('BAT_FONCTION')) then
      begin
        Result := TT;
        Break;
      end;
    end;
  end;

var SQl : string;
    ZZ : WideString;
    QQ : TQuery;
    TOBT,TT,TOBL,ThePhase,TOBD,TA,TB : TOB;
    II,JJ,KK : Integer;
    NumOrdre,NumPhase : Integer;
    DDay,FDay : TDateTime;
    Titre : string;
    Tn,RootTN: TTreeNode;
begin
  TTVARBO.Items.Clear; // reinit de l'arborescence
  Titre := self.Caption;
  RootTN := TTVARBO.items.add(nil, Titre);

  NumOrdre := 0;
  NumPhase := 0;
  ThePhase := nil;
  TOBD := TOB.Create('UNE GROOOOOSE TOB',nil,-1);
  TOBT := TOB.Create ('UNE TOB',nil,-1);
  // constitution de la structure du document
  SQl := 'SELECT * '+
         'FROM BTPLANNING '+
         'WHERE '+
         'GL_AFFAIRE="'+Chantier+'" '+
         'ORDER BY GL_NUMLIGNE,FONCTION';
  QQ := OpenSQL(SQl,True,-1,'',true);
  TOBT.LoadDetailDB('UNE LIGNE','','',QQ,false);
  Ferme(QQ);
  //
  for  II := 0 to TOBT.Detail.Count -1 do
  begin
    TT := TOBT.detail[II];
    if TT.getString('GL_TYPELIGNE')='TP1' then
    begin
      NumOrdre := 0;
      ThePhase := nil;
    end else if TT.getString('GL_TYPELIGNE')='DP1' then
    begin
      Inc(NumPhase);
      NumOrdre := TT.GetInteger ('GL_NUMORDRE');
      TOBL := TOB.Create ('UNE LIGNE',TobAffiches,-1);
      AddChampsPlanningCha (TOBL,0,NumPhase,NumOrdre,TT.GetString('GL_LIBELLE'),0,0);
      ThePhase := TOBL;
      Tn := TTVARBO.Items.AddChild(RootTN, TT.GetString('GL_LIBELLE'));
      Tn.Data := TOBL;
    end else
    begin
      if NumOrdre = 0 then
      begin
        Inc(NumPhase);
        NumOrdre := TT.GetInteger ('GL_NUMORDRE');
        TOBL := TOB.Create ('UNE LIGNE',TobAffiches,-1);
        AddChampsPlanningCha (TOBL,0,NumPhase,NumOrdre,'Hors Phase',16777215,0);
        ThePhase := TOBL;
        Tn := TTVARBO.Items.AddChild(RootTN, TT.GetString('GL_LIBELLE'));
        Tn.Data := TOBL;
      end;
      TOBL :=  FonctionExist(TT.GetString('FONCTION'),NumPhase);
      if TOBL = nil then
      begin
        TOBL := TOB.Create ('UNE LIGNE',TobAffiches,-1);
        TOBL.Data := ThePhase;
        AddChampsPlanningCha (TOBL,1,NumPhase,NumOrdre,TT.GetString('LIBELLE'),TT.GetInteger('BCOLOR'),TT.GetInteger('COLOR'));
        PositionneChamps (TOBL,TT);
        AddTOBFilles(TOBL);
      end;
      TOBL.SetDouble('BUDGET',TOBL.GetDouble('BUDGET')+TT.GetDouble('QTEBUDGET'));
    end;
  end;
  TOBT.free;
  TTVARBO.FullExpand;
  //
  ZZ := ConstitueConsoDetPlann (DateDepart);
  // recup et positionnemnt du détail
  SQl := 'SELECT B1.BAT_AFFAIRE,B1.BAT_NUMORDRE,B1.BAT_FONCTION,SUM(B1.BAT_NBHRS) AS DEJAPLANN, '+
          ZZ +' '+
         'FROM BAFFECTCHANT B1 '+
         'WHERE B1.BAT_AFFAIRE="'+Chantier+'" '+
         'GROUP BY B1.BAT_AFFAIRE,B1.BAT_NUMORDRE,B1.BAT_FONCTION';

  QQ := OpenSQL(SQl,True,-1,'',true);
  TOBD.LoadDetailDB('UNE TOB','','',QQ,false);
  Ferme(QQ);
  for II :=0 To TOBD.Detail.count -1 do
  begin
    TT := TOBD.detail[II];
    for JJ := 0 to TobAffiches.Detail.count -1 do
    begin
      TA := TobAffiches.detail[JJ];
      if (TA.GetInteger('NUMORDRE')=TT.GetInteger('BAT_NUMORDRE')) and
         (TA.Getstring('FONCTION')=TT.Getstring('BAT_FONCTION')) then
      begin
        TA.SetDouble('DEJAPLANN',TT.GetDouble('DEJAPLANN'));
        for KK := 1 to NBjours do
        begin
          ZZ := 'D'+InttoStr(KK);
          if TT.getDouble(ZZ)<>0 then TA.SetDouble(ZZ,TT.GetDouble(ZZ));
        end;
      end;
      TA.SetDouble('RESTE',TA.GetDouble('BUDGET')-TA.GetDouble('DEJAPLANN'));
      //
    end;
  end;
  TOBD.free;
  for II := 0 to TobAffiches.Detail.count -1 do
  begin
    TA := TobAffiches.detail[II];
    if TA.GetString('FONCTION')='' then continue;
    TB := TOB(TA.Data);
    if TB <> nil then
    begin
      // calcul sur la totalisation
      TB.SetDouble('BUDGET',TB.GetDouble('BUDGET')+TA.GetDouble('BUDGET'));
      TB.SetDouble('DEJAPLANN',TB.GetDouble('DEJAPLANN')+TA.GetDouble('DEJAPLANN'));
      TB.SetDouble('RESTE',TB.GetDouble('BUDGET')-TB.GetDouble('DEJAPLANN'));
      for KK := 1 to NBjours do
      begin
        ZZ := 'D'+InttoStr(KK);
        if TA.getDouble(ZZ)<>0 then TB.SetDouble(ZZ,TB.GetDouble(ZZ)+TA.GetDouble(ZZ));
      end;
    end;
  end;
  // recup du detail des plannification sur la période en vue de la mise a jour
  DDay := StrToDateTime(DateToStr(DateDepart)+' 00:00:00');
  FDay := StrToDateTime(DateToStr(DateFin)+' 23:59:59');
  SQl := 'SELECT *,1 AS DBPRESENT '+
         'FROM BAFFECTCHANT '+
         'WHERE BAT_AFFAIRE="'+Chantier+'" AND '+
         'BAT_DATE >= "'+USDATETIME(DDay)+'" AND '+
         'BAT_DATE <= "'+USDATETIME(FDay)+'" ';
  QQ := OpenSQL(SQl ,True,-1,'',true);
  TOBAffects.LoadDetailDB('BAFFECTCHANT','','',QQ,false);
  II := 0;
  if TOBAffects.detail.count > 0 then
  begin
    repeat
      TA := TOBAffects.Detail[II];
      TB := TrouveLigne (TOBAffiches,TA);
      if TB <> nil then
      begin
        TA.ChangeParent(TB.detail[0],-1);
      end else Inc(II);
    until II>= TOBAffects.Detail.Count ;
    //
  end;
  TOBAffects.ClearDetail;
  ferme (QQ);

end;

procedure TFplanningBTP.NBRRESSChange(Sender: TObject);
begin
  NBHrs.value := ARRONDI(NBRRESS.value * DureeJournee, V_PGI.OkDecQ);
end;

procedure TFplanningBTP.NBHRSExit(Sender: TObject);
begin
  NBRRESS.OnChange := nil;
  NBRRESS.Value := Ceil(NBHRS.Value / DureeJournee);
  NBRRESS.OnChange := NBRRESSChange;
end;

procedure TFplanningBTP.BREFRESHPClick(Sender: TObject);
begin
  if DetailModif then
  begin
    if PgiAsk ('ATTENTION : Vous allez perdre vos modifications.#13#10 Confirmez-vous ?') <> mryes then exit;
    detailModif := false;
  end;

  DateDepart := StrToDate(DATED.Text);
  DateFin := StrToDateTime(DATED_.Text+' 23:59:59');
  //
  NbJours :=  DaysBetween (DateDepart,DateFin)+1;
  BRefreshClick(self); 
end;

procedure TFplanningBTP.BPARAMSClick(Sender: TObject);
begin
  if TTPARAMS.visible then
  begin
    TTParams.visible := false;
    BPARAMS.Down := false;
  end else
  begin
    TTParams.visible := true;
    TTParams.BringToFront;
    BPARAMS.Down := true;
  end;
end;

procedure TFplanningBTP.TTPARAMSClose(Sender: TObject);
begin
  BPARAMS.Down := false;
end;

procedure TFplanningBTP.FormResize(Sender: TObject);
begin
  ReajusteEcran;
end;

procedure TFplanningBTP.TBVIEWARBOClick(Sender: TObject);
begin
  if TTARBO.visible then
  begin
    TTARBO.Visible := false;
    TBVIEWARBO.down := false;
  end else
  begin
    TBVIEWARBO.down := true;
    TTARBO.Visible := true;
  end;
end;

procedure TFplanningBTP.TTVARBOClick(Sender: TObject);
var Tn: TTreeNode;
  TOBL: TOB;
  Arow: integer;
  cancel : boolean;
begin
  Tn := TTVARBO.selected;
  TOBL := TOB(Tn.data);
  if (TOBL <> nil) then
  begin
    ARow := TOBL.GetIndex;
    G1.Row := Arow;
    G1RowEnter(self,Arow,cancel,false);
  end;
end;

procedure TFplanningBTP.TTARBOClose(Sender: TObject);
begin
  TBVIEWARBO.down := false;
end;

procedure TFplanningBTP.CAffHeuresClick(Sender: TObject);
begin
  G2.Invalidate;
end;

procedure TFplanningBTP.RBmodeAffichageClick(Sender: TObject);
begin
  if TRadioButton(Sender).Name = 'RBHeures' then
  begin
    ModeAffichage := TmAHeure;
    TLIBTOTALHRS.Caption := 'Total des heures';
    TLIBHRSPREV.Caption := 'Heures Prévues';
  end else
  begin
    ModeAffichage := TmaJours;
    TLIBHRSPREV.Caption := 'Jours Prévus';
    TLIBTOTALHRS.Caption := 'Total des jours';
  end;
  G2.Invalidate;
  GTHeures.Invalidate;
  G1.Invalidate;
  GFONCTION.Invalidate;
end;

procedure TFplanningBTP.GFONCTIONDrawCell(Sender: TObject; ACol,ARow: Integer; Rect: TRect; State: TGridDrawState);
var grid : THGrid;
    DD1,DD2,DDR : double;
    St,TheText : string;
    TT : TOB;
    TheColor : Tcolor;
begin
  if TOBTOTFONC.detail.count = 0 then exit;
  TT := TOBTOTFONC.detail[0];
  grid := (Sender as THgrid);
  if not Grid.visible then exit;
  Grid.Canvas.FillRect(Rect);
  ClearCanvas (grid,ACol,Rect);
  DrawButton(Grid,Rect,true);
  
  if IsDayClose (lesColonnesG2.items[Acol],JourFermeture) and not(gdSelected in state) then
  begin
    Grid.Canvas.Brush.Color := $f4f2a5;
    Grid.Canvas.FillRect(rect);
  end else if (gdSelected in state)  then
  begin
    Grid.Canvas.Brush.Color := clHighlight ;
    Grid.Canvas.FillRect(rect);
  end;
  DrawButton(Grid,Rect,true);

  St := 'PJ'+InttoStr(Acol+1);
  DD1 := TT.getDouble(St);
  IF IsDayClose (lesColonnesG2.items[Acol],JourFermeture) then thetext := ''
                                                          else TheText := GetDureeSt (DD1);
  OwnerDrawText(Grid,State,Acol,Rect,TheText,DT_SINGLELINE or DT_CENTER or DT_TOP or DT_EDITCONTROL);
  St := 'CJ'+InttoStr(Acol+1);
  DD2 := TT.getDouble(St);
  IF IsDayClose (lesColonnesG2.items[Acol],JourFermeture) then thetext := ''
                                                          else TheText := GetDureeSt (DD2);
  OwnerDrawText(Grid,State,Acol,Rect,TheText,DT_SINGLELINE or DT_CENTER or DT_VCENTER or DT_EDITCONTROL);
  //
  DDR := GetDureeDD(DD1) - GetDureeDD(DD2);
  //
  IF IsDayClose (lesColonnesG2.items[Acol],JourFermeture) then thetext := ''
                                                          else TheText := Format('%5.2f',[DDR]);
  if DDR < 0 then TheColor := clred else TheColor := 0;
  OwnerDrawText(Grid,State,Acol,Rect,TheText,DT_SINGLELINE or DT_CENTER or DT_BOTTOM or DT_EDITCONTROL,TheColor );
end;


function TFplanningBTP.ConstitueInfoFonction (Fonction: string) : string;

  procedure SetAbsencesSalaries (Fonction : string;  DateDeb,DateFin,DDay,FDay : TdateTime);
  var SQl : string;
      TOBConges,TOBC,TT : TOB;
      II, JJ : integer;
      Duree : double;
      ODayDep,ODayFin,Heure,DD : Tdatetime;
  begin
    if TOBTOTFONC.detail.count = 0 then exit;
    TT := TOBTOTFONC.detail[0];
    TOBConges := TOB.Create ('LES CONGES', nil,-1);
    TRY
      // issu du planning
      Sql := 'SELECT PCN_DATEDEBUTABS,PCN_DATEFINABS,PCN_DEBUTDJ,PCN_FINDJ,0 AS DATEDEBR, 0 AS DATEFINR FROM ABSENCESALARIE WHERE '+
             'PCN_SALARIE  IN '+
             '('+
             'SELECT ARS_SALARIE FROM RESSOURCE WHERE ARS_TYPERESSOURCE = "SAL" AND '+
             'ARS_FONCTION1 = "'+Fonction+'" AND ARS_FERME<>"X"';
      if Etablissement <> '' then
      begin
        Sql := Sql + ' AND ARS_ETABLISSEMENT="'+Etablissement+'"';
      end;
      SQL := SQL + ') '+
             'AND  PCN_ETATPOSTPAIE="VAL" AND '+
             '( '+ // 1
             '(PCN_DATEDEBUTABS>="' + UsDateTime(DDay) + '" AND PCN_DATEDEBUTABS<="' + UsDateTime(FDay) + '") OR '+
             '(PCN_DATEFINABS>="' + USDateTime(DDay) + '" AND PCN_DATEFINABS<="' + USDateTime(FDay) + '") OR ' +
             '(PCN_DATEDEBUTABS<="' + UsDateTime(DDay) + '" AND PCN_DATEFINABS>="' + UsDateTime(FDay) + '")' +
             ')'; //1
      TOBConges.LoadDetailDBFromSQL('ABSENCES',Sql,false);
      for II := 0 To TOBConges.detail.count -1 do
      begin
        TOBC := TOBConges.detail[II];
        
        if TOBC.GetDouble('DATEDEBR')=0 then
        begin
          DD := StrToDate(DateToStr(TOBC.getValue('PCN_DATEDEBUTABS')));
          if TOBC.getValue('PCN_DEBUTDJ')='MAT' then
          begin
            heure := GetDebutMatinee;
          end else
          begin
            heure := GetDebutApresMidi;
          end;
          TOBC.SetDateTime ('DATEDEBR',DD+heure);
          //
          DD := StrToDate(DateToStr(TOBC.getValue('PCN_DATEFINABS')));
          if TOBC.getValue('PCN_FINDJ')='MAT' then
          begin
            heure := GetFinMatinee;
          end else
          begin
            heure := GetFinApresMidi;
          end;
          TOBC.SetDateTime ('DATEFINR',DD+heure);
        end;
        OdayDep := DateDeb;
        for JJ := 1 to NbJours do
        begin
          DD := StrToDateTime(DateToStr(ODayDep));
          if (DD >= TOBC.GetDateTime('PCN_DATEDEBUTABS')) and (DD <= TOBC.GetDateTime('PCN_DATEFINABS')) then
          begin
            ODayDep := StrToDateTime(DateToStr(DD))+GetDebutMatinee;
            ODayFin := StrToDateTime(DateToStr(DD))+GetFinApresMidi;
            if DD = TOBC.GetDateTime('PCN_DATEDEBUTABS') then ODayDep := TOBC.GetDateTime('DATEDEBR');
            if DD = TOBC.GetDateTime('PCN_DATEFINABS') then ODayFin := TOBC.GetDateTime('DATEFINR');
            Duree := Arrondi(CalculDureeEvenement (ODayDep,ODayFin)/60,2);
            TT.SetDouble('CJ'+InttoStr(JJ),TT.GetDouble('CJ'+InttoStr(JJ))+Duree);
            TT.SetDouble('ABS'+InttoStr(JJ),TT.GetDouble('ABS'+InttoStr(JJ))+Duree);
          end;
          ODayDep := IncDay(ODayDep,1);
        end;
      end;
    FINALLY
      TOBConges.free;
    END;
  end;

  procedure SetEvtGRC (Fonction : string;  DateDeb,DateFin,DDay,FDay : TdateTime);
  var SQl : string;
      TOBGRC,TOBC,TT : TOB;
      II,JJ : integer;
      Duree : double;
      ODayDep,ODayFin,Heure,DD : Tdatetime;
  begin
    if TOBTOTFONC.detail.count = 0 then exit;
    TT := TOBTOTFONC.detail[0];
    TOBGRC := TOB.Create ('EVT GRC', nil,-1);
    TRY
      Sql := 'SELECT RAC_DATEACTION,RAC_HEUREACTION,RAC_DUREEACTION,0 AS DATEDEBR, 0 AS DATEFINR FROM ACTIONS WHERE '+
             'RAC_INTERVENANT IN ('+
             'SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_TYPERESSOURCE = "SAL" '+
             'AND ARS_FONCTION1 = "'+Fonction+'" AND ARS_FERME<>"X"';
      if Etablissement <> '' then
      begin
        Sql := Sql + ' AND ARS_ETABLISSEMENT="'+Etablissement+'"';
      end;
      SQl := SQL + ') '+
        'AND RAC_DATEACTION>="' + USDATETime(DateDeb) + '" AND RAC_DATEACTION<="' + USDATETime(DateFin) + '"';
      TOBGRC.LoadDetailDBFromSQL('ACTIONS',Sql,false);
      for II := 0 To TOBGRC.detail.count -1 do
      begin
        TOBC := TOBGRC.detail[II];

        if TOBC.GetDouble('DATEDEBR')=0 then
        begin
          DD := StrToDate(DateToStr(TOBC.getValue('RAC_DATEACTION')));
          heure := TOBC.getValue('RAC_HEUREACTION');
          TOBC.SetDateTime ('DATEDEBR',DD+heure);
          //
          TOBC.SetDateTime ('DATEFINR',AjouteDuree(TOBC.GetDateTime ('DATEDEBR'),TOBC.getInteger('RAC_DUREEACTION')));
        end;
        OdayDep := DateDeb;
        for JJ := 1 to NbJours do
        begin
          DD := StrToDateTime(DateToStr(ODayDep));
          if (DD = TOBC.GetDateTime('RAC_DATEACTION')) then
          begin
            ODayDep := TOBC.GetDateTime('DATEDEBR');
            ODayFin := TOBC.GetDateTime('DATEFINR');
            Duree := Arrondi(CalculDureeEvenement (ODayDep,ODayFin)/60,2);
            TT.SetDouble('CJ'+InttoStr(JJ),TT.GetDouble('CJ'+InttoStr(JJ))+Duree);
            TT.SetDouble('GRC'+InttoStr(JJ),TT.GetDouble('GRC'+InttoStr(JJ))+Duree);
          end;
          ODayDep := IncDay(ODayDep,1);
        end;
      end;
    FINALLY
      TOBGRC.free;
    END;
  end;

  procedure SetEvtInterventions (Fonction : string;  DateDeb,DateFin,DDay,FDay : TdateTime);
  var SQl : string;
      TOBINT,TOBC,TT : TOB;
      II,JJ : integer;
      Duree : double;
      ODayDep,ODayFin,Heure,DD : Tdatetime;
  begin
    if TOBTOTFONC.detail.count = 0 then exit;
    TT := TOBTOTFONC.detail[0];
    TOBINT := TOB.Create ('EVT APP', nil,-1);
    TRY
      Sql := 'SELECT BEP_DATEDEB,BEP_HEUREDEB,BEP_DATEFIN,BEP_HEUREFIN,0 AS DATEDEBS,0 AS DATEFINS,0 AS DATEDEBR, 0 AS DATEFINR FROM BTEVENPLAN WHERE '+
             'BEP_RESSOURCE IN ('+
             'SELECT ARS_RESSOURCE FROM RESSOURCE WHERE '+
             'AND ARS_FONCTION1 = "'+Fonction+'" AND ARS_FERME<>"X"';
      if Etablissement <> '' then
      begin
        Sql := Sql + ' AND ARS_ETABLISSEMENT="'+Etablissement+'"';
      end;
      SQl := SQL + ') AND '+
            '('+
            '(BEP_DATEDEB>="' + UsDateTime(DDay) + '" AND BEP_DATEDEB<="'  + UsDateTime(FDay) + '") '+
            'OR (BEP_DATEFIN>="'  + USDateTime(DDay) + '" AND BEP_DATEFIN<="'  + USDateTime(FDay) + '") '+
            'OR (BEP_DATEDEB<="'  + UsDateTime(DDay) + '" AND BEP_DATEFIN>="'  + UsDateTime(FDay) + '")'+
            ')';
      TOBINT.LoadDetailDBFromSQL('BTEVENPLAN',Sql,false);
      for II := 0 To TOBINT.detail.count -1 do
      begin
        TOBC := TOBINT.detail[II];

        if TOBC.GetDouble('DATEDEBR')=0 then
        begin
          DD := StrToDate(DateToStr(TOBC.getValue('BEP_DATEDEB')));
          heure := TOBC.getValue('BEP_HEUREDEB');
          TOBC.SetDateTime ('DATEDEBR',DD+heure);
          TOBC.SetDateTime ('DATEDEBS',DD);
          //
          DD := StrToDate(DateToStr(TOBC.getValue('BEP_DATEFIN')));
          heure := TOBC.getValue('BEP_HEUREFIN');
          TOBC.SetDateTime ('DATEFINR',DD+heure);
          TOBC.SetDateTime ('DATEFINS',DD);
        end;
        OdayDep := DateDeb;
        for JJ := 1 to NbJours do
        begin
          DD := StrToDate(DateToStr(ODayDep));
          if (DD >= TOBC.GetDateTime('DATEDEBS')) and (DD <= TOBC.GetDateTime('DATEFINS')) then
          begin
            ODayDep := StrToDateTime(DateToStr(DD))+GetDebutMatinee;
            ODayFin := StrToDateTime(DateToStr(DD))+GetFinApresMidi;
            if DD = TOBC.GetDateTime('DATEDEBS') then ODayDep := TOBC.GetDateTime('DATEDEBR');
            if DD = TOBC.GetDateTime('DATEFINS') then ODayFin := TOBC.GetDateTime('DATEFINR');
            Duree := Arrondi(CalculDureeEvenement (ODayDep,ODayFin)/60,2);
            TT.SetDouble('CJ'+InttoStr(JJ),TT.GetDouble('CJ'+InttoStr(JJ))+Duree);
            TT.SetDouble('INT'+InttoStr(JJ),TT.GetDouble('INT'+InttoStr(JJ))+Duree);
          end;
          ODayDep := IncDay(ODayDep,1);
        end;
      end;
    FINALLY
      TOBINT.free;
    END;
  end;

  procedure SetEvtMateriel (Fonction : string;  DateDeb,DateFin,DDay,FDay : TdateTime);
  var SQl : string;
      TOBMAT,TOBC,TT : TOB;
      II,JJ : integer;
      Duree : double;
      ODayDep,ODayFin,Heure,DD : Tdatetime;
  begin
    if TOBTOTFONC.detail.count = 0 then exit;
    TT := TOBTOTFONC.detail[0];
    TOBMAT := TOB.Create ('EVT MAT', nil,-1);
    TRY
      Sql := 'SELECT BEM_DATEDEB,BEM_NBHEURE,0 AS DATEDEBS, 0 AS DATEFINS,0 AS DATEDEBR, 0 AS DATEFINR FROM BTEVENTMAT WHERE '+
             'BEM_CODEMATERIEL IN ( SELECT DISTINCT BMA_CODEMATERIEL FROM BTMATERIEL WHERE BMA_RESSOURCE IN ('+
             'SELECT ARS_RESSOURCE FROM RESSOURCE WHERE '+
             'AND ARS_FONCTION1 = "'+Fonction+'" AND ARS_FERME<>"X"';
      if Etablissement <> '' then
      begin
        Sql := Sql + ' AND ARS_ETABLISSEMENT="'+Etablissement+'"';
      end;
      SQl := SQL + ') AND BMA_RESSOURCE <> '') AND '+
            '('+
            '(BEM_DATEDEB>="' + UsDateTime(DDay) + '" AND BEM_DATEFIN<="'  + UsDateTime(FDay) + '") '+
            'OR (BEM_DATEDEB>="'  + USDateTime(DDay) + '" AND BEM_DATEFIN<="'  + USDateTime(FDay) + '") '+
            'OR (BEM_DATEDEB<="'  + UsDateTime(DDay) + '" AND BEM_DATEFIN>="'  + UsDateTime(FDay) + '")'+
            ')';
      TOBMAT.LoadDetailDBFromSQL('BTEVENTMAT',Sql,false);
      for II := 0 To TOBMAT.detail.count -1 do
      begin
        TOBC := TOBMAT.detail[II];

        if TOBC.GetDouble('DATEDEBR')=0 then
        begin
          DD := StrToDate(DateToStr(TOBC.getValue('BEM_DATEDEB')));
          heure := GetDebutMatinee;
          TOBC.SetDateTime ('DATEDEBR',DD+heure);
          TOBC.SetDateTime ('DATEDEBS',DD);
          //
          DD := TOBC.GetDateTime('DATEDEBR');
          heure := TOBC.getValue('BEM_NBHEURE');
          TOBC.SetDateTime ('DATEFINR',DD+heure);
          TOBC.SetDateTime ('DATEFINS',StrToDate(DateToStr(TOBC.GetDateTime ('DATEFINR'))));
        end;
        OdayDep := DateDeb;
        for JJ := 1 to NbJours do
        begin
          DD := StrToDate(DateToStr(ODayDep));
          if (DD >= TOBC.GetDateTime('DATEDEBS')) and (DD <= TOBC.GetDateTime('DATEFINS')) then
          begin
            ODayDep := StrToDateTime(DateToStr(DD))+GetDebutMatinee;
            ODayFin := StrToDateTime(DateToStr(DD))+GetFinApresMidi;
            if DD = TOBC.GetDateTime('DATEDEBS') then ODayDep := TOBC.GetDateTime('DATEDEBR');
            if DD = TOBC.GetDateTime('DATEFINS') then ODayFin := TOBC.GetDateTime('DATEFINR');
            Duree := Arrondi(CalculDureeEvenement (ODayDep,ODayFin)/60,2);
            TT.SetDouble('CJ'+InttoStr(JJ),TT.GetDouble('CJ'+InttoStr(JJ))+Duree);
            TT.SetDouble('MAT'+InttoStr(JJ),TT.GetDouble('MAT'+InttoStr(JJ))+Duree);
          end;
          ODayDep := IncDay(ODayDep,1);
        end;
      end;
    FINALLY
      TOBMAT.free;
    END;
  end;

  procedure SetEvtPlanChantier (Fonction : string);
  var SQl,Zone : Widestring;
      II,JJ: integer;
      TOBPLAN,TT : TOB;
  begin
    if TOBTOTFONC.detail.count = 0 then exit;
    TT := TOBTOTFONC.detail[0];
    TOBPLAN := TOB.Create ('UN TOTAL',nil,-1);
    TRY
      Sql :='SELECT ' + ConstitueCumFonctionJour (Fonction);
      //
      TOBPLAN.LoadDetailDBFromSQL ('UNE PLANNIF',Sql);
      for II := 0 to TOBPlan.detail.count -1 do
      begin
        for JJ := 1 to NbJours do
        begin
          Zone := 'TJ'+InttoStr(JJ);
          TT.SetDouble('CJ'+InttoStr(JJ),TT.GetDouble('CJ'+InttoStr(JJ))+TOBPlan.detail[II].GetDouble(Zone));
          TT.SetDouble('CHA'+InttoStr(JJ),TT.GetDouble('CHA'+InttoStr(JJ))+TOBPlan.detail[II].GetDouble(Zone));
        end;
      end;
    FINALLY
      TOBPLAN.free;
    END;
  end;


var II : Integer;
    DDay,FDay,DateD,DateF : tdatetime;
    QQ : Tquery;
    MAXHeures : double;
    TT : TOB;
begin
  if TOBTOTFONC.detail.count = 0 then exit;
  TT := TOBTOTFONC.detail[0];
  DDay := DateDepart;
  DDay := StrToDateTime(DateToStr(DDay)+' 00:00:00');
  fDay := StrToDateTime(DateToStr(IncDay(DDay,nbJours))+' 23:59:59');
  DateD := StrToDate(DateToStr(DateDepart));
  DateF := StrToDate(DateToStr(IncDay(DateDepart,nbJours)));
  For II := 1 to NBjours do
  begin
    TT.SetDouble('CJ'+InttoStr(II),0);
  end;
  //
  // Recup des absences salaries
  SetAbsencesSalaries (fonction ,DateD,DateF,DDay,FDay);
  // Recup des Evt GRC
  SetEvtGRC (fonction ,DateD,DateF,DDay,FDay);
  // Recup des Evt Interventions
  SetEvtInterventions (fonction ,DateD,DateF,DDay,FDay);
  // Recup des Evt Interventions
  SetEvtMateriel (fonction , DateD,DateF,DDay,FDay);
  // Recup des plannif plan de charge
  SetEvtPlanChantier (fonction);
  //
  MaxHeures := 0;
  QQ := OpenSql ('SELECT COUNT(ARS_SALARIE) NBELT FROM RESSOURCE WHERE ARS_FONCTION1="'+Fonction+'"',True,1,'',true);
  if not QQ.eof then
  begin
    MAXHeures := QQ.fields[0].AsInteger * DureeJournee ;
  end;
  ferme (QQ);
  //
  For II := 1 to NBjours do
  begin
    //
    TT.SetDouble('PJ'+InttoStr(II),MaxHeures);
    TT.SetDouble('RJ'+InttoStr(II),TT.GetDouble('PJ'+InttoStr(II))- TT.GetDouble('CJ'+InttoStr(II)));
  end;
end;

procedure TFplanningBTP.ConstitueCumulFonction(Fonction: string);
var II : integer;
    TT : TOB;
begin
  if TOBTOTFONC.detail.count = 0 then exit;
  TT := TOBTOTFONC.detail[0];
  // Deduire absencesalarie ,rac (grc) et interv
  if (Fonction <> '---') and (Fonction <> '') then
  begin
    for II := 1 to NbJours do
    begin
      TT.SetDouble('PJ'+InttoStr(II),0); // plannifiable
      TT.SetDouble('CJ'+InttoStr(II),0); // consommé (plannifié)
      TT.SetDouble('RJ'+InttoStr(II),0); // Reste
      TT.SetDouble('ABS'+InttoStr(II),0); // Reste
      TT.SetDouble('GRC'+InttoStr(II),0); // Reste
      TT.SetDouble('INT'+InttoStr(II),0); // Reste
      TT.SetDouble('MAT'+InttoStr(II),0); // Reste
      TT.SetDouble('CHA'+InttoStr(II),0); // Reste
    end;
    //
    ConstitueInfoFonction (Fonction);
    //
  end;
end;

procedure TFplanningBTP.ShowThisHint(var HintStr: String; var CanShow: Boolean; var HintInfo: THintInfo);

  function ConstitueLibelleHint(TT : TOB; Colonne : integer) : string;
  begin
    result := '';
    if TT.GetDouble('ABS'+InttoStr(Colonne))<>0 then result := result + 'Absence         --> '+GetDureeSt(TT.GetDouble('ABS'+InttoStr(Colonne)));
    if TT.GetDouble('GRC'+InttoStr(Colonne))<>0 then
    begin
      if result <> '' then result := result + #13#10;
      result := result + 'Action GRC     - '+GetDureeSt(TT.GetDouble('GRC'+InttoStr(Colonne)));
    end;
    if TT.GetDouble('INT'+InttoStr(Colonne))<>0 then
    begin
      if result <> '' then result := result + #13#10;
      result := result + 'Interventions  - '+GetDureeSt(TT.GetDouble('INT'+InttoStr(Colonne)));
    end;
    if TT.GetDouble('MAT'+InttoStr(Colonne))<>0 then
    begin
      if result <> '' then result := result + #13#10;
      result := result + 'Plan. Matériel - '+GetDureeSt(TT.GetDouble('MAT'+InttoStr(Colonne)));
    end;
    if TT.GetDouble('CHA'+InttoStr(Colonne))<>0 then
    begin
      if result <> '' then result := result + #13#10;
      result := result + 'Chantier       - '+GetDureeSt(TT.GetDouble('CHA'+InttoStr(Colonne)));
    end;
  end;

var Arow,Acol : Integer;
    TT : TOB;
begin
  if HintInfo.HintControl = GFONCTION then
  begin
    if TOBTOTFONC.detail.count = 0 then exit;
    TT := TOBTOTFONC.detail[0];
    Arow := 0;
    Acol := 0;
    GFonction.MouseToCell(HintInfo.CursorPos.X, HintInfo.CursorPos.Y, Acol, Arow);
    HintInfo.CursorRect := GFonction.CellRect(Acol,Arow);
    HintInfo.HintPos.X := HintInfo.HintPos.X + 20;
    HintStr := ConstitueLibelleHint(TT,(Acol+1));
    canshow := true;
  end;
end;

procedure TFplanningBTP.GFONCTIONCellenter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  cancel := true;
end;

procedure TFplanningBTP.BImprimerClick(Sender: TObject);
begin
  PrintThisScreen(Tform(self),self.Caption);
end;

procedure TFplanningBTP.AffectRessourcesClick(Sender: TObject);
var TT : TOB;
    chantier : string;
    Phase : integer;
begin
  TT := GetTOBLigne(TobAffiches,G1.row); if TT = nil then Exit;
  if (Modegestion = TmPcPlanningCha)then
  begin
    Chantier := self.Chantier;
  end else
  begin
    chantier := TT.GetString('AFFAIRE');
  end;
  Phase := TT.GetInteger('NUMORDRE');
  OuvrePlanningRessources(Chantier,Phase,DateDepart,taModif);
  if (Modegestion = TmPcPlanningCha)then
  begin
    ReloadPlannifChantier(Chantier);
  end else
  begin
    RefreshChantier(TT);
  end;
end;

procedure TFplanningBTP.ReloadPlannifChantier(Chantier : string);

  procedure  ReinitThisPlannif (TT : TOB);
  var JJ : integer;
  begin
    for JJ := 1 to NbJours do
    begin
      TT.SetDouble('PLA'+InttoStr(JJ),0);
    end;
  end;
var II,I : integer;
    TT : TOB;
begin
  For ii := 0 to TOBAffiches.Detail.count -1 do
  begin
    TT := TOBAffiches.Detail[II];
    ReinitThisPlannif (TT);
    G2HauteurLigne (II);
  end;
  AjoutePlannifieChantier (nil,Chantier);

  G2.Invalidate;
end;

procedure TFplanningBTP.AjoutePlannifieChantier (TOBAdd : TOB = nil; ChantierF : string = '';FonctionF : string='');

procedure AffectePlannif (TT,TOBC : TOB ;DateD : TdateTime);
var ODayDep,ODayFin : TdateTime;
    JJ : integer;
    DD : TDateTime;
    Duree : double;
begin
  OdayDep := DateD;
  //
  for JJ := 1 to NbJours do
  begin
    DD := StrToDate(DateToStr(ODayDep));
    if (DD >= TOBC.GetDateTime('DATEDEBS')) and (DD <= TOBC.GetDateTime('DATEFINS')) then
    begin
      ODayDep := StrToDateTime(DateToStr(DD))+GetDebutMatinee;
      ODayFin := StrToDateTime(DateToStr(DD))+GetFinApresMidi;
      if DD = TOBC.GetDateTime('DATEDEBS') then ODayDep := TOBC.GetDateTime('DATEDEBR');
      if DD = TOBC.GetDateTime('DATEFINS') then ODayFin := TOBC.GetDateTime('DATEFINR');
      Duree := Arrondi(CalculDureeEvenement (ODayDep,ODayFin)/60,2);
      TT.SetDouble('PLA'+InttoStr(JJ),TT.GetDouble('PLA'+InttoStr(JJ))+Duree);
    end;
    ODayDep := IncDay(ODayDep,1);
  end;
end;

var SQl : string;
    TOBINT,TOBC,TT,TOBAT : TOB;
    II : integer;
    Heure,DD,DateD,DDay,FDay : Tdatetime;
begin
  DDay := StrToDate(DateToStr(DateDepart));
  fDay := StrToDate(DateToStr(IncDay(DDay,nbJours)));
  DateD := StrToDate(DateToStr(DateDepart));
  if TOBADD <> nil then TOBAT := TOBADD
                   else TOBAT := TOBaffiches;  
  TOBINT := TOB.Create ('EVT APP', nil,-1);
  TRY
    Sql := 'SELECT BEC_AFFAIRE,BEC_NUMPHASE,BEC_FONCTION,BEC_DATEDEB,BEC_HEUREDEB,BEC_DATEFIN,BEC_HEUREFIN,0 AS DATEDEBS,0 AS DATEFINS,0 AS DATEDEBR, 0 AS DATEFINR '+
           'FROM BTEVENTCHA '+
           'WHERE ';
    if (Modegestion = TmPcPlanningCha) then
    begin
      Sql := Sql + 'BEC_AFFAIRE="'+Chantier+'" AND ';
    end else
    begin
      if ChantierF <> '' then SQl := Sql + 'BEC_AFFAIRE="'+Chantierf+'" AND ';
      if FonctionF <> '' then Sql := Sql + 'BEC_FONCTION="'+Fonctionf +'" AND ';
      if (ChantierF = '') then Sql := Sql + SetListeChantiers ('BEC')+ ' AND ';
    end;
    SQl := Sql + '('+
          '(BEC_DATEDEB>="' + UsDateTime(DDay) + '" AND BEC_DATEDEB<="'  + UsDateTime(FDay) + '") '+
          'OR (BEC_DATEFIN>="'  + USDateTime(DDay) + '" AND BEC_DATEFIN<="'  + USDateTime(FDay) + '") '+
          'OR (BEC_DATEDEB<="'  + UsDateTime(DDay) + '" AND BEC_DATEFIN>="'  + UsDateTime(FDay) + '")'+
          ')';
    TOBINT.LoadDetailDBFromSQL('BTEVENTCHA',Sql,false);
    for II := 0 To TOBINT.detail.count -1 do
    begin
      TOBC := TOBINT.detail[II];

      if TOBC.GetDouble('DATEDEBR')=0 then
      begin
        DD := StrToDate(DateToStr(TOBC.getValue('BEC_DATEDEB')));
        heure := TOBC.getValue('BEC_HEUREDEB');
        TOBC.SetDateTime ('DATEDEBR',DD+heure);
        TOBC.SetDateTime ('DATEDEBS',DD);
        //
        DD := StrToDate(DateToStr(TOBC.getValue('BEC_DATEFIN')));
        heure := TOBC.getValue('BEC_HEUREFIN');
        TOBC.SetDateTime ('DATEFINR',DD+heure);
        TOBC.SetDateTime ('DATEFINS',DD);
      end;
      if (Modegestion = TmPcPlanningCha) then
      begin
        if TOBAT.detail.count > 0 then TT := TOBAT.FindFirst(['NUMORDRE','FONCTION'],[TOBC.GetString('BEC_NUMPHASE'),''],true)
                                  else TT := TOBAT;
        if TT <> nil then
        begin
          AffectePlannif (TT,TOBC,DateD);
        end;
        if TOBC.GetString('BEC_FONCTION') <> '' then
        begin
          TT := TOBAT.FindFirst(['NUMORDRE','FONCTION'],[TOBC.GetString('BEC_NUMPHASE'),TOBC.GetString('BEC_FONCTION')],true);
          if TT <> nil then
          begin
            AffectePlannif (TT,TOBC,DateD);
          end;
        end;
      end else if (Modegestion = TmPcPlanCharge) then  // plan de charge par chantier
      begin
        if TOBAT.detail.count > 0 then TT := TOBAT.FindFirst(['AFFAIRE','FONCTION'],[TOBC.GetString('BEC_AFFAIRE'),''],true)
                                  else TT := TOBAT;
        if TT <> nil then
        begin
          AffectePlannif (TT,TOBC,DateD);
        end;
        if TOBC.GetString('BEC_FONCTION') <> '' then
        begin
          TT := TOBAT.FindFirst(['AFFAIRE','FONCTION'],[TOBC.GetString('BEC_AFFAIRE'),TOBC.GetString('BEC_FONCTION')],true);
          if TT <> nil then
          begin
            AffectePlannif (TT,TOBC,DateD);
          end;
        end;
      end else if (Modegestion = TmcPlanCharfnc) then // Plan de charge par fonction
      begin
        if TOBAT.detail.count > 0 then TT := TOBAT.FindFirst(['AFFAIRE','FONCTION'],['',TOBC.GetString('BEC_FONCTION')],true)
                                  else TT := TOBAT;
        if TT <> nil then
        begin
          AffectePlannif (TT,TOBC,DateD);
        end;
        if TOBC.GetString('BEC_FONCTION') <> '' then
        begin
          TT := TOBAT.FindFirst(['AFFAIRE','FONCTION'],[TOBC.GetString('BEC_AFFAIRE'),TOBC.GetString('BEC_FONCTION')],true);
          if TT <> nil then
          begin
            AffectePlannif (TT,TOBC,DateD);
          end;
        end;
      end;
    end;
  FINALLY
    TOBINT.free;
  END;
end;

end.
