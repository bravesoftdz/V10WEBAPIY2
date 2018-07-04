unit SaisiePL;

interface

uses Windows,
  Messages,
  SysUtils,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  Hent1,
  Hctrls,
  Hplanning,
  {$ifndef eaglclient}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$endif}
  uTob,
  HTB97,
  ExtCtrls,
  HPanel,
  HmsgBox,
  StdCtrls,
  Mask,
  UtilXls,
  HStatus;

const
  XPWM_APPSTARTUP = WM_USER + 1;
  MAX_LEVEL_ETAT = 5;

type

  TTypeCumul = (tcJour, tcPourcentage);

  TTypeUserPlanning = (upUser, upResponsable);

  TTypePlanning = (tpPrestation, tpAffaire, tpRessource, tpAffairePrestation, tpAffaireRessource, tpPrestationRessource, tpAffairePrestationRessource);

  TRecordParamsPlanning = record
    LesCriteres: string ;
    T: TTypePLanning;
    Ressource, Affaire, Prestation: string;
    D, F: TDateTime;
    CLP: Boolean;
    TypeCUmul: TTypeCumul;
    AvecLesConges: Boolean;
    LesPrestations: TOB;
    LesAffaires: TOB;
    LesRessources: TOB;
    Responsable: Boolean;
    FClient: string;
  end;

  TFichePlanning = class(TForm)
    PanelPlanning: THPanel;
    Panel1: TToolWindow97;
    BAide: TToolbarButton97;
    Bannuler: TToolbarButton97;
    Bprint: TToolbarButton97;
    Bperso: TToolbarButton97;
    BNextMonth: TToolbarButton97;
    BPrevMonth: TToolbarButton97;
    LabelDate: THLabel;
    Bexcel: TToolbarButton97;
    Dock971: TDock97;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure BpersoClick(Sender: TObject);
    procedure BannulerClick(Sender: TObject);
    procedure BPrevMonthClick(Sender: TObject);
    procedure BNextMonthClick(Sender: TObject);
    procedure BprintClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BexcelClick(Sender: TObject);
  PRIVATE
    iValues: array of double;
    function LoadData: Boolean;
    procedure LoadConfigPlanning;
    procedure InitPlanning;
    procedure GoPlanning;
    procedure WMAppStartup(var msg: TMessage); MESSAGE XPWM_APPSTARTUP;
    procedure PlanningDblClick(Sender: TObject);
    procedure AvertirApplication(Sender: TObject; FromItem: TOB; ToItem: TOB; Actions: THPlanningAction);
    procedure CalculTotalColonne;
    procedure UpdateDateDuJour;
    procedure AjoutAffaire(Code, Code1: string; Valeur: Double = 0.00);
    procedure AjoutPrestation(Code, Code1: string; Valeur: Double = 0.00);
    procedure AjoutRessource(Code, Code1: string; Valeur: Double = 0.00);
    function InitTobRessource: TOB;
    { Déclarations privées }
  PUBLIC
    { Déclarations publiques }
    LaRessource: string;
    LaAffaire: string;
    LaPrestation: string;
    LaDateDuJour: TDateTime;
    LaDateDeFin: TDateTime;
    LePlanning: THPlanning;
    LeClient: string;
    TypePlanning: TTypePlanning;
    DateDeDebut: TDateTime;
    DateDeFin: TDateTime;

    { Les Tob de travail }
    LesRessources: TOB;
    LesEtats: TOB;
    LesItems: TOB;
    LesPrestations: TOB;
    LesAffaires: TOB;
    LesActivites: TOB;
    LesSalaries: TOB;
    LesCumuls: TOB;
    LesCriteres: string ;

    { Dates de chargement }
    DateDebutLoad, DateFinLoad: TDateTime;

    { Couleur de début et de fin des états }
    eColorStart: TColor;
    eColorEnd: TColor;

    TypeUser: TTypeUserPlanning;

    TmpDate: TDateTime;

    CLinePLanning: Boolean;

    AvecLesConges: Boolean;
    TCumul: TTypeCumul;

    FFormat: string;

    Multiplicateur: Double;
    Responsable: Boolean;

  end;

  (***
  procedure eCongePlanning(T: TTypePlanning; Ressource: string; D: TDateTime; TypeCumul: TTypeCumul = tcPourcentage); overload;
  ***)
procedure eCongePlanning(P: TRecordParamsPlanning); overload;

implementation

{$R *.DFM}
uses SaisiePersoPlanning, eTempsDetail;

type
  TTypeFunc = procedure(Code, Code1: string; Valeur: Double = 0) of object;


function MakeChaine(Src: string): string;
var
  S: string;
begin
  Result := '';
  S := ReadTokenSt(Src);
  while S <> '' do
  begin
    if Result = '' then Result := '"' + S + '"' else Result := Result + ',"' + S + '"';
    S := ReadTokenSt(Src);
  end;
end;


procedure eCongePlanning(P: TRecordParamsPlanning); OVERLOAD;
begin
  with TFichePlanning.Create(Application) do
  begin
    TCumul := P.TypeCumul;
    TypePlanning := P.T;
    CLinePlanning := P.CLP;
    DateDeDebut := P.D;
    LeClient := P.Fclient;
    Responsable := P.Responsable;
    DateDeFin := P.F;
    LaRessource := MakeChaine(P.Ressource);
    LaAffaire := MakeChaine(P.Affaire);
    LaPrestation := MakeChaine(P.Prestation);
    LaDateDeFin := P.F;
    AvecLesConges := P.AvecLesConges;
    LaDateDuJour := DateDeDebut;
    TypeUser := upResponsable;
    LesAffaires := P.LesAffaires;
    LesPrestations := P.LesPrestations;
    LesSalaries := P.LesRessources;
    LesCriteres := P.LesCriteres ;

    { Tri }
    if LesSalaries <> nil then if LesSalaries.Detail.Count > 0 then LesSalaries.Detail.Sort('SALARIE');

    ShowModal;
    Free;
  end;
end;

procedure TFichePlanning.FormCreate(Sender: TObject);
begin
  CLinePlanning := False;
  LaRessource := '';
  LaAffaire := '';
  LaPrestation := '';
  LesRessources := TOB.Create('Les_ressources', nil, -1);
  LesEtats := TOB.Create('Les_états', nil, -1);
  LesItems := TOB.Create('Les_items', nil, -1);
  LesActivites := TOB.Create('Les_activités', nil, -1);
  LesCumuls := TOB.Create('Les_cumuls', nil, -1);
  LeClient := '' ;
  LePlanning := nil;

  BExcel.Visible := OfficeExcelDispo;
end;

procedure TFichePlanning.FormShow(Sender: TObject);
var
  Q: TQuery;
  S: string;
begin
  FFormat := '0.##';
  Multiplicateur := 1;
  if TCumul = tcPourCentage then
  begin
    FFormat := '##0';
    Multiplicateur := 100;
  end;

  S := '';
  if TypeUSer <> upResponsable then
  begin
    { Récupération de la ressource complète }
    Q := OpenSQl('SELECT ARS_LIBELLE,ARS_LIBELLE2 FROM RESSOURCE WHERE ARS_SALARIE="' + V_PGI.UserSalarie + '"', True, 1);
    if not Q.EOF then S := Q.FindField('ARS_LIBELLE').AsString + ' ' + Q.FindField('ARS_LIBELLE2').AsString;
    Ferme(Q);
    S := ' pour ' + s;
  end;

  { Le caption de la forme }
  Caption := 'Planning ';
  case TypePlanning of
    tpPrestation: Caption := Caption + 'par prestation' + S;
    tpAffaire: Caption := Caption + 'par affaire' + S;
    tpRessource: Caption := Caption + 'par ressource' + S;
    tpAffairePrestation: Caption := Caption + 'croisé Affaire/Prestation' + s;
  end;

  { Création du planning }
  InitPlanning;

  DateDebutLoad := DateDeDebut;
  DateFinLoad := DateDeFin;

  { Chargement des données }
  if not LoadData then
  begin
    HShowMessage('0;' + Caption + ';Aucune activité disponible.;I;O;O;O', '', '');
    PostMessage(Handle, XPWM_APPSTARTUP, 0, 0);
  end else

    TmpDate := DateDeDebut;

  { Affichage du planning }
  GoPlanning;
end;

procedure TFichePlanning.WMAppStartup(var msg: TMessage);
begin
  CLose;
end;

procedure TFichePlanning.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TFichePlanning.FormDestroy(Sender: TObject);
begin
  LesCumuls.Free;
  LesRessources.Free;
  LesEtats.Free;
  LesItems.Free;
  LesActivites.Free;

  LePlanning.Free;
end;

{ ***************************************************************** }

procedure TFichePlanning.LoadConfigPlanning;
var
  S, ST: string;
  L: TStrings;
  I: Integer;
begin
  L := TStringList.Create;
  S := GetSynRegKey('ConfigPlanningTemps', S, True);
  if S <> '' then
  begin
    L.Text := S;

    { Les couleurs Fond+Fériés}
    S := L.Strings[0];
    LePlanning.ColorBackground := StringToColor(ReadTokenSt(S));
    LePlanning.ColorJoursFeries := StringToColor(ReadTokenSt(S));
    St := ReadTokenSt(S);
    if St <> '' then
    begin
      try
        LePlanning.ColorOfSaturday := StringToColor(St);
        LePlanning.ColorOfSunday := LePlanning.ColorOfSaturday;
      except
        on E: Exception do ;
      end;
    end
    else
    begin
      LePlanning.ColorOfSaturday := LePlanning.ColorJoursFeries;
      LePlanning.ColorOfSunday := LePlanning.ColorJoursFeries;
    end;
    St := ReadTokenSt(S);
    if St <> '' then
    begin
      try
        LePlanning.ColorSelection := StringToCOlor(St) ;
      except
        LePlanning.ColorSelection := LePlanning.ColorBackground;
      end;
    end else LePlanning.ColorSelection := LePlanning.ColorBackground; // XP 30-06-2005

    { La forme graphique }
    S := L.Strings[1];
    LePlanning.FormeGraphique := AglPlanningStringToFormeGraphique(ReadTokenSt(S));

    { Hauteur,Largeur, cumuldate et mode de cumul }
    S := L.Strings[2];
    LePlanning.RowSizeData := ValeurI(ReadTokenSt(S));
    LePlanning.ColSizeData := ValeurI(ReadTokenSt(S));
    LePlanning.ActiveLigneGroupeDate := ReadTokenSt(S) = '1';
    LePlanning.CumulInterval := AglPlanningStringToCumulInterval(ReadTokenSt(S));
    St := ReadTokenSt(S);
    if St = '' then
      if (LePlanning.ColSizeData = 15) and (LePlanning.CumulInterval = pciMois) then St := 'dd' else St := 'dd/mm';
    LePlanning.DateFormat := St;
    LePlanning.RefreshEtats;

    { Hauteur,Largeur, cumuldate et mode de cumul }
    S := L.Strings[3];
    if S <> '' then
      { Les couleurs des niveaux }
      for i := 0 to LesEtats.Detail.Count - 1 do
        if not (Copy(LesEtats.Detail[i].GetValue('E_CODE'), 1, 1) = 'C') then
          LesEtats.Detail[i].PutValue('E_COULEURFOND', ReadTokenSt(S));
  end
  else
  begin
    { Les valeurs par défaut }
    LePlanning.ColorBackground := clBtnFace;
    LePlanning.ColorJoursFeries := clInfoBk;
    LePlanning.ColorOfSaturday := clInfoBk;
    LePlanning.ColorOfSunday := clInfoBk;
  end;
  L.Free;
end;

procedure TFichePlanning.InitPlanning;
begin
  { Nettoyage des TOBs }
  LesRessources.ClearDetail;
  LesItems.ClearDetail;
  LesActivites.ClearDetail;
  LesEtats.ClearDetail;

  { Création du planning }
  if LePlanning <> nil then LePlanning.Free;
  LePlanning := THplanning.Create(Self);

  { Les évènements }
  LePlanning.OnDblClick := PlanningDblClick;
  //  LePlanning.OnDblClickSpec:= TestSouad ;
  LePlanning.OnAvertirApplication := AvertirApplication;

  { Le parent }
  LePlanning.Parent := PanelPLanning;

  { Agrandissement }
  LePlanning.Align := alClient;

  LePlanning.ActiveSaturday := True;
  LePlanning.ActiveSunday := True;

  { C'est commun à tous les plannings }
  LePlanning.ActiveLigneGroupeDate := True;
  LePlanning.Interval := piJour;
  LePlanning.CumulInterval := pciSemaine;
  LePlanning.GestionJoursFeriesActive := True;
  LePlanning.Evidence := true;
  LePlanning.ResizeFixedCols := True;

  { Le graphique }
  LePlanning.FormeGraphique := pgRectangle;
  LePlanning.RowSizeData := 20;
  LePlanning.ColorSelection := LePlanning.ColorBackground;

  { Les champs des états }
  LePlanning.EtatChampBackGroundColor := 'E_COULEURFOND';
  LePlanning.EtatChampCode := 'E_CODE';
  LePlanning.EtatChampFontColor := 'E_COULEURFONTE';
  LePlanning.EtatChampFontName := 'E_NOMFONTE';
  LePlanning.EtatChampFontSize := 'E_TAILLEFONTE';
  LePlanning.EtatChampFontStyle := 'E_STYLEFONTE';
  LePlanning.EtatChampIcone := 'E_ICONE';
  LePlanning.EtatChampLibelle := 'E_LIBELLE';

  { Les champs Items }
  LePlanning.ChampLineID := 'I_CODE';
  LePlanning.ChampLibelle := 'I_LIBELLE';
  LePlanning.ChampdateDebut := 'I_DATEDEBUT';
  LePlanning.ChampDateFin := 'I_DATEFIN';
  LePlanning.ChampEtat := 'I_ETAT';
  LePlanning.ChampColor := 'I_COLOR';
  LePlanning.ChampHint := 'I_HINT';
  LePlanning.ChampIcone := 'I_ICONE';

  { Les champs des ressources }
  LePlanning.ResChampColor := '';
  LePlanning.ResChampFixedColor := '';
  LePlanning.ResChampID := 'R_CODE';
  LePlanning.ResChampReadOnly := '';

  LePlanning.EdgeEnteteColFixed := True;

  { Les autorisations }
  LePlanning.Autorisation := [];

  { Le Zoom }
  LePlanning.Zoom := True;

  { Pas de sélection }
  LePlanning.NoSelectCell := True;
end;

procedure TFichePlanning.GoPlanning;

  procedure AddEtat(Code, Libelle, Cfond, Cfonte, NFonte: string; TFOnte: Integer; SFonte: string; Icone: Integer);
  begin
    if LesEtats.FindFirst(['E_CODE'], [Code], False) = nil then
      with TOB.Create('Un état', LesEtats, -1) do
      begin
        AddChampSupValeur('E_CODE', Code);
        AddChampSupValeur('E_LIBELLE', Libelle);
        AddChampSupValeur('E_COULEURFOND', CFond);
        AddChampSupValeur('E_COULEURFONTE', CFonte);
        AddChampSupValeur('E_NOMFONTE', NFonte);
        AddChampSupValeur('E_TAILLEFONTE', TFonte);
        AddChampSupValeur('E_STYLEFONTE', SFonte);
        AddChampSupValeur('E_ICONE', ICone);
      end;
  end;

var
  I: Integer;
  cStart: Integer; { Couleur de début par défaut }
  cEnd: Integer; { Couleur de fin par défaut }
  cSaut: Integer; { Incrément ou décrément de couleur }
begin
  LePlanning.Activate := False;

  { Création des états }
  cStart := RGB(220, 220, 220);
  cEnd := Rgb(19, 29, 138);
  cSaut := (cEnd - cStart) div MAX_LEVEL_ETAT;
  for i := 0 to MAX_LEVEL_ETAT - 1 do
  begin
    AddEtat(Format('%3.3d', [i + 1]), '', ColorToString(cStart), 'clBlack', 'Arial', 8, '', -1);
    cStart := cStart + cSaut;
  end;

  { Etats spécifique au cumul }
  cStart := RGB(220, 220, 220);
  cEnd := Rgb(19, 29, 138);
  cSaut := (cEnd - cStart) div MAX_LEVEL_ETAT;
  for i := 0 to MAX_LEVEL_ETAT - 1 do
  begin
    AddEtat(Format('C%2.2d', [i + 1]), '', ColorToString(cStart), 'clBlack', 'Arial', 8, 'I', -1);
    cStart := cStart + cSaut;
  end;

  { Les dates }
  LePlanning.IntervalDebut := DateDeDebut;
  LePlanning.IntervalFin := DateDeFin;

  { Position par défaut }
  LePlanning.DateOfStart := DebutDeMois(TmpDate);

  { Les items }
  LePlanning.TobItems := LesItems;

  { Les états }
  LePlanning.TobEtats := LesEtats;

  { Les ressources }
  LePlanning.TobRes := LesRessources;

  { Chargement de la configuration du planning }
  LoadConfigPlanning;

  { Suppression des items dont la valeur I_DATA=0 }
  for I := LesItems.Detail.Count - 1 downto 0 do
    if LesItems.Detail[I].GetValue('I_DATA') = 0.00 then LesItems.Detail[I].Free;

  if TypePlanning in [tpPrestation, tpAffaire, tpAffairePrestation] then CalculTotalColonne;

  { Activation du planning }
  LePlanning.Activate := True;

  UpdateDateDuJour;
end;

function TFichePlanning.InitTobRessource: TOB;
begin
  Result := TOB.Create('Les Ressources', LesRessources, -1);
  with Result do
  begin
    AddChampSupValeur('R_CODE', '');
    AddChampSupValeur('R_LIBAFFAIRE', '');
    AddChampSupValeur('R_LIBPRESTATION', '');
    AddChampSupValeur('R_LIBRESSOURCE', '');
    AddChampSupValeur('R_CUMUL', 0.00);
    AddChampSupValeur('R_LIBCUMUL', '');
  end;
end;

procedure TFichePlanning.AjoutAffaire(Code, Code1: string; Valeur: Double = 0.00);
var
  T, TT: TOB;
begin
  TT := LesRessources.FindFirst(['R_CODE'], [Code1], FALSE);
  if TT = nil then
  begin
    with InitTobRessource do
    begin
      PutValue('R_CODE', Code1);
      T := LesAffaires.FindFirst(['AFF_AFFAIRE'], [Code], False);
      if T <> nil then PutValue('R_LIBAFFAIRE', T.GetValue('AFF_LIBELLE'));
      PutValue('R_CUMUL', Valeur);
    end;
  end
  else TT.PutValue('R_CUMUL', Double(TT.GetValue('R_CUMUL')) + Valeur);
end;

procedure TFichePlanning.AjoutPrestation(Code, Code1: string; Valeur: Double = 0.00);
var
  T, TT: TOB;
begin
  TT := LesRessources.FindFirst(['R_CODE'], [Code1], FALSE);
  if TT = nil then
  begin
    with InitTobRessource do
    begin
      PutValue('R_CODE', Code1);
      T := LesPrestations.FindFirst(['GA_CODEARTICLE'], [Code], False);
      if T <> nil then PutValue('R_LIBPRESTATION', Copy(T.GetValue('GA_CODEARTICLE'), 5, 200) + ' : ' + T.GetValue('GA_LIBELLE'));
      PutValue('R_CUMUL', Valeur);
    end;
  end
  else TT.PutValue('R_CUMUL', Double(TT.GetValue('R_CUMUL')) + Valeur);
end;

procedure TFichePlanning.AjoutRessource(Code, Code1: string; Valeur: Double = 0.00);
var
  T, TT: TOB;
  I: INTEGER;
  s: string;
begin
  TT := LesRessources.FindFirst(['R_CODE'], [Code1], FALSE);
  if TT = nil then
  begin
    with InitTobRessource do
    begin
      PutValue('R_CODE', Code1);
      if LesSalaries <> nil then
      begin
        s := format('%10.10d', [valeuri(code)]);
        T := LesSalaries.FindFirst(['SALARIE'], [s], False);
        if T <> nil then
        begin
          I := ValeurI(T.GetValue('SALARIE'));
          PutValue('R_LIBRESSOURCE', Format('%4.4d - %s', [I, T.GetValue('NOMPRENOM')]));
        end
        else
        begin
          PutValue('R_LIBRESSOURCE', s);
          ShowMessage('Le salarié : ' + Code + ' n''est pas référencé dans la base.');
        end;
      end;
      PutValue('R_CUMUL', Valeur);
    end;
  end
  else TT.PutValue('R_CUMUL', Double(TT.GetValue('R_CUMUL')) + Valeur);
end;


{***********A.G.L.***********************************************
Auteur  ...... : DMD
Créé le ...... : 13/02/2002
Modifié le ... :   /  /
Description .. : Cette procédure prend en charge le chargement des
Suite ........ : donénes depusi la base en tenant compte de la ressource,
Suite ........ : du type de planning, de la période.
Suite ........ : Elle nettoie au préalable les TOBs de travail.
Mots clefs ... :
*****************************************************************}
function TFichePlanning.LoadData: Boolean;
var
  iMaxData: Double;
  I, J: Integer; { variable de boucle }
  PrevRessource, PrevAffaire, PrevArticle: string;
  ValueQ: Double;
  Tressource, Tprestation, Taffaire: TOB;
  DateActivite: TDAteTime;
  PrevDate: TDateTime;
  PrevValueDateAffaire, PrevValueDate, PrevValueAffaire, PrevValueRessource, PrevValuePrestation: Double;
  iMax, Indice: Integer;
  Fini: Boolean;
  N1, N2, N3, N4: Double;
  TValues, TValues2: array of double;
  TobTemp: TOB;
  T: TOB;

  procedure LoadActivity;
  var
    Q: TQuery;
    S: string;
    St: string;
  begin
    Q := nil;
    try
      case TypePlanning of
        tpPrestation: St := 'ACT_CODEARTICLE';
        tpAffaire: St := 'ACT_AFFAIRE';
        tpRessource: St := 'ACT_RESSOURCE';
        tpAffairePrestation: St := 'ACT_AFFAIRE,ACT_CODEARTICLE';
        tpAffaireRessource: St := 'ACT_AFFAIRE,ACT_RESSOURCE';
        tpPrestationRessource: St := 'ACT_CODEARTICLE,ACT_RESSOURCE';
        tpAffairePrestationRessource: St := 'ACT_AFFAIRE,ACT_CODEARTICLE,ACT_RESSOURCE';
      end;

      { Pour tout le monde }
      S := 'SELECT ' + St + ',ACT_DATEACTIVITE,SUM(ACT_QTE) AS Q,MAX(ACT_LIBELLE) AS L FROM ACTIVITE WHERE';

      if LaAffaire <> '' then
      begin
        if AvecLesConges then
        begin
          S := S + ' ACT_AFFAIRE IN ("AZDE1005987    00","AZDE1015265    00",' + LaAffaire + ') AND';
        end
        else
        begin
          S := S + ' ACT_AFFAIRE IN (' + LaAffaire + ') AND';
        end;
      end
      else
      begin
        if not AvecLesConges then S := S + ' ACT_AFFAIRE<>"AZDE1005987    00" AND ACT_AFFAIRE<>"AZDE1015265    00" and';
      end;

      if LaPrestation <> '' then S := S + ' ACT_CODEARTICLE IN (' + LaPrestation + ') AND';
      if LaRessource <> '' then S := S + ' ACT_RESSOURCE IN (' + LaRessource + ') AND';
      if LeClient <> '' then S := S + ' ACT_TIERS="' + LeClient + '" AND';
      // XP 23.03.2006 S := S + ' ACT_TYPEACTIVITE="REA" AND (ACT_TYPERESSOURCE="SAL" OR ACT_TYPERESSOURCE="ST") AND ACT_DATEACTIVITE>="' + UsDateTime(DateDebutLoad) +
      S := S + ' ACT_TYPEACTIVITE="REA" AND ACT_DATEACTIVITE>="' + UsDateTime(DateDebutLoad) +
        '" AND ACT_DATEACTIVITE<="' + UsDateTime(DateFinLoad) + '"';
      S := S + ' GROUP BY ' + St + ',ACT_DATEACTIVITE';

      Q := OpenSql(S, True);
      if not Q.Eof then LesActivites.LoadDetailDB('ACTIVITE', '', '', Q, True, True);
    finally
      if Q <> nil then Ferme(Q);
    end;
  end;

  function AddItem(TypeItem: string; Code, Libelle: string; D1, D2: TDateTime; Etat, C, H: string; Icone: Integer; Donnee: Double; CodeAffaire, CodeArticle,
    CodeRessource: string): TOB;
  begin
    Result := TOB.Create('un item', LesItems, -1);
    with Result do
    begin
      AddChampSupValeur('I_CODE', Code);
      AddChampSupValeur('I_LIBELLE', Libelle);
      AddChampSupValeur('I_DATEDEBUT', D1);
      AddChampSupValeur('I_DATEFIN', D2);
      AddChampSupValeur('I_ETAT', Etat);
      AddChampSupValeur('I_RESSOURCE', CodeRessource);
      AddChampSupValeur('I_TYPE', TypeItem);

      if Donnee > iMaxData then iMaxData := Donnee;

      if TypeItem <> 'C' then
        if Donnee <= 0.20 then PutValue('I_ETAT', '001')
        else if Donnee <= 0.40 then PutValue('I_ETAT', '002')
        else if Donnee <= 0.60 then PutValue('I_ETAT', '003')
        else if Donnee <= 0.80 then PutValue('I_ETAT', '004')
        else PutValue('I_ETAT', '005');

      AddChampSupValeur('I_LIBELLE', FormatFloat(FFormat, Donnee * Multiplicateur));

      AddChampSupValeur('I_COLOR', C);
      AddChampSupValeur('I_HINT', H);
      AddChampSupValeur('I_ICONE', Icone);
      AddChampSupValeur('I_DATA', Donnee);
    end;
  end;

  procedure PutColonneFixe(Champs, Sizes, Aligns: string);
  begin
    LePlanning.TokenFieldColFixed := Champs;
    LePlanning.TokenSizeColFixed := Sizes;
    LePlanning.TokenAlignColFixed := Aligns;
  end;

  procedure AddCumul;
  begin
    LePlanning.TokenFieldColFixed := LePlanning.TokenFieldColFixed + ';R_LIBCUMUL';
    LePlanning.TokenSizeColFixed := LePlanning.TokenSizeColFixed + ';60';
    LePlanning.TokenAlignColFixed := LePlanning.TokenAlignColFixed + ';C';
  end;

var
  MaxValues: Integer;
  F1: TTypeFUnc;
  F2: TTypeFUnc;
  CHAMPS1: string;
  CHAMPS2: string;
  PrevCode1: string;
  PrevCode2: string;
  PrevValue1: Double;
  PrevValue2: Double;
begin
  { L'optimisme est de rigeur }
  Result := True;

  { Chargement des activités }
  LoadActivity;

  iMaxData := 0.00;

  MaxValues := Trunc((DateDeFin - DateDeDebut)) + 1;
  SetLength(TValues, MaxValues + 1);
  SetLength(TValues2, MaxValues + 1);

  TobTemp := TOB.Create('', nil, -1);

  { Init }
  InitMove(100, '');
  PrevValuePrestation := 0.00;
  PrevValueDate := 0.00;
  PrevValueAffaire := 0.00;
  PrevValueRessource := 0.00;
  PrevValueDateAffaire := 0.00;

  Indice := 0;
  iMax := LesActivites.Detail.Count div 100;

  F1 := nil;
  F2 := nil;

  if LesActivites.Detail.Count = 0 then Result := False
  else
  begin
    if (TypePlanning = tpPrestation) or (TypePlanning = tpAffaire) or (TypePlanning = tpRessource) then
    begin
      case TypePlanning of
        tpPrestation:
          begin
            PutColonneFixe('R_LIBPRESTATION', '150', 'L');
            CHAMPS1 := 'ACT_CODEARTICLE';
            F1 := AjoutPrestation;
          end;
        tpAffaire:
          begin
            PutColonneFixe('R_LIBAFFAIRE', '150', 'L');
            CHAMPS1 := 'ACT_AFFAIRE';
            F1 := AjoutAffaire;
          end;
        tpRessource:
          begin
            PutColonneFixe('R_LIBRESSOURCE', '150', 'L');
            CHAMPS1 := 'ACT_RESSOURCE';
            F1 := AjoutRessource;
          end;
      end;
      LesActivites.Detail.Sort(CHAMPS1 + ';ACT_DATEACTIVITE');
      PrevValue1 := 0.00;
      while LesActivites.Detail.Count > 0 do
      begin
        PrevCode1 := LesActivites.Detail[0].GetValue(CHAMPS1);
        while (LesActivites.Detail.Count > 0) and (PrevCode1 = LesActivites.Detail[0].GetValue(CHAMPS1)) do
        begin
          PrevDate := LesActivites.Detail[0].GetValue('ACT_DATEACTIVITE');
          while (LesActivites.Detail.Count > 0) and (PrevDate = LesActivites.Detail[0].GetValue('ACT_DATEACTIVITE')) and (PrevCode1 =
            LesActivites.Detail[0].GetValue(CHAMPS1)) do
          begin
            ValueQ := LesActivites.Detail[0].GetValue('Q');
            PrevValueDate := PrevValueDate + ValueQ;
            PrevValue1 := PrevValue1 + ValueQ;
            if Indice = iMax then
            begin
              indice := 0;
              MoveCur(1);
            end else Inc(indice);
            LesActivites.Detail[0].ChangeParent(TobTemp, -1);
          end;
          T := AddItem('', PrevCode1, '', PrevDate, PrevDate, '', '', '', -1, PrevValueDate, '', '', '');
          while TobTemp.Detail.Count > 0 do TobTemp.Detail[0].ChangeParent(T, -1);
          PrevValueDate := 0.00;
        end;
        if Assigned(F1) then F1(PrevCode1, PrevCode1, PrevValue1);
        PrevValue1 := 0.00;
      end;
    end
    else if (TypePlanning = TpAffairePrestation) or (TypePlanning = tpAffaireRessource) or (Typeplanning = tpPrestationRessource) then
    begin
      case TypePlanning of
        TpAffairePrestation:
          begin
            PutColonneFixe('R_LIBAFFAIRE;R_LIBPRESTATION', '150;150', 'L;L');
            CHAMPS1 := 'ACT_AFFAIRE';
            CHAMPS2 := 'ACT_CODEARTICLE';
            F1 := AjoutAffaire;
            F2 := AjoutPrestation;
          end;
        tpAffaireRessource:
          begin
            PutColonneFixe('R_LIBAFFAIRE;R_LIBRESSOURCE', '150;150', 'L;L');
            CHAMPS1 := 'ACT_AFFAIRE';
            CHAMPS2 := 'ACT_RESSOURCE';
            F1 := AjoutAffaire;
            F2 := AjoutRessource;
          end;
        tpPrestationRessource:
          begin
            PutColonneFixe('R_LIBPRESTATION;R_LIBRESSOURCE', '150;150', 'L;L');
            CHAMPS1 := 'ACT_CODEARTICLE';
            CHAMPS2 := 'ACT_RESSOURCE';
            F1 := AjoutPrestation;
            F2 := AjoutRessource;
          end;
      end;
      { TRI }
      LesActivites.Detail.Sort(CHAMPS1 + ';' + CHAMPS2 + ';ACT_DATEACTIVITE');

      Fini := False;
      PrevValue1 := 0.00;
      while LesActivites.Detail.Count > 0 do
      begin
        PrevCode1 := LesActivites.Detail[0].GetValue(CHAMPS1);
        PrevValue2 := 0.00;
        while (LesActivites.Detail.count > 0) and (PrevCode1 = LesActivites.Detail[0].GetValue(CHAMPS1)) do
        begin
          PrevCode2 := LesActivites.Detail[0].GetValue(CHAMPS2);
          while (LesActivites.Detail.count > 0) and (PrevCode1 = LesActivites.Detail[0].GetValue(CHAMPS1)) and (PrevCode2 =
            LesActivites.Detail[0].GetValue(CHAMPS2)) do
          begin
            PrevDate := LesActivites.Detail[0].GetValue('ACT_DATEACTIVITE');
            while (LesActivites.Detail.count > 0) and (PrevDate = LesActivites.Detail[0].GetValue('ACT_DATEACTIVITE')) and (PrevCode1 =
              LesActivites.Detail[0].GetValue(CHAMPS1)) and (PrevCode2 = LesActivites.Detail[0].GetValue(CHAMPS2)) do
            begin
              ValueQ := LesActivites.Detail[0].GetValue('Q');
              TValues[Trunc((PrevDate - DateDeDebut)) + 1] := TValues[Trunc((PrevDate - DateDeDebut)) + 1] + ValueQ;
              PrevValueDate := PrevValueDate + ValueQ;
              PrevValue2 := PrevValue2 + ValueQ;
              PrevValue1 := PrevValue1 + ValueQ;
              LesActivites.Detail[0].ChangeParent(TobTemp, -1);
            end;
            T := AddItem('', PrevCode1 + '¤' + PrevCode2, '', PrevDate, PrevDate, '', '', '', -1, PrevValueDate, '', '', '');
            while TobTemp.Detail.Count > 0 do TobTemp.Detail[0].ChangeParent(T, -1);
            PrevValueDate := 0.00;
          end;
          if Assigned(F2) then F2(PrevCode2, PrevCode1 + '¤' + PrevCode2, PrevValue2);
          PrevValue2 := 0.00;
        end;
        if Assigned(F1) then F1(PrevCode1, PrevCode1, PrevValue1);
        PrevValue1 := 0.00;
        for J := Trunc(DateDeDebut) to Trunc(DateDeFin) do
          if TValues[Trunc((j - DateDeDebut)) + 1] <> 0 then
          begin
            AddItem('C', PrevCode1, '', J, J, '', '', '', -1, TValues[Trunc((j - DateDeDebut)) + 1], '', '', '');
            TValues[Trunc((j - DateDeDebut)) + 1] := 0.00;
          end;
      end;
    end
    else if Typeplanning = tpAffairePrestationRessource then
    begin
      PutColonneFixe('R_LIBAFFAIRE;R_LIBPRESTATION;R_LIBRESSOURCE', '150;150;150', 'L;L;L');

      { TRI }
      LesActivites.Detail.Sort('ACT_AFFAIRE;ACT_CODEARTICLE;ACT_RESSOURCE;ACT_DATEACTIVITE');

      while LesActivites.Detail.Count > 0 do
      begin
        PrevAffaire := LesActivites.Detail[0].GetValue('ACT_AFFAIRE');
        while (LesActivites.Detail.Count > 0) and (PrevAffaire = LesActivites.Detail[0].GetValue('ACT_AFFAIRE')) do
        begin
          PrevArticle := LesActivites.Detail[0].GetValue('ACT_CODEARTICLE');
          while (LesActivites.Detail.Count > 0) and (PrevAffaire = LesActivites.Detail[0].GetValue('ACT_AFFAIRE')) and (PrevArticle =
            LesActivites.Detail[0].GetValue('ACT_CODEARTICLE')) do
          begin
            PrevRessource := LesActivites.Detail[0].GetValue('ACT_RESSOURCE');
            while (LesActivites.Detail.Count > 0) and (PrevRessource = LesActivites.Detail[0].GetValue('ACT_RESSOURCE')) and (PrevAffaire =
              LesActivites.Detail[0].GetValue('ACT_AFFAIRE')) and (PrevArticle = LesActivites.Detail[0].GetValue('ACT_CODEARTICLE')) do
            begin
              PrevDate := LesActivites.Detail[0].GetValue('ACT_DATEACTIVITE');
              while (LesActivites.Detail.Count > 0) and (PrevDate = LesActivites.Detail[0].GetValue('ACT_DATEACTIVITE')) and (PrevRessource =
                LesActivites.Detail[0].GetValue('ACT_RESSOURCE')) and (PrevAffaire = LesActivites.Detail[0].GetValue('ACT_AFFAIRE')) and (PrevArticle =
                LesActivites.Detail[0].GetValue('ACT_CODEARTICLE')) do
              begin
                ValueQ := LesActivites.Detail[0].GetValue('Q');
                TValues[Trunc((PrevDate - DateDeDebut)) + 1] := TValues[Trunc((PrevDate - DateDeDebut)) + 1] + ValueQ;
                TValues2[Trunc((PrevDate - DateDeDebut)) + 1] := TValues2[Trunc((PrevDate - DateDeDebut)) + 1] + ValueQ;
                PrevValueDate := PrevValueDate + ValueQ;
                PrevValuePrestation := PrevValuePrestation + ValueQ;
                PrevValueAffaire := PrevValueAffaire + ValueQ;
                PrevValueRessource := PrevValueRessource + ValueQ;
                if Indice = iMax then
                begin
                  indice := 0;
                  MoveCur(1);
                end else Inc(indice);
                LesActivites.Detail[0].ChangeParent(TobTemp, -1);
              end;
              T := AddItem('', PrevAffaire + '¤' + PrevArticle + '¤' + PrevRessource, '', PrevDate, PrevDate, '', '', '', -1, PrevValueDate, '', '', '');
              while TobTemp.Detail.Count > 0 do TobTemp.Detail[0].ChangeParent(T, -1);
              PrevValueDate := 0.00;
            end;
            AjoutRessource(PrevRessource, PrevAffaire + '¤' + PrevArticle + '¤' + PrevRessource, PrevValueRessource);
            PrevValueRessource := 0.00;
          end;
          AjoutPrestation(PrevArticle, PrevAffaire + '¤' + PrevArticle, PrevValuePrestation);
          PrevValuePrestation := 0.00;
          for J := Trunc(DateDeDebut) to Trunc(DateDeFin) do
            if TValues2[Trunc((j - DateDeDebut)) + 1] <> 0 then
            begin
              AddItem('C', PrevAffaire + '¤' + PrevArticle, '', J, J, '', '', '', -1, TValues2[Trunc((j - DateDeDebut)) + 1], '', '', '');
              TValues2[Trunc((j - DateDeDebut)) + 1] := 0.00;
            end;
        end;
        AjoutAffaire(PrevAffaire, PrevAffaire, PrevValueAffaire);
        PrevValueAffaire := 0.00;
        for J := Trunc(DateDeDebut) to Trunc(DateDeFin) do
          if TValues[Trunc((j - DateDeDebut)) + 1] <> 0 then
          begin
            AddItem('C', PrevAffaire, '', J, J, '', '', '', -1, TValues[Trunc((j - DateDeDebut)) + 1], '', '', '');
            TValues[Trunc((j - DateDeDebut)) + 1] := 0.00;
          end;
      end;
    end;
  end;

  LesRessources.Detail.Sort('R_CODE');

  FiniMove;

  if CLinePLanning then
  begin
    AddCumul;
    for i := 0 to LesRessources.Detail.Count - 1 do
      with LesRessources.Detail[i] do
        PutValue('R_LIBCUMUL', FormatFloat(FFormat, GetValue('R_CUMUL') * Multiplicateur));
  end;

  { Calcul des bornes des niveaux }
  N1 := iMaxData / 5;
  N2 := 2 * N1;
  N3 := 3 * N1;
  N4 := 4 * N1;

  { Boucle pour mettre à joutr les états des lignes de cumul ou I_TYPE='C' }
  for I := 0 to LesItems.Detail.Count - 1 do
    if LesItems.Detail[I].GetValue('I_TYPE') = 'C' then
      with LesItems.Detail[I] do
        if GetValue('I_DATA') <= N1 then PutValue('I_ETAT', 'C01')
        else if GetValue('I_DATA') <= N2 then PutValue('I_ETAT', 'C02')
        else if GetValue('I_DATA') <= N3 then PutValue('I_ETAT', 'C03')
        else if GetValue('I_DATA') <= N4 then PutValue('I_ETAT', 'C04')
        else PutValue('I_ETAT', 'C05');

  SetLength(TValues, 0);
  SetLength(TValues2, 0);
  TobTemp.Free;
end;

procedure TFichePlanning.BpersoClick(Sender: TObject);
begin
  eCongePersoPlanning('ConfigPlanningTemps', LePlanning, eColorStart, eColorEnd, True);
end;

procedure TFichePlanning.BannulerClick(Sender: TObject);
begin
  Close;
end;

procedure TFichePlanning.BPrevMonthClick(Sender: TObject);
var
  D, DD: TDateTime;
begin
  { CALCUL DU MOIS PRECEDENT EN FONCTION DE LaDateDuJour }
  D := LaDateDuJour;
  DD := DebutDeMois(PlusMois(D, -1));
  if DD < LePlanning.IntervalDebut then
  begin
    DateDeDebut := DebutDeMois(PlusMois(LePlanning.IntervalDebut, -1));
    LaDateDuJour := DateDeDebut;
    UpdateDateDuJour;
    TmpDate := DateDeDebut;
    DateDebutLoad := DateDeDebut;
    DateFinLoad := LePLanning.IntervalDebut - 1;
    LoadData;
    GoPlanning;
  end
  else
  begin
    LePlanning.Activate := False;
    LePlanning.DateOfStart := DebutDeMois(DD);
    LaDateDuJour := DebutDeMois(DD);
    UpdateDateDuJour;
    LePlanning.Activate := True;
  end;
end;

procedure TFichePlanning.BNextMonthClick(Sender: TObject);
var
  D, DD: TDateTime;
begin
  { CALCUL DU MOIS SUIVANT EN FONCTION DE LaDateDuJour }
  D := LaDateDuJour;
  DD := DebutDeMois(PlusMois(D, 1));
  if DD > LePlanning.IntervalFin then
  begin
    DateDeFin := FinDeMois(PlusMois(LePlanning.IntervalFin, 1));
    TmpDate := DebutDeMois(DateDeFin);
    LaDateDuJour := LePLanning.IntervalFin + 1;
    UpdateDateDuJour;
    DateDebutLoad := LePLanning.IntervalFin + 1;
    DateFinLoad := DateDeFin;
    LoadData;
    GoPlanning;
  end
  else
  begin
    LePlanning.Activate := False;
    LePlanning.DateOfStart := DebutDeMois(DD);
    LaDateDuJour := DebutDeMois(DD);
    UpdateDateDuJour;
    LePlanning.Activate := True;
  end;
end;

procedure TFichePlanning.BprintClick(Sender: TObject);
var
  S: string;
{$ifdef aucasou}
  Sr: string;
{$endif}
begin
  LePlanning.TypeEtat := 'E';
  LePlanning.NatureEtat := 'XDT';
  LePlanning.CodeEtat := 'PLA';

  // extended

{$ifdef aucasou}
  { Les ressources }
  SR := StringReplace(LaRessource, '"', '', [rfReplaceAll]);
  if SR = '' then SR := 'Toutes les ressources rattachées';
  S := 'RESSOURCE=' + SR;

  { Les affaires }
  SR := StringReplace(LaAffaire, '"', '', [rfReplaceAll]);
  if SR = '' then SR := 'Toutes les affaires';
  S := S + '`AFFAIRE=' + SR;

  { Les prestations }
  SR := StringReplace(LaPrestation, '"', '', [rfReplaceAll]);
  if SR = '' then SR := 'Toutes les prestations';
  S := S + '`PRESTATION=' + SR;

  { Le reste }
  S := S + '`DATEDEB=' + FormatDateTime('dd/mm/yyyy', DateDeDebut) + '^E' + IntToStr(Byte(otDate));
  S := S + '`DATEFIN=' + FormatDateTime('dd/mm/yyyy', DateDeFin) + '^E' + IntToStr(Byte(otDate));

  // ajout du client
  s := s + '`CLIENT='+LeClient ;

  LePlanning.CriteresToPrint := S;
{$endif}

  { Le reste }
  S := '`DATEDEB=' + FormatDateTime('dd/mm/yyyy', DateDeDebut) + '^E' + IntToStr(Byte(otDate));
  S := S + '`DATEFIN=' + FormatDateTime('dd/mm/yyyy', DateDeFin) + '^E' + IntToStr(Byte(otDate));

  LePlanning.CriteresToPrint := LesCriteres + S ;

  LePlanning.Print;
end;

procedure TFichePlanning.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    Key := 0;
    Close;
  end;
end;

procedure TFichePlanning.AvertirApplication(Sender: TObject; FromItem, ToItem: TOB; Actions: THPlanningAction);
begin
  if Actions = paClickLeft then
  begin
    LaDateDuJour := LePlanning.GetDateOfCol(LePlanning.Col);
    UpdateDateDuJour;
  end;
end;

procedure TFichePlanning.UpdateDateDuJour;
begin
  LabelDate.Caption := 'Date sélectionnée : ' + FormatDateTime('dd/mm/yyyy', LaDateDuJour);
end;

procedure TFichePlanning.CalculTotalColonne;
var
  I: Integer;
  T: TOB;
  Index: Integer;
  Values: array of string;
  Fonte: TFont;
  S: string;
begin
  SetLength(iValues, 0);
  SetLength(iValues, Trunc(DateDeFin - DateDeDebut) + 1);
  SetLength(Values, Trunc(DateDeFin - DateDeDebut) + 1);
  for i := 0 to LesItems.Detail.Count - 1 do
  begin
    T := LesItems.Detail[i];
    S := T.GetValue('I_ETAT');
    if S[1] <> 'C' then
    begin
      Index := Trunc(T.GetValue('I_DATEDEBUT') - DateDeDebut);
      iValues[Index] := iValues[Index] + T.GetValue('I_DATA');
      Values[Index] := FormatFloat(FFormat, iValues[Index] * Multiplicateur);
    end;
  end;

  Fonte := TFont.Create;
  Fonte.Name := 'Arial';
  Fonte.Size := 10;
  Fonte.Style := [];
  Fonte.Color := clBlack;

  LePlanning.SupprimeLigneFixe;
  LePlanning.AjoutLigneFixe(Fonte, clBtnFace, Values);

  Fonte.Free;
  SetLength(Values, 0);
end;

procedure TFichePlanning.BexcelClick(Sender: TObject);
var
  C: TSaveDialog;
begin
  C := TSaveDialog.Create(Self);
  try
    C.Options := [ofOverwritePrompt];
    C.Filter := 'Microsoft Excel (*.xls)|*.XLS';
    if C.Execute then LePlanning.ExportToExcel(True, C.FileName);
  finally
    C.Free;
  end;
end;

procedure TFichePlanning.PlanningDblCLick(Sender: TObject);
{$ifdef aucasou}
var
  T: TOB;
{$endif}
begin
{$ifdef aucasou}
  T := LePlanning.GetCurItem;
  if (T <> nil) and (T.Detail.Count > 0) then
    ET_TempsDetail(Responsable, T.GetValue('I_DATEDEBUT'), TypePlanning, T, LesAffaires, LesPrestations, LesSalaries);
{$endif}
end;


end.

