unit SaisiePersoPlanning;

interface

uses Windows,
  Messages,
  SysUtils,
  Classes,
  ExtCtrls,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  HTB97,
  Hent1,
  HPanel,
  ComCtrls,
  StdCtrls,
  Hctrls,
  HPlanning,
  HColor,
  Spin,
  HmsgBox,
  uTob, Mask;

procedure eCongePersoPlanning(PConfigName: string; P: THPlanning; var eColorStart, eColorEnd: TColor; WithLevel: Boolean = False);

type
  TSaisieConfigPlanning = class(TForm)
    HPanel1: THPanel;
    GroupBox3: TGroupBox;
    HLabel11: THLabel;
    HLabel13: THLabel;
    LabelRegroupement: TLabel;
    Hligne: TSpinEdit;
    Lcolonne: TSpinEdit;
    CcumulDate: TCheckBox;
    Cregroupement: TComboBox;
    GroupBox4: TGroupBox;
    HLabel12: THLabel;
    HLabel8: THLabel;
    Cfond: TPaletteButton97;
    Cferies: TPaletteButton97;
    HLabel10: THLabel;
    CFormeGraphique: TComboBox;
    Baide: TToolbarButton97;
    Bannuler: TToolbarButton97;
    Bvalider: TToolbarButton97;
    MaskFormat: THCritMaskEdit;
    Gniveaux: TGroupBox;
    HLabel1: THLabel;
    C1: TPaletteButton97;
    C2: TPaletteButton97;
    C3: TPaletteButton97;
    C4: TPaletteButton97;
    C5: TPaletteButton97;
    HLabel2: THLabel;
    CsamDim: TPaletteButton97;
    HLabel3: THLabel;
    CSelection: TPaletteButton97;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormDestroy(Sender: TObject);
    procedure Changecouleur(Sender: TObject);
    procedure BvaliderClick(Sender: TObject);
    procedure BannulerClick(Sender: TObject);
    procedure CcumulDateClick(Sender: TObject);
    procedure CFormeGraphiqueChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Déclarations privées }
    procedure Sauve;
  public
    { Déclarations publiques }
    LePlanning: THPlanning;

    { Couleur du fond avant modification }
    PrevColorFond: TColor;

    AvecLesNiveaux: BOolean;

    ConfigName: string;
  end;

implementation

{$R *.DFM}

uses SaisiePL;

procedure eCongePersoPlanning(PConfigName: string; P: THPlanning; var eColorStart, eColorEnd: TColor; WithLevel: Boolean = False);
begin
  with TSaisieConfigPlanning.Create(Application) do
  begin
    ConfigName := PConfigName;
    AvecLesNiveaux := WithLevel;
    LePlanning := P;
    ShowModal;
    Free;
  end;
end;

procedure TSaisieConfigPlanning.FormCreate(Sender: TObject);
begin
  //
end;

procedure TSaisieConfigPlanning.FormShow(Sender: TObject);
var
  i: integer;
  Index: Integer;
begin
  { Les couleurs }
  Index := 0;
  if AvecLesNiveaux then
    for I := 0 to LePlanning.TobEtats.Detail.Count - 1 do
      if not (Copy(LepLanning.TobEtats.Detail[i].GetValue('E_CODE'), 1, 1) = 'C') then
      begin
        Inc(Index);
        TPaletteButton97(FindComponent('C' + IntToStr(Index))).CurrentChoix := StringToColor(LePlanning.TobEtats.Detail[i].GetValue('E_COULEURFOND'));
      end;

  { La forme graphique }
  CFormeGraphique.ItemIndex := Integer(LePlanning.FormeGraphique);

  { Couleur du fond }
  PrevColorFond := LePlanning.ColorBackground;

  { Couleur du fond, de la sélection et des jours fériés }
  CFond.CurrentChoix := LePLanning.ColorBackground;
  Cferies.CurrentChoix := LePlanning.ColorJoursFeries;
  CSamDim.CurrentChoix := LePlanning.ColorOfSaturday;
  CSelection.CurrentChoix := LePlanning.ColorSelection ;

  { Les lignes/Colonnes }
  Hligne.Value := LePlanning.RowSizeData;
  LColonne.Value := LePlanning.ColSizeData;

  { Le mask de la date }
  MaskFormat.Text := LePlanning.DateFormat;

  { Ligne de cumul des dates }
  CcumuLDate.Checked := LePlanning.ActiveLigneGroupeDate;
  if CcumulDate.Checked then
    if LePlanning.CumulInterval = pciSemaine then Cregroupement.ItemIndex := 0
    else if LePlanning.CumulInterval = pciMois then Cregroupement.ItemIndex := 1;

  { Pour ne pas afficher les niveaux depuis les plannings congés }
  GNiveaux.Visible := AvecLesNiveaux;

  // xp le 21-07-2003
  BValider.Enabled := False;
end;

procedure TSaisieConfigPlanning.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Action := caFree;
end;

procedure TSaisieConfigPlanning.FormDestroy(Sender: TObject);
begin
  //
end;


{ ************************************************* }
procedure TSaisieConfigPlanning.Changecouleur(Sender: TObject);
begin
  Bvalider.Enabled := True;

  if Sender = Cregroupement then
    if Cregroupement.ItemIndex = 1 then
    begin
      MaskFormat.Text := 'dd';
      lColonne.Value := 22;
    end
    else if Cregroupement.ItemIndex = 0 then
    begin
      MaskFormat.Text := 'dd mmm';
      lColonne.Value := 40;
    end;
end;

procedure TSaisieConfigPlanning.BvaliderClick(Sender: TObject);
var
  I: Integer;
  T: TOB;
  Index: Integer;
  P: TPaletteButton97;
begin
  { Couleur du fond, de la sélection et des jours fériés }
  LePLanning.ColorBackground := CFond.CurrentChoix;
  LePlanning.ColorSelection := CSelection.CurrentChoix ; // XP 30-06-2005 LePlanning.ColorBackground;
  LePlanning.ColorJoursFeries := Cferies.CurrentChoix;
  LePlanning.ColorOfSaturday := CsamDim.CurrentChoix;
  LePlanning.ColorOfSunday := CsamDim.CurrentChoix;

  { La forme graphique }
  LePlanning.FormeGraphique := THPlanningFormeGraphique(CFormeGraphique.ItemIndex);

  { Les lignes/Colonnes }
  LePlanning.RowSizeData := Hligne.Value;
  LePlanning.ColSizeData := LColonne.Value;

  { Ligne de cumul des dates }
  LePlanning.ActiveLigneGroupeDate := CcumuLDate.Checked;
  if CcumulDate.Checked then
    if Cregroupement.ItemIndex = 0 then LePlanning.CumulInterval := pciSemaine
    else if Cregroupement.ItemIndex = 1 then LePlanning.CumulInterval := pciMois;

  { Le mask de la date }
  if MaskFormat.Text <> '' then LePlanning.DateFormat := MaskFormat.Text;

  { Modification de la TOB des états }
  if AvecLesNiveaux then
  begin
    for I := 0 to LePlanning.TobEtats.Detail.Count - 1 do
    begin
      T := LePlanning.TobEtats.Detail[I];
      if not (Copy(T.GetValue('E_CODE'), 1, 1) = 'C') then
        T.PutValue('E_COULEURFOND', ColorToString(TPaletteButton97(FindComponent('C' + IntToStr(i + 1))).CurrentChoix));
    end;
    Index := 0;
    for I := 0 to LePlanning.TobEtats.Detail.Count - 1 do
    begin
      T := LePlanning.TobEtats.Detail[I];
      if Copy(T.GetValue('E_CODE'), 1, 1) = 'C' then
      begin
        Inc(Index);
        P := TPaletteButton97(FindComponent('C' + IntToStr(Index)));
        if P <> nil then
          T.PutValue('E_COULEURFOND', ColorToString(P.CurrentChoix));
      end;
    end;
    LePlanning.RefreshEtats;
  end;

  { Remise du bouton à false }
  BValider.Enabled := False;

  { A faire à la fin uniquement }
  LePlanning.Invalidate;

  Sauve;
end;

procedure TSaisieConfigPlanning.BannulerClick(Sender: TObject);
begin
  if Bvalider.Enabled then
    if HShowMessage('0;' + Caption + ';Voulez-vous enregistrer les modifications ?;Q;YN;Y;Y', '', '') = mrOk then
    begin
      BvaliderClick(nil);
    end;
  Close;
end;

procedure TSaisieConfigPlanning.CcumulDateClick(Sender: TObject);
begin
  LabelRegroupement.Visible := TCheckBox(Sender).Checked;
  CRegroupement.Visible := TCheckBox(Sender).Checked;
  BValider.Enabled := True;
end;

function AglVraiFaux(E, V, F: Variant): Variant;
begin
  if E then Result := V else Result := F;
end;

procedure TSaisieConfigPlanning.Sauve;
var
  S: string;
  L: TStrings;
  I: Integer;
  P: TPaletteButton97;
begin
  { Création de la liste des données }
  L := TStringList.Create;

  { Les couleurs Fond+Fériés}
  L.Add(ColorToString(Cfond.CurrentChoix) + ';' + ColorToString(Cferies.CurrentChoix) + ';' + ColorToString(CSamDim.CurrentChoix) + ';' + ColorToString(CSelection.CurrentChoix));

  { La forme graphique }
  L.Add(AglPlanningFormeGraphiqueToString(THPlanningFormeGraphique(Integer(cFormeGraphique.ItemIndex))));

  { Hauteur,Largeur, cumuldate et mode de cumul }
  L.Add(Format('%d;%d;%s;%s;%s', [LePlanning.RowSizeData, LePlanning.ColSizeData, string(AglVraiFaux(CCumulDate.Checked, '1', '0')),
    AglPlanningCumulIntervalToString(LePLanning.CumulInterval), LePlanning.DateFormat]));

  { Les couleurs des niveaux }
  S := '';
  for i := 0 to MAX_LEVEL_ETAT - 1 do
  begin
    P := TPaletteButton97(FindComponent('C' + IntToStr(i + 1)));
    if P <> nil then
      if S = '' then S := ColorToString(P.CurrentChoix) else S := S + ';' + ColorToString(P.CurrentChoix);
  end;
  L.Add(S);
  L.Add('');

  SaveSynRegKey(ConfigName, L.Text, True);

  L.Free;
end;

procedure TSaisieConfigPlanning.CFormeGraphiqueChange(Sender: TObject);
begin
  Bvalider.Enabled := True;
end;

procedure TSaisieConfigPlanning.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (key = vk_escape) then BannulerClick(Bannuler)
  else if (key = vk_f10) and (bvalider.enabled) then BvaliderClick(BValider) ;
end;

end.

