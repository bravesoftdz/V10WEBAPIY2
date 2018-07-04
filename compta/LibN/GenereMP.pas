unit GenereMP;

interface

uses
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hcompte, Mask, Hctrls, ExtCtrls, ComCtrls, Menus, HSysMenu,
  HTB97, ENT1, HENT1, hmsgbox, Saisutil, Encutil, {Filtre, LettUtil,
  DBCtrls,}SaisTaux1, GenerMP, UObjFiltres;

function ParamsMP(PourGeneration: Boolean; var GMP: tGenereMP): Boolean;
function ParamsMPAuto(var GMP: tGenereMP): Boolean;

type
  TFGenereMP = class(TForm)
    Pages: TPageControl;
    Dock: TDock97;
    PanelBouton: TToolWindow97;
    BOuvrir: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    FFiltres: THValComboBox;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    PGenere: TTabSheet;
    Bevel7: TBevel;
    H_MODEGENERE: TLabel;
    HDateReg: THLabel;
    MODEGENERE: THValComboBox;
    DATEGENERATION: THCritMaskEdit;
    DATEECHEANCE: THCritMaskEdit;
    CChoixDate: TCheckBox;
    HCpteGen: THLabel;
    CompteGeneration: THCritMaskEdit;
    GroupBox1: TGroupBox;
    TREFECR2: THLabel;
    REFECR2: THCritMaskEdit;
    HLabel3: THLabel;
    LIBECR2: THCritMaskEdit;
    GroupBox2: TGroupBox;
    LIBECR1: THCritMaskEdit;
    REFECR1: THCritMaskEdit;
    TREFECR1: THLabel;
    HLabel1: THLabel;
    HM: THMsgBox;
    HLabel4: THLabel;
    NUMENCADECA: THCritMaskEdit;
    Label1: TLabel;
    GROUPEENCADECA: THValComboBox;
    HJournalGenere: TLabel;
    JournalGenere: THValComboBox;
    LAnal: TLabel;
    ANAL: THValComboBox;
    BFiltre: TToolbarButton97;
    PEmission: TTabSheet;
    HLabel10: THLabel;
    TER_DOCUMENT: THLabel;
    TER_ENVOITRANS: THLabel;
    EXPORTCFONB: TCheckBox;
    FORMATCFONB: THValComboBox;
    EnvoiBORDEREAU: TCheckBox;
    ModeleBordereau: THValComboBox;
    ModeTeletrans: THValComboBox;
    Bevel1: TBevel;
    TTauxDev: TLabel;
    TAUXDEV: THNumEdit;
    HDevise: TLabel;
    TDevise: TLabel;
    AlerteEcheMP: TCheckBox;
    HLabel5: THLabel;
    REFEXT1: THCritMaskEdit;
    HLabel6: THLabel;
    REFLIB1: THCritMaskEdit;
    HLabel7: THLabel;
    HLabel8: THLabel;
    REFLIB2: THCritMaskEdit;
    REFEXT2: THCritMaskEdit;
    BAssist: TToolbarButton97;
    Dupliquer1: TMenuItem;
    LettrageAuto: TCheckBox;

    procedure FormShow(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure CChoixDateClick(Sender: TObject);
    procedure CompteGenerationExit(Sender: TObject);
    procedure JournalGenereExit(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TAUXDEVKeyPress(Sender: TObject; var Key: Char);
    procedure MODEGENEREChange(Sender: TObject);
    procedure BAssistClick(Sender: TObject);
    procedure REFECR1Change(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
    ObjFiltre: TObjFiltre;
    FNomFiltre: string;
    Cat: string;
    GMP: tGenereMP;
    PourGeneration: Boolean;
    LoadLeFiltre: Boolean;
    procedure Init;
    procedure Charge;
    procedure Sauve;
    function CompteOk: Boolean;
  public
    { Déclarations publiques }
    CodeJournal : string; {JP 22/02/05 : contient le journal correspondant à la banque prévisionnelle}

    function  GereBanquePrevi : Boolean;{JP 22/02/05 : Gestion de la banque prévisionnelle}
    procedure ApresChargement;{JP 22/02/05 : Pour annuler le chargement du journal et du général}
    function  TesteGeneral(Cpte : string) : Boolean; {JP 05/06/07 : FQ 18660 : on s'assure que le compte n'est pas fermé}
  end;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  utilPGI, MulSMPUtil, GuidTool, UTob;

{$R *.DFM}

function ParamsMP(PourGeneration: Boolean; var GMP: tGenereMP): Boolean;
var
  FGenereMP: TFGenereMP;

begin
  Result := FALSE;
  FGenereMP := TFGenereMP.Create(Application);
  try
    FGenereMP.FNomFiltre := 'G' + GMP.NomFSelect;
    FGenereMP.Cat := GMP.Cat;
    FGenereMP.PourGeneration := PourGeneration;
    FGenereMP.GMP := GMP;
    FGenereMP.GMP.TauxVolatil := False; {JP 13/06/06 : FQ 18319} 
    if FGenereMP.ShowModal = mrOk then
    begin
      Result := TRUE;
      GMP := FGenereMP.GMP;
    end;
  finally
    FGenereMP.Free;
  end;
  Screen.Cursor := SyncrDefault;
end;

function TrouveMPG: string;
var
  i: Integer;
  St, MPLu, Acc, CatLu: string;

begin
  Result := '';
  for i := 0 to VH^.MPACC.Count - 1 do
  begin
    St := VH^.MPACC[i];
    MPLu := ReadtokenSt(St);
    Acc := ReadtokenSt(St);
    CatLu := ReadTokenSt(St);
    if (CatLu = 'LCR') and ((Acc = 'TRA') or (Acc = 'NON') or (Acc = 'NAC') or (Acc = 'ACC')) then
    begin
      Result := MPLu;
      Break;
    end;
  end;
end;

function ParamsMPAuto(var GMP: tGenereMP): Boolean;
var
  i, j: Integer;
  Jal, Cpt: string;
  TobL: TOB;
begin
  Result := FALSE;
  i := TrouveJalCptEff(GMP.Smp);
  if i < 0 then Exit;
  TOBL := VH^.TOBJalEffet.Detail[i];
  Jal := TOBL.GetValue('J_JOURNAL');
  Cpt := TOBL.GetValue('J_CONTREPARTIE');
  GMP.CptG := Cpt;
  GMP.MPG := TrouveMPG;
  if GMP.MPG = '' then Exit;
  GMP.JalG := Jal;
  GMP.NatureG := TOBL.GetValue('G_NATUREGENE');
  GMP.NatJalG := TOBL.GetValue('J_NATUREJAL');
  GMP.GroupeEncadeca := 'ECH';
  GMP.ModeAnal := 'ATT';
  GMP.NumEncaDeca := '';
  GMP.DateC := V_PGI.DateEntree; {JP 08/06/07 : FQ 17846 : la date d'entrée et non la date système}
  GMP.DateE := GMP.DateC;
  GMP.ForceEche := TRUE;
  GMP.GenColl := TOBL.GetValue('G_COLLECTIF') = 'X';
  GMP.GenPoint := TOBL.GetValue('G_POINTABLE') = 'X';
  GMP.GenLett := TOBL.GetValue('G_LETTRABLE') = 'X';
  GMP.GenVent := TOBL.GetValue('G_VENTILABLE') = 'X';
  for j := 1 to MaxAxe do GMP.Ventils[j] := TOBL.GetValue('G_VENTILABLE' + IntToStr(j)) = 'X';
  GMP.Ref1 := '';
  GMP.Lib1 := '';
  GMP.Ref2 := '';
  GMP.Lib2 := '';
  GMP.RefExt1 := '';
  GMP.RefLib1 := '';
  GMP.RefExt2 := '';
  GMP.RefLib2 := '';
  GMP.Cat := '';
  GMP.NomFSelect := '';
  GMP.Dev.Code := V_PGI.DevisePivot;
  GMP.Dev.Taux := 1.00;
  GMP.ExportCFONB := FALSE;
  GMP.EnvoiBordereau := FALSE;
  GMP.FormatCFONB := '';
  GMP.ModeleBordereau := '';
  GMP.ModeTeletrans := '';
  GMP.AlerteEcheMP := FALSE;
  GMP.TIDTIC := (TOBL.GetValue('G_NATUREGENE') = 'TID') or (TOBL.GetValue('G_NATUREGENE') = 'TIC');
  GMP.Lettrauto := TRUE; // YMO 25/11/2005 : Lettrage désactivé pour les reglts en devise
  Result := TRUE;

end;

procedure TFGenereMP.Charge;
var
  dateTaux: TDateTime;
begin
  CompteGeneration.Text := GMP.CptG;
  ModeGenere.Value := GMP.MPG;
  JournalGenere.Value := GMP.JalG;
  GroupeEncaDeca.Value := GMP.GroupeEncadeca;
  NumEncaDeca.Text := GMP.NumEncadeca;
  if GMP.ModeAnal <> '' then Anal.Value := GMP.ModeAnal;
  if GMP.DateC = 0 then DateGeneration.Text := DateToStr(V_PGI.DateEntree)
  else DateGeneration.Text := DateToStr(GMP.DateC);
  CChoixDate.Checked := GMP.ForceEche;
  DateEcheance.Text := DateToStr(GMP.DateE);
  RefEcr1.Text := GMP.Ref1;
  LibEcr1.Text := GMP.Lib1;
  RefEcr2.Text := GMP.Ref2;
  LibEcr2.Text := GMP.Lib2;
  RefExt1.Text := GMP.RefExt1;
  RefLib1.Text := GMP.RefLib1;
  RefExt2.Text := GMP.RefExt2;
  RefLib2.Text := GMP.RefLib2;
  ExportCFONB.Checked := GMP.ExportCFONB;
  EnvoiBordereau.Checked := GMP.EnvoiBordereau;
  if GMP.ExportCFONB then FormatCFONB.Value := GMP.FormatCFONB else FormatCFONB.ItemIndex := -1;
  if GMP.EnvoiBordereau then ModeleBordereau.Value := GMP.ModeleBordereau else ModeleBordereau.ItemIndex := -1;
  if GMP.ModeTeletrans = '' then ModeTeletrans.ItemIndex := -1 else ModeTeletrans.Value := GMP.ModeTeletrans;
  if not VH^.OldTeleTrans then ModeTeletrans.ItemIndex := -1;
  if GMP.Dev.Code = '' then GMP.Dev.Code := V_PGI.DevisePivot;
  TDEVISE.Caption := RechDom('ttDevisetoutes', GMP.DEV.Code, False);
  TauxDEV.visible := FALSE;
  TTauxDev.visible := FALSE;
  if (GMP.Dev.Code = V_PGI.DevisePivot) or (VH^.TenueEuro and (GMP.Dev.Code = V_PGI.DeviseFongible)) then
  begin
    HDEVISE.Visible := FALSE;
    TDEVISE.Visible := FALSE;
    TauxDev.Value := 1.00;
    TauxDev.Enabled := FALSE;
  end else
  begin
    if Arrondi(GMP.Dev.Taux, 6) = 0 then GMP.Dev.Taux := GETTAUX(GMP.Dev.Code, DateTaux, StrToDate(DateGeneration.Text));
    TauxDev.Value := GMP.Dev.Taux / V_PGI.TauxEuro;
    TauxDev.Enabled := TRUE;
    TauxDev.visible := TRUE;
    TTauxDev.visible := TRUE;
    LettrageAuto.visible := TRUE;
  end;
  AlerteEcheMP.Checked := GMP.AlerteEcheMP;
  GMP.LettrAuto:=TRUE;
  LettrageAuto.Checked := GMP.LettrAuto;
end;

procedure TFGenereMP.Sauve;
var
  DateTaux: TDateTime;
begin
  GMP.CptG := CompteGeneration.Text;
  GMP.MPG := ModeGenere.Value;
  GMP.NatureG := 'OD';
  if GMP.NatJalG = 'BQE' then
  begin
    if IsEnc(GMP.smp) then GMP.NatureG := 'RC' else GMP.NatureG := 'RF';
  end;
  GMP.ModeAnal := Anal.Value;
  GMP.JalG := JournalGenere.Value;
  GMP.GroupeEncadeca := GroupeEncaDeca.Value;
  GMP.NumEncadeca := NumEncaDeca.Text;
  GMP.DateC := StrToDate(DateGeneration.Text);
  GMP.DateE := StrToDate(DateEcheance.Text);
  GMP.ForceEche := CChoixDate.Checked;
  GMP.Ref1 := RefEcr1.Text;
  GMP.Lib1 := LibEcr1.Text;
  GMP.Ref2 := RefEcr2.Text;
  GMP.Lib2 := LibEcr2.Text;
  GMP.RefExt1 := RefExt1.Text;
  GMP.RefLib1 := RefLib1.Text;
  GMP.RefExt2 := RefExt2.Text;
  GMP.RefLib2 := RefLib2.Text;
  GMP.ExportCFONB := ExportCFONB.Checked;
  GMP.EnvoiBordereau := EnvoiBordereau.Checked;
  GMP.LettrAuto := LettrageAuto.Checked; // YMO 25/11/2005 : Lettrage désactivé si reglts en devise
  if GMP.ExportCFONB then GMP.FormatCFONB := FormatCFONB.Value else GMP.FormatCFONB := '';
  if GMP.EnvoiBordereau then GMP.ModeleBordereau := ModeleBordereau.Value else GMP.ModeleBordereau := '';
  if ModeTeletrans.ItemIndex = -1 then GMP.ModeTeletrans := '' else GMP.ModeTeletrans := ModeTeletrans.Value;
  if not PEmission.TabVisible then
  begin
    GMP.ExportCFONB := FALSE;
    GMP.EnvoiBordereau := FALSE;
    GMP.FormatCFONB := '';
    GMP.ModeleBordereau := '';
    GMP.ModeTeletrans := '';
  end;
  CreerCodeLot(NumEncaDeca.Text);
  if (GMP.Dev.Code = V_PGI.DevisePivot) or (VH^.TenueEuro and (GMP.Dev.Code = V_PGI.DeviseFongible)) then
  begin
    GMP.Dev.Taux := 1.00;
  end else
  begin
    GMP.Dev.Taux := TauxDev.VALUE * V_PGI.TauxEuro;
  end;
  GMP.AlerteEcheMP := AlerteEcheMP.Checked;
end;

procedure TFGenereMP.Init;
begin
  case GMP.smp of
    smpEncPreBqe: CompteGeneration.DataType := 'TZGBANQUE';
    smpEncTraEdt: CompteGeneration.DataType := 'TZGENCAIS';
    smpEncTraPor: CompteGeneration.DataType := 'TZGENCAIS';
    smpEncChqPor: CompteGeneration.DataType := 'TZGENCAIS';
    smpEncCBPor: CompteGeneration.DataType := 'TZGENCAIS';
    smpEncTraEdtNC: CompteGeneration.DataType := 'TZGENCAIS';
    smpEncTraEnc: CompteGeneration.DataType := 'TZGENCAIS';
    smpEncTraEsc: CompteGeneration.DataType := 'TZGENCAIS';
    smpEncChqBqe: CompteGeneration.DataType := 'TZGBANQUE';
    smpEncCBBqe: CompteGeneration.DataType := 'TZGBANQUE';
    smpEncTraBqe: CompteGeneration.DataType := 'TZGBANQUE';
    smpEncPreEdt: CompteGeneration.DataType := 'TZGBANQUE';
    smpEncDiv: CompteGeneration.DataType := 'TZGBILAN';
    smpEncTous: CompteGeneration.DataType := 'TZGBILAN';
   { YMO 02/01/2006 FQ10237 Changement de la tablette pour ajouter les comptes TIC et DIVERS
   (édition des lettres-chèque et édition comptabilisation des lettres-chèque }
    smpDecChqEdt: begin
                      CompteGeneration.plus := 'G_NATUREGENE = "BQE" OR G_NATUREGENE = "TIC" OR G_NATUREGENE = "DIV"';
                      CompteGeneration.DataType := 'TZGENERAL';
                  end;
    smpDecChqEdtNC: CompteGeneration.DataType := 'TZGBANQUE';
    smpDecVirBqe: CompteGeneration.DataType := 'TZGBANQUE';
    smpDecVirEdt: CompteGeneration.DataType := 'TZGBANQUE';
    smpDecVirEdtNC: CompteGeneration.DataType := 'TZGBANQUE';
    smpDecVirInBqe: CompteGeneration.DataType := 'TZGBANQUE';
    smpDecVirInEdt: CompteGeneration.DataType := 'TZGBANQUE';
    smpDecVirInEdtNC: CompteGeneration.DataType := 'TZGBANQUE';
    smpDecBorEdt: CompteGeneration.DataType := 'TZGDECAIS';
    smpDecBorEdtNC: CompteGeneration.DataType := 'TZGDECAIS';
    smpDecTraPor: CompteGeneration.DataType := 'TZGDECAIS';
    smpDecBorDec: CompteGeneration.DataType := 'TZGBANQUE';
    smpDecBorEsc: CompteGeneration.DataType := 'TZGBANQUE';
    smpDecTraBqe: CompteGeneration.DataType := 'TZGBANQUE';
    smpDecDiv: CompteGeneration.DataType := 'TZGBILAN';
    smpDecTous: CompteGeneration.DataType := 'TZGBILAN';
  else CompteGeneration.DataType := 'TZGBILAN';
  end;
  //YMO 07/06/2006 FQ17524 Déplacement hors du filtre pour le filtre par défaut
  LettrageAuto.Checked := GMP.LettrAuto;
end;

procedure TFGenereMP.FormShow(Sender: TObject);
begin
  LoadLeFiltre := FALSE;
  Init;
  CatToMP(Cat, MODEGENERE.Items, MODEGENERE.Values, tslAucun, TRUE);
  GroupeEncadeca.ItemIndex := 0;
  ModeGenere.itemIndex := 0;
  Anal.ItemIndex := 2;
  //RRO le 10/03/20003 modif fonctionnelle pour smpEncDiv, smpDevDiv
  PEmission.TabVisible := GMP.smp in [smpEncDiv, smpDecDiv, smpEncPreBqe, smpEncPreEdt, smpEncTraEnc, smpEncTraEsc, smpEncTraBqe,
    smpDecVirBqe, smpDecVirEdt, smpDecVirInBqe, smpDecVirInEdt, smpDecBorDec, smpDecBorEsc, smpDecTraBqe];
  Charge;
  LoadLeFiltre := TRUE;
  {JP 19/08/04 : FQ 14075 : nouvelle gestion des filtres}
  ObjFiltre.FFI_TABLE := FNomFiltre;
  ObjFiltre.Charger;

  LoadLeFiltre := FALSE;
  if GMP.smp in [smpDecChqEdtNC, smpDecChqEdt] then
  begin
    REFECR1.Enabled := FALSE;
    TREFECR1.Enabled := FALSE;
    REFECR2.Enabled := FALSE;
    TREFECR2.Enabled := FALSE;
    REFECR1.Text := '';
    REFECR2.Text := '';
  end;
  if not VH^.OldTeleTrans then
  begin
    ModeTeleTrans.Visible := FALSE;
    TER_ENVOITRANS.Visible := FALSE;
  end;
  {JP 22/02/05 : Gestion de la banque prévisionnelle
   JP 09/08/05 : FQ 15815 : gestion de la banque prévisionnelle dans la remise en banque des BOR (smpDecBorDec)}
  if (GMP.smp in [smpEncDiv, smpDecDiv, smpDecBorDec]) then begin
    if not GereBanquePrevi then
      {Si on n'a pas réussi à récupérer le journal, on sort}
      PostMessage(Self.Handle, WM_CLOSE, 1, 0);
  end;

  //SDA le 27/12/2007 version belge - éléments CFONB rendus invisibles si contexte belge
  if VH^.PaysLocalisation<>CodeISOFR then
  begin
    EXPORTCFONB.Visible := False;
    HLABEL10.Visible := False;
    FORMATCFONB.Visible := False;
  end;
  //Fin SDA le 27/12/2007

end;

function TFGenereMP.CompteOk: Boolean;
var
  Q: tQuery;
  Err: Integer;
  St, NatCpt: string;
begin
  Err := 0; //Result:=FALSE ;
  GMP.GenColl := FALSE;
  GMP.GenPoint := FALSE;
  GMP.GenLett := FALSE;
  GMP.GenVent := FALSE;
  GMP.Ventils[1] := FALSE;
  GMP.Ventils[2] := FALSE;
  GMP.Ventils[3] := FALSE;
  GMP.Ventils[4] := FALSE;
  GMP.Ventils[5] := FALSE;
  GMP.NatJalG := '';
  NatCpt := '';
  St := 'SELECT G_COLLECTIF, G_POINTABLE, G_LETTRABLE, G_VENTILABLE, G_VENTILABLE1, '
    + 'G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5, G_NATUREGENE FROM GENERAUX '
    + 'WHERE G_GENERAL="' + CompteGeneration.Text + '" ';
  Q := OpenSQL(St, TRUE);
  if not Q.Eof then
  begin
    GMP.GenColl := Q.FindField('G_COLLECTIF').AsString = 'X';
    GMP.GenPoint := Q.FindField('G_POINTABLE').AsString = 'X';
    GMP.GenLett := Q.FindField('G_LETTRABLE').AsString = 'X';
    GMP.GenVent := Q.FindField('G_VENTILABLE').AsString = 'X';
    GMP.Ventils[1] := Q.FindField('G_VENTILABLE1').AsString = 'X';
    GMP.Ventils[2] := Q.FindField('G_VENTILABLE2').AsString = 'X';
    GMP.Ventils[3] := Q.FindField('G_VENTILABLE3').AsString = 'X';
    GMP.Ventils[4] := Q.FindField('G_VENTILABLE4').AsString = 'X';
    GMP.Ventils[5] := Q.FindField('G_VENTILABLE5').AsString = 'X';
    NatCpt := Q.FindField('G_NATUREGENE').AsString;
  end else
  begin
    Err := 3;
    Pages.ActivePage := PGenere;
    CompteGeneration.SetFocus;
  end;
  Ferme(Q);
  if Err = 0 then
  begin
    Q := OpenSQL('SELECT J_JOURNAL,J_CONTREPARTIE, J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="' + JournalGenere.Value + '" ', TRUE);
    if not Q.Eof then
    begin
      if (Q.FindField('J_NATUREJAL').AsString = 'BQE') and (Q.FindField('J_CONTREPARTIE').AsString <> CompteGeneration.Text) then
      begin
        Err := 6;
        Pages.ActivePage := PGenere;
        CompteGeneration.SetFocus;
      end;
      GMP.NatJalG := Q.FindField('J_NATUREJAL').AsString;
    end else
    begin
      Err := 4;
      Pages.ActivePage := PGenere;
      JournalGenere.SetFocus;
    end;
    Ferme(Q);
  end;
  if Err = 0 then
  begin
    if GMP.GenColl and ((GroupeEncaDeca.Value = 'GLO') or (GroupeEncaDeca.Value = 'ECT')) then
    begin
      Err := 5;
      Pages.ActivePage := PGenere;
      GroupeEncaDeca.SetFocus;
    end;
  end;
  if Err = 0 then
  begin
    if GMP.smp in [smpEncPreBqe..smpEncDiv] then
    begin
      if (NatCpt = 'COF') or (NatCpt = 'TIC') then Err := 14;
    end else if GMP.smp in [smpDecVirBqe..smpDecDiv] then
    begin
      if (NatCpt = 'COC') or (NatCpt = 'TID') then Err := 14;
    end;
  end;
  if Err = 0 then
  begin
    if GMP.GenColl and GMP.TIDTIC then Err := 15;
  end;
  if Err <> 0 then
  begin
    Result := FALSE;
    HM.Execute(Err, Caption, '');
    Exit;
  end {Else Result:=TRUE };
  Result := Err = 0;
  {05/06/07 : FQ 18660 : On relance le test sur le général s'il y a un filtre}
  if Result then Result := (FFiltres.Value = '') or TesteGeneral(CompteGeneration.Text);
end;

procedure TFGenereMP.BOuvrirClick(Sender: TObject);
var i :integer ;
begin
  if LettrageAuto.Checked then i:=2 else i:=17 ; //YMO 12/10/2006 message d'avertissement en devise si choix de non lettrage
  if PourGeneration then if HM.Execute(i, Caption, '') <> mrYes then Exit;
  if not CompteOk then Exit;
  Sauve;
  if PourGeneration then
  begin
    if ParamsMPSup(GMP) then ModalResult := mrOk else Exit;
  end else ModalResult := mrOk;
end;

procedure TFGenereMP.CChoixDateClick(Sender: TObject);
begin
  if CChoixDate.Checked then
  begin
    DateEcheance.Enabled := True;
    if DateEcheance.Text = StDate1900 then DateEcheance.Text := DateToStr(V_PGI.DateEntree);
    if ModeGenere.ItemIndex > 0 then
    begin
      AlerteEcheMP.Enabled := FALSE;
      AlerteEcheMP.Checked := FALSE;
    end else AlerteEcheMP.Enabled := TRUE;
  end else
  begin
    DateEcheance.Text := StDate1900;
    DateEcheance.Enabled := False;
  end;
end;

{JP 05/06/07 : FQ 18660 : on s'assure que le compte n'est pas fermé
               Par la même occasion, on fait le test sur le journal
{---------------------------------------------------------------------------------------}
procedure TFGenereMP.CompteGenerationExit(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Q: TQuery;
begin
  if not TesteGeneral(CompteGeneration.Text) then CompteGeneration.Text := ''; {FQ 18660}

  if (CompteGeneration.Text <> '') and (JournalGenere.Value = '') then begin
    Q := OpenSQL('SELECT J_JOURNAL, J_FERME FROM JOURNAL WHERE J_CONTREPARTIE="' + CompteGeneration.Text + '" ', TRUE);
    if not Q.Eof then begin
      if Q.Fields[1].AsString = 'X' then
        PGIError(TraduireMemoire('Il est impossible d''utiliser un journal de génération fermé'), Caption)
      else
        JournalGenere.Value := Q.Fields[0].AsString;
    end;
    Ferme(Q);
  end;
end;

{JP 05/06/07 : FQ 18660 : on s'assure que le compte n'est pas fermé
               Par la même occasion, on fait le test sur le journal
{---------------------------------------------------------------------------------------}
procedure TFGenereMP.JournalGenereExit(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Q: TQuery;
begin
  if (CompteGeneration.Text = '') and (JournalGenere.Value <> '') then begin
    Q := OpenSQL('SELECT J_CONTREPARTIE, J_FERME FROM JOURNAL WHERE J_JOURNAL="' + JournalGenere.Value + '" ', TRUE);
    if not Q.Eof then begin
      if Q.Fields[1].AsString = 'X' then begin
        JournalGenere.ItemIndex := -1;
        PGIError(TraduireMemoire('Il est impossible d''utiliser un journal de génération fermé'), Caption);
      end
      else
        CompteGeneration.Text := Q.Fields[0].AsString;
    end;
    Ferme(Q);
    if not TesteGeneral(CompteGeneration.Text) then CompteGeneration.Text := ''; {FQ 18660}
  end;
end;

procedure TFGenereMP.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  {JP 12/08/05 : FQ 10445 : Validation sur F10}
  if Key=VK_F10 then begin
    Key := 0;
    BOuvrirClick(BOuvrir);
  end
  else
  if ((TauxDev.Focused) and (Key = VK_F5)) then
  begin
    Key := 0;
    if SaisieNewTaux2000(GMP.DEV, StrToDate(DateGeneration.Text)) then begin
      TauxDev.Value := GMP.Dev.Taux / V_PGI.TauxEuro;
      {JP 13/06/06 : FQ 18319 : Pour ne pas reprendre le taux de la table Chancell dans GenerMP}
      GMP.TauxVolatil := True;
    end;
  end;
end;

procedure TFGenereMP.TAUXDEVKeyPress(Sender: TObject; var Key: Char);
begin
  Key := #0;
end;

procedure TFGenereMP.MODEGENEREChange(Sender: TObject);
begin
  if (MODEGENERE.ItemIndex > 0) and CCHoixDate.Checked then
  begin
    AlerteEcheMP.Enabled := FALSE;
    AlerteEcheMP.Checked := FALSE;
  end else AlerteEcheMP.Enabled := TRUE;
end;



procedure TFGenereMP.BAssistClick(Sender: TObject);
var
  St, StC: string;
  TE: tEdit;
begin
  inherited;
  TE := TEdit(Sender);
  St := ChoixChampZone(0, 'LIB');
  if St = '' then Exit;
  Stc := TE.Text;
  if Length(Stc + ' ' + St) > 100 then HM.Execute(16, Caption, '');
  if TE.SelLength > 0 then Delete(StC, TE.SelStart + 1, TE.SelLength);
  if StC = '' then StC := St else StC := StC + ' ' + St;
  TE.Text := StC;
end;

procedure TFGenereMP.REFECR1Change(Sender: TObject);
var
  THS, THD: tHCritMaskEdit;
begin
  if ObjFiltre.InChargement then Exit; //FQ 14879 SG6 03/11/2004
  if LoadLeFiltre then Exit;
  THS := tHCritMaskEdit(Sender);
  if THS = nil then Exit;
  case THS.Tag of
    1: THD := REFECR2;
    2: THD := LIBECR2;
    3: THD := REFEXT2;
    4: THD := REFLIB2;
  else Exit;
  end;
  THD.Text := THS.Text;
end;

procedure TFGenereMP.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

{JP 19/08/04 : FQ 14075 : nouvelle gestion des filtres
{---------------------------------------------------------------------------------------}
procedure TFGenereMP.FormCreate(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Composants: TControlFiltre;
begin
  Composants.PopupF := POPF;
  Composants.Filtres := FFILTRES;
  Composants.Filtre := BFILTRE;
  Composants.PageCtrl := PAGES;
  {A cet endroit FNomFiltre est vide, il sera mis à jour dans le FormShow}
  ObjFiltre := TObjFiltre.Create(Composants, FNomFiltre);
  ObjFiltre.ApresChangementFiltre := ApresChargement;
end;

{JP 19/08/04 : FQ 14075 : nouvelle gestion des filtres
{---------------------------------------------------------------------------------------}
procedure TFGenereMP.FormClose(Sender: TObject; var Action: TCloseAction);
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ObjFiltre) and (Action = caFree) then FreeAndNil(ObjFiltre);
end;

{JP 22/02/05 : Gestion de la banque prévisionnelle
{---------------------------------------------------------------------------------------}
function TFGenereMP.GereBanquePrevi : Boolean;
{---------------------------------------------------------------------------------------}
var
  Q     : TQuery;
  Multi : Boolean;
begin
  Result := True;
  if GMP.BanquePrevi <> '' then begin
    {Récupération du journal en fonction de la banque prévisionnelle
     JP 05/06/06 : FQ 17388 : ajout de J_MODESAISIE qui figure dans la clause Plus de JournalGenere}
    Q := OpenSQL('SELECT J_JOURNAL FROM JOURNAL WHERE J_CONTREPARTIE = "' + GMP.BanquePrevi + '" AND J_MODESAISIE = "-"', True);
    Multi := Q.RecordCount > 1;
    if not Q.EOF then
      CodeJournal := Q.Fields[0].AsString
    else begin
      HShowMessage('0;' + Caption + ';Impossible de récupérer le journal;W;O;O;O;', '', '');
      Result := False;
    end;
    Ferme(Q);

    {JP 05/06/06 : FQ 17388 : si plusieurs journaux, on laisse la possiblité de sélectionner le journal}
    if Multi then begin
      JournalGenere.Plus    := 'AND J_MODESAISIE = "-" AND J_CONTREPARTIE = "' + GMP.BanquePrevi + '" ';
      JournalGenere.Enabled := True;
    end else begin
      JournalGenere.Enabled := False;
    end;
    JournalGenere.Value   := CodeJournal;
    {Le compte est forcé et désactivé}
    CompteGeneration.Enabled := False;
    CompteGeneration.Text := GMP.BanquePrevi;
  end;
end;

{JP 22/02/05 : Pour annuler le chargement du journal et de la banque prévisionnelle si
               celle-ci est renseignée lors du chargement des filtres dans les Enc/Dec
{---------------------------------------------------------------------------------------}
procedure TFGenereMP.ApresChargement;
{---------------------------------------------------------------------------------------}
begin
  {JP 09/08/05 : FQ 15815 : gestion de la banque prévisionnelle dans la remise en banque des BOR (smpDecBorDec)}
  if (GMP.BanquePrevi <> '') and (GMP.smp in [smpEncDiv, smpDecDiv, smpDecBorDec]) then begin
    JournalGenere.Value   := CodeJournal;
    CompteGeneration.Text := GMP.BanquePrevi;
  end;
end;

{JP 05/06/07 : FQ 18660 : on s'assure que le compte n'est pas fermé
{---------------------------------------------------------------------------------------}
function TFGenereMP.TesteGeneral(Cpte : string) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := ExisteSQL('SELECT G_FERME FROM GENERAUX WHERE (G_FERME <> "X" OR G_FERME IS NULL) AND G_GENERAL = "' + Cpte + '"');
  if not Result then PGIError(TraduireMemoire('Il est impossible d''utiliser un compte de génération fermé'), Caption);
end;

end.

