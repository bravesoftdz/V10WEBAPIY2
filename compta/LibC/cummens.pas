unit CUMMENS;

//================================================================================
// Interface
//================================================================================
interface

uses
  SysUtils,
  WinTypes,
  WinProcs,
  Messages,
  Classes,
  Graphics,
  Controls,
  Forms,
  Dialogs,
  Grids,
  Hctrls,
  StdCtrls,
  Buttons,
  ExtCtrls,
  {$IFDEF EAGLCLIENT}
  UTOB,
  {$ELSE}
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  DBGrids,
  HQry,
  {$ENDIF}
  MajTable,
  TabNotBk,
  Ent1,
  Hcompte,
  DBCtrls,
  HEnt1,
  SAISUTIL,
  CPTEUTIL,
  hmsgbox,
  HSysMenu,
  HStatus,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ELSE}
  tCalcCum,
  {$ENDIF MODENT1}
  {$IFDEF COMPTA}
  UtilEdt, // InitCritEdt
  {$ENDIF}
  AGLInit, // TheData
  UentCommun
  ;

//==================================================
// Externe
//==================================================
procedure CumulCpteMensuel(NatureCpte: TFichierBase; Cpte: String17; Lib: string; Exercice: TexoDate);

//==================================================
// Definition de class
//==================================================
type
  TFCumMens = class(TForm)

    Fliste: THGrid;
    CBDebDat: TComboBox;
    CBFinDate: TComboBox;
    LAnouv: TLabel;
    LTotaux: TLabel;
    CbExo: THValComboBox;
    CbDevise: THValComboBox;
    CbSoldCum: TCheckBox;
    CbTypMvt: THValComboBox;
    NumEditSold: THNumEdit;
    CbEtab: THValComboBox;
    MsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    Rev: TCheckBox;
    Dexo2: TEdit;
    Dexo1: TEdit;
    BGraph: THBitBtn;
    BValider: THBitBtn;

    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);

    procedure BChercherClick(Sender: TObject);
    procedure FlisteDblClick(Sender: TObject);

    procedure BImprimerClick(Sender: TObject);
    procedure BGraphClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);

    procedure CbSoldCumClick(Sender: TObject);
    procedure CbDeviseChange(Sender: TObject);
    procedure CbExoChange(Sender: TObject);
    procedure CbTypMvtChange(Sender: TObject);
    procedure CbEtabChange(Sender: TObject);

  private
    QuelCode: String17;
    Lib: string;
    ExerciceEntree: TExoDate;
    LeTypeFic: TFichierBase;
    LeTypeTab: TZoomTable;
    CodExo, CodEtb: String3;
    CodMvt, FAxe: String3;
    Devise: RDevise;
    FiltreDev, FiltreEtab, FiltreExo: Boolean;
    LeMvt: Byte;
    WMinX, WMinY: Integer;

    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    procedure CumulMensuel;
    procedure CumulMensuel_UDF;
    procedure ChoixTable;
    {$IFNDEF EAGLCLIENT}
    procedure CreateViewResult;
    {$ENDIF}
    procedure AffichePeriode(NbrMois: Integer);
  public
    { Public-déclarations }
  end;

  //================================================================================
  // Implementation
  //================================================================================
implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  ULibExercice,
  CPProcMetier,
  CPVersion,
  {$ENDIF MODENT1}
  {$IFDEF EAGLCLIENT}
   utileAgl, {LanceEtatTob}
  {$ELSE}
  PrintDBG,
  {JP 09/03/07 : Unité inutile, me semble-t-il ?
  CpteSav,}
  Graph, {FQ 14046}
  {$ENDIF}
  CritEdt,
  UtilPGI
  {$IFDEF COMPTA}
{$IFNDEF CMPGIS35}
  , uTofCPGLAna {JP 19/08/05 : FQ 16485 : Branchement du nouveau grand livre au lieu de QRGLANA}
  , uTofCPGLGene
  , uTofCPGLAuxi
{$ENDIF}
  {$ENDIF}
  ;

//==================================================
// fonctions hors class : Point d'entré
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CumulCpteMensuel(NatureCpte: TFichierBase; Cpte: String17; Lib: string; Exercice: TexoDate);
var
  FCumMens: TFCumMens;
begin
  if (_Blocage(['nrCloture'], true, 'nrAucun')) then exit;

  FCumMens := TFCumMens.Create(Application);

  try
    FCumMens.LeTypeFic := NatureCpte;
    FCumMens.QuelCode := Cpte;
    FCumMens.Lib := Lib;
    FCumMens.ExerciceEntree := Exercice;
    FCumMens.Showmodal;
  finally
    FCumMens.Free;
  end;

  Screen.Cursor := SyncrDefault;
end;

//==================================================
// Fonction de la class : Evenement de la fiche
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.FormCreate(Sender: TObject);
begin
  PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
  WMinX := Width;
  WMinY := Height;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.FormShow(Sender: TObject);
begin
  NumEditSold.Visible := False;
    {$IFDEF CCS3}
  Rev.Visible := False;
  {$ELSE}
  Rev.Visible := V_PGI.Controleur;
  {$ENDIF}

  {$IFDEF EAGLCLIENT}
  BGraph.Visible := false;
  {$ENDIF}
  if (Rev.Visible) then
  begin
    Rev.Checked := TRUE;
    Rev.State := cbGrayed;
  end
  else
  begin
    Rev.Checked := FALSE;
    Rev.State := cbUnchecked;
  end;

  ChoixTable;

  case (LeTypeFic) of
    fbGene:
      begin
        Caption := MsgBox.Mess[0] + QuelCode + ' ' + Lib;
        HelpContext := 7109060;
      end;
    fbAux:
      begin
        Caption := MsgBox.Mess[1] + QuelCode + ' ' + Lib;
        HelpContext := 7142060;
      end;
    fbJal:
      begin
        {JP 24/10/05 : FQ 16857 : le grand livre n'a pas de raison d'être sur les journaux}
        BValider.Visible := False;
        
        Caption := MsgBox.Mess[2] + QuelCode + ' ' + Lib;
        HelpContext := 7208050;
      end;
    fbAxe1..fbAxe5:
      begin
        Caption := MsgBox.Mess[3] + QuelCode + ' ' + Lib;
        HelpContext := 7175040;
      end;
  end;

  FiltreDev := False;
  FiltreEtab := False;
  FiltreExo := False;
  CbDevise.ItemIndex := 0;
  CbExo.Value := ExerciceEntree.Code;
  CbTypMvt.Value := 'NOR';
//  LeMvt := 4;
  //Sg6 13/01/05 FQ 14777
  LeMvt := 0;  //pour ecriture normale LeMvt = 0 voir procedure CbTypMvtChange 
  CbEtab.ItemIndex := 0;
  PositionneEtabUser(cbEtab);

  if (CbDevise.Value = '') then Devise.Code := ''
  else Devise.Code := CbDevise.Value;
  GetInfosDevise(Devise);
  {$IFNDEF EAGLCLIENT}
  if ((V_PGI.UDF) and (V_PGI.Driver <> dbMSACCESS)) then CreateViewResult;
  {$ENDIF}
  if ((V_PGI.UDF) and (V_PGI.Driver <> dbMSACCESS)) then CumulMensuel_UDF
  else CumulMensuel;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.BChercherClick(Sender: TObject);
begin
  ChangeMask(NumEditSold, Devise.Decimale, '');

  if ((V_PGI.UDF) and (V_PGI.Driver <> dbMSACCESS)) then CumulMensuel_UDF
  else CumulMensuel;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.BImprimerClick(Sender: TObject);
  {$IFDEF EAGLCLIENT}
  var    {Lek 210706}
    TobGrille,
    TobLigne : TOB;
    i  : integer;
  {$ENDIF}
begin
  {$IFDEF EAGLCLIENT}
  {Lek 210706}
  TobGrille := Tob.Create ('cumul mensuel',nil,-1);
  for i:=1 to FListe.RowCount - 1 do begin
    TobLigne := Tob.Create ('un ligne',TobGrille,-1);
    TobLigne.AddChampSupValeur ('_PERIODE',FListe.CellValues[0,i]);
    TobLigne.AddChampSupValeur ('_DEBIT',FListe.CellValues[1,i]);
    TobLigne.AddChampSupValeur ('_CREDIT',FListe.CellValues[2,i]);
    TobLigne.AddChampSupValeur ('_SOLDE',FListe.CellValues[3,i]);
    end;
  LanceEtatTob ('E','CST','CMS',TobGrille,true,false,false,nil,'',Caption,false);
  TobGrille.Free;
  {$ELSE}
  PrintDBGrid(FListe, nil, Caption, '');
  {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.BGraphClick(Sender: TObject);
begin
  {$IFDEF EAGLCLIENT}
  {FQ 14046}
// VisuGraph('CUMMENS' + IntToStr(Integer(LeTypeFic)), Caption, FListe, CBDebDat.Items, 0, 2, FListe.RowCount - 2, TRUE, [3, 1, 2], nil, nil);
  {$ELSE}
  VisuGraph('CUMMENS' + IntToStr(Integer(LeTypeFic)), Caption, FListe, CBDebDat.Items, 0, 2, FListe.RowCount - 2, TRUE, [3, 1, 2], nil, nil);
  {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.BFermeClick(Sender: TObject);
begin
  Close;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.WMGetMinMaxInfo(var MSG: Tmessage);
begin
  with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do
  begin
    X := WMinX;
    Y := WMinY;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... : 17/02/2004
Description .. : GCO - 17/02/2004
Suite ........ : -> Branchement du nouveau GL avec le CritEDT
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.FlisteDblClick(Sender: TObject);
{$IFDEF COMPTA}
var
  lCritEdt: ClassCritEdt;
  {$ENDIF}
begin
  {$IFDEF COMPTA}
  if (Fliste.Row = 1) then Exit;
  lCritEdt := ClassCritEdt.Create;
  // Init du Record
  Fillchar(lCritEdt.CritEdt, SizeOf(lCritEdt.CritEdt), #0);

  if (Fliste.Row <> Fliste.RowCount - 1) then
  begin
    lCritEdt.CritEdt.Date1 := StrToDate(CBDebDat.Items[Fliste.Row]);
    lCritEdt.CritEdt.Date2 := StrToDate(CBFinDate.Items[Fliste.Row]);
  end
  else
  begin
    lCritEdt.CritEdt.Date1 := StrToDate(CBDebDat.Items[2]);
    lCritEdt.CritEdt.Date2 := StrToDate(CBFinDate.Items[CBFinDate.Items.Count - 1]);
  end;

  // Préparation du CritEdt en fonction de la natuer de l'état
  lCritEdt.CritEdt.NatureEtat := neGL;
  InitCritEdt(lCritEdt.CritEdt);

  lCritEdt.CritEdt.Cpt1 := QuelCode;
  lCritEdt.CritEdt.Cpt2 := QuelCode;
  lCritEdt.CritEdt.DateDeb := lCritEdt.CritEdt.Date1;
  lCritEdt.CritEdt.DateFin := lCritEdt.CritEdt.Date2;

  lCritEdt.CritEdt.Valide := '';
  lCritEdt.CritEdt.ModeRevision := Rev.State;
  lCritEdt.CritEdt.Etab := CbEtab.Value;
  lCritEdt.CritEdt.DeviseSelect := CbDevise.Value;

  if cbTypMvt.Value <> 'TOU' then
  begin
    if cbTypMvt.Value = 'NOR' then lCritEdt.CritEdt.QualifPiece := 'N;'
    else
      if cbTypMvt.Value = 'NSS' then lCritEdt.CritEdt.QualifPiece := 'N;S;U;'
    else
      if cbTypMvt.Value = 'PRE' then lCritEdt.CritEdt.QualifPiece := 'P;'
    else
      if cbTypMvt.Value = 'SSI' then lCritEdt.CritEdt.QualifPiece := 'S;U;'
  end;
  TheData := lCritEdt;

  {$IFNDEF IMP}
  case LeTypeFic of
    // Grand Livre Général
    fbGene: CPLanceFiche_CPGLGene('');
    // Grand Livre Auxiliaire
    fbAux: CPLanceFiche_CPGLAuxi('');
    {$IFNDEF CCMP}
    {$IFDEF EAGLCLIENT}

    {$ELSE}
    fbAxe1..fbAxe5:
      begin
        lCritEdt.CritEdt.Gl.Axe := fbToAxe(LeTypeFic);
        {JP 19/08/05 : FQ 16485 : Branchement du nouveau grand livre ce qui permet de
                       contourner le problème de requête sur les écritures de révision
                       qui dépend d'une fonction AGL HQRY.TraduitNatureEcr
         GLAnalZoom(lCritEdt.CritEdt);}
        CPLanceFiche_CPGLANA;
      end;
    {$ENDIF}
    {$ENDIF}
  end;
  {$ENDIF}

  TheData := nil;
  FreeAndNil(lCritEdt);
  {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.CbSoldCumClick(Sender: TObject);
var
  i: Byte;
  D1, C1, MC, MD: Extended;
  PremMois, PremAnnee, NbrMois: Word;
begin
  NombrePerExo(ExerciceEntree, PremMois, PremAnnee, NbrMois);

  if (CbSoldCum.Checked) then
  begin
    Md := 0;
    MC := 0;
    for i := 1 to NbrMois + 1 do
    begin
      D1 := Valeur(Fliste.Cells[1, i]);
      C1 := Valeur(Fliste.Cells[2, i]);
      AfficheLeSolde(NumEditSold, MD + D1, MC + C1);
      FListe.Cells[3, i] := NumEditSold.Text;
      NumEditSold.Text := Fliste.Cells[3, i];
      if (NumEditSold.Debit) then
      begin
        MC := 0;
        MD := NumEditSold.Value;
      end
      else
      begin
        MC := NumEditSold.Value;
        MD := 0;
      end;
    end;
  end
  else
  begin
    for i := 1 to NbrMois + 1 do
    begin
      AfficheLeSolde(NumEditSold, Valeur(Fliste.Cells[1, i]), Valeur(Fliste.Cells[2, i]));
      FListe.Cells[3, i] := NumEditSold.Text;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.CbDeviseChange(Sender: TObject);
begin
  if (CbDevise.Value = '') then
  begin
    Devise.Code := '';
    FiltreDev := False;
  end
  else
  begin
    Devise.Code := CbDevise.Value;
    FiltreDev := True;
  end;
  GetInfosDevise(Devise);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.CbExoChange(Sender: TObject);
begin
  CodExo := CbExo.Value;
  FiltreExo := (CodExo <> '');
  if (CodExo <> '') then ExoToDates(CodExo, Dexo1, Dexo2)
  else CodExo := VH^.Entree.Code;
  ExerciceEntree.Code := CodExo;
  ExerciceEntree.Deb := StrToDate(Dexo1.Text);
  ExerciceEntree.Fin := StrToDate(Dexo2.Text);
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.CbEtabChange(Sender: TObject);
begin
  CodEtb := CbEtab.Value;
  FiltreEtab := (CodEtb <> '');
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.CbTypMvtChange(Sender: TObject);
begin
  CodMvt := CbTypMvt.Value;

  if (CodMvt = 'NOR') then LeMvt := 0
  else if (CodMvt = 'NSS') then LeMvt := 1
  else if (CodMvt = 'PRE') then LeMvt := 2
  else if (CodMvt = 'SSI') then LeMvt := 3
  else if (CodMvt = 'TOU') then LeMvt := 4
    ;
end;

//==================================================
// Fonction de la class : Autres fonctions
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.ChoixTable;
begin
  case LeTypeFic of
    fbGene: LeTypeTab := tzGeneral;
    fbAux: LeTypeTab := tzTiers;
    fbJal: LeTypeTab := tzJournal;
    fbAxe1..fbAxe5: LeTypeTab := tzSection;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.CumulMensuel;
var
  Q: TQuery;
  j: Integer;
  Debit, Credit, SDebit, SCredit, DebANew, CredANew: extended;
  DebDate, FinDate: TDateTime;
  PremMois, PremAnnee, NbrMois: Word;
  StM: string;
  ToTal, TotalEuro: TabTot;
begin
  NombrePerExo(ExerciceEntree, PremMois, PremAnnee, NbrMois);

  if EstSpecif('51185') then
  begin
    if (LeTypeFic <> fbJal) then Q := PrepareTotCpt(LeTypeFic, [], FiltreDev, FiltreEtab, FiltreExo, false, Un)
    else Q := PrepareTotJal(QuelCode, [], FiltreDev, FiltreEtab, FiltreExo, false);
  end;

  Debit := 0;
  Credit := 0;
  SDebit := 0;
  SCredit := 0;
  Fliste.RowCount := 3;
  CBFinDate.Items.Clear;
  CBDebDat.Items.Clear;
  CBDebDat.Items.Add('');
  CBFinDate.Items.Add(''); //Titre
  CBDebDat.Items.Add('');
  CBFinDate.Items.Add(''); //Anouveau
  DebDate := ExerciceEntree.Deb;

  //Pour les X_ECRANOUVEAU
  FinDate := FindeMois(DebDate);
  if (not EstSpecif('51185')) then
  begin
    if (LeTypeFic <> fbJal) then
      ExecuteTotCpt(Q, QuelCode, DebDate, FinDate, Devise.Code, CodEtb, CodExo, ToTal, TotalEuro, false, Devise.Decimale, V_PGI.OkDecE, LeTypeFic, [], Un, false)
    else ExecuteTotJal(Q, QuelCode, DebDate, FinDate, Devise.Code, CodEtb, CodExo, ToTal, TotalEuro, false, Devise.Decimale, V_PGI.OkDecE, [], false);
  end
  else ExecuteTotCpt(Q, QuelCode, DebDate, FinDate, Devise.Code, CodEtb, CodExo, ToTal, TotalEuro, false, Devise.Decimale, V_PGI.OkDecE, LeTypeFic, [], Un, false);

  FListe.Cells[0, 1] := LAnouv.Caption;
  DebANew := Total[0].TotDebit;
  CredANew := Total[0].TotCredit;
  if (LeTypeFic = fbJal) then
  begin
    FListe.Cells[1, 1] := StrFMontant(DebANew, 15, Devise.Decimale, Devise.Symbole, TRUE);
    FListe.Cells[2, 1] := StrFMontant(CredANew, 15, Devise.Decimale, Devise.Symbole, TRUE);
  end;

  AfficheLeSolde(NumEditSold, DebANew, CredAnew);
  FListe.Cells[3, 1] := NumEditSold.Text;
  if (LeTypeFic <> fbJal) then
  begin
    if (DebANew - CredANew <= 0) then
    begin
      DebANew := 0;
      FListe.Cells[2, 1] := StrFMontant(Abs(Valeur(FListe.Cells[3, 1])), 15, Devise.Decimale, Devise.Symbole, true);
      FListe.Cells[1, 1] := StrFMontant(DebANew, 15, Devise.Decimale, Devise.Symbole, true);
    end
    else
    begin
      CredANew := 0;
      FListe.Cells[1, 1] := StrFMontant(Abs(Valeur(FListe.Cells[3, 1])), 15, Devise.Decimale, Devise.Symbole, true);
      FListe.Cells[2, 1] := StrFMontant(CredANew, 15, Devise.Decimale, Devise.Symbole, true);
    end;
  end;

  FListe.ColAligns[0] := taCenter;
  InitMove(NbrMois, '');
  for j := 1 to NbrMois do
  begin
    if j = NbrMois then
      FinDate := ExerciceEntree.Fin // On est sur le dernier mois de l'exercice
    else
      FinDate := FindeMois(DebDate);
    //  ExecuteTotCpt(Q,QuelCode,DebDate,FinDate,Devise.Code,CodEtb,CodExo,Total,TotalEuro,FALSE,Devise.Decimale,V_PGI.OkDecE) ;
    if (not EstSpecif('51185')) then
    begin
      if (LeTypeFic <> fbJal) then
        ExecuteTotCpt(Q, QuelCode, DebDate, FinDate, Devise.Code, CodEtb, CodExo, ToTal, TotalEuro, false, Devise.Decimale, V_PGI.OkDecE, LeTypeFic, [], Un, false)
      else ExecuteTotJal(Q, QuelCode, DebDate, FinDate, Devise.Code, CodEtb, CodExo, ToTal, TotalEuro, false, Devise.Decimale, V_PGI.OkDecE, [], false);
    end
    else ExecuteTotCpt(Q, QuelCode, DebDate, FinDate, Devise.Code, CodEtb, CodExo, Total, TotalEuro, false, Devise.Decimale, V_PGI.OkDecE, LeTypeFic, [], Un, false);

    case LeMvt of
      (*NOR*)0:
        begin
          if (Rev.State = cbUnchecked) then
          begin
            Debit := Arrondi(Total[1].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[1].TotCredit, Devise.Decimale);
          end
          else if (Rev.State = cbGrayed) then
          begin
            Debit := Arrondi(Total[1].TotDebit + Total[5].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[1].TotCredit + Total[5].TotCredit, Devise.Decimale);
          end
          else if (Rev.State = cbChecked) then
          begin
            Debit := Arrondi(Total[5].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[5].TotCredit, Devise.Decimale);
          end;
        end;
      (*NSS*)1:
        begin
          if (Rev.State = cbUnchecked) then
          begin
            Debit := Arrondi(Total[1].TotDebit + Total[2].TotDebit + Total[4].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[1].TotCredit + Total[2].TotCredit + Total[4].TotCredit, Devise.Decimale);
          end
          else if (Rev.State = cbGrayed) then
          begin
            Debit := Arrondi(Total[1].TotDebit + Total[2].TotDebit + Total[4].TotDebit + Total[5].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[1].TotCredit + Total[2].TotCredit + Total[4].TotCredit + Total[5].TotCredit, Devise.Decimale);
          end
          else if (Rev.State = cbChecked) then
          begin
            Debit := Arrondi(Total[5].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[5].TotCredit, Devise.Decimale);
          end;
        end;
      (*PRE*)2:
        begin
          if (Rev.State = cbUnchecked) then
          begin
            Debit := Arrondi(Total[3].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[3].TotCredit, Devise.Decimale);
          end
          else if (Rev.State = cbGrayed) then
          begin
            Debit := Arrondi(Total[3].TotDebit + Total[5].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[3].TotCredit + Total[5].TotCredit, Devise.Decimale);
          end
          else if (Rev.State = cbChecked) then
          begin
            Debit := Arrondi(Total[5].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[5].TotCredit, Devise.Decimale);
          end;
        end;
      (*SSI*)3:
        begin
          if (Rev.State = cbUnchecked) then
          begin
            Debit := Arrondi(Total[2].TotDebit + Total[4].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[2].TotCredit + Total[4].TotCredit, Devise.Decimale);
          end
          else if (Rev.State = cbGrayed) then
          begin
            Debit := Arrondi(Total[2].TotDebit + Total[4].TotDebit + Total[5].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[2].TotCredit + Total[4].TotCredit + Total[5].TotCredit, Devise.Decimale);
          end
          else if (Rev.State = cbChecked) then
          begin
            Debit := Arrondi(Total[5].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[5].TotCredit, Devise.Decimale);
          end;
        end;
      (*TOU*)4:
        begin
          if (Rev.State = cbUnchecked) then
          begin
            Debit := Arrondi(Total[1].TotDebit + Total[2].TotDebit + Total[3].TotDebit + Total[4].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[1].TotCredit + Total[2].TotCredit + Total[3].TotCredit + Total[4].TotCredit, Devise.Decimale);
          end
          else if (Rev.State = cbGrayed) then
          begin
            Debit := Arrondi(Total[1].TotDebit + Total[2].TotDebit + Total[3].TotDebit + Total[4].TotDebit + Total[5].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[1].TotCredit + Total[2].TotCredit + Total[3].TotCredit + Total[4].TotCredit + Total[5].TotCredit, Devise.Decimale);
          end
          else if (Rev.State = cbChecked) then
          begin
            Debit := Arrondi(Total[5].TotDebit, Devise.Decimale);
            Credit := Arrondi(Total[5].TotCredit, Devise.Decimale);
          end;
        end;
    end;

    SDebit := SDebit + Debit;
    SCredit := SCredit + Credit;
    CBDebDat.Items.Add(DatetoStr(DebDate));
    CBFinDate.Items.Add(DatetoStr(FinDate));
    StM := FormatDateTime('mmmm yyyy', DebDate);
    StM := FirstMajuscule(StM);
    FListe.Cells[0, j + 1] := StM;
    Fliste.Cells[1, j + 1] := StrFMontant(Debit, 15, Devise.Decimale, Devise.Symbole, TRUE);
    Fliste.Cells[2, j + 1] := StrFMontant(Credit, 15, Devise.Decimale, Devise.Symbole, TRUE);
    AfficheLeSolde(NumEditSold, Debit, Credit);
    FListe.Cells[3, j + 1] := NumEditSold.Text;
    Fliste.RowCount := Fliste.RowCount + 1;
    // --> Premier jour du mois suivant - GCO - 03/2004
    DebDate := DebutDeMois(PlusMois(DebDate, 1));
    // DebDate := PlusMois(DebDate, 1);
    MoveCur(False);
  end;

  FListe.Cells[0, NbrMois + 2] := LTotaux.Caption;
  SDebit := SDebit + Valeur(FListe.Cells[1, 1]);
  SCredit := SCredit + Valeur(FListe.Cells[2, 1]);
  Fliste.Cells[1, NbrMois + 2] := StrFMontant(SDebit, 15, Devise.Decimale, Devise.Symbole, TRUE);
  Fliste.Cells[2, NbrMois + 2] := StrFMontant(SCredit, 15, Devise.Decimale, Devise.Symbole, TRUE);
  AfficheLeSolde(NumEditSold, SDebit, SCredit);
  FListe.Cells[3, NbrMois + 2] := NumEditSold.Text;
  CbSoldCumClick(nil);

  if (Q <> nil) then Ferme(Q);
  FiniMove;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.AffichePeriode(NbrMois: Integer);
var
  i: Byte;
  DebDate, FinDate: TDateTime;
  StM: string;
begin
  DebDate := ExerciceEntree.Deb;
  Fliste.RowCount := 3;
  CBDebDat.Items.Clear;
  CBFinDate.Items.Clear;
  CBDebDat.Items.Add('');
  CBFinDate.Items.Add('');
  CBDebDat.Items.Add('');
  CBFinDate.Items.Add('');
  FListe.Cells[0, 1] := LAnouv.Caption;

  for i := 1 to NbrMois do
  begin
    FinDate := FindeMois(DebDate);
    CBDebDat.Items.Add(DatetoStr(DebDate));
    CBFinDate.Items.Add(DatetoStr(FinDate));
    StM := FormatDateTime('mmmm yyyy', DebDate);
    StM[1] := upcase(StM[1]);
    FListe.Cells[0, i + 1] := StM;
    Fliste.RowCount := Fliste.RowCount + 1;
    // --> Premier jour du mois suivant - GCO - 03/2004
    DebDate := DebutDeMois(PlusMois(DebDate, 1));
    // DebDate := PlusMois(DebDate, 1);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{$IFNDEF EAGLCLIENT}
procedure TFCumMens.CreateViewResult;
var
  St, Pre, NomEcr, NomTable, NomChamp, NomCompte: string;
begin
  FAxe := '';

  case LeTypeFic of
    fbGene:
      begin
        NomCompte := 'E_GENERAL';
        NomEcr := 'ECRITURE';
        Pre := 'E';
        NomTable := 'MOISECRGEN';
        NomChamp := 'E_MOIS';
      end;
    fbAux:
      begin
        NomCompte := 'E_AUXILIAIRE';
        NomEcr := 'ECRITURE';
        Pre := 'E';
        NomTable := 'MOISECRAUX';
        NomChamp := 'E_MOIS';
      end;
    fbJal:
      begin
        FAxe := JAlODA(QuelCode);
        if (FAxe = '') then
        begin
          NomCompte := 'E_JOURNAL';
          NomEcr := 'ECRITURE';
          Pre := 'E';
          NomTable := 'MOISECRJAL';
          NomChamp := 'E_MOIS';
        end
        else
        begin
          NomCompte := 'Y_JOURNAL,Y_AXE';
          NomEcr := 'ANALYTIQ';
          Pre := 'Y';
          NomTable := 'MOISECRJALODA';
          NomChamp := 'Y_MOIS';
        end;
      end;
    fbAxe1..fbAxe5:
      begin
        FAxe := fbToAxe(LeTypeFic);
        NomCompte := 'Y_SECTION,Y_AXE';
        NomEcr := 'ANALYTIQ';
        Pre := 'Y';
        NomTable := 'MOISECRSEC';
        NomChamp := 'Y_MOIS';
      end;
  end;

  if TableExiste(NomTable) then
  begin
    if _ChampExiste(NomTable, NomChamp) then exit
    else
    begin
      BeginTrans;
      ExecuteSQL('DROP VIEW ' + NomTable);
      CommitTrans;
    end;
  end;

  St := 'CREATE VIEW ' + NomTable + ' ' + '(' + NomCompte + ',' + Pre + '_MOIS,' + Pre + '_DEBIT,' + Pre + '_CREDIT,' + Pre + '_DEBITDEV,' + Pre + '_CREDITDEV,' + Pre +
    '_DATECOMPTABLE,' + Pre + '_EXERCICE,' + Pre + '_ECRANOUVEAU,' + Pre + '_QUALIFPIECE,' + Pre + '_DEVISE,' + Pre + '_ETABLISSEMENT) ' + 'AS SELECT ' + NomCompte + ',' +
    DB_Month(Pre + '_DATECOMPTABLE') + ',' + Pre + '_DEBIT,' + Pre + '_CREDIT,' + Pre + '_DEBITDEV,' + Pre + '_CREDITDEV,' + Pre + '_DATECOMPTABLE,' + Pre + '_EXERCICE,' + Pre +
    '_ECRANOUVEAU,' + Pre + '_QUALIFPIECE,' + Pre + '_DEVISE,' + Pre + '_ETABLISSEMENT ' + 'FROM ' + NomEcr;
  BeginTrans;
  ExecuteSQL(St);
  CommitTrans;
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCumMens.CumulMensuel_UDF;
var
  Q: TQuery;
  j: Integer;
  Debit, Credit, SDebit, SCredit, DebANew, CredANew: Extended;
  PremMois, PremAnnee, NbrMois: Word;
  Pre, NomTable, NomCompte: string;
  StAxe, StSDev, StQul, StEtb, StDev: string;
begin
  NombrePerExo(ExerciceEntree, PremMois, PremAnnee, NbrMois);
  AffichePeriode(NbrMois);

  SDebit := 0;
  SCredit := 0;
  FListe.ColAligns[0] := taCenter;
  DebANew := 0;
  CredANew := 0;
  AfficheLeSolde(NumEditSold, 0, 0);

  for j := 2 to NbrMois + 1 do
  begin
    FListe.Cells[1, j] := NumEditSold.Text;
    FListe.Cells[2, j] := NumEditSold.Text;
    FListe.Cells[3, j] := NumEditSold.Text;
  end;

  case LeTypeFic of
    fbGene:
      begin
        Pre := 'E';
        NomTable := 'MOISECRGEN';
        NomCompte := 'E_GENERAL';
      end;
    fbAux:
      begin
        Pre := 'E';
        NomTable := 'MOISECRAUX';
        NomCompte := 'E_AUXILIAIRE';
      end;
    fbJal:
      begin
        if (FAxe = '') then
        begin
          Pre := 'E';
          NomTable := 'MOISECRJAL';
          NomCompte := 'E_JOURNAL';
        end
        else
        begin
          Pre := 'Y';
          NomTable := 'MOISECRJALODA';
          NomCompte := 'Y_JOURNAL';
        end;
      end;
    fbAxe1..fbAxe5:
      begin
        Pre := 'Y';
        NomTable := 'MOISECRSEC';
        NomCompte := 'Y_SECTION';
      end;
  end;

  if (FAxe <> '') then StAxe := ' AND Y_AXE="' + FAxe + '"'
  else StAxe := '';

  if (FiltreDev) then
  begin
    StSDev := 'SUM(' + Pre + '_DEBITDEV),SUM(' + Pre + '_CREDITDEV)';
    StDev := 'AND ' + Pre + '_DEVISE="' + Devise.Code + '" '
  end
  else
  begin
    StSDev := 'SUM(' + Pre + '_DEBIT),SUM(' + Pre + '_CREDIT)';
    StDev := '';
  end;

  if (FiltreEtab) then StEtb := 'AND ' + Pre + '_ETABLISSEMENT="' + CodEtb + '" '
  else StEtb := '';

  StQul := WhereSupp(Pre + '_', WhatTypeEcr(CbTypMvt.Value, V_PGI.Controleur, Rev.State));

  Q := OpenSQL('SELECT ' + NomCompte + ',' + Pre + '_MOIS,' + StSDev + ',' + Pre + '_ECRANOUVEAU ' + 'FROM ' + NomTable + ' ' + 'WHERE ' + NomCompte + '="' + QuelCode + '" AND ' +
    Pre + '_EXERCICE="' + ExerciceEntree.Code + '" ' + StAxe + StDev + StEtb + StQul + 'GROUP BY ' + NomCompte + ',' + Pre + '_MOIS' + ',' + Pre + '_ECRANOUVEAU', true);

  InitMove(NbrMois, '');

  while (not Q.EOF) do
  begin
    if (Q.Fields[4].AsString <> 'N') then
    begin
      DebANew := DebANew + Q.Fields[2].AsFloat;
      CredANew := CredANew + Q.Fields[3].AsFloat;
    end
    else
    begin
      j := Q.Fields[1].AsInteger;
      Debit := Q.Fields[2].AsFloat;
      Credit := Q.Fields[3].AsFloat;
      SDebit := SDebit + Debit;
      SCredit := SCredit + Credit;
      Fliste.Cells[1, j + 1] := StrFMontant(Debit, 15, Devise.Decimale, Devise.Symbole, TRUE);
      Fliste.Cells[2, j + 1] := StrFMontant(Credit, 15, Devise.Decimale, Devise.Symbole, TRUE);
      AfficheLeSolde(NumEditSold, Debit, Credit);
      FListe.Cells[3, j + 1] := NumEditSold.Text;
      MoveCur(False);
    end;
    Q.Next;
  end;

  FListe.Cells[0, NbrMois + 2] := LTotaux.Caption;
  AfficheLeSolde(NumEditSold, DebANew, CredAnew);
  FListe.Cells[3, 1] := NumEditSold.Text;

  if (LeTypeFic <> fbJal) then
  begin
    if (DebANew - CredANew <= 0) then
    begin
      DebAnew := 0;
      FListe.Cells[2, 1] := StrFMontant(Abs(Valeur(FListe.Cells[3, 1])), 15, Devise.Decimale, Devise.Symbole, true);
      FListe.Cells[1, 1] := StrFMontant(0, 15, Devise.Decimale, Devise.Symbole, true);
    end
    else
    begin
      CredAnew := 0;
      FListe.Cells[1, 1] := StrFMontant(Abs(Valeur(FListe.Cells[3, 1])), 15, Devise.Decimale, Devise.Symbole, true);
      FListe.Cells[2, 1] := StrFMontant(0, 15, Devise.Decimale, Devise.Symbole, true);
    end;
  end
  else
  begin
    FListe.Cells[1, 1] := StrFMontant(DebANew, 15, Devise.Decimale, Devise.Symbole, TRUE);
    FListe.Cells[2, 1] := StrFMontant(CredANew, 15, Devise.Decimale, Devise.Symbole, TRUE);
  end;

  SDebit := SDebit + DebANew;
  SCredit := SCredit + CredANew;
  Fliste.Cells[1, NbrMois + 2] := StrFMontant(SDebit, 15, Devise.Decimale, Devise.Symbole, TRUE);
  Fliste.Cells[2, NbrMois + 2] := StrFMontant(SCredit, 15, Devise.Decimale, Devise.Symbole, TRUE);
  AfficheLeSolde(NumEditSold, SDebit, SCredit);
  FListe.Cells[3, NbrMois + 2] := NumEditSold.Text;
  CbSoldCumClick(nil);
  Ferme(Q);
  FiniMove;
end;

end.

