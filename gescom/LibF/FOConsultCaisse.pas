{***********UNITE*************************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 04/03/2004
Description .. : Affichage de la Situation Flash de la caisse de travail
Mots clefs ... : FO;SITUATION;FLASH
*****************************************************************}
unit FOConsultCaisse;

interface

uses
  Forms, hmsgbox, HSysMenu, Grids, Hctrls, TeEngine, Series, TeeProcs,
  Chart, HPanel, ExtCtrls, StdCtrls, Controls, HTB97, Classes, Graphics,
  Windows, SysUtils,
{$IFDEF EAGLCLIENT}
  Maineagl,
{$ELSE}
  dbtables, FE_Main,
{$ENDIF}
  HEnt1, UIUtil, ParamSoc, hstatus, UTOB
{$IFDEF VOIRTOB}
  ,VoirTob
{$ENDIF VOIRTOB}
  ;

procedure FOConsultationCaisse(Caisse, DeNumeroZ, ANumeroZ: string; Inside: THPanel; MultiCais: Boolean);

type
  TFConsultCaisse = class(TForm)
    HmTrad: THSystemMenu;
    HM: THMsgBox;
    DockBottom: TDock97;
    Panel1: TPanel;
    PG1: THPanel;
    GRV: TChart;
    Series1: TPieSeries;
    PG2: THPanel;
    GRF: TChart;
    PieSeries1: TPieSeries;
    GSS: THGrid;
    GSV: THGrid;
    FNomCaisse: THLabel;
    Ouvertele: THLabel;
    GSR: THGrid;
    Bevel1: TBevel;
    ToolWindow971: TToolWindow97;
    BImprimer: TToolbarButton97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    Famille: THValComboBox;
    BParametrer: TToolbarButton97;
    NUMCAISSE: THValComboBox;
    LNumCaisse: THLabel;
    procedure OnChangeFamille(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure GSSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure BAbandonClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BParametrerClick(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure NUMCAISSEChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    TobRDAnt : TOB; // JTR
    TotGalVente : double; //JTR

    procedure ChargeArguments(Caisse, DeNumeroZ, ANumeroZ: string; MultiCais: Boolean);
    procedure ParamVisu;
    procedure ChargeParamCaisse;
    function FabricWhereMdpEcart(NotIn: Boolean): string;
    function FabricWhereFDCaisse(NotIn: Boolean): string;
    function FabricWhereCaisse(Prefixe: string): string;
    function GereVendeur: Boolean;
    procedure ChargeValeurs;
    procedure ChargeReglements;
    procedure AfficheReglements;
    procedure ChargeVentes;
    procedure AfficheVentes;
    function CalculStatistiques(TOBStatistiques: TOB; stType, stChamp: string; i: integer): integer;
    procedure GetCellCanvasGSV(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure PostDrawCellGSR(Acol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
    procedure GetCellCanvasGSR(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GetCellCanvasGSS(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure DessineGraphs(ACol: integer);
    function SymboleDevise(Devise: string): string;
    procedure AppliquerParametrage;
    procedure ParamListeCaisse;
    procedure TabGridVendeurs(Grid: THGrid);
    procedure TabGridCateg(Grid: THGrid);
    procedure GridGriseCellPV(GS: THGrid; Acol, ARow: integer; Canvas: TCanvas);
    procedure Progression;
    procedure ChangeTitreCaisse;
    procedure ChangeTitle(TC: TChart; LibFamille, Titre: string);
    procedure MajTitreGSS;
  public
    CodeCaisse, sMdpEcart, OldFamilleNiv, FamilleNiv, sOldCategorie: string;
    LigneDebutVendeurs, LigneFinVendeurs, LigneDebutFamilles, LigneFinFamilles: Integer;
    sOldVendeurs, sOldStats, sOldTypeCateg, sOldReglts: Integer;
    sOldCaisseBtq: boolean;
    DeNumZ, ANumZ, Jauge: Integer;
    TOBDev: TOB; // TOB des devises
    TOBPCaisse: TOB; // TOB des caisses
    TOBReg: TOB; // TOB de la grille des règlements
    TOBVen: TOB; // TOB de la grille des ventes
    DetailOpFi: boolean;
    DetailPreDiv: boolean;
    GereFDCaisse: Boolean;
    MultiCaisse: Boolean;
    FdCaisOuv: boolean;
  end;

implementation

uses
  FactUtil, FOUtil, FODefi, Ent1;

{$R *.DFM}

const
  GSV_C_CODE = 0;
  GSV_C_LIBELLE = 1;
  GSV_C_NOMBRE = 2;
  GSV_C_MONTANT = 3;

  GSR_TYPE = 0;
  GSR_CODE = 1;
  GSR_LIBELLE = 2;
  GSR_MTTENCAIS = 3;
  GSR_MTTDEVISE = 4;
  GSR_FONDS = 5;

  GSS_TYPE = 0;
  GSS_CODE = 1;
  GSS_LIBELLE = 2;
  GSS_CATTC = 3;
  GSS_NBART = 4;
  GSS_NBVTC = 5;
  GSS_REMISE = 6;
  GSS_PANIERTTC = 7;
  GSS_PANIERQTE = 8;

  Espacement = 4;
  ColorTotal = TColor($00B18663);
  ColorCumul = TColor($00C9AC94);
  ColorLigne = TColor($00E1D2C6);


  {***********A.G.L.Privé.*****************************************
  Auteur  ...... : O. TARCY
  Créé le ...... : 05/06/2001
  Modifié le ... : 23/07/2001
  Description .. : Affichage de la Situation Flash de la caisse de travail
  Mots clefs ... : FO
  *****************************************************************}

procedure FOConsultationCaisse(Caisse, DeNumeroZ, ANumeroZ: string; Inside: THPanel; MultiCais: Boolean);
var x: TFConsultCaisse;
begin
  if trim(Caisse) = '' then // la caisse n'est pas paramétrée
  begin
    PGIBox('Vouz devez paramétrer la caisse.', 'Situation Flash');
    Exit;
  end;
  SourisSablier;
  X := TFConsultCaisse.Create(Application);
  X.ChargeArguments(Caisse, DeNumeroZ, ANumeroZ, MultiCais);
  if Inside = nil then
  begin
    try
      X.showmodal;
    finally
      X.Free;
    end;
  end else
  begin
    InitInside(X, Inside);
    x.Show;
  end;
  SourisNormale;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 17/05/2002
Description .. : Lit une valeur d'une clé de registre
Mots clefs ... : FO
*****************************************************************}

function LitValeurRegistre(NomCle: string; Valeur: Variant): Variant;
begin
  Result := FOGetFromRegistry(REGSITUFLASH, NomCle, Valeur);
  if (NomCle = REGCAISSEBTQ) and (EtabForce <> '') then
    Result := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 23/07/2001
Description .. : Initialisation en fonction des arguments d'appel
Mots clefs ... : FO
*****************************************************************}

procedure TFConsultCaisse.ChargeArguments(Caisse, DeNumeroZ, ANumeroZ: string; MultiCais: Boolean);
var NivFam: string;
begin
  CodeCaisse := Caisse;
  NumCaisse.Value := CodeCaisse;
  MultiCaisse := MultiCais;
  if not MultiCaisse then
  begin
    DeNumZ := StrToInt(DeNumeroZ);
    ANumZ := StrToInt(ANumeroZ);
  end else
  begin
    DeNumZ := 0;
    ANumZ := 0;
  end;
  // niveau de famille à traiter
  NivFam := LitValeurRegistre(REGCATEGORIEPARDEFAUT, 'LF1');
  if ((not GetParamSoc('SO_ARTLOOKORLI')) and (StrToInt(copy(NivFam, 3, 1)) > 3)) or (RechDom('GCLIBFAMILLE', NivFam, False) = '.-') then
    NivFam := 'LF1';
  FamilleNiv := NivFam;
  OldFamilleNiv := NivFam;
  Famille.Value := FamilleNiv;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/11/2001
Modifié le ... : 23/11/2001
Description .. : Récupération du symbole de la devise
Mots clefs ... : FO;
*****************************************************************}

function TFConsultCaisse.SymboleDevise(Devise: string): string;
var TOBL: TOB;
begin
  TOBL := TOBDev.FindFirst(['D_DEVISE'], [Devise], False);
  if TOBL = nil then
  begin
    TOBL := TOB.Create('DEVISE', TOBDev, -1);
    FOChargeDevise(TOBL, Devise);
  end;
  if TOBL <> nil then result := TOBL.GetValue('D_SYMBOLE') else result := '';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 27/09/2001
Modifié le ... : 17/05/2002
Description .. : Changement de caisse
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.NUMCAISSEChange(Sender: TObject);
var Stg: string;
  Ind: Integer;
begin
  Stg := NumCaisse.Value;
  if NumCaisse.Value = NumCaisse.VideString then
  begin
    Stg := '';
    for Ind := 0 to NumCaisse.Values.Count - 1 do
    begin
      if Ind > 0 then Stg := Stg + ';' + NumCaisse.Values[Ind]
      else Stg := Stg + NumCaisse.Values[Ind];
    end;
  end;
  if Stg <> CodeCaisse then
  begin
    CodeCaisse := Stg;
    Jauge := 0;
    ParamVisu;
    ChangeTitreCaisse;
    ChargeValeurs;
    DessineGraphs(GSS.Col);
    AppliquerParametrage;
    //GSV.Setfocus;
    ChangeTitle(GRV, '', HM.Mess[12]);
    ChangeTitle(GRF, FamilleNiv, '');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 27/09/2001
Modifié le ... : 27/09/2001
Description .. : Changement du paramètre catégorie
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.OnChangeFamille(Sender: TObject);
var Q: TQuery;
  Stg: string;
  ifin, isuppr, irow: integer;
begin
  FamilleNiv := Famille.Value;
  if (OldFamilleNiv <> FamilleNiv) then
  begin
    GSS.SynEnabled := False;
    // Affichage de toute la catégorie
    for irow := LigneDebutFamilles to LigneFinFamilles do
      if GSS.RowHeights[irow] = -1 then GSS.RowHeights[irow] := GSS.DefaultRowHeight;

    ChangeTitle(GRF, FamilleNiv, '');
    Stg := TraduireMemoire(RechDom('GCLIBFAMILLE', FamilleNiv, False));
    irow := LigneDebutFamilles;
    Q := OpenSQL('select CC_TYPE, CC_CODE, CC_LIBELLE from CHOIXCOD where CC_TYPE="FN' + copy(FamilleNiv, 3, 1) + '"', True);
    while not Q.EOF do
    begin
      // Ajout d'une ligne si fin de tableau
      if irow > LigneFinFamilles then
        GSS.InsertRow(irow);

      GSS.Cells[GSS_TYPE, irow] := Q.FindField('CC_TYPE').AsString;
      GSS.Cells[GSS_CODE, irow] := Q.FindField('CC_CODE').AsString;
      GSS.Cells[GSS_LIBELLE, irow] := StringOfChar(' ', 2*Espacement) + Q.FindField('CC_LIBELLE').AsString;

      // Incrémentation des limites de tableau
      if (LigneDebutFamilles > LigneDebutVendeurs) and (irow > LigneFinFamilles) then
        Inc(LigneFinFamilles) else
        if (LigneDebutFamilles < LigneDebutVendeurs) and (irow > LigneFinFamilles) then
      begin
        Inc(LigneFinFamilles);
        Inc(LigneDebutVendeurs);
        Inc(LigneFinVendeurs);
      end;

      Inc(irow);
      Q.Next;
    end;
    Ferme(Q);

    // Ajout d'une ligne en fin de tableau pour les articles sans famille
    if irow > LigneFinFamilles then
      GSS.InsertRow(irow);
    GSS.Cells[GSS_TYPE, irow] := 'FN' + copy(FamilleNiv, 3, 1);
    GSS.Cells[GSS_CODE, irow] := '';
    GSS.Cells[GSS_LIBELLE, irow] := StringOfChar(' ', 2*Espacement) + TraduireMemoire('Sans ') + Stg;
    if (LigneDebutFamilles > LigneDebutVendeurs) and (irow > LigneFinFamilles) then
      Inc(LigneFinFamilles) else
      if (LigneDebutFamilles < LigneDebutVendeurs) and (irow > LigneFinFamilles) then
    begin
      Inc(LigneFinFamilles);
      Inc(LigneDebutVendeurs);
      Inc(LigneFinVendeurs);
    end;
    Inc(irow);

    // Suppression éventuelle des lignes en trop
    ifin := LigneFinFamilles;
    for isuppr := irow to ifin do
    begin
      GSS.DeleteRow(irow);
      if LigneDebutFamilles > LigneDebutVendeurs then
        Dec(LigneFinFamilles) else
      begin
        Dec(LigneFinFamilles);
        Dec(LigneDebutVendeurs);
        Dec(LigneFinVendeurs);
      end;
    end;
    MajTitreGSS;
    GSS.Cells[GSS_LIBELLE, LigneDebutFamilles - 1] := StringOfChar(' ', Espacement) + RechDom('GCLIBFAMILLE', FamilleNiv, False);

    ChargeValeurs;
    DessineGraphs(GSS.Col);
    OldFamilleNiv := FamilleNiv;
    GSS.SynEnabled := True;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 05/10/2001
Modifié le ... : 05/10/2001
Description .. : Change le titre du tableau des statistiques
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.ChangeTitle(TC: TChart; LibFamille, Titre: string);
var Stg: string;
begin
  Stg := HM.Mess[4] + ' ';
  if LibFamille = '' then Stg := Stg + TraduireMemoire(Titre)
  else Stg := Stg + RechDom('GCLIBFAMILLE', LibFamille, False);
  with TC.Title.Text do
  begin
    Clear;
    Add(Stg);
  end;
  TC.Repaint;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 28/09/2001
Modifié le ... : 28/09/2001
Description .. : Mise à jour de l'affichage du titre du tableau des statistiques
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.MajTitreGSS;
var
  sCrit2, sCrit1, sTitre: string;
begin
  sCrit2 := RechDom('GCLIBFAMILLE', FamilleNiv, False);
  sCrit1 := trim(GSS.Cells[GSS_LIBELLE, LigneDebutVendeurs - 1]);
  sTitre := HM.Mess[4] + ' ' + sCrit1 + ' ' + HM.Mess[5] + ' ' + sCrit2;
  GSS.Cells[GSS_LIBELLE, 0] := sTitre;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 05/10/2001
Modifié le ... : 17/05/2002
Description .. : Change le titre de la fiche avec le code de la caisse et la
Suite ........ : date d'ouverture
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.ChangeTitreCaisse;
var TOBL: TOB;
  Ind: Integer;
  Separateur: string;
begin
  BImprimer.Enabled := False;
  if (TOBPCaisse.Detail.Count = 1) and (Pos(';', CodeCaisse) = 0) then
  begin
    TOBL := TOBPCaisse.Detail[0];
    //   FNomCaisse.Caption := RechDom('GCPCAISSE', CodeCaisse, FALSE) ;
    FNomCaisse.Caption := RechDom('GCPCAISSE', TOBL.GetValue('NOCAIS'), FALSE);
    if DeNumZ = ANumZ then
    begin
      // Clôture n°xx du xx/xx/xxxx
      Ouvertele.Caption := HM.Mess[0] + ' ' + HM.Mess[1] + IntToStr(TOBL.GetValue('NUMZ')) + ' '
        + HM.Mess[2] + ' ' + DateToStr(TOBL.GetValue('DATEOUV'));
      {$IFDEF FOS5}
      BImprimer.Enabled := True;
      {$ENDIF}
    end else
    begin
      // Clôtures n°xx du xx/xx/xxxx - n°xx du xx/xx/xxxx
      Ouvertele.Caption := HM.Mess[13] + ' ' + HM.Mess[1] + IntToStr(TOBL.GetValue('NUMZ')) + ' '
        + HM.Mess[2] + ' ' + DateToStr(TOBL.GetValue('DATEOUV')) + ' - '
        + HM.Mess[1] + IntToStr(TOBL.GetValue('NUMZFIN')) + ' '
        + HM.Mess[2] + ' ' + DateToStr(TOBL.GetValue('DATEFIN'));
    end;
  end else
  begin
    Ouvertele.Caption := '';
    // Caisses: xxx (n°xx du xx/xx/xxxx), xxx (n°xx du xx/xx/xxxx)
    FNomCaisse.Caption := HM.Mess[3];
    for Ind := 0 to TOBPCaisse.Detail.Count - 1 do
    begin
      TOBL := TOBPCaisse.Detail[Ind];
      if Ind > 0 then Separateur := ', ' else Separateur := ': ';
      FNomCaisse.Caption := FNomCaisse.Caption + Separateur + TOBL.GetValue('NOCAIS')
        + ' (' + HM.Mess[1] + IntToStr(TOBL.GetValue('NUMZ'))
        + ' ' + HM.Mess[2] + ' ' + DateToStr(TOBL.GetValue('DATEOUV')) + ')';
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 05/10/2001
Modifié le ... : 17/05/2002
Description .. : Affichage de la fiche
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.FormShow(Sender: TObject);
var Stg: string;
begin
  GSV.Ctl3D := False;
  GSR.Ctl3D := False;
  GSS.Ctl3D := False;

  sOldCaisseBtq := LitValeurRegistre(REGCAISSEBTQ, True);
  ParamListeCaisse;
  Stg := CodeCaisse;
  CodeCaisse := '';
  if Pos(';', Stg) > 0 then
    NumCaisse.Value := NumCaisse.VideString
  else
    NumCaisse.Value := Stg;
  // Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 05/10/2001
Modifié le ... : 17/05/2002
Description .. : Fermeture de la fiche
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if IsInside(Self) then Action := caFree;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 05/10/2001
Modifié le ... : 17/05/2002
Description .. : Création de la fiche
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.FormCreate(Sender: TObject);
var TOBL: TOB;
begin
  FdCaisOuv := False;
  DetailOpFi := True;
  DetailPreDiv := True;
  sOldCategorie := '';
  sOldStats := -1;
  sOldVendeurs := -1;
  sOldTypeCateg := -1;
  sOldReglts := -1;
  sOldCaisseBtq := True;
  // TOB des devises
  TOBDev := TOB.Create('Les devises', nil, -1);
  TOBL := TOB.Create('DEVISE', TOBDev, -1);
  FOChargeDevise(TOBL, V_PGI.DevisePivot);
  // TOB des caisses
  TOBPCaisse := TOB.Create('Les caisses', nil, -1);
  // TOB des règlements
  TOBReg := TOB.Create('Les règlements', nil, -1);
  // TOB des ventes
  TOBVen := TOB.Create('Les ventes', nil, -1);
  {$IFDEF GESCOM}
  if not (FOExisteClavierEcran) then
  begin
    Align := alNone;
    BorderIcons := [biSystemMenu, biMaximize];
    BorderStyle := bsSizeToolWin;
  end;
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 05/10/2001
Modifié le ... : 17/05/2002
Description .. : Destruction de la fiche
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.FormDestroy(Sender: TObject);
begin
  if TOBDev <> nil then FreeAndNil(TOBDev);
  if TOBPCaisse <> nil then FreeAndNil(TOBPCaisse);
  if TOBReg <> nil then FreeAndNil(TOBReg);
  if TOBVen <> nil then FreeAndNil(TOBVen);
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 21/11/2001
Modifié le ... : 23/11/2001
Description .. : Formatage des grids
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.ParamVisu;
var
  FmtV, FmtN: string;
begin
  FmtV := StrfMask(V_PGI.OkDecV, '', TRUE);
  FmtN := StrfMask(0, '', TRUE);
  ChargeParamCaisse;

  // Grid ventes
  GSV.VidePile(TRUE);
  GSV.HideSelectedWhenInactive := True;
  GSV.ColWidths[GSV_C_CODE] := -1;
  GSV.ColAligns[GSV_C_NOMBRE] := taRightJustify;
  GSV.ColFormats[GSV_C_NOMBRE] := FmtN;
  GSV.ColAligns[GSV_C_MONTANT] := taRightJustify;
  GSV.ColFormats[GSV_C_MONTANT] := FmtV;
  GSV.GetCellCanvas := GetCellCanvasGSV;

  // Grid règlements
  GSR.VidePile(TRUE);
  GSR.HideSelectedWhenInactive := True;
  GSR.ColWidths[GSR_TYPE] := -1;
  GSR.ColWidths[GSR_CODE] := -1;
  GSR.ColWidths[GSR_MTTDEVISE] := -1;
//  GSR.ColAligns[GSR_MTTDEVISE] := taRightJustify;
//  GSR.ColFormats[GSR_MTTDEVISE] := FMTV;
  GSR.ColAligns[GSR_MTTENCAIS] := taRightJustify;
  GSR.ColFormats[GSR_MTTENCAIS] := FMTV;
  GSR.ColAligns[GSR_FONDS] := taRightJustify;
  GSR.ColFormats[GSR_FONDS] := FMTV;
  GSR.PostDrawCell := PostDrawCellGSR;
  GSR.GetCellCanvas := GetCellCanvasGSR;

  // Grid statistiques
  GSS.VidePile(TRUE);
  GSS.HideSelectedWhenInactive := True;
  GSS.ColWidths[GSS_TYPE] := -1;
  GSS.ColWidths[GSS_CODE] := -1;
  GSS.ColAligns[GSS_NBART] := taRightJustify;
  GSS.ColFormats[GSS_NBART] := FmtN;
  GSS.ColAligns[GSS_CATTC] := taRightJustify;
  GSS.ColFormats[GSS_CATTC] := FmtV;
  GSS.ColAligns[GSS_NBVTC] := taRightJustify;
  GSS.ColFormats[GSS_NBVTC] := FmtN;
  GSS.ColAligns[GSS_REMISE] := taRightJustify;
  GSS.ColFormats[GSS_REMISE] := FmtV;
  GSS.ColAligns[GSS_PANIERQTE] := taRightJustify;
  GSS.ColFormats[GSS_PANIERQTE] := FmtV;
  GSS.ColAligns[GSS_PANIERTTC] := taRightJustify;
  GSS.ColFormats[GSS_PANIERTTC] := FmtV;
  GSS.GetCellCanvas := GetCellCanvasGSS;
  GSS.Cells[GSS_TYPE, GSS.RowCount - 1] := 'TOTAL';
  GSS.Cells[GSS_CODE, GSS.RowCount - 1] := 'TOTAL';
  GSS.Cells[GSS_LIBELLE, GSS.RowCount - 1] := HM.Mess[8]; // Total

  // Constitution de l'ordre du tableau d'après le paramètre sauvegardé
  sOldStats := LitValeurRegistre(REGSTATSPARDEFAUT, STATSINI);
  if sOldStats = STATSINI then
  begin
    TabGridVendeurs(GSS);
    TabGridCateg(GSS);
  end else
  begin
    TabGridCateg(GSS);
    TabGridVendeurs(GSS);
  end;

  TobRDAnt := TOB.Create('',nil,-1); //Contient les Reste dû antérieurs (utilisé dans ventes et rgt)
  ChargeVentes;
  if not GereFDCaisse then GSR.ColWidths[GSR_FONDS] := -1;
  ChargeReglements;

  AfficheVentes;
  AfficheReglements;
  if TobRDAnt <> nil then FreeAndNil(TobRDAnt);

  MajTitreGSS;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N ACHINO
Créé le ...... : 26/02/2004
Modifié le ... : 26/02/2004
Description .. : Chargement des règlements
Mots clefs ... :
*****************************************************************}

procedure TFConsultCaisse.ChargeReglements;
var
  QQ: TQuery;
  sSql: string;
  Ind: Integer;
  TOBMdp, TOBR, TOBL: TOB;
  sCodeMdp: string;
  TotGen, CumEncais, CumAutre, Mnt: double;
  // JTR Reste dû réglé
  CumEncaisCredit, CumEncaisCreditAnt : double;
  TobTmp, TobRDAnterieur : TOB;
  Lib, NumPieceLien : string;
  DateFin, DatePCe : TDateTime;
  // Fin JTR Reste dû réglé
begin
  TotGen := 0;
  CumEncais := 0;
  CumAutre := 0;
  CumEncaisCredit := 0;
  if TOBReg.Detail.Count > 0 then TOBReg.ClearDetail;
  TobRDAnterieur := nil;

  // Gestion du fond de caisse
  if not GereFDCaisse then GSR.ColWidths[GSR_FONDS] := -1;

  // Recherche des modes de paiements
  sSql := 'SELECT MP_TYPEMODEPAIE,MP_MODEPAIE,MP_DEVISEFO,MP_LIBELLE,CO_LIBRE FROM MODEPAIE'
    + ' LEFT JOIN COMMUN ON CO_TYPE="TMP" AND CO_CODE=MP_TYPEMODEPAIE'
    + ' WHERE (MP_UTILFO="X" OR ' + FabricWhereMdpEcart(False) + ')'
    + ' ORDER BY CO_LIBRE,MP_TYPEMODEPAIE,MP_LIBELLE';

  TOBReg.LoadDetailFromSQL( sSql);
  if TOBReg.Detail.Count > 0 then
  begin
    TOBR := TOBReg.Detail[0];
    TOBR.AddChampSupValeur('MONTANTENCAIS', 0, True);
    TOBR.AddChampSupValeur('MONTANTECHE', 0, True);
    TOBR.AddChampSupValeur('ANTERIEUR', '-', True);
    if GereFDCaisse then TOBR.AddChampSupValeur('FONDCAISSE', 0, True);
    // JTR Reste dû réglé 
    TobTmp := TOBReg.FindFirst(['MP_TYPEMODEPAIE'],['007'],True);
    if TobTmp <> nil  then
    begin
      TobTmp.PutValue('CO_LIBRE','3');
      Lib := TobTmp.GetValue('MP_LIBELLE');
      TobTmp.PutValue('MP_LIBELLE', TobTmp.GetValue('MP_LIBELLE')+TraduireMemoire(' depuis ouverture'));
      // Création d'une fille qui contiendra reste dû antérieurs
      TobRDAnterieur := TOB.Create('',TOBReg,-1);
      TobRDAnterieur.Dupliquer(TobTmp,True,True);
      TobRDAnterieur.PutValue('CO_LIBRE','4');
      TobRDAnterieur.PutValue('MP_LIBELLE', Lib+TraduireMemoire(' antérieur'));
      TobRDAnterieur.PutValue('ANTERIEUR', 'X');
    end;
  end;

  if GereFDCaisse then
  begin
    if FdCaisOuv then
      // fond de caisse à l'ouverture
      sSql := 'SELECT SUM(GJM_FDCAISSEDEV) as FDCAISSEDEV,GJM_MODEPAIE AS MODEPAIE FROM CTRLCAISMT ' +
        'WHERE GJM_FDCAISSEDEV>0 AND ' + FabricWhereFDCaisse(False) + 'GROUP BY GJM_MODEPAIE'
    else
      // fond de caisse versé à l'ouverture et en cours de journée
      sSql := 'SELECT GPE_MODEPAIE AS MODEPAIE,SUM(GPE_MONTANTENCAIS) AS FDCAISSEDEV FROM PIEDECHE ' +
        'WHERE ' + FabricWhereCaisse('GPE') + ' AND EXISTS (SELECT GL_NUMLIGNE FROM LIGNE ' +
        'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE WHERE GPE_NATUREPIECEG=GL_NATUREPIECEG ' +
        'AND GPE_SOUCHE=GL_SOUCHE AND GPE_NUMERO=GL_NUMERO AND GPE_INDICEG=GL_INDICEG ' +
        'AND GL_TYPEARTICLE="FI" AND GA_TYPEARTFINAN="' + TYPEOPCFDCAISSE + '") GROUP BY GPE_MODEPAIE';
    QQ := OpenSQL(sSql, True);
    while not QQ.EOF do
    begin
      sCodeMdp := QQ.Findfield('MODEPAIE').AsString;
      TOBR := TOBReg.FindFirst(['MP_MODEPAIE', 'MP_TYPEMODEPAIE'], [sCodeMdp, TYPEPAIEESPECE], False);
      if TOBR <> nil then TOBR.PutValue('FONDCAISSE', QQ.Findfield('FDCAISSEDEV').AsFloat);
      QQ.Next;
    end;
    Ferme(QQ);
  end;

  //Règlements du jour
  if (TOBPCaisse <> nil) and (TOBPCaisse.Detail.Count > 0) then
  begin
    sSql := 'SELECT GPE_MODEPAIE, SUM(GPE_MONTANTENCAIS) AS MONTANTENCAIS, SUM(GPE_MONTANTECHE) AS MONTANTECHE,'
          + ' GOC_NUMPIECELIEN'
          + ' FROM PIEDECHE'
          + ' LEFT JOIN PIECE ON GPE_NATUREPIECEG=GP_NATUREPIECEG AND GPE_SOUCHE=GP_SOUCHE'
          +   ' AND GPE_NUMERO=GP_NUMERO AND GPE_INDICEG=GP_INDICEG '
          + ' LEFT JOIN OPERCAISSE ON GPE_NATUREPIECEG=GOC_NATUREPIECEG AND GPE_SOUCHE=GOC_SOUCHE'
          +   ' AND GPE_NUMERO=GOC_NUMERO AND GPE_INDICEG=GOC_INDICEG'
          +   ' AND GPE_INDICEG=GOC_INDICEG AND GOC_PREFIXE="GPE"'
          + ' WHERE ' + FabricWhereCaisse('GP')
          + ' GROUP BY GPE_MODEPAIE, GOC_NUMPIECELIEN';
    TOBMdp := TOB.Create('', nil, -1);
    TOBMdp.LoadDetailFromSQL(sSql, False, True);

    for Ind := 0 to TOBMdp.Detail.Count -1 do
    begin
      TOBL := TOBMdp.Detail[Ind];
      sCodeMdp := TOBL.GetValue('GPE_MODEPAIE');

      // DBR Agl580
      NumPieceLien := TOBL.GetString ('GOC_NUMPIECELIEN');
      {
      if VarIsNull(TOBL.GetValue('GOC_NUMPIECELIEN')) then
        NumPieceLien := ''
        else
        NumPieceLien := TOBL.GetValue('GOC_NUMPIECELIEN');
      if VarIsNull(NumPieceLien) then NumPieceLien := '';
      }
      if NumPieceLien = '' then
        TOBR := TOBReg.FindFirst(['MP_MODEPAIE', 'ANTERIEUR'], [sCodeMdp, '-'], False)
        else
        TOBR := TOBReg.FindFirst(['MP_MODEPAIE', 'ANTERIEUR'], [sCodeMdp, 'X'], False);
      if TOBR <> nil then
      begin
        Mnt := TOBL.GetValue('MONTANTECHE'); // JTR Reste dû réglé
        TOBR.PutValue('MONTANTECHE', Mnt);
        TOBR.PutValue('MONTANTENCAIS', TOBL.GetValue('MONTANTENCAIS'));
        if TOBR.GetValue('CO_LIBRE') = '1' then
          CumEncais := CumEncais + Mnt
        else if TOBR.GetValue('CO_LIBRE') = '2' then
          CumAutre := CumAutre + Mnt
        else if TOBR.GetValue('CO_LIBRE') = '3' then
          CumEncaisCredit := CumEncaisCredit + Mnt
        else if TOBR.GetValue('CO_LIBRE') = '4' then
        begin
          TOBR.PutValue('MONTANTECHE', TOBR.GetValue('MONTANTECHE')-Mnt);
          TOBR.PutValue('MONTANTENCAIS', TOBR.GetValue('MONTANTENCAIS')-Mnt);
          Mnt := 0;
        end;
        TotGen := TotGen + Mnt;
      end;
    end;
    if TOBMdp <> nil then FreeAndNil(TOBMdp);

    // Reste dû antérieur non encaissés
    CumEncaisCreditAnt := 0;
    if TobRDAnt.detail.count > 0 then
    begin
      for Ind := 0 to TobRDAnt.Detail.count-1 do
      begin
        DateFin := TOBPCaisse.detail[0].GetValue('DATEFIN');
        DatePce := TobRDAnt.Detail[Ind].GetValue('GOC_DATEPIECE');
        if ((DatePce = DateFin) and
           (TobRDAnt.Detail[Ind].GetValue('GOC_NUMZCAISSE') < IntToStr(TOBPCaisse.detail[0].GetValue('NUMZ')))) or
           (DatePce < DateFin) then
          CumEncaisCreditAnt := CumEncaisCreditAnt + TobRDAnt.Detail[Ind].GetValue('GOC_TOTALTTC');
      end;
    end;
    TobRDAnterieur.PutValue('MONTANTECHE', CumEncaisCreditAnt);
    TobRDAnterieur.PutValue('MONTANTENCAIS', CumEncaisCreditAnt);
    TotGen := TotGen + CumEncaisCreditAnt;

    TOBR := TOB.Create('', TOBReg, 0);
    TOBR.AddChampSupValeur('MP_TYPEMODEPAIE', 'TOTAL');
    TOBR.AddChampSupValeur('MP_MODEPAIE', '1');
    TOBR.AddChampSupValeur('MP_DEVISEFO', '');
    TOBR.AddChampSupValeur('MP_LIBELLE', HM.Mess[8]); // Total
    TOBR.AddChampSupValeur('CO_LIBRE', '');
    if GereFDCaisse then TOBR.AddChampSupValeur('FONDCAISSE', 0);
    TOBR.AddChampSupValeur('MONTANTECHE', TotGen);
    TOBR.AddChampSupValeur('MONTANTENCAIS', 0);

    TOBR := TOB.Create('', TOBReg, 1);
    TOBR.AddChampSupValeur('MP_TYPEMODEPAIE', 'CUM');
    TOBR.AddChampSupValeur('MP_MODEPAIE', '2');
    TOBR.AddChampSupValeur('MP_DEVISEFO', '');
    TOBR.AddChampSupValeur('MP_LIBELLE', HM.Mess[16]); // Total encaissements
    TOBR.AddChampSupValeur('CO_LIBRE', '');
    if GereFDCaisse then TOBR.AddChampSupValeur('FONDCAISSE', 0);
    TOBR.AddChampSupValeur('MONTANTECHE', CumEncais);
    TOBR.AddChampSupValeur('MONTANTENCAIS', 0);

    // JTR Reste dû réglé
    for Ind := 0 to TOBReg.Detail.Count -1 do
    begin
      TOBR := TOBReg.Detail[Ind];
      if TOBR.GetValue('CO_LIBRE') = '3' then Break;
    end;

    TOBR := TOB.Create('', TOBReg, Ind);
    TOBR.AddChampSupValeur('MP_TYPEMODEPAIE', 'CUM');
    TOBR.AddChampSupValeur('MP_MODEPAIE', '4');
    TOBR.AddChampSupValeur('MP_DEVISEFO', '');
    TOBR.AddChampSupValeur('MP_LIBELLE', HM.Mess[22]); // Total encaissements de crédit
    TOBR.AddChampSupValeur('CO_LIBRE', '');
    if GereFDCaisse then TOBR.AddChampSupValeur('FONDCAISSE', 0);
    TOBR.AddChampSupValeur('MONTANTECHE', CumEncaisCredit + CumEncaisCreditAnt);
    TOBR.AddChampSupValeur('MONTANTENCAIS', 0);

    TOBR := TOB.Create('', TOBReg, -1);
    TOBR.AddChampSupValeur('MP_TYPEMODEPAIE', 'CUM');
    TOBR.AddChampSupValeur('MP_MODEPAIE', '3');
    TOBR.AddChampSupValeur('MP_DEVISEFO', '');
    TOBR.AddChampSupValeur('MP_LIBELLE', HM.Mess[17]); // Total autres règlements
    TOBR.AddChampSupValeur('CO_LIBRE', '');
    if GereFDCaisse then TOBR.AddChampSupValeur('FONDCAISSE', 0);
    TOBR.AddChampSupValeur('MONTANTECHE', CumAutre);
    TOBR.AddChampSupValeur('MONTANTENCAIS', 0);

    TOBMdp.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N ACHINO
Créé le ...... : 26/02/2004
Modifié le ... : 26/02/2004
Description .. : Affichage des règlements
Mots clefs ... :
*****************************************************************}

procedure TFConsultCaisse.AfficheReglements;
var
  Ind, ARow: integer;
  TOBL: TOB;
  sCodeMdp, sTypeMdp, sLib, sDev, Symbole: string;
  Mnt: double;
  EcartOk: boolean;
begin
  ARow := GSR.FixedRows;
  GSR.RowCount := GSR.FixedRows + TOBReg.Detail.Count;

  for Ind := 0 to TOBReg.Detail.Count -1 do
  begin
    TOBL := TOBReg.Detail[Ind];
    sCodeMdp := TOBL.GetValue('MP_MODEPAIE');
    // DBR Agl580
    sTypeMdp := TOBL.GetString ('MP_TYPEMODEPAIE');
    Mnt := TOBL.GetDouble ('MONTANTECHE');
    {
    if VarIsNull(TOBL.GetValue('MP_TYPEMODEPAIE')) then
      sTypeMdp := ''
      else
      sTypeMdp := TOBL.GetValue('MP_TYPEMODEPAIE');
    Mnt := TOBL.GetValue('MONTANTECHE');
    }

    if (sOldReglts = REGLTSINI) and (Mnt = 0) then Continue
    else if (sOldReglts = REGLTSAUCUN) and (sTypeMdp <> 'TOTAL') then Continue;

    EcartOk := FOStrCmp(sCodeMdp, sMdpEcart);

    GSR.Cells[GSR_TYPE, ARow] := sTypeMdp;
    GSR.Cells[GSR_MTTENCAIS, ARow] := FormatFloat(GSR.ColFormats[GSR_MTTENCAIS], Mnt);
    GSR.Cells[GSR_CODE, ARow] := sCodeMdp;

    if sTypeMdp = 'TOTAL' then sLib := ''
    else if sTypeMdp = 'CUM' then sLib := StringOfChar(' ', Espacement)
    else sLib := StringOfChar(' ', 2*Espacement);
    GSR.Cells[GSR_LIBELLE, ARow] := sLib + TOBL.GetValue('MP_LIBELLE');

    sDev := TOBL.GetString('MP_DEVISEFO');
    if sDev = '' then Symbole := V_PGI.SymbolePivot else Symbole := SymboleDevise(sDev);
    if (sTypeMdp = 'TOTAL') or (sTypeMdp = 'CUM') or (EcartOk) then
      GSR.Cells[GSR_MTTDEVISE, ARow] := ''
    else
      GSR.Cells[GSR_MTTDEVISE, ARow] := StrFMontant(TOBL.GetValue('MONTANTENCAIS'), 12, 2, Symbole, True);

    if (GereFDCaisse) and (sTypeMdp = TYPEPAIEESPECE) and not (EcartOk) then
      GSR.Cells[GSR_FONDS, ARow] := FormatFloat(GSR.ColFormats[GSR_FONDS], TOBL.GetValue('FONDCAISSE'))
    else
      GSR.Cells[GSR_FONDS, ARow] := '';

    Inc(Arow);
  end;
  GSR.RowCount := ARow;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N ACHINO
Créé le ...... : 26/02/2004
Modifié le ... : 26/02/2004
Description .. : Chargement des ventes
Mots clefs ... :
*****************************************************************}

procedure TFConsultCaisse.ChargeVentes;
var
  sSql, sType: string;
  TOBStat, TOBS, TOBOpFi, TOBPre: TOB;
  Ind, Nb: integer;
  NbVentes, NbRetours, NbOpFi, NbPreDiv, NbRDAnt, NbEncRDAnt: Integer;
  TotVentes, TotRetours, NbArt, CAOpFi, CAPreDiv, Nbre, Val: Double;
  NumZ : integer;
  TotalRDAnt, TotalEncRDAnt : double;
  DateFin, DatePce : TDateTime;
  Qry : TQuery;

  function AjouteLigne(Code, Libelle: string; Nbre, Mnt: double; Pos : integer): TOB;
  begin
    Result := TOB.Create('', TOBVen, Pos);
    Result.AddChampSupValeur('CODE', Code);
    Result.AddChampSupValeur('LIB', Libelle);
    Result.AddChampSupValeur('NOMBRE', Nbre);
    Result.AddChampSupValeur('MONTANT', Mnt);
  end;

begin
  if TOBVen.Detail.Count > 0 then TOBVen.ClearDetail;

  // Ventes et Retours
  sSql := 'SELECT GL_TYPEARTICLE, GL_NUMERO,SUM(GL_MONTANTTTC) AS MNT,SUM(GL_QTEFACT) AS QTE, COUNT(GL_MONTANTTTC) AS NB FROM PIECE'
    + ' LEFT JOIN LIGNE ON GL_NATUREPIECEG=GP_NATUREPIECEG'
    + ' AND GL_SOUCHE=GP_SOUCHE AND GL_NUMERO=GP_NUMERO AND GL_INDICEG=GP_INDICEG'
    + ' WHERE ' + FabricWhereCaisse('GP') + ' AND GP_VENTEACHAT="VEN"'
    + ' AND GL_TYPEARTICLE IN ("MAR","NOM","PRE") AND GL_TYPELIGNE="ART"'
    + ' GROUP BY GL_NUMERO, GL_TYPEARTICLE';
  TOBStat := TOB.Create('', nil, -1);
  TOBStat.LoadDetailFromSQL(sSql);
  NbVentes := 0;  TotVentes := 0;
  NbRetours := 0; TotRetours := 0;
  NbArt := 0;
  TOBS := TOBStat.FindFirst(['!GL_TYPEARTICLE'],['PRE'],True);
  while TOBS <> nil do
  begin
    Val := Arrondi(TOBS.GetValue('MNT'), 2);
    if Val > 0 then
    begin
      Inc(NbVentes);
      TotVentes := TotVentes + Val
    end else
    begin
      Inc(NbRetours);
      TotRetours := TotRetours - Val
    end;
    NbArt := NbArt + Arrondi(TOBS.GetValue('QTE'), 2);
    TOBS := TOBStat.FindNext(['!GL_TYPEARTICLE'],['PRE'],True);
  end;
  // Ajoute Panier moyen
  {
  if (NbVentes + NbRetours) = 0 then
  begin
    Nbre := 0;
    Val := 0;
  end else
  begin
    Nbre := NbArt / (NbVentes + NbRetours);
    Val := (TotVentes - TotRetours) / (NbVentes + NbRetours);
  end;
  AjouteLigne('PANIER', StringOfChar(' ', Espacement) + HM.Mess[9], Nbre, Val, -1);
  }
  // Ajoute Ventes et retours
  AjouteLigne('CUM', StringOfChar(' ', Espacement) + HM.Mess[6], (NbVentes - NbRetours), (TotVentes - TotRetours), -1);
  if TotVentes > 0 then AjouteLigne('1', StringOfChar(' ', 2*Espacement) + HM.Mess[15], NbVentes, TotVentes, -1);
  if TotRetours > 0 then AjouteLigne('2', StringOfChar(' ', 2*Espacement) + HM.Mess[7], NbRetours, TotRetours, -1);
  //
  // Prestations
  //
  TOBPre := nil;
  NbPreDiv := 0; CAPreDiv := 0;
  TOBS := TOBStat.FindFirst(['GL_TYPEARTICLE'],['PRE'],True);
  while TOBS <> nil do
  begin
    sType  := TOBS.GetValue('GL_TYPEARTICLE');
    Val := TOBS.GetValue('MNT');
    Nb  := TOBS.GetValue('NB');
    if TOBPre = nil then TOBPre := AjouteLigne('LIG', StringOfChar(' ', 2*Espacement) + HM.Mess[11], 0, 0, -1);
    Inc(NbPreDiv, Nb);
    CAPreDiv := Arrondi(CAPreDiv + Val, 6);
    AjouteLigne(sType, StringOfChar(' ', Espacement) + TOBS.GetValue('LIB'), Nb, Val, -1);
    TOBS := TOBStat.FindNext(['GL_TYPEARTICLE'],['PRE'],True);
  end;
  if TOBPre <> nil then
  begin
    TOBPre.PutValue('NOMBRE', NbPreDiv);
    TOBPre.PutValue('MONTANT', CAPreDiv);
  end;
  TOBStat.Free;
  //
  // Opérations Financières
  //
  // JTR Reste dû réglé
  // Recherche des modes de paiement (avec reste dû et s'ils sont réglés)
  sSql := 'SELECT GL_TYPEARTICLE, GL_CODEARTICLE, MAX(GL_LIBELLE) AS LIB, COUNT(GL_MONTANTTTC) AS NB,'
        +   ' SUM(GL_MONTANTTTC) AS CA, SUM(GOC_TOTALTTC) AS RD'
        + ' FROM PIECE'
        + ' LEFT JOIN LIGNE ON GL_NATUREPIECEG=GP_NATUREPIECEG AND GL_SOUCHE=GP_SOUCHE'
        +   ' AND GL_NUMERO=GP_NUMERO AND GL_INDICEG=GP_INDICEG'
        + ' LEFT JOIN OPERCAISSE ON GL_NATUREPIECEG=GOC_NATUREPIECEG AND GL_SOUCHE=GOC_SOUCHE'
        +   ' AND GL_NUMERO=GOC_NUMERO AND GL_INDICEG=GOC_INDICEG AND GL_NUMLIGNE=GOC_NUMLIGNE'
        +   ' AND GOC_PREFIXE="GL" AND GOC_NUMPIECELIEN<>""'
        + ' WHERE '+ FabricWhereCaisse('GP') +' AND GL_TYPEARTICLE ="FI"'
        + ' GROUP BY GL_TYPEARTICLE, GL_CODEARTICLE, GOC_TOTALTTC'
        + ' ORDER BY GL_TYPEARTICLE DESC, GL_CODEARTICLE';
  TOBStat := TOB.Create('', nil, -1);
  TOBStat.LoadDetailFromSQL(sSql);
  TOBOpFi := nil;
  NbOpFi := 0;
  CAOpFi := 0;
  for Ind := 0 to TOBStat.Detail.Count - 1 do
  begin
    TOBS := TOBStat.Detail[Ind];
    // DBR Agl580
    Val := TOBS.GetDouble ('CA');
    Nb := TOBS.GetInteger ('NB');
    sType := TOBS.GetString ('GL_TYPEARTICLE');
    {
    Val := TOBS.GetValue('CA');
    Nb := TOBS.GetValue('NB');
    sType := TOBS.GetValue('GL_TYPEARTICLE');
    }
    // DBR Agl580
    Val := Val - TobS.GetDouble ('RD');
    {
    if VarIsNull(TOBS.GetValue('RD')) then
      TOBS.PutValue('RD',0)
      else
      Val := Val - Valeur(TOBS.GetValue('RD'));
    }
    if Val = 0 then continue; // Pour traiter les encaissement de crédit réglé
    if TOBOpFi = nil then TOBOpFi := AjouteLigne('CUM', StringOfChar(' ', Espacement) + HM.Mess[10], 0, 0, -1);
    Inc(NbOpFi, Nb);
    CAOpFi := Arrondi(CAOpFi + Val, 6);
    AjouteLigne(sType, StringOfChar(' ', Espacement) + TOBS.GetValue('LIB'), Nb, Val, -1);
  end;
  TOBStat.Free;

  if TOBOpFi = nil then TOBOpFi := AjouteLigne('CUM', StringOfChar(' ', Espacement) + HM.Mess[10], 0, 0, -1);
  TOBOpFi.PutValue('NOMBRE', NbOpFi);
  TOBOpFi.PutValue('MONTANT', CAOpFi);

  //Reste dû antérieurs non réglés
  NumZ := TOBPCaisse.detail[0].GetValue('NUMZ');
  sSql := 'SELECT GOC_CAISSE, GOC_NUMZCAISSE, GOC_DATEPIECE, GOC_TOTALTTC FROM OPERCAISSE'
        + ' LEFT OUTER JOIN MODEPAIE ON  MP_MODEPAIE=GOC_MODEPAIE'
        + ' WHERE GOC_CAISSE ="'+ TOBPCaisse.detail[0].GetValue('NOCAIS') +'"'
        +  ' AND GOC_NUMZCAISSE <='+ IntToStr(NumZ)
        +  ' AND GOC_MODEPAIE <>"" AND GOC_TOTALDISPOTTC > 0'
//          +  ' AND GOC_NATUREPIECEG = "'+  + '"'
        +  ' AND MP_TYPEMODEPAIE = "'+ TYPEPAIERESTEDU + '"'
        +  ' AND GOC_ETABLISSEMENT = "'+ TOBPCaisse.detail[0].GetValue('ETAB') +'"'
        +  ' AND GOC_DATEPIECE <= "'+ USDateTime(TOBPCaisse.detail[0].GetValue('DATEFIN'))+'"';
  TobRDAnt.LoadDetailFromSQL(sSql);

  TotalRDAnt := 0;
  NbRDAnt := 0;

  if TobRDAnt.detail.count > 0 then
  begin
    for Ind := 0 to TobRDAnt.Detail.count-1 do
    begin
      DateFin := TOBPCaisse.detail[0].GetValue('DATEFIN');
      DatePce := TobRDAnt.Detail[Ind].GetValue('GOC_DATEPIECE');
      if ((DatePce = DateFin) and (TobRDAnt.Detail[Ind].GetValue('GOC_NUMZCAISSE') < IntToStr(NumZ))) or
         (DatePce < DateFin) then
      begin
        NbRDAnt := NbRDAnt + 1;
        TotalRDAnt := TotalRDAnt + TobRDAnt.Detail[Ind].GetValue('GOC_TOTALTTC');
      end;
    end;
    AjouteLigne('CUM', StringOfChar(' ', Espacement) + HM.Mess[21], NbRDAnt, TotalRDAnt, -1);
  end;

  //Reste dû antérieurs réglés ce jour
//  TotalEncRDAnt := 0;
//  NbEncRDAnt := 0;
  sSql := 'SELECT COUNT(*) AS NB, SUM(GOC_TOTALTTC) AS MT FROM OPERCAISSE'
        + ' WHERE GOC_CAISSE ="'+ TOBPCaisse.detail[0].GetValue('NOCAIS') +'"'
        +  ' AND GOC_NUMZCAISSE <='+ IntToStr(NumZ)
//          +  ' AND GOC_NATUREPIECEG = "'+  + '"'
        +  ' AND GOC_ETABLISSEMENT = "'+ TOBPCaisse.detail[0].GetValue('ETAB') +'"'
        +  ' AND GOC_DATEPIECE <= "'+ USDateTime(TOBPCaisse.detail[0].GetValue('DATEFIN'))+'"'
        +  ' AND GOC_ANTERIEUR = "X"';
  Qry := OpenSQL(sSQL, true);
  NbEncRDAnt := Qry.FindField('NB').AsInteger;
  TotalEncRDAnt := Qry.FindField('MT').AsFloat;
  Ferme(Qry);
  if NbEncRDAnt > 0 then
    AjouteLigne('CUM', StringOfChar(' ', Espacement) + HM.Mess[20], NbEncRDAnt, TotalEncRDAnt, -1);

  // Total général
  Nbre := (NbVentes - NbRetours) + NbPreDiv + NbOpFi + NbRDAnt + NbEncRDAnt;
  Val := (TotVentes - TotRetours) + CAPreDiv + CAOpFi + TotalRDAnt + TotalEncRDAnt;
  TotGalVente := Val;
  AjouteLigne('TOT', HM.Mess[8], Nbre, Val, 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N ACHINO
Créé le ...... : 26/02/2004
Modifié le ... : 26/02/2004
Description .. : Affichage des ventes
Mots clefs ... :
*****************************************************************}

procedure TFConsultCaisse.AfficheVentes;
var
  Ind, ARow: integer;
  TOBL: TOB;
  sCode: string;
begin
  ARow := GSV.FixedRows;
  GSV.RowCount := GSV.FixedRows + TOBVen.Detail.Count;
  for Ind := 0 to TOBVen.Detail.Count -1 do
  begin
    TOBL := TOBVen.Detail[Ind];
    sCode := TOBL.GetValue('CODE');

    if not (DetailPreDiv) and (sCode = 'PRE') then Continue;
    if not (DetailOpFi) and (sCode = 'FI') then Continue;

    GSV.Cells[GSV_C_CODE, ARow] := sCode;
    GSV.Cells[GSV_C_LIBELLE, ARow] := TOBL.GetValue('LIB');
    if sCode = 'PANIER' then
    begin
      GSV.Cells[GSV_C_NOMBRE, ARow] := FormatFloat(GSR.ColFormats[GSV_C_MONTANT], TOBL.GetValue('NOMBRE'))+ ' ';
      GSV.Cells[GSV_C_MONTANT, ARow] := FormatFloat(GSR.ColFormats[GSV_C_MONTANT], TOBL.GetValue('MONTANT'))+ ' ';
    end else
    begin
      GSV.Cells[GSV_C_NOMBRE, ARow] := FormatFloat(GSR.ColFormats[GSV_C_NOMBRE], TOBL.GetValue('NOMBRE'));
      GSV.Cells[GSV_C_MONTANT, ARow] := FormatFloat(GSR.ColFormats[GSV_C_MONTANT], TOBL.GetValue('MONTANT'));
    end;

    Inc(Arow);
  end;
  GSV.RowCount := ARow;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 27/09/2001
Modifié le ... : 17/05/2002
Description .. : Chargement des paramètres des caisses choisies
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.ChargeParamCaisse;
var sSql, Stg: string;
  TOBP: TOB;
  Ind: Integer;
  DateOuv: TDateTime;
begin
  GereFDCaisse := False;
  sMdpEcart := '';
  if MultiCaisse then
  begin
    sSql := 'SELECT GJC_CAISSE AS NOCAIS,MAX(GJC_NUMZCAISSE) AS NUMZ,'
      + ' MAX(GPK_GEREFONDCAISSE) AS FDCAIS,MAX(GPK_MDPECART) AS ECART,MAX(GPK_ETABLISSEMENT) AS ETAB,'
      + ' MAX(GPK_VENDSAISIE) AS VENDENT,MAX(GPK_VENDSAISLIG) AS VENDLIG'
      + ' FROM JOURSCAISSE,PARCAISSE'
      + ' WHERE ' + FOFabriqueSQLIN(CodeCaisse, 'GJC_CAISSE', False, False, False)
      + ' AND PARCAISSE.GPK_CAISSE=JOURSCAISSE.GJC_CAISSE'
      + ' GROUP BY GJC_CAISSE';
  end else
  begin
    sSql := 'SELECT GPK_CAISSE AS NOCAIS,GPK_GEREFONDCAISSE AS FDCAIS,GPK_MDPECART AS ECART,'
      + ' GPK_ETABLISSEMENT AS ETAB,GPK_VENDSAISIE AS VENDENT,GPK_VENDSAISLIG AS VENDLIG'
      + ' FROM PARCAISSE'
      + ' WHERE ' + FOFabriqueSQLIN(CodeCaisse, 'GPK_CAISSE', False, False, False);
  end;
  TOBPCaisse.LoadDetailFromSQL(sSql);
  for Ind := 0 to TOBPCaisse.Detail.Count - 1 do
  begin
    TOBP := TOBPCaisse.Detail[Ind];
    if MultiCaisse then
    begin
      DeNumZ := TOBP.GetValue('NUMZ');
      ANumZ := DeNumZ;
      TOBP.AddChampSupValeur('NUMZFIN', ANumZ);
      Stg := TOBP.GetValue('NOCAIS');
      DateOuv := FOGetDateOuv(Stg, DeNumZ);
      TOBP.AddChampSupValeur('DATEOUV', DateOuv);
      if ANumZ <> DeNumZ then DateOuv := FOGetDateOuv(Stg, ANumZ);
      TOBP.AddChampSupValeur('DATEFIN', DateOuv);
    end else
    begin
      Stg := TOBP.GetValue('NOCAIS');
      TOBP.AddChampSupValeur('NUMZ', DeNumZ);
      DateOuv := FOGetDateOuv(Stg, DeNumZ);
      TOBP.AddChampSupValeur('DATEOUV', DateOuv);
      TOBP.AddChampSupValeur('NUMZFIN', ANumZ);
      if ANumZ <> DeNumZ then DateOuv := FOGetDateOuv(Stg, ANumZ);
      TOBP.AddChampSupValeur('DATEFIN', DateOuv);
    end;
    if TOBP.GetValue('FDCAIS') = 'X' then GereFDCaisse := True;
    if sMdpEcart = '' then sMdpEcart := TOBP.GetValue('ECART') else
      if FOStrCmp(Stg, sMdpEcart) then sMdpEcart := sMdpEcart + ';' + TOBP.GetValue('ECART');
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 27/09/2001
Modifié le ... : 17/05/2002
Description .. : Fabrique une clause WHERE de sélection des modes de
Suite ........ : paiement pour les écarts de change.
Mots clefs ... : FO;
*****************************************************************}

function TFConsultCaisse.FabricWhereMdpEcart(NotIn: Boolean): string;
begin
  Result := FOFabriqueSQLIN(sMdpEcart, 'MP_MODEPAIE', False, NotIn, False);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 27/09/2001
Modifié le ... : 17/05/2002
Description .. : Fabrique une clause WHERE de sélection des montants du
Suite ........ : fond de caisse
Mots clefs ... : FO;
*****************************************************************}

function TFConsultCaisse.FabricWhereFDCaisse(NotIn: Boolean): string;
var TOBP: TOB;
  Ind: Integer;
  Stg: string;
begin
  Result := '';
  if TOBPCaisse = nil then Exit;
  for Ind := 0 to TOBPCaisse.Detail.Count - 1 do
  begin
    TOBP := TOBPCaisse.Detail[Ind];
    Stg := '(GJM_CAISSE="' + TOBP.GetValue('NOCAIS') + '" AND ' +
      'GJM_NUMZCAISSE=' + IntToStr(TOBP.GetValue('NUMZ')) + ')';
    if Result <> '' then Result := Result + ' OR ';
    Result := Result + Stg;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 27/09/2001
Modifié le ... : 17/05/2002
Description .. : Fabrique une clause WHERE de sélection à partir des
Suite ........ : éléments de la caisse
Mots clefs ... : FO;
*****************************************************************}

function TFConsultCaisse.FabricWhereCaisse(Prefixe: string): string;
var
  TOBP: TOB;
  Ind: Integer;
  Stg: string;
begin
  Result := '';
  if TOBPCaisse = nil then Exit;
  for Ind := 0 to TOBPCaisse.Detail.Count - 1 do
  begin
    TOBP := TOBPCaisse.Detail[Ind];
    if Result = '' then
      Result := '('
      else
      Result := Result + ' OR ';
    Stg := '(' + Prefixe + '_NATUREPIECEG=' + FOGetNatureTicket(False, True)
           + ' AND '+ Prefixe + '_CAISSE="' + TOBP.GetValue('NOCAIS') + '"'
           + ' AND ' + Prefixe + '_DATEPIECE>="' + USDateTime(TOBP.GetValue('DATEOUV')) + '"'
           + ' AND ' + Prefixe + '_DATEPIECE<="' + USDateTime(TOBP.GetValue('DATEFIN')) + '"'
           + ' AND ' + Prefixe + '_NUMZCAISSE>=' + IntToStr(TOBP.GetValue('NUMZ'))
           + ' AND ' + Prefixe + '_NUMZCAISSE<=' + IntToStr(TOBP.GetValue('NUMZFIN'));
    Stg := Stg + ')';
    Result := Result + Stg;
  end;
  if Result <> '' then Result := Result + ')';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 05/06/2001
Description .. :
Mots clefs ... :
*****************************************************************}

function TFConsultCaisse.GereVendeur: Boolean;
var TOBP: TOB;
  Ind: Integer;
begin
  Result := False;
  if TOBPCaisse = nil then Exit;
  for Ind := 0 to TOBPCaisse.Detail.Count - 1 do
  begin
    TOBP := TOBPCaisse.Detail[Ind];
    Result := ((TOBP.GetValue('VENDENT') = 'X') or (TOBP.GetValue('VENDLIG') = 'X'));
    if Result then Break;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 05/06/2001
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFConsultCaisse.ChargeValeurs;
var i: Integer;
  sSQL1, sChamp, sJointure, sType: string;
  QQ: TQuery;
  TOBStatistiques: TOB;
begin
  if (TOBPCaisse = nil) or (TOBPCaisse.Detail.Count <= 0) then Exit;
  //
  // Grid des statistiques
  //
  if StrToInt(Copy(FamilleNiv, 3, 1)) < 4 then
  begin
    sChamp := 'GL_FAMILLENIV' + Copy(FamilleNiv, 3, 1);
    sJointure := '';
  end else
  begin
    sChamp := 'GA2_FAMILLENIV' + Copy(FamilleNiv, 3, 1);
    sJointure := ' LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GL_ARTICLE';
  end;

  sSQL1 := 'SELECT GL_NUMERO,GL_MONTANTTTC,GL_QTEFACT,GL_TOTREMLIGNEDEV,'
    + ' GL_REPRESENTANT,' + sChamp
    + ' FROM PIECE LEFT JOIN LIGNE ON GL_NATUREPIECEG=GP_NATUREPIECEG'
    + ' AND GL_SOUCHE=GP_SOUCHE AND GL_NUMERO=GP_NUMERO AND GL_INDICEG=GP_INDICEG'
    + sJointure
    + ' WHERE '+ FabricWhereCaisse('GP') +'AND GL_TYPEARTICLE IN ("MAR","NOM")';

  QQ := OpenSQL(sSQL1, True);
  TOBStatistiques := TOB.Create('_FLASH', nil, -1);
  TOBStatistiques.LoadDetailDB('_FLASH', '', '', QQ, False);
  Ferme(QQ);

  // les totaux
  if GSS.Cells[GSS_TYPE, GSS.FixedRows] = 'TOTAL' then
  begin
    sChamp := '';
    sType := '';
    CalculStatistiques(TOBStatistiques, sType, sChamp, GSS.FixedRows);
  end;

  // les vendeurs
  //if LowerCase (Trim (GSS.Cells[GSS_LIBELLE, GSS.FixedRows + 1])) = 'vendeur' then
  if sOldStats = STATSINI then
  begin
    i := GSS.FixedRows + 2;
    sChamp := 'GL_REPRESENTANT';
    sType := 'VEN';
    i := CalculStatistiques(TOBStatistiques, sType, sChamp, i);
    // puis les familles
    Inc(i);
    if StrToInt(copy(FamilleNiv, 3, 1)) < 4 then
      sChamp := 'GL_FAMILLENIV' + copy(FamilleNiv, 3, 1) else
      sChamp := 'GA2_FAMILLENIV' + copy(FamilleNiv, 3, 1);
    sType := 'FN' + copy(FamilleNiv, 3, 1);
    CalculStatistiques(TOBStatistiques, sType, sChamp, i);
  end else

    // les familles
  begin
    i := GSS.FixedRows + 2;
    if StrToInt(copy(FamilleNiv, 3, 1)) < 4 then
      sChamp := 'GL_FAMILLENIV' + copy(FamilleNiv, 3, 1) else
      sChamp := 'GA2_FAMILLENIV' + copy(FamilleNiv, 3, 1);
    sType := 'FN' + copy(FamilleNiv, 3, 1);
    i := CalculStatistiques(TOBStatistiques, sType, sChamp, i);
    // puis les vendeurs
    Inc(i);
    sChamp := 'GL_REPRESENTANT';
    sType := 'VEN';
    CalculStatistiques(TOBStatistiques, sType, sChamp, i);
  end;
  TOBStatistiques.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 05/06/2001
Description .. :
Mots clefs ... :
*****************************************************************}

function TFConsultCaisse.CalculStatistiques(TOBStatistiques: TOB; stType, stChamp: string; i: integer): integer;
var nbtickets, j: integer;
  sTickets: string;
begin
  if (stType <> '') and (stChamp <> '') then
    while GSS.Cells[GSS_TYPE, i] = stType do
      // cas des vendeurs et des catégories
    begin
      nbtickets := 0;
      sTickets := '';
      GSS.Cells[GSS_CATTC, i] := FormatFloat(GSS.ColFormats[GSS_CATTC], TOBStatistiques.Somme('GL_MONTANTTTC', [stChamp], [GSS.Cells[GSS_CODE, i]], True));
      GSS.Cells[GSS_NBART, i] := FormatFloat(GSS.ColFormats[GSS_NBART], TOBStatistiques.Somme('GL_QTEFACT', [stChamp], [GSS.Cells[GSS_CODE, i]], True));

      for j := 0 to TobStatistiques.Detail.Count - 1 do
        if TobStatistiques.Detail[j].GetValue(stChamp) = GSS.Cells[GSS_CODE, i] then
          if pos(TobStatistiques.Detail[j].GetValue('GL_NUMERO'), sTickets) = 0 then
          begin
            Inc(nbtickets);
            sTickets := sTickets + ';' + IntToStr(TobStatistiques.Detail[j].GetValue('GL_NUMERO'));
          end;

      GSS.Cells[GSS_NBVTC, i] := FormatFloat(GSS.ColFormats[GSS_NBVTC], nbtickets);
      GSS.Cells[GSS_REMISE, i] := FormatFloat(GSS.ColFormats[GSS_REMISE], TOBStatistiques.Somme('GL_TOTREMLIGNEDEV', [stChamp], [GSS.Cells[GSS_CODE, i]],
        True));
      if nbtickets > 0 then // pour traiter le cas de la division par 0
      begin
        GSS.Cells[GSS_PANIERQTE, i] := FormatFloat(GSS.ColFormats[GSS_PANIERQTE], Valeur(GSS.Cells[GSS_NBART, i]) / Valeur(GSS.Cells[GSS_NBVTC, i]));
        GSS.Cells[GSS_PANIERTTC, i] := FormatFloat(GSS.ColFormats[GSS_PANIERTTC], Valeur(GSS.Cells[GSS_CATTC, i]) / Valeur(GSS.Cells[GSS_NBVTC, i]));
      end else
      begin
        GSS.Cells[GSS_PANIERQTE, i] := FormatFloat(GSS.ColFormats[GSS_PANIERQTE], 0);
        GSS.Cells[GSS_PANIERTTC, i] := FormatFloat(GSS.ColFormats[GSS_PANIERTTC], 0);
      end;
      Progression;
      Inc(i);
    end else
    // cas de la ligne de totaux
  begin
    nbtickets := 0;
    sTickets := '';
    GSS.Cells[GSS_CATTC, i] := FormatFloat(GSS.ColFormats[GSS_CATTC], TOBStatistiques.Somme('GL_MONTANTTTC', [''], [''], True));
    GSS.Cells[GSS_NBART, i] := FormatFloat(GSS.ColFormats[GSS_NBART], TOBStatistiques.Somme('GL_QTEFACT', [''], [''], True));

    for j := 0 to TobStatistiques.Detail.Count - 1 do
      if pos(TobStatistiques.Detail[j].GetValue('GL_NUMERO'), sTickets) = 0 then
      begin
        Inc(nbtickets);
        sTickets := sTickets + ';' + IntToStr(TobStatistiques.Detail[j].GetValue('GL_NUMERO'));
      end;

    GSS.Cells[GSS_NBVTC, i] := FormatFloat(GSS.ColFormats[GSS_NBVTC], nbtickets);
    GSS.Cells[GSS_REMISE, i] := FormatFloat(GSS.ColFormats[GSS_REMISE], TOBStatistiques.Somme('GL_TOTREMLIGNEDEV', [''], [''], True));
    if nbtickets > 0 then // pour traiter le cas de la division par 0
    begin
      GSS.Cells[GSS_PANIERQTE, i] := FormatFloat(GSS.ColFormats[GSS_PANIERQTE], Valeur(GSS.Cells[GSS_NBART, i]) / Valeur(GSS.Cells[GSS_NBVTC, i]));
      GSS.Cells[GSS_PANIERTTC, i] := FormatFloat(GSS.ColFormats[GSS_PANIERTTC], Valeur(GSS.Cells[GSS_CATTC, i]) / Valeur(GSS.Cells[GSS_NBVTC, i]));
    end else
    begin
      GSS.Cells[GSS_PANIERQTE, i] := FormatFloat(GSS.ColFormats[GSS_PANIERQTE], 0);
      GSS.Cells[GSS_PANIERTTC, i] := FormatFloat(GSS.ColFormats[GSS_PANIERTTC], 0);
    end;
    Progression;
  end;

  Result := i;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 05/06/2001
Description .. :
Mots clefs ... :
*****************************************************************}


procedure TFConsultCaisse.GetCellCanvasGSV(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow < GSV.FixedRows then Exit;
  if (GSV.Focused) and ((gdSelected in AState) or (gdFocused in AState)) then Exit;

  // mise en évidence des cumuls et totaux
  if GSV.Cells[GSV_C_CODE, ARow] = 'LIG' then
  begin
    Canvas.Brush.Color := ColorLigne;
  end else
  if GSV.Cells[GSV_C_CODE, ARow] = 'CUM' then
  begin
    Canvas.Brush.Color := ColorCumul;
    Canvas.Font.Style := [fsBold];
  end else
  if GSV.Cells[GSV_C_CODE, ARow] = 'TOT' then
  begin
    Canvas.Brush.Color := ColorTotal;
    Canvas.Font.Style := [fsBold];
    Canvas.Font.Color := clWhite;
  end else
  begin
    Canvas.Brush.Color := clWindow;
    if (GSV.Cells[GSV_C_CODE, ARow] = 'PRE') or (GSV.Cells[GSV_C_CODE, ARow] = 'FI') then
      Canvas.Font.Color := clBlue
    else if GSV.Cells[GSV_C_CODE, ARow] = 'PANIER' then
      Canvas.Font.Style := [fsItalic];
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 05/06/2001
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFConsultCaisse.PostDrawCellGSR(ACol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
begin
  GSR.Font.Style := [fsItalic];
  if (ARow >= GSR.FixedRows) and (ACol >= GSR.FixedCols) then
  begin
    if (((GSR.Cells[GSR_TYPE, ARow] = 'TOTAL') or (GSR.Cells[GSR_TYPE, ARow] = 'CUM')) and (ACol = GSR_MTTDEVISE)) or
      ((FOStrCmp(GSR.Cells[GSR_CODE, ARow], sMdpEcart)) and ((ACol = GSR_MTTDEVISE) or (ACol = GSR_FONDS))) or
      ((GSR.Cells[GSR_TYPE, ARow] <> TYPEPAIEESPECE) and (ACol = GSR_FONDS)) then
    begin
      GridGriseCellPV(GSR, ACol, ARow, Canvas);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 05/06/2001
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFConsultCaisse.GetCellCanvasGSR(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow < GSR.FixedRows then Exit;
  if (GSR.Focused) and ((gdSelected in AState) or (gdFocused in AState)) then Exit;

  // mise en évidence des cumuls et totaux
  if GSR.Cells[GSR_TYPE, ARow] = 'CUM' then
  begin
    Canvas.Brush.Color := ColorCumul;
    Canvas.Font.Style := [fsBold];
  end else
  if GSR.Cells[GSR_TYPE, ARow] = 'TOTAL' then
  begin
    Canvas.Brush.Color := ColorTotal;
    Canvas.Font.Style := [fsBold];
    Canvas.Font.Color := clWhite;
  end else
  begin
    Canvas.Brush.Color := clWindow;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 05/06/2001
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFConsultCaisse.GetCellCanvasGSS(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
begin
  if ARow < GSS.FixedRows then Exit;
  if (GSS.Focused) and ((gdSelected in AState) or (gdFocused in AState)) then Exit;

  // mise en évidence des cumuls et totaux
  if GSS.Cells[GSS_TYPE, ARow] = 'CUM' then
  begin
    Canvas.Brush.Color := ColorCumul;
    Canvas.Font.Style := [fsBold];
  end else
  if GSS.Cells[GSS_TYPE, ARow] = 'TOTAL' then
  begin
    Canvas.Brush.Color := ColorTotal;
    Canvas.Font.Style := [fsBold];
    Canvas.Font.Color := clWhite;
  end else
  begin
    Canvas.Brush.Color := clWindow;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 05/06/2001
Modifié le ... : 05/06/2001
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TFConsultCaisse.DessineGraphs(ACol: integer);
var
  Ind: integer;
  Mnt: double;
begin
  // Vendeurs
  if GereVendeur then
  begin
    PG1.Visible := TRUE;
    Ind := GSS.Left + GSS.Width - PG2.Width;
    if Ind > 0 then PG2.Left := Ind;
    Series1.Clear;
    for Ind := LigneDebutVendeurs to LigneFinVendeurs do
    begin
      Mnt := Valeur(GSS.Cells[ACol, Ind]);
      if Mnt = 0 then continue;
      Series1.AddY(Mnt, Trim(GSS.Cells[GSS_LIBELLE, Ind]), clTeeColor);
    end;
  end else
  begin
    PG1.Visible := FALSE;
    Ind := GSS.Left + ((GSS.Width - PG2.Width) div 2);
    if Ind > 0 then PG2.Left := Ind;
  end;
  // Familles
  PieSeries1.Clear;
  for Ind := LigneDebutFamilles to LigneFinFamilles do
  begin
    Mnt := Valeur(GSS.Cells[ACol, Ind]);
    if Mnt = 0 then continue;
    PieSeries1.AddY(Mnt, Trim(GSS.Cells[GSS_LIBELLE, Ind]), clTeeColor);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 08/08/2001
Modifié le ... : 24/08/2001
Description .. : Lors de l'entrée dans une cellule
Mots clefs ... : FO;SITUATION;FLASH
*****************************************************************}

procedure TFConsultCaisse.GSSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  Ind: integer;
begin
  if GSS.RowHeights[GSS.Row] < 0 then
  begin
    if GSS.Row > ARow then
    begin
      for Ind := GSS.Row to GSS.RowCount -1 do
        if GSS.RowHeights[Ind] > 0 then Break;
    end else
    begin
      for Ind := GSS.Row downto GSS.FixedRows do
        if GSS.RowHeights[Ind] > 0 then Break;
    end;
    if (Ind >= GSS.FixedRows) and (Ind < GSS.RowCount) then GSS.Row := Ind else Cancel := True;
  end;

  if not Cancel then DessineGraphs(GSS.Col);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 08/08/2001
Modifié le ... : 24/08/2001
Description .. : Bouton abandon
Mots clefs ... : FO;SITUATION;FLASH
*****************************************************************}

procedure TFConsultCaisse.BAbandonClick(Sender: TObject);
begin
  if IsInside(Self) then
  begin
    Close;
    THPanel(Parent).CloseInside;
  end
  else
    ModalResult := mrCancel;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 08/08/2001
Modifié le ... : 24/08/2001
Description .. : Bouton d'aide
Mots clefs ... : FO;SITUATION;FLASH
*****************************************************************}

procedure TFConsultCaisse.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 08/08/2001
Modifié le ... : 24/08/2001
Description .. : Lancement de l'impression du ticket X ou Z
Mots clefs ... : FO;SITUATION;FLASH
*****************************************************************}

procedure TFConsultCaisse.BImprimerClick(Sender: TObject);
begin
  {$IFDEF FOS5}
  if not BImprimer.Enabled then Exit;
  if TOBPCaisse.Detail.Count > 1 then Exit;
  if (FOEtatCaisse(CodeCaisse, False) = 'OUV') and (DeNumZ = FOGetNumZCaisse(CodeCaisse, 'MAX')) then
  begin
    if FOJaiLeDroit(101) then
      FOLanceImprimeLP(efoTicketX, FOMakeWhereTicketX('GP', CodeCaisse), True, nil);
  end else
  begin
    if FOJaiLeDroit(102) then
      FOLanceImprimeLP(efoTicketZ, FOMakeWhereTicketZ('GP', CodeCaisse, DeNumz, ANumZ), True, nil);
  end;
  {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 08/08/2001
Modifié le ... : 24/08/2001
Description .. : Lancement de la fiche de paramétrage de la situation flash
Mots clefs ... : FO;SITUATION;FLASH
*****************************************************************}

procedure TFConsultCaisse.BParametrerClick(Sender: TObject);
var sValidation: string;
begin
  if not BParametrer.Enabled then Exit;
  if not FOJaiLeDroit(83) then Exit;
  // Récupération des paramètres
  sOldReglts := LitValeurRegistre(REGREGLTSPARDEFAUT, REGLTSINI);
  sOldCategorie := LitValeurRegistre(REGCATEGORIEPARDEFAUT, Famille.Value);
  sOldStats := LitValeurRegistre(REGSTATSPARDEFAUT, STATSINI);
  sOldVendeurs := LitValeurRegistre(REGVENDEURSPARDEFAUT, VENDEURSINI);
  sOldTypeCateg := LitValeurRegistre(REGTYPECATEGPARDEFAUT, TYPECATEGINI);
  sOldCaisseBtq := LitValeurRegistre(REGCAISSEBTQ, True);
  // Lancement de la fiche de paramétrage de la situation flash
  sValidation := AGLLanceFiche('MFO', 'PARAMSITUFLASH', '', '', '');
  if sValidation = 'validation' then
    AppliquerParametrage;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 27/08/2001
Modifié le ... : 25/09/2001
Description .. : Applique le paramétrage de la situation flash
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.AppliquerParametrage;
var i, j, LDV, LFV, LDF, LFF, nbVendeurs, nbFamilles, stVendeurs, stCategorie, stStats, stReglts: integer;
  Ok, Ok2: boolean;
  st: string;
begin

  //
  //  Actualisation de la liste des caisses
  //
  Ok := LitValeurRegistre(REGCAISSEBTQ, True);
  if Ok <> sOldCaisseBtq then
  begin
    sOldCaisseBtq := Ok;
    ParamListeCaisse;
  end;
  //
  //  Actualisation du détail des prestations et des opérations de caisse
  //
  Ok := LitValeurRegistre(REGDETAILPRESTATION, False);
  Ok2 := LitValeurRegistre(REGDETAILOPCAISSE, False);
  if (Ok <> DetailPreDiv) or (Ok2 <> DetailOpFi) then
  begin
    DetailPreDiv := Ok;
    DetailOpFi := Ok2;
    AfficheVentes;
  end;

  stVendeurs := LitValeurRegistre(REGVENDEURSPARDEFAUT, VENDEURSINI);
  stCategorie := LitValeurRegistre(REGTYPECATEGPARDEFAUT, TYPECATEGINI);
  stStats := LitValeurRegistre(REGSTATSPARDEFAUT, STATSINI);
  stReglts := LitValeurRegistre(REGREGLTSPARDEFAUT, REGLTSINI);

  //
  // Actualisation de la catégorie
  //
  st := LitValeurRegistre(REGCATEGORIEPARDEFAUT, Famille.Value);
  if sOldCategorie <> st then
  begin
    sOldTypeCateg := -1;
    sOldVendeurs := -1;
    if ((not GetParamSoc('SO_ARTLOOKORLI')) and (StrToInt(copy(st, 3, 1)) > 3)) or (RechDom('GCLIBFAMILLE', st, False) = '.-') then
      Famille.Value := 'LF1' else
      Famille.Value := st;
  end;

  //
  // Modification de l'ordre d'affichage des statistiques
  //

  // stockage des limtes des tableaux
  LDV := LigneDebutVendeurs;
  LFV := LigneFinVendeurs;
  LDF := LigneDebutFamilles;
  LFF := LigneFinFamilles;

  GSS.SynEnabled := False;
  if sOldStats <> stStats then
  begin

    // Affichage des tableaux dans leur totalité
    for i := LigneDebutVendeurs to LigneFinVendeurs do
      if (GSS.Cells[GSS_NBVTC, i] = '0') and (GSS.RowHeights[i] = -1) then
        GSS.RowHeights[i] := GSS.DefaultRowHeight;
    for i := LigneDebutFamilles to LigneFinFamilles do
      if (GSS.Cells[GSS_NBVTC, i] = '0') and (GSS.RowHeights[i] = -1) then
        GSS.RowHeights[i] := GSS.DefaultRowHeight;

    nbVendeurs := LFV - LDV;
    nbFamilles := LFF - LDF;

    // Les vendeurs en tête de tableau
    if (stStats = STATSINI) and (LDV > LDF) then
    begin
      // insertion des lignes manquantes
      if nbVendeurs < nbFamilles then
      begin
        for i := 1 to nbFamilles - nbVendeurs do
        begin
          GSS.InsertRow(LFV + 1);
          Inc(LFV);
        end;
      end else
        if nbVendeurs > nbFamilles then
        for i := 1 to nbVendeurs - nbFamilles do
        begin
          GSS.InsertRow(LFF + 1);
          Inc(LFF);
          Inc(LDV);
          Inc(LFV);
        end;

      // Permutation des lignes
      if nbVendeurs >= nbFamilles then
        j := nbVendeurs else
        j := nbFamilles;
      for i := -1 to j do
        GSS.ExchangeRow(LDV + i, LDF + i);

      // Suppression des lignes rajoutées
      if nbVendeurs < nbFamilles then
      begin
        for i := 1 to nbFamilles - nbVendeurs do
        begin
          GSS.DeleteRow(LFF);
          Dec(LFF);
          Dec(LDV);
          Dec(LFV);
        end;
      end else
        if nbVendeurs > nbFamilles then
      begin
        for i := 1 to nbVendeurs - nbFamilles do
        begin
          GSS.DeleteRow(LFV);
          Dec(LFV);
        end;
      end;
    end;

    // La catégorie en tête de tableau
    if (stStats = STATSCATEG) and (LDF > LDV) then
    begin
      // insertion des lignes manquantes
      if nbVendeurs < nbFamilles then
      begin
        for i := 1 to nbFamilles - nbVendeurs do
        begin
          GSS.InsertRow(LFV + 1);
          Inc(LDF);
          Inc(LFF);
        end;
      end else
        if nbVendeurs > nbFamilles then
        for i := 1 to nbVendeurs - nbFamilles do
        begin
          GSS.InsertRow(LFF + 1);
          Inc(LFF);
        end;

      // Permutation des lignes
      if nbVendeurs >= nbFamilles then
        j := nbVendeurs else
        j := nbFamilles;
      for i := -1 to j do
        GSS.ExchangeRow(LDV + i, LDF + i);

      // Suppression des lignes rajoutées
      if nbVendeurs < nbFamilles then
      begin
        for i := 1 to nbFamilles - nbVendeurs do
        begin
          GSS.DeleteRow(LFF);
          Dec(LFF);
          Inc(LFV);
        end;
      end else
        if nbVendeurs > nbFamilles then
      begin
        for i := 1 to nbVendeurs - nbFamilles do
        begin
          GSS.DeleteRow(LFV);
          Dec(LFV);
          Dec(LDF);
          Dec(LFF);
        end;
      end;
    end;

    // Réaffectation des limites
    LigneDebutFamilles := LDV;
    LigneFinFamilles := LFV;
    LigneDebutVendeurs := LDF;
    LigneFinVendeurs := LFF;
    // pour forcer l'affichage des vendeurs et des familles
    sOldVendeurs := -1;
    sOldTypeCateg := -1;
  end;

  //
  // Modification de l'affichage des vendeurs
  //
  if sOldVendeurs <> stVendeurs then
    if stVendeurs = VENDEURSINI then
    begin
      for i := LigneDebutVendeurs to LigneFinVendeurs do
        if GSS.Cells[GSS_NBVTC, i] = '0' then
          GSS.RowHeights[i] := -1;
    end else
    begin
      for i := LigneDebutVendeurs to LigneFinVendeurs do
        if GSS.Cells[GSS_NBVTC, i] = '0' then
          GSS.RowHeights[i] := GSS.DefaultRowHeight;
    end;

  //
  // Modification de l'affichage de la catégorie
  //
  if sOldTypeCateg <> stCategorie then
    if stCategorie = TYPECATEGINI then
    begin
      for i := LigneDebutFamilles to LigneFinFamilles do
        if GSS.Cells[GSS_NBVTC, i] = '0' then
          GSS.RowHeights[i] := -1;
    end else
    begin
      for i := LigneDebutFamilles to LigneFinFamilles do
        if GSS.Cells[GSS_NBVTC, i] = '0' then
          GSS.RowHeights[i] := GSS.DefaultRowHeight;
    end;
  GSS.SynEnabled := True;

  // Modification de l'affichage des règlements
  if sOldReglts <> stReglts then
  begin
    sOldReglts := stReglts;
    AfficheReglements;
  end;

  FormResize(nil);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 05/10/2001
Modifié le ... : 17/05/2002
Description .. : Paramétrage de la liste des caisses
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.ParamListeCaisse;
var Stg: string;
begin
  Stg := 'AND GPK_FERME="-"';
  if sOldCaisseBtq then
    Stg := Stg + ' AND GPK_ETABLISSEMENT="' + FOGetParamCaisse('GPK_ETABLISSEMENT') + '"';
  NumCaisse.Plus := Stg;
  if NumCaisse.ItemIndex < 0 then NumCaisse.Value := FOCaisseCourante;

  NumCaisse.Visible := (NumCaisse.Values.Count > 2);
  LNumCaisse.Visible := NumCaisse.Visible;
  NumCaisse.Enabled := MultiCaisse;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 25/09/2001
Modifié le ... : 25/09/2001
Description .. : Sous-tableau vendeurs du grid des statistiques
Mots clefs ... : FO;
*****************************************************************}

procedure TFConsultCaisse.TabGridVendeurs(Grid: THGrid);
var Q: TQuery;
  sSql, Stg: string;
  stEtab: string;
  TOBP: TOB;
  Ind: Integer;
  DateSupp: TDateTime;
  DateOuv: TDateTime;
begin
  GSS.RowCount := GSS.RowCount + 1;
  GSS.Cells[GSS_LIBELLE, GSS.RowCount - 1] := StringOfChar(' ', Espacement) + HM.Mess[12]; // Vendeur
  GSS.Cells[GSS_TYPE, GSS.RowCount - 1] := 'CUM';
  LigneDebutVendeurs := GSS.RowCount;
  // recheche de la date d'ouverture la plus haute et des établissements concernés
  stEtab := '';
  Stg := '';
  DateSupp := Date;
  if TOBPCaisse <> nil then
  begin
    for Ind := 0 to TOBPCaisse.Detail.Count - 1 do
    begin
      TOBP := TOBPCaisse.Detail[Ind];
      if stEtab = '' then stEtab := TOBP.GetValue('ETAB') else
        if FOStrCmp(Stg, stEtab) then stEtab := stEtab + ';' + TOBP.GetValue('ETAB');
      if Ind = 0 then DateSupp := TOBP.GetValue('DATEOUV');
      DateOuv := TOBP.GetValue('DATEOUV');
      if DateOuv > DateSupp then DateSupp := DateOuv;
    end;
  end;
  // recherche des vendeurs
  sSql := 'select GCL_COMMERCIAL,GCL_LIBELLE from COMMERCIAL where GCL_TYPECOMMERCIAL="VEN" and GCL_DATESUPP>"' +
    USDateTime(DateSupp) + '" AND ' + FOFabriqueSQLIN(stEtab, 'GCL_ETABLISSEMENT', False, False, False);
  Q := OpenSQL(sSql, True);
  while not Q.EOF do
  begin
    GSS.RowCount := GSS.RowCount + 1;
    GSS.Cells[GSS_TYPE, GSS.RowCount - 1] := 'VEN';
    GSS.Cells[GSS_CODE, GSS.RowCount - 1] := Q.FindField('GCL_COMMERCIAL').AsString;
    GSS.Cells[GSS_LIBELLE, GSS.RowCount - 1] := StringOfChar(' ', 2*Espacement) + Q.FindField('GCL_LIBELLE').AsString;
    Q.Next;
  end;
  Ferme(Q);
  // Ajout d'une ligne en fin de tableau pour les lignes sans vendeur
  GSS.RowCount := GSS.RowCount + 1;
  GSS.Cells[GSS_TYPE, GSS.RowCount - 1] := 'VEN';
  GSS.Cells[GSS_CODE, GSS.RowCount - 1] := '';
  GSS.Cells[GSS_LIBELLE, GSS.RowCount - 1] := StringOfChar(' ', 2*Espacement) + TraduireMemoire('Sans vendeur');

  LigneFinVendeurs := GSS.RowCount - 1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 25/09/2001
Modifié le ... : 25/09/2001
Description .. : Sous-tableau catégorie du grid des statistiques
Mots clefs ... :
*****************************************************************}

procedure TFConsultCaisse.TabGridCateg(Grid: THGrid);
var Q: TQuery;
  Stg: string;
begin
  Stg := TraduireMemoire(RechDom('GCLIBFAMILLE', Famille.Value, False));
  GSS.RowCount := GSS.RowCount + 1;
  GSS.Cells[GSS_LIBELLE, GSS.RowCount - 1] := StringOfChar(' ', Espacement) + Stg; // le libellé change selon la catégorie choisie
  GSS.Cells[GSS_TYPE, GSS.RowCount - 1] := 'CUM';
  LigneDebutFamilles := GSS.RowCount;
  Q := OpenSQL('select CC_TYPE, CC_CODE, CC_LIBELLE from CHOIXCOD where CC_TYPE="FN' + copy(FamilleNiv, 3, 1) + '"', True);
  while not Q.EOF do
  begin
    GSS.RowCount := GSS.RowCount + 1;
    GSS.Cells[GSS_TYPE, GSS.RowCount - 1] := Q.FindField('CC_TYPE').AsString;
    GSS.Cells[GSS_CODE, GSS.RowCount - 1] := Q.FindField('CC_CODE').AsString;
    GSS.Cells[GSS_LIBELLE, GSS.RowCount - 1] := StringOfChar(' ', 2*Espacement) + Q.FindField('CC_LIBELLE').AsString;
    Q.Next;
  end;
  Ferme(Q);
  // Ajout d'une ligne en fin de tableau pour les articles sans famille
  GSS.RowCount := GSS.RowCount + 1;
  GSS.Cells[GSS_TYPE, GSS.RowCount - 1] := 'FN' + copy(FamilleNiv, 3, 1);
  GSS.Cells[GSS_CODE, GSS.RowCount - 1] := '';
  GSS.Cells[GSS_LIBELLE, GSS.RowCount - 1] := StringOfChar(' ', 2*Espacement) + TraduireMemoire('Sans ') + Stg;
  LigneFinFamilles := GSS.RowCount - 1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 24/07/2001
Modifié le ... : 21/11/2001
Description .. : Grise une cellule d'une grille identifiée
Suite ........ : par son numéro de ligne et de colonne
Mots clefs ... : FO
*****************************************************************}

procedure TFConsultCaisse.GridGriseCellPV(GS: THGrid; Acol, ARow: integer; Canvas: TCanvas);
var R: TRect;
begin
  Canvas.Brush.Color := $00E4E4E4;
  Canvas.Brush.Style := bsDiagCross;
  Canvas.Pen.Color := Canvas.Brush.Color;
  Canvas.Pen.Mode := pmMask;
  Canvas.Pen.Style := psDot;
  Canvas.Pen.Width := 1;
  R := GS.CellRect(ACol, ARow);
  Canvas.Rectangle(R.Left - 1, R.Top - 1, R.Right + 1, R.Bottom + 1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 08/08/2001
Modifié le ... : 24/08/2001
Description .. : Modification de la taille de la fenêtre
Mots clefs ... : FO;SITUATION;FLASH
*****************************************************************}

procedure TFConsultCaisse.FormResize(Sender: TObject);
begin
  HmTrad.ResizeGridColumns(GSV);
  HmTrad.ResizeGridColumns(GSR);
  HmTrad.ResizeGridColumns(GSS);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 08/08/2001
Modifié le ... : 24/08/2001
Description .. : Incrémente la barre de progression
Mots clefs ... : FO;SITUATION;FLASH
*****************************************************************}

procedure TFConsultCaisse.Progression;
begin
  MoveCur(False);
  Inc(Jauge);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 08/08/2001
Modifié le ... : 24/08/2001
Description .. : Gestion des touches du clavier
Mots clefs ... : FO;SITUATION;FLASH
*****************************************************************}

procedure TFConsultCaisse.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if Key = VK_ESCAPE then
  begin
    Key := 0;
    BAbandonClick(nil);
  end;
end;

end.
