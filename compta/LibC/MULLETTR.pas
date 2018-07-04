unit MULLETTR;

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
{$IFDEF EAGLCLIENT}

{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  QRGrid,
{$ENDIF}
  Grids,
  StdCtrls,
  Mask,
  Hctrls,
  Hcompte,
  SaisUtil,
  ComCtrls,
  Buttons,
  ExtCtrls,
  Hqry,
  Ent1,
  LettUtil, //InitTablesLibresTiers
  Lettrage,
  hmsgbox,
  HEnt1,
  Menus,
  Paramdat,
  HSysMenu,
  HTB97,
  HPanel,
  UiUtil,
  UTOB,
  LookUp,
  Filtre,     // VideFiltre
  uTXML,      // XMLDecodeSt
  uListByUser // TListByUser
  ;

procedure LanceLettrage;
procedure LanceLettrageMP(Qui: tProfilTraitement);

type
  TFMulLettre = class(TForm)
    Pages: TPageControl;
    Princ: TTabSheet;
    Complements: TTabSheet;
    GL: THGrid;
    Avances: TTabSheet;
    FindLettre: TFindDialog;
    HMessMulL: THMsgBox;
    Cache: THCpteEdit;
    PFiltres: TToolWindow97;
    FFiltres: THValComboBox;
    BChercher: TToolbarButton97;
    HLabel4: THLabel;
    E_GENERAL: THCritMaskEdit;
    HLabel5: THLabel;
    E_GENERAL_: THCritMaskEdit;
    BFeuVert: THBitBtn;
    E_AUXILIAIRE_: THCritMaskEdit;
    HLabel2: THLabel;
    HLabel1: THLabel;
    T_LIBELLE: TEdit;
    Label5: THLabel;
    Bevel1: TBevel;
    Label12: THLabel;
    E_REFINTERNE: THCritMaskEdit;
    LETTREPARTIEL: TCheckBox;
    HLabel10: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    HLabel7: THLabel;
    E_DATEECHEANCE_: THCritMaskEdit;
    Bevel2: TBevel;
    Label7: THLabel;
    E_REFLETTRAGE: THCritMaskEdit;
    Label3: THLabel;
    TOTAL: THCritMaskEdit;
    Label4: THLabel;
    TOTAL_: THCritMaskEdit;
    LETTRAGE_: THCritMaskEdit;
    HLabel11: THLabel;
    LETTRAGE: THCritMaskEdit;
    HLabel9: THLabel;
    Bevel3: TBevel;
    HMTrad: THSystemMenu;
    Plibres: TTabSheet;
    Bevel4: TBevel;
    TT_TABLE0: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE4: THCritMaskEdit;
    T_TABLE3: THCritMaskEdit;
    T_TABLE2: THCritMaskEdit;
    T_TABLE1: THCritMaskEdit;
    T_TABLE0: THCritMaskEdit;
    T_TABLE5: THCritMaskEdit;
    T_TABLE6: THCritMaskEdit;
    T_TABLE7: THCritMaskEdit;
    T_TABLE8: THCritMaskEdit;
    T_TABLE9: THCritMaskEdit;
    POPF: TPopupMenu;
    BCREER: TMenuItem;
    BENREGISTRE: TMenuItem;
    BSUPPRIMER: TMenuItem;
    BRENOMMER: TMenuItem;
    BRECHERCHER: TMenuItem;
    BFiltre: TToolbarButton97;
    HLabel8: THLabel;
    E_DEVISE: THValComboBox;
    TE_EXERCICE: THLabel;
    E_EXERCICE: THValComboBox;
    HLabel3: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    HLabel6: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    Label1: THLabel;
    E_NUMEROPIECE: THCritMaskEdit;
    Label2: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    BAgrandir: TToolbarButton97;
    BRecherche: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BReduire: TToolbarButton97;
    HPB: TToolWindow97;
    Dock: TDock97;
    Dock971: TDock97;
    CFacNotPay: TCheckBox;
    XX_WHERETP: TEdit;
    Label6: THLabel;
    T_TIERS: TEdit;
    TE_ETABLISSEMENT: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    Label14: THLabel;
    T_NATUREAUXI: THValComboBox;
    PRefs: TTabSheet;
    Bevel5: TBevel;
    TE_REFRELEVE: THLabel;
    E_REFRELEVE: THCritMaskEdit;
    E_REFRELEVE_: THCritMaskEdit;
    TE_REFRELEVE_: THLabel;
    TE_REFLIBRE: THLabel;
    E_REFLIBRE: THCritMaskEdit;
    TE_REFLIBRE_: THLabel;
    E_REFLIBRE_: THCritMaskEdit;
    TE_REFEXTERNE: THLabel;
    E_REFEXTERNE: THCritMaskEdit;
    TE_REFEXTERNE_: THLabel;
    E_REFEXTERNE_: THCritMaskEdit;
    iCritGlyph: TImage;
    iCritGlyphModified: TImage;
    E_AUXILIAIRE: THCritMaskEdit;
    BDUPLIQUER: TMenuItem;
    Distinguer: TCheckBox;
    procedure BChercherClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BENREGISTREClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure GLDblClick(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure FindLettreFind(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);


    procedure E_EXERCICEChange(Sender: TObject);
    procedure E_GENERALExit(Sender: TObject);
    procedure BFeuVertClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
    procedure BSUPPRIMERClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BRECHERCHERClick(Sender: TObject);
    procedure BCREERClick(Sender: TObject);
    procedure BRENOMMERClick(Sender: TObject);
    procedure E_AUXILIAIREExit(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure CFacNotPayClick(Sender: TObject);
    procedure E_GENERALElipsisClick(Sender: TObject);
    procedure E_AUXILIAIRE2ElipsisClick(Sender: TObject);
    procedure BDUPLIQUERClick(Sender: TObject);
  private
    QL: TQuery;
    WMinX, WMinY: Integer;
    Qui: tProfilTraitement;
    Loading, FCritModified: boolean;
    FFiltreByUser : TListByUser;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    procedure RechercheEcritures;
    function WhereCrit(AvecTiers: boolean): string;
    function SelectQL: string;
    procedure ChargeAVide;
    procedure AttribCollectif;
    procedure ControleFeuVert;
    procedure NextAuxi(Suiv: boolean);
    procedure NextGene(Suiv: boolean);
    procedure InitCriteres;
    function EtudieModeLettrage(var R: RLETTR): TModalResult;
    procedure SetCritModified(Value: Boolean);
    procedure CritereChange(Sender: TObject);

    // Code du MUL pour la nouvelle Gestion des FILTRES
    procedure InitAddFiltre     (T: TOB);
    procedure InitGetFiltre     (T: TOB);
    procedure InitSelectFiltre  (T: TOB);
    procedure ParseParamsFiltre (Params: HTStrings);
    procedure UpgradeFiltre     (T: TOB);
    // FIN gestion FILTRES

  protected
    procedure Loaded; override;
  public
    FindFirst: boolean;
    property CritModified: Boolean read FCritModified write SetCritModified;
  end;

implementation

uses
{$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  CPTypeCons,

{$ENDIF MODENT1}
  ParamDBG,
{$IFNDEF EAGLCLIENT}
  PrintDBG,
{$ENDIF}
  UtilPGI,
  uLibWindows // TestJocker FQ 20381 BVE 28.05.07
  ;

{$R *.DFM}

procedure LanceLettrage;
var
  X: TFMulLettre;
  PP: THPanel;
begin
  if _Blocage(['nrCloture'], False, 'nrAucun') then Exit;
  PP := FindInsidePanel;
  X := TFMulLettre.Create(Application);
  if PP = nil then
  begin
    try
      X.Qui := prAucun;
      X.ShowModal;
    finally
      X.Free;
    end;
    Screen.Cursor := SyncrDefault;
  end else
  begin
    InitInside(X, PP);
    X.Qui := prAucun;
    X.Show;
  end;
end;


procedure LanceLettrageMP(Qui: tProfilTraitement);
var X: TFMulLettre;
  PP: THPanel;
begin
  if _Blocage(['nrCloture'], False, 'nrAucun') then Exit;
  PP := FindInsidePanel;
  X := TFMulLettre.Create(Application);
  if PP = nil then
  begin
    try
      X.Qui := Qui;
      X.ShowModal;
    finally
      X.Free;
    end;
    Screen.Cursor := SyncrDefault;
  end else
  begin
    InitInside(X, PP);
    X.Qui := Qui;
    X.Show;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : redimensionnement de la grille
Mots clefs ... :
*****************************************************************}

procedure TFMulLettre.FormShow(Sender: TObject);
begin
    // proprieté de la grille
  GL.ListeParam := 'MULLETEUR';
    // nb D
  GL.ColColors[4] := clBlue;
    // debit
  GL.ColColors[5] := clBlue;
    // nb c
  GL.ColColors[6] := clRed;
    // credit
  GL.ColColors[7] := clRed;
    // sold
  GL.ColColors[8] := clGreen;

  // Libellé devise
  Distinguer.Caption := Distinguer.Caption+' '+GetColonneSQL('DEVISE','D_LIBELLE','D_DEVISE="'+V_PGI.DevisePivot+'"');
  InitCriteres;
  InitTablesLibresTiers(PLibres);
  E_DATEECHEANCE.Text := StDate1900;
  E_DATEECHEANCE_.Text := StDate2099;
  PositionneEtabUser(E_ETABLISSEMENT, False); // 15086
{$IFNDEF CCS3}
  if (not VH^.OuiTP) then
{$ENDIF}
    CFacNotPay.Visible := false;

  if FFiltreByUser <> nil then
    FFiltreByUser.LoadDB( 'LETTRAGE' );
    
  if (FFiltres.Text = 'DEFAUT') then BChercherClick(nil)
  else ChargeAVide;

  HMTrad.ResizeGridColumns(GL);
  if E_AUXILIAIRE.CanFocus then E_AUXILIAIRE.SetFocus;

{$IFDEF CCSTD}
  FFiltreByUser.ForceAccessibilite := faPublic;
{$ENDIF}
  Pages.ActivePageIndex := 0;
  Loading := False;
end;

procedure TFMulLettre.BCREERClick(Sender: TObject);
begin
  FFiltreByUser.Creer;
  FFiltreByUser.Save;
end;

procedure TFMulLettre.BENREGISTREClick(Sender: TObject);
begin
  FFiltreByUser.Save;
  CritModified := False;
end;

procedure TFMulLettre.BDUPLIQUERClick(Sender: TObject);
begin
  FFiltreByUser.Duplicate;
  FFiltreByUser.Save;
end;

procedure TFMulLettre.BSUPPRIMERClick(Sender: TObject);
begin
  VideFiltre(FFiltres, Pages) ;
  FFiltreByUser.Delete;
  FFiltreByUser.Save;
end;

procedure TFMulLettre.BRENOMMERClick(Sender: TObject);
begin
  FFiltreByUser.Rename;
  FFiltreByUser.Save;
end;

procedure TFMulLettre.BRECHERCHERClick(Sender: TObject);
begin
  VideFiltre(FFiltres, Pages);
  InitCriteres;
end;

procedure TFMulLettre.FFiltresChange(Sender: TObject);
begin
  FFiltreByUser.Select(FFiltres.Value);
  ControleFeuVert;
end;

function TFMulLettre.WhereCrit(AvecTiers: boolean): string;
var j, i: integer;
  P: TTabSheet;
  Nam, StM, StC, OldWhere: string;
  StCli, StEcr : HString;
  C: TControl;
  QT: THQuery;
begin
  OldWhere := XX_WHERETP.Text;
  if (not AvecTiers) then XX_WHERETP.Text := '';
  StCli := ''; StEcr := ''; StM := ''; StC := '';
  QT := THQuery.Create(Self);
  for j := 0 to Pages.PageCount - 1 do
  begin
    P := Pages.Pages[j];
    for i := 0 to P.ControlCount - 1 do
    begin
      C := P.Controls[i];
      if (C.Tag = 1) then continue;
      Nam := C.Name;
      if (Nam[Length(Nam)] = '_') then System.Delete(Nam, Length(Nam), 1);
            // recup des control pour en faire des critere
      if Copy(Nam, 1, 2) = 'T_' then
      begin
        if (AvecTiers) then QT.Control2Criteres(Nam, StCli, C, P);
      end
      else QT.Control2Criteres(Nam, StEcr, C, P);
    end;
  end;

  if (Copy(StEcr, 2, 3) = 'AND') then System.Delete(StEcr, 1, 5);
  if (Copy(StEcr, 2, 2) = 'OR') then System.Delete(StEcr, 1, 4);
  if (Copy(StCli, 2, 3) = 'AND') then System.Delete(StCli, 1, 5);
  if (Copy(StCli, 2, 2) = 'OR') then System.Delete(StCli, 1, 4);

  if ((Valeur(TOTAL.Text) > -99999999.0) and (TOTAL.Text <> '')) or ((Valeur(TOTAL_.Text) < 99999999.0) and (TOTAL_.Text <> '')) then StM := ' AND ' + Fourchette('(E_DEBIT+E_CREDIT)', TOTAL.Text, TOTAL_.Text, ftFloat);

  if Trim(StEcr) <> '' then
  begin
    if LettrePartiel.Checked then StEcr := ' AND ((' + StEcr + ' AND E_ETATLETTRAGE="AL") OR (E_ETATLETTRAGE="PL")) '
    else StEcr := ' AND ' + StEcr + ' AND E_ETATLETTRAGE="AL" ';
  end;

  StC := Fourchette('UPPER(E_LETTRAGE)', LETTRAGE.Text, LETTRAGE_.Text, ftString);
  if (STC <> '') then StC := ' AND ' + StC;
  if StCli <> '' then StCli := ' AND ' + StCli;
  Result := StEcr + StM + StC + StCli;
  XX_WHERETP.Text := OldWhere;
  FreeAndNil(QT);
end;

function TFMulLettre.SelectQL: string;
begin
  Result := 'Select E_GENERAL, E_AUXILIAIRE, Max(T_LIBELLE), '
    + 'Sum(' + DSurDC + ') DSURDC, Sum(E_DEBIT), '
    + 'Sum(' + CSurDC + ') CSURDC, Sum(E_CREDIT), Sum(E_DEBIT-E_CREDIT), Max(E_DEVISE) '
    + 'From Ecriture Left Outer Join Tiers On E_AUXILIAIRE=T_AUXILIAIRE ';
  if Qui <> prAucun then Result := Result + ' Left Outer Join GENERAUX On E_GENERAL=G_GENERAL ';
end;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 04/06/2007
Modifié le ... : 24/08/2007
Description .. : - LG - FB 19757 - LWhereCompte2 est remplace apr 
Suite ........ : LWhereCompte
Suite ........ : - LG - 24/08/2007 - FB 15056 - ajout d'un order by
Mots clefs ... : 
*****************************************************************}
procedure TFMulLettre.RechercheEcritures;
var
  Query: string;
begin
  Screen.Cursor := crHourglass;

  Ferme(QL);

  Query := SelectQL + 'WHERE ' + LWhereBase(true, true, false, false) + LWhereComptes(Self) + WhereCrit(true);
  if (MonProfilOk(Qui)) then Query := Query + ' AND ' + WhereProfilUser(QL, Qui);
  Query := Query + ' GROUP BY E_GENERAL, E_AUXILIAIRE, E_DEVISE';
  Query := Query + ' ORDER BY E_GENERAL, E_AUXILIAIRE ' ;

  QL := OpenSQL(Query, true);

  Screen.Cursor := SyncrDefault;
//LG* 18/06/2002
  HMTrad.ResizeGridColumns(GL);
end;

procedure TFMulLettre.ChargeAVide;
var
  Query: string;
begin
  Ferme(QL);

  Query := SelectQL + 'WHERE ' + LWhereVide + ' GROUP BY E_GENERAL, E_AUXILIAIRE';

  QL := OpenSQL(Query, true);
end;

procedure TFMulLettre.BChercherClick(Sender: TObject);
{$IFNDEF EAGLCLIENT}
var
  tmpTOB: TOB;
{$ENDIF}
begin
  if ((Sender <> nil) and (StrToDate(E_DATECOMPTABLE.Text) > StrToDate(E_DATECOMPTABLE_.Text))) then
  begin
    HMessMulL.Execute(4, '', '');
    exit;
  end;
  RechercheEcritures;
{$IFDEF EAGLCLIENT}
  QL.PutGridDetail(GL, false, false, '', true);
{$ELSE}
  tmpTOB := TOB.Create('', nil, -1);
  tmpTOB.LoadDetailDB('', '', '', QL, false, false);
  tmpTOB.PutGridDetail(GL, false, false, '', true);

  FreeAndNil(tmpTOB);
{$ENDIF}
  HMTrad.ResizeGridColumns(GL);
  CritModified := False;
end;

procedure TFMulLettre.E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin
  ParamDate(Self, Sender, Key);
end;

procedure TFMulLettre.FormCreate(Sender: TObject);
begin
  Loading := TRUE;
  PopUpMenu := ADDMenuPop(PopUpMenu, '', '');
  WMinX := Width;
  WMinY := 242;

  // Nouvelle Gestion des Filtres ( Création du TlisteByUSer )
  FFiltreByUser := TListByUser.Create(FFiltres, BFiltre, toFiltre, False);
  with FFiltreByUser do
  begin
    OnSelect  := InitSelectFiltre;
    OnInitGet := InitGetFiltre;
    OnInitAdd := InitAddFiltre;
    OnUpgrade := UpgradeFiltre;
    OnParams  := ParseParamsFiltre;
  end;

end;

procedure TFMulLettre.InitCriteres;
begin
  if VH^.Precedent.Code <> '' then E_DATECOMPTABLE.Text := DateToStr(VH^.Precedent.Deb)
  else E_DATECOMPTABLE.Text := DateToStr(VH^.Encours.Deb);
  E_DATECOMPTABLE_.Text := DateToStr(V_PGI.DateEntree);
end;

procedure TFMulLettre.BFermeClick(Sender: TObject);
begin
  if (IsInside(Self)) then CloseInsidePanel(self);
  close;
end;

procedure TFMulLettre.BRechercheClick(Sender: TObject);
begin
  FindFirst := true;
  FindLettre.Execute;
end;

procedure TFMulLettre.FindLettreFind(Sender: TObject);
begin
  Rechercher(GL, FindLettre, FindFirst);
end;

procedure TFMulLettre.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  Ferme(QL);
  FreeAndNil(FFiltreByUser);
{$IFNDEF EAGLCLIENT}
  if (Parent is THPanel) then Action := caFree;
{$ENDIF}
end;

procedure TFMulLettre.E_EXERCICEChange(Sender: TObject);
begin
  if (E_EXERCICE.ItemIndex > 0) then ExoToDates(E_EXERCICE.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_)
  else
  begin
    E_DATECOMPTABLE.Text := StDate1900;
    E_DATECOMPTABLE_.Text := StDate2099;
  end;
end;

procedure TFMulLettre.AttribCollectif;
var Q: TQuery;
  DEV: String3;
begin
  if ((E_AUXILIAIRE_.Text <> '') or (E_GENERAL.Text <> '') or (E_GENERAL_.Text <> '')) then Exit;
  Q := OpenSQL('SELECT T_COLLECTIF, T_DEVISE  FROM TIERS WHERE T_AUXILIAIRE="' + E_AUXILIAIRE.Text + '"', true);
  DEV := E_DEVISE.Value;
  if not Q.EOF then
  begin
    E_GENERAL.Text := Q.Fields[0].AsString;
    DEV := Q.Fields[1].AsString;
  end;
  Ferme(Q);
  ControleFeuVert;
  if BFeuVert.Visible then
  begin
    if ((DEV <> '') and (DEV <> E_DEVISE.Value)) then E_DEVISE.Value := DEV;
  end;
end;

procedure TFMulLettre.ControleFeuVert;
var OkOk: boolean;
begin
  OkOk := True;
  if ((E_GENERAL_.Text <> '') or (E_AUXILIAIRE_.Text <> '')) then Okok := False;
  if ((E_GENERAL.Text = '') or (E_AUXILIAIRE.Text = '')) then Okok := False;

  if (BFeuVert.Visible <> Okok) then BFeuVert.Visible := Okok;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 27/06/2006
Modifié le ... :   /  /
Description .. : - LG - FB 15017 - gestion des comptes confidentiels
Mots clefs ... :
*****************************************************************}
procedure TFMulLettre.E_GENERALExit(Sender: TObject);
begin
 if EstSQLConfidentiel('GENERAUX', THCritMaskEdit(Sender).Text) then
  begin
    HMessMulL.Execute(1, '', '');
    THCritMaskEdit(Sender).Text := '' ;
    if THCritMaskEdit(Sender).Canfocus then
     THCritMaskEdit(Sender).SetFocus ;
    Exit;
  end;
  { FQ 20381 BVE 28.05.07 }
  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if TestJoker(THEdit(Sender).Text) then Exit;

  // Complétion auto du numéro de compte si possible
  if ExisteSql('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + BourreEtLess(THEdit(Sender).Text, fbGene ) + '" AND (G_LETTRABLE="X" OR G_COLLECTIF="X")') then
  begin
     if Length(THEdit(Sender).Text) < VH^.Cpta[fbGene].Lg then
        THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbGene);
  end
  else
     THEdit(Sender).ElipsisClick(nil);
  { END FQ 20381 }

  ControleFeuVert;
end;

function TFMulLettre.EtudieModeLettrage(var R: RLETTR): TModalResult;
var ii: integer;
begin
  result := mrYes;
  if R.DeviseMvt = V_PGI.DevisePivot then
  begin
   {Paquet en Franc ou Euro}
    if R.Distinguer then
    begin
      R.CritDev := V_PGI.DevisePivot;
    end;
  end else
  begin
    if R.DeviseMvt = R.CritDev then
    begin
      {Paquet en devise --> Lettrage devise}
      R.LettrageDevise := True;
    end else
    begin
      {Paquet en devise (non explicite) --> poser la question}
      ii := HMessMulL.Execute(0, '', ' ' + RechDom('ttDevise', R.DeviseMvt, False)); if ii = mrCancel then Exit;
      if (ii = mrCancel) then
      begin
        result := mrCancel;
        exit;
      end;

      R.LettrageDevise := (ii = mrYes);
      if (R.LettrageDevise) then R.CritDev := R.DeviseMvt
      else R.Distinguer := false;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 16/06/2004
Modifié le ... :   /  /    
Description .. : - LG - 16/06/2004 - FB 10647 - ON passe la requet a la 
Suite ........ : fenetre de lettrage, pour ooiuvoir changer de compte 
Suite ........ : directement ds celle ci
Mots clefs ... : 
*****************************************************************}
procedure TFMulLettre.GLDblClick(Sender: TObject);
var R: RLETTR;
begin
  if ((QL.EOF) and (QL.BOF)) then Exit;
  FillChar(R, Sizeof(R), #0);
  R.General := GL.Cells[1, GL.Row];
  if EstSQLConfidentiel('GENERAUX', R.General) then begin
    HMessMulL.Execute(1, '', '');
    Exit;
  end;
  R.Auxiliaire := GL.Cells[2, GL.Row];
  R.Appel := tlMenu;
  R.CritDev := E_DEVISE.Value;
  R.GL := nil; //GL; // Inutiles
  R.CritMvt := WhereCrit(False);
  R.DeviseMvt := GL.Cells[9, GL.Row];
  EtudieModeLettrage(R);
{$IFDEF CCMP}
  LettrageManuelMP(R, True, taModif, Qui);
{$ELSE}
  LettrageManuel(R, True, taModif,QL);
{$ENDIF}
end;

procedure TFMulLettre.BFeuVertClick(Sender: TObject);
var Q: TQuery;
  Existe, Coll: Boolean;
  R: RLETTR;
begin
  Screen.Cursor := crHourglass;
  Coll := false;

    // controle auxi
  Q := OpenSQL('SELECT T_AUXILIAIRE, T_CONFIDENTIEL FROM TIERS WHERE T_AUXILIAIRE="' + E_AUXILIAIRE.Text + '"', true);
  if not Q.EOF then
  begin
    if (EstConfidentiel(Q.Fields[1].AsString)) then Existe := false
    else Existe := true;
  end
  else Existe := false;
  Ferme(Q);

  if (not Existe) then
  begin
    Screen.Cursor := SyncrDefault;
    HMessMulL.Execute(2, '', '');
    E_AUXILIAIRE.SetFocus;
    exit;
  end;

    // controle gene
  Q := OpenSQL('SELECT G_COLLECTIF, G_CONFIDENTIEL FROM GENERAUX WHERE G_GENERAL="' + E_GENERAL.Text + '"', true);
  if not Q.EOF then
  begin
    Coll := (Q.Fields[0].AsString = 'X');
    if (EstConfidentiel(Q.Fields[1].AsString)) then Existe := false
    else Existe := true;
  end
  else Existe := false;
  Ferme(Q);

  if (not Existe) then
  begin
    Screen.Cursor := SyncrDefault;
    HMessMulL.Execute(1, '', '');
    E_GENERAL.SetFocus;
    exit;
  end;
  if (not Coll) then
  begin
    Screen.Cursor := SyncrDefault;
    HMessMulL.Execute(3, '', '');
    E_GENERAL.SetFocus;
    exit;
  end;

  E_GENERAL_.Text := E_GENERAL.Text;
  E_AUXILIAIRE_.Text := E_AUXILIAIRE.Text;

    // constitution lettrage
  FillChar(R, Sizeof(R), #0);
  R.General := E_GENERAL.Text;
  R.Auxiliaire := E_AUXILIAIRE.Text;
  R.Appel := tlMenu;
  R.CritDev := E_DEVISE.Value;
  R.DeviseMvt := V_PGI.DevisePivot;
  R.GL := nil;
  R.CritMvt := WhereCrit(false);

  Screen.Cursor := SyncrDefault;

  if (EtudieModeLettrage(R) <> mrcancel) then
  begin
{$IFDEF CCMP}
  LettrageManuelMP(R, True, taModif, Qui);
{$ELSE}
  LettrageManuel(R, True, taModif);
{$ENDIF}
  end;
  Screen.Cursor := crDefault;
  E_AUXILIAIRE_.Text := '';
  E_GENERAL_.Text := '';
end;

procedure TFMulLettre.NextAuxi(Suiv: boolean);
var Q: TQuery;
  StConf: string;
begin
  if ((E_GENERAL_.Text <> '') or (E_AUXILIAIRE_.Text <> '') or (E_AUXILIAIRE.Text = '')) then exit;

{$IFDEF EAGLCLIENT}
    // A FAIRE
{$ELSE}
  StConf := SQLConf('TIERS');
  if (StConf <> '') then StConf := ' AND ' + StConf;
{$ENDIF}
  if (Suiv) then Q := OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE>"' + E_AUXILIAIRE.Text + '" ' + StConf + ' ORDER BY T_AUXILIAIRE', true)
  else Q := OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE<"' + E_AUXILIAIRE.Text + '" ' + StConf + ' ORDER BY T_AUXILIAIRE DESC', true);

  if (not Q.EOF) then
  begin
    E_AUXILIAIRE.Text := Q.Fields[0].AsString;
    AttribCollectif;
  end;

  Ferme(Q);
end;

procedure TFMulLettre.NextGene(Suiv: boolean);
var Q: TQuery;
  Existe: boolean;
  Nat: String3;
  StConf: string;
begin
  if (E_GENERAL.Text = '') then exit;

  Q := OpenSQL('SELECT G_NATUREGENE FROM GENERAUX WHERE G_GENERAL="' + E_GENERAL.Text + '"', true);
  Existe := (not Q.EOF);
  if (Existe) then
  begin
    Nat := Q.Fields[0].AsString;
    Ferme(Q);
  end
  else
  begin
    Ferme(Q);
    exit;
  end;

{$IFDEF EAGLCLIENT}
    // A FAIRE
{$ELSE}
  StConf := SQLConf('GENERAUX');
  if (StConf <> '') then StConf := ' AND ' + StConf;
{$ENDIF}
  if (Suiv) then Q := OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL>"' + E_GENERAL.Text + '" AND G_NATUREGENE="' + Nat + '" ' + StConf + ' ORDER BY G_GENERAL', true)
  else Q := OpenSQL('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL<"' + E_GENERAL.Text + '" AND G_NATUREGENE="' + Nat + '" ' + StConf + ' ORDER BY G_GENERAL DESC', true);

  if (not Q.EOF) then E_GENERAL.Text := Q.Fields[0].AsString;
  Ferme(Q);

  ControleFeuVert;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/09/2007
Modifié le ... :   /  /    
Description .. : - LG - FB 21411 - on traite le F que sur la grille
Mots clefs ... : 
*****************************************************************}
procedure TFMulLettre.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var S: Boolean;
begin
  S := (Shift = [ssShift]);
  case Key of
    VK_F3:
      begin
        Key := 0;
        if (S) then NextGene(false)
        else NextAuxi(false);
      end;
    VK_F4:
      begin
        Key := 0;
        if (S) then NextGene(true)
        else NextAuxi(true);
      end;
    // FQ 12990 BVE 04.06.07 - FB 21411 - LG
    VK_F5:
      if GL.Focused then
      begin
        Key := 0;
        BValiderClick(nil);
      end;
    70:      
      // CTRL + F
      begin
        if Shift = [ssCtrl] then
           BRechercheClick(nil); 
      end;
    { END FQ 12990 }
    VK_F9:
      begin
        Key := 0;
        BChercherClick(nil);
      end;
    VK_F10:
      begin
        Key := 0;
        BValiderClick(nil);
      end;
  end;
end;

procedure TFMulLettre.BAgrandirClick(Sender: TObject);
begin
  ChangeListeCrit(Self, true);
end;

procedure TFMulLettre.BReduireClick(Sender: TObject);
begin
  ChangeListeCrit(Self, false);
end;

procedure TFMulLettre.BImprimerClick(Sender: TObject);
begin
  if CritModified then
  begin
    PgiInfo('Vous n''avez pas appliqué les critères !');
    exit;
  end;
{$IFDEF EAGLCLIENT}
  // A FAIRE
{$ELSE}
  PrintGrid([GL],Caption);
{$ENDIF}
end;

procedure TFMulLettre.BValiderClick(Sender: TObject);
begin
  GLDblClick(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 27/06/2006
Modifié le ... : 03/09/2007
Description .. : - LG - FB 15017 - gestion des comptes confidentiels
Suite ........ : - LG - FB 15017 - on test le compte confidentiel apres le 
Suite ........ : bourrage
Mots clefs ... : 
*****************************************************************}
procedure TFMulLettre.E_AUXILIAIREExit(Sender: TObject);
begin

  { FQ 20381 BVE 28.05.07 }
  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if TestJoker(THEdit(Sender).Text) then Exit;

  // Complétion auto du numéro de compte si possible
  if ExisteSql('SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE = "' + BourreEtLess(THEdit(Sender).Text, fbAux) + '"') then
  begin
     if Length(THEdit(Sender).Text) < VH^.Cpta[fbAux].Lg then
        THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbAux);
  end
  else
     THEdit(Sender).ElipsisClick(nil);
  { END FQ 20381 }


 if EstSQLConfidentiel('TIERS', THCritMaskEdit(Sender).Text) then
  begin
    HMessMulL.Execute(1, '', '');
    THCritMaskEdit(Sender).Text := '' ;
    if THCritMaskEdit(Sender).Canfocus then
     THCritMaskEdit(Sender).SetFocus ;
    Exit;
  end;

  //AttribCollectif;
  //ControleFeuVert;
end;

procedure TFMulLettre.WMGetMinMaxInfo(var MSG: Tmessage);
begin
  with (PMinMaxInfo(MSG.lparam)^.ptMinTrackSize) do
  begin
    X := WMinX;
    Y := WMinY;
  end;
end;

procedure TFMulLettre.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFMulLettre.POPFPopup(Sender: TObject);
begin
  UpdatePopFiltre( BEnregistre, BSupprimer, BRenommer, FFiltres) ;
end;

procedure TFMulLettre.CFacNotPayClick(Sender: TObject);
begin
  if (not VH^.OuiTP) then
  begin
    XX_WHERETP.Text := '';
    exit;
  end;

  case CFacNotPay.State of
    cbChecked: XX_WHERETP.Text := 'E_TIERSPAYEUR="" OR E_TIERSPAYEUR IS NULL';
    cbUnChecked: XX_WHERETP.Text := 'E_TIERSPAYEUR<>""';
    cbGrayed: XX_WHERETP.Text := '';
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 01/12/2004
Modifié le ... :   /  /
Description .. : - LG - 01/12/2004 - FB 15017 - gestion des confidentiels
Mots clefs ... : 
*****************************************************************}
procedure TFMulLettre.E_GENERALElipsisClick(Sender: TObject);
begin
  LookUpList(TControl(Sender), 'Recherche d''un compte général', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', '( G_LETTRABLE="X" OR G_COLLECTIF="X") AND ' + CGenereSQLConfidentiel('G') , 'G_GENERAL', true, -1);
  ControleFeuVert;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 01/12/2004
Modifié le ... :   /  /    
Description .. : - LG - 01/12/2004 - FB 15017 - gestion des confidentiels
Mots clefs ... : 
*****************************************************************}
procedure TFMulLettre.E_AUXILIAIRE2ElipsisClick(Sender: TObject);
begin
  if (LookUpList(TControl(Sender), 'Recherche d''un compte auxiliaire', 'TIERS', 'T_AUXILIAIRE', 'T_LIBELLE', CGenereSQLConfidentiel('T'), 'T_AUXILIAIRE', true, -1)) then AttribCollectif;
  ControleFeuVert;
end;

procedure TFMulLettre.CritereChange(Sender: TObject);
begin
  if Loading then exit;
  CritModified := true;
end;

procedure TFMulLettre.Loaded;
var
  i: Integer;
begin
  inherited;
  if V_PGI.AglDesigning then exit;
  for i := 0 to ComponentCount - 1 do
  begin
    if (Components[i] is TControl) and (TControl(Components[i]).Parent is TTabSheet) then
      if (Components[i] is TControl) and TControl(Components[i]).Visible and (TControl(Components[i]).Enabled) then
      begin
        if (Components[i] is THCritMaskEdit) and (not Assigned(THCritMaskEdit(Components[i]).OnEnter)) then
          THCritMaskEdit(Components[i]).OnEnter := V_PGI.EgaliseOnEnter;

        if (Components[i] is TEdit) and not assigned(TEdit(Components[i]).OnChange) then TEdit(Components[i]).OnChange := CritereChange;
        if (Components[i] is THValComboBox) and not assigned(THValComboBox(Components[i]).OnClick) then THValComboBox(Components[i]).OnClick := CritereChange;
        if (Components[i] is TCheckBox) and not assigned(TCheckBox(Components[i]).OnClick) then TCheckBox(Components[i]).OnClick := CritereChange;
        if (Components[i] is THCritMaskEdit) and not assigned(THCritMaskEdit(Components[i]).OnChange) then THCritMaskEdit(Components[i]).OnChange := CritereChange;
        if (Components[i] is THRadioGroup) and not assigned(THRadioGroup(Components[i]).OnClick) then THRadioGroup(Components[i]).OnClick := CritereChange;
      end;
  end;
end;

procedure TFMulLettre.SetCritModified(Value: Boolean);
begin
  FCritModified := value;
  if Value then
    BChercher.Glyph := iCritGlyphModified.Picture.BitMap
  else
    BChercher.Glyph := iCritGlyph.Picture.BitMap;
end;


procedure TFMulLettre.InitAddFiltre(T: TOB);
var Lines : HTStrings ;
begin
  Lines := HTStringList.Create ;
  SauveCritMemoire(Lines, Pages) ;
  FFiltreByUser.AffecteTOBFiltreMemoire(T, Lines);
  Lines.Free ;
end;

procedure TFMulLettre.InitGetFiltre(T: TOB);
var Lines : HTStrings ;
begin
  Lines := HTStringList.Create ;
  SauveCritMemoire(Lines, Pages) ;
  FFiltreByUser.AffecteTOBFiltreMemoire(T, Lines);
  Lines.Free ;
end;

procedure TFMulLettre.InitSelectFiltre(T: TOB);
var Lines : HTStrings ;
    i : integer;
    stChamp, stVal : string;
begin
  if T = nil then exit;
  Lines := HTStringList.Create ;
  for i:=0 to T.Detail.Count-1 do
  begin
    stChamp := T.Detail[i].GetValue('N');
    stVal   := T.Detail[i].GetValue('V');
    Lines.Add(stChamp + ';' + stVal);
  end;
  VideFiltre (FFiltres, Pages, False);
  ChargeCritMemoire(Lines, Pages) ;
  Lines.Free ;
end;

procedure TFMulLettre.ParseParamsFiltre(Params: HTStrings);
var T : TOB;
begin
  FFiltreByUser.AddVersion;
  T := FFiltreByUser.Add;
  //en position 0 de Params se trouve le nom du filtre
  T.PutValue('NAME', XMLDecodeSt(Params[0])) ;
  T.PutValue('USER','---') ;
  Params.Delete(0);
  FFiltreByUser.AffecteTOBFiltreMemoire(T, Params);
end;

procedure TFMulLettre.UpgradeFiltre(T: TOB);
begin
  //
end;

end.

