unit CPMULLETTR;

//================================================================================
// Interface
//================================================================================
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
{$IFNDEF EAGLCLIENT}
    DB,
    DBTables,
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
    LettUtil,
    Lettrage,
    hmsgbox,
    HEnt1,
    Menus,
    Paramdat,
    UObjFiltres, //SG6 12/01/05 Gestion Filtres V6 FQ 15255
    HSysMenu,
    HTB97,
    HPanel,
    UiUtil,
    UTOB,
    LookUp
    ;

{$IFDEF EAGLCLIENT}
    Type TQuery = TOB;
{$ENDIF}

//==================================================
// Externe
//==================================================
procedure LanceLettrage;
{$IFDEF CCMP}
procedure LanceLettrageMP(Qui : tProfilTraitement);
{$ENDIF}

//==================================================
// Definition de class
//==================================================
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
    BFeuVert: TBitBtn;
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
    Distinguer: TCheckBox;
    iCritGlyph: TImage;
    iCritGlyphModified: TImage;
    E_AUXILIAIRE: THCritMaskEdit;
    E_AUXILIAIRE_: THCritMaskEdit;{JP 26/10/05 : FQ 16944 : manquait "_"}

        // evenement par default
        procedure FormCreate(Sender : TObject);
        procedure FormShow(Sender : TObject);
        procedure FormClose(Sender : TObject; var Action : TCloseAction);
        // keyDown
        procedure FormKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
        // click de la bar
        procedure BAgrandirClick(Sender : TObject);
        procedure BReduireClick(Sender : TObject);
        procedure BRechercheClick(Sender : TObject);
        procedure BImprimerClick(Sender : TObject);
        procedure BValiderClick(Sender : TObject);
        procedure BFermeClick(Sender : TObject);
        procedure BAideClick(Sender : TObject);
        // Click de la bar de recherche
        procedure BChercherClick(Sender : TObject);
        // autre click
        procedure BFeuVertClick(Sender : TObject);
        procedure CFacNotPayClick(Sender : TObject);
        // dbl click
        procedure GLDblClick(Sender : TObject);
        // elipsis Click
        procedure E_GENERALElipsisClick(Sender : TObject);
        procedure E_AUXILIAIRE2ElipsisClick(Sender : TObject);
        // change
        procedure E_EXERCICEChange(Sender : TObject);
        // exit
        procedure E_GENERALExit(Sender : TObject);
        procedure E_AUXILIAIRE2Exit(Sender : TObject);
        // autre evenement
        procedure E_DATECOMPTABLEKeyPress(Sender : TObject ; var Key : Char);
        procedure FindLettreFind(Sender : TObject);
    private
        ObjFiltre : TObjFiltre; //SG6 12/01/05 Gestion Filtres V6 FQ 15255
        QL : TQuery;
        WMinX,WMinY : Integer ;
        Qui : tProfilTraitement ;
        Loading, FCritModified: boolean;

        // message
        procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
        // creation de la requete vide
        procedure ChargeAVide;
        // creation de la requete
        procedure RechercheEcritures;
        function  SelectQL : String;
        procedure InitCriteres;
        function  WhereCrit(AvecTiers : boolean) : String;
        // control
        procedure ControleFeuVert;

        procedure AttribCollectif;
        procedure NextAuxi(Suiv : boolean);
        procedure NextGene(Suiv : boolean);
        function  EtudieModeLettrage(var R : RLETTR) : TModalResult;
        procedure SetCritModified(Value: Boolean);
        procedure CritereChange(Sender: TObject);

  protected
    procedure Loaded; override;
  public
    FindFirst: boolean;
    property CritModified: Boolean read FCritModified write SetCritModified;
  end;

implementation

uses
    ParamDBG,
{$IFNDEF EAGLCLIENT}
    PrintDBG,
{$ENDIF}
    UtilPGI
    ;

{$R *.DFM}

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LanceLettrage;
var
    X  : TFMulLettre;
    PP : THPanel;
begin
    if (_Blocage(['nrCloture'],false,'nrAucun')) then exit;

    PP := FindInsidePanel;
    X := TFMulLettre.Create(Application);

    if (PP = nil) then
    begin
        try
            X.Qui := prAucun;
            X.ShowModal;
        finally
            X.Free;
        end;

        Screen.Cursor := SyncrDefault;
    end
    else
    begin
        InitInside(X,PP);
        X.Qui := prAucun;
        X.Show;
    end;
end;

{$IFDEF CCMP}
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LanceLettrageMP(Qui : tProfilTraitement);
var
    X  : TFMulLettre;
    PP : THPanel;
begin
    if (_Blocage(['nrCloture'],false,'nrAucun')) then exit;

    PP := FindInsidePanel;
    X := TFMulLettre.Create(Application);

    if (PP = nil) then
    begin
        try
            X.Qui := Qui;
            X.ShowModal;
        finally
            X.Free;
        end;

        Screen.Cursor := SyncrDefault;
    end
    else
    begin
        InitInside(X,PP);
        X.Qui := Qui;
        X.Show;
    end;
end;
{$ENDIF}

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

  InitCriteres;
  InitTablesLibresTiers(PLibres);
  E_DATEECHEANCE.Text := StDate1900;
  E_DATEECHEANCE_.Text := StDate2099;
  PositionneEtabUser(E_ETABLISSEMENT, False); // 15086
{$IFNDEF CCS3}
  if (not VH^.OuiTP) then
{$ENDIF}
    CFacNotPay.Visible := false;

    //SG6 12/01/05 Gestion Filtres V6 FQ 15255
    ObjFiltre.Charger;

    //YMO 16/11/2005 FQ17015
    if (FFiltres.Text = 'DEFAUT') then BChercherClick(nil)
    else ChargeAVide;

    HMTrad.ResizeGridColumns(GL);
    if E_AUXILIAIRE.CanFocus then E_AUXILIAIRE.SetFocus;
    Pages.ActivePageIndex := 0;
    Loading := False;
end;

//==================================================
// Evenements de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TFMulLettre.WhereCrit(AvecTiers : boolean) : String;
var
    j,i : integer;
    P : TTabSheet;
    StCli,StEcr,Nam,StM,StC,OldWhere : String;
    C : TControl;

    QT : THQuery;
begin
  OldWhere := XX_WHERETP.Text;
  if (not AvecTiers) then XX_WHERETP.Text := '';
    StCli := '';
    StEcr := '';
    StM := '';
    StC := '';

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

function TFMulLettre.SelectQL : String;
begin
   // result := 'SELECT E_GENERAL, E_AUXILIAIRE, MAX(T_LIBELLE), SUM(' + DSurDC + '),
   //SUM(E_DEBIT), SUM(' + CSurDC + '), SUM(E_CREDIT), SUM(E_DEBIT-E_CREDIT), MAX(E_DEVISE)
   //FROM ECRITURE LEFT OUTER JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE ';
  {JP 26/10/05 : FQ 16944 : reprise du code de MULLETTR}
  Result := 'Select E_GENERAL, E_AUXILIAIRE, Max(T_LIBELLE), '
    + 'Sum(' + DSurDC + '), Sum(E_DEBIT), '
    + 'Sum(' + CSurDC + '), Sum(E_CREDIT), Sum(E_DEBIT-E_CREDIT), Max(E_DEVISE) '
    + 'From Ecriture Left Outer Join Tiers On E_AUXILIAIRE=T_AUXILIAIRE ';
  if Qui <> prAucun then Result := Result + ' Left Outer Join GENERAUX On E_GENERAL=G_GENERAL ';
end ;

procedure TFMulLettre.RechercheEcritures;
var
    Query : String;
begin
  Screen.Cursor := crHourglass;

  Ferme(QL);
  {JP 26/10/05 : FQ 16944 : LWhereComptes2 et non LWhereComptes}
  Query := SelectQL + 'WHERE ' + LWhereBase(true, true, false, false) + LWhereComptes2(Self) + WhereCrit(true);
  if (MonProfilOk(Qui)) then Query := Query + ' AND ' + WhereProfilUser(QL, Qui);
  Query := Query + ' GROUP BY E_GENERAL, E_AUXILIAIRE, E_DEVISE';

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

procedure TFMulLettre.E_DATECOMPTABLEKeyPress(Sender : TObject ; var Key : Char);
begin
  ParamDate(Self, Sender, Key);
end;

procedure TFMulLettre.FormCreate(Sender : TObject);
var
  Composants : TControlFiltre;
begin
    Composants.PopupF   := POPF;
    Composants.Filtres  := FFILTRES;
    Composants.Filtre   := BFILTRE;
    Composants.PageCtrl := Pages;
    ObjFiltre := TObjFiltre.Create(Composants,'LETTRAGE');

    Loading := TRUE;
    PopUpMenu := ADDMenuPop(PopUpMenu,'','');
    WMinX := Width;
    WMinY := 242;
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

procedure TFMulLettre.FormClose(Sender : TObject ; var Action : TCloseAction);
begin
    //SG6 12/01/05 Gestion Filtres V6 FQ 15255
    FreeAndnil(ObjFiltre);
    Ferme(QL);
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
var
    Q : TQuery;
    DEV : String3;
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
var
    OkOk : boolean;
begin
  OkOk := True;
  if ((E_GENERAL_.Text <> '') or (E_AUXILIAIRE_.Text <> '')) then Okok := False;
  if ((E_GENERAL.Text = '') or (E_AUXILIAIRE.Text = '')) then Okok := False;

  if (BFeuVert.Visible <> Okok) then BFeuVert.Visible := Okok;
end;

procedure TFMulLettre.E_GENERALExit(Sender: TObject);
begin
  ControleFeuVert;
end;

function  TFMulLettre.EtudieModeLettrage(var R : RLETTR) : TModalResult;
var
    ii : integer;
begin
    result := mrYes;

    R.LettrageDevise := false;

    if (not (R.DeviseMvt = V_PGI.DevisePivot)) then
    {begin // Paquet en Franc ou Euro
        if (R.Distinguer) then
        begin
            R.CritDev := V_PGI.DevisePivot;
        end;
    end
    else}
    begin
        if (R.DeviseMvt = R.CritDev) then R.LettrageDevise := true // Paquet en devise --> Lettrage devise
        else
        begin //Paquet en devise (non explicite) --> poser la question
            ii := HMessMulL.Execute(0,'',' ' + RechDom('ttDevise',R.DeviseMvt,false));
            if (ii = mrCancel) then
            begin
                result := mrCancel;
                exit;
            end;

            R.LettrageDevise := (ii = mrYes);
            if (R.LettrageDevise) then R.CritDev := R.DeviseMvt
            {else R.Distinguer := false};
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulLettre.GLDblClick(Sender : TObject);
var
    R  : RLETTR ;
begin

    if ((QL.EOF) and (QL.BOF)) then exit;

    FillChar(R,Sizeof(R),#0);
    R.General := GL.Cells[1,GL.Row];
    if (EstSQLConfidentiel('GENERAUX',R.General)) then
    begin
        HMessMulL.Execute(1,'','');
        exit;
    end;
    R.Auxiliaire := GL.Cells[2,GL.Row];
    R.Appel := tlMenu;
    R.CritDev := E_DEVISE.Value;
    R.CritMvt := WhereCrit(false);
    R.DeviseMvt := GL.Cells[9,GL.Row];
//    R.Distinguer := Distinguer.Checked;

    if (EtudieModeLettrage(R) <> mrcancel) then
    begin
{$IFDEF CCMP}
        LettrageManuelMP(R,true,taModif,Qui);
{$ELSE}
        LettrageManuel(R,true,taModif);
{$ENDIF}
    end;
end;

procedure TFMulLettre.BFeuVertClick(Sender: TObject);
var
  Q : TQuery;
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

procedure TFMulLettre.NextAuxi(Suiv : boolean);
var
    Q : TQuery;
    StConf : String;
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
var
    Q : TQuery;
    Existe : boolean;
    Nat    : String3;
    StConf : String;
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

procedure TFMulLettre.FormKeyDown(Sender : TObject ; var Key : Word ; Shift : TShiftState);
var
    S : Boolean;
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

procedure TFMulLettre.E_AUXILIAIRE2Exit(Sender : TObject);
begin
  AttribCollectif;
  ControleFeuVert;
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
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulLettre.E_GENERALElipsisClick(Sender : TObject);
begin
{$IFDEF EAGLCLIENT}
    LookUpList(TControl(Sender),'Recherche d''un compte général','GENERAUX','G_GENERAL','G_LIBELLE','G_LETTRABLE="X" OR G_COLLECTIF="X"','G_GENERAL',true,-1);
{$ELSE}
    AppelGeneral(Sender,Cache);
{$ENDIF}
    ControleFeuVert;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFMulLettre.E_AUXILIAIRE2ElipsisClick(Sender : TObject);
begin
{$IFDEF EAGLCLIENT}
    if (LookUpList(TControl(Sender),'Recherche d''un compte auxiliaire','TIERS','T_AUXILIAIRE','T_LIBELLE','','T_AUXILIAIRE',true,-1)) then AttribCollectif;
{$ELSE}
    if (AppelAuxi(Sender,Cache)) then AttribCollectif;
{$ENDIF}
    ControleFeuVert;
end;

procedure TFMulLettre.CritereChange(Sender: TObject);
begin
  if Loading then exit;
  CritModified := true;
end;

procedure TFMulLettre.Loaded;
var
  i : Integer;
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

end.
