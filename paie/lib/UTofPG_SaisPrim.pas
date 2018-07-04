{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 21/05/2002
Modifié le ... :   /  /
Description .. : Unit de gestion de la saisie des primes
Mots clefs ... : PAIE;SAISPRIM
*****************************************************************}
{
---- PH 10/08/2005 Suppression directive de compil $IFDEF AGL550B ----
JL 27/04/2007 Modification chargement données
PT1   : 03/03/2008 FL V_8  EManager / L'ouverture de l'écran en mode
                           consultation vérouille la grille
PT2   : 22/04/2008 VG V_80 EManager / En mode consultation, historique de la
                           saisie des primes actif                           
}
unit UTofPG_SAISPRIM;

interface
uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
  Windows,
  StdCtrls,
  Controls,
  Classes,
  Graphics,
  forms,
  sysutils,
  Comctrls,
  HTB97,
  Grids,
  HCtrls,
  HEnt1,
  EntPaie,
  HMsgBox,
  HSysMenu,
  UTOF,
  UTOB,
  Vierge,
{$IFNDEF EAGLCLIENT}
  DBCtrls,
  HDB,
  HQry,
{$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
  PrintDbg,
  HQuickRP,
  QRGrid,
  FE_Main,
{$ELSE}
  HQry,
  MaineAGL,
  UtileAgl,
{$ENDIF}
  ParamSoc,
  uPaieRemunerations,
{$IFDEF EPRIMES}

  UTofPG_MulSaisRub,
{$ENDIF EPRIMES}
  AGLInit,
  uListByUSer,
  ed_tools;
type
  TOF_PG_SAISPRIM = class(TOF)
  private
    LesMontants: array[1..35] of double;
    LesDecimales: array[1..35] of integer;
    LesAReporter: array[1..35] of string;
    LesRub: array[1..35] of string;
    LesTyp: array[1..35] of string;
    CodeCal: array[1..35] of string;
    LesMasq: array[1..5] of string; // codes des masques traités
    LesDecal: array[1..5] of integer; // nbre de colonnes de décalage (décalge cumulé)
    NbCol: array[1..5] of integer; // nbre de colonnes ds chq masque
    PrimAffic: array[1..5] of Boolean; // colonnes elts de salaires affichables
    PgNbSalLib, NbDecal: Integer;
    LibColRem: array[1..5] of string;
    QMul: TQUERY; // Query recuperee du mul
    LeTitre, LeCode: string; // Code Salarie en traitement et Titre de la forme
    Modifier: Boolean;
    Ceg: Boolean;
    DateDebut, DateFin: TDateTime; // Date de debut exercice social
    DebutM1, FinM1: TDateTime; // Date de debut et date de fin du mois précédent
    LaGrille, LaGrille1: THGrid; // Grilles de saisie des elts et totalisation
{$IFDEF EPRIMES}
    LeMasque: TOB; // TOB du masque de la saisie par rubrique
    TSal: TOB; // TOB des salaries et des remunerations avec champ variable
{$ELSE EPRIMES}
    LeMasque, TOB_Masque: TOB; // TOB du masque de la saisie par rubrique
    TSal, LesRem: TOB; // TOB des salaries et des remunerations avec champ variable
    T_RemMere, T_RemFille: TOB; // TOb gestion historiques salariés = liste des remunerations contenues dans les masques
{$ENDIF EPRIMES}
    Ligne: Integer;
    AA, MM, JJ: WORD;
    SalAnn1: string;
    HMTrad: THSystemMenu;
    BCherche: TToolbarButton97;
    BoutonFiltre: TToolbarButton97;
    ComboFiltre: THValComboBox;
    PanelCriteres: TPageControl;
    bEffaceAvance: TToolbarButton97;
    LesFiltres: TListByUser;
    QZ: THQuery;
    FFicheName: string;
    procedure BNouvRechClick(Sender: TObject);
    procedure InitAddFiltre(T: TOB);
    procedure InitGetFiltre(T: TOB);
    procedure InitSelectFiltre(T: TOB);
    procedure ParseParamsFiltre(Params: HTStrings);
    procedure UpgradeFiltre(T: TOB);
    procedure BChercherClick(Sender: TObject);
    procedure bEffaceAvanceClick(Sender: TObject);
    procedure InitChargeChampSQL;
    procedure AfterShow;

    function OnSauve: boolean;
    procedure ValiderClick(Sender: TObject);
    procedure FermeClick(Sender: TObject);
    procedure AccesSal(Sender: TObject);
    procedure HistSal(Sender: TObject);
    procedure HistoSalarie(Sender: TObject);
    function NbreDecimale(NbD: Integer): string;
    function RendSaisieRub(Salarie: string; DateDebut, DateFin: TDateTime; Decalage, NbreCols: Integer; LeMasque: TOB; var libelle: string): Boolean;
    function RendSaisiePrec(Salarie: string; DateDebut, DateFin: TDateTime; Decalage, NbreCols: Integer; LeMasque: TOB): Boolean;
    procedure GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GrilleCopierColler;
    function IdentSalarie(Salarie: string): Integer;
    procedure ImpClik(Sender: TObject);
    procedure PostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GrilleCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GrilleRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure CalculTotaux;
    procedure IdentCol; // indentification des codes calcul et indice zonz elt de salaire
    procedure TotClik(Sender: TObject);
    procedure AideCalcul(Sender: TObject);
    procedure GrilleDblClick(Sender: TObject);
    procedure RempliGrille;
    procedure AfficheTitre;
    procedure ExporterGrille(Sender: TObject);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnClose; override;
  end;

implementation

uses P5Def, P5Util, Filtre, UtXML;

procedure TOF_PG_SAISPRIM.BChercherClick(Sender: TObject);
begin
  inherited;
  QZ.UpdateCriteres;
  //  PGIInfo(QZ.Criteres);
  RempliGrille;
  SetControlEnabled('BCherche', FALSE);
  HMTrad.ResizeGridColumns(THGrid(GetControl('GRILLE')));
  HMTrad.ResizeGridColumns(THGrid(GetControl('GRILLE1')));
end;

procedure TOF_PG_SAISPRIM.AfterShow;
begin
  exit;
  if ((HMTrad.ActiveResize) and (HMTrad.ResizeDBGrid) and (V_PGI.OutLook)) then
  begin
    HMTrad.ResizeGridColumns(THGrid(GetControl('GRILLE')));
    HMTrad.ResizeGridColumns(THGrid(GetControl('GRILLE1')));
  end;
end;


procedure TOF_PG_SAISPRIM.BNouvRechClick(Sender: TObject);
begin
  if V_PGI.AGLDesigning then exit;
  VideFiltre(ComboFiltre, PanelCriteres);
  LesFiltres.New;
end;

procedure TOF_PG_SAISPRIM.InitGetFiltre(T: TOB);
var
  Lines: HTStrings;
  //  Item : TMenuItem;
begin
  if V_PGI.AGLDesigning then exit;
  Lines := nil;
  try
    Lines := HTStringList.Create;
    SauveCritMemoire(Lines, PanelCriteres);
    LesFiltres.AffecteTOBFiltreMemoire(T, Lines);
  finally
    FreeAndNil(Lines);
  end;
end;

procedure TOF_PG_SAISPRIM.InitAddFiltre(T: TOB);
var
  Lines: HTStrings;
begin
  if V_PGI.AGLDesigning then exit;
  Lines := HTStringList.Create;
  SauveCritMemoire(Lines, PanelCriteres);
  LesFiltres.AffecteTOBFiltreMemoire(T, Lines);
  Lines.Free;
end;

procedure TOF_PG_SAISPRIM.InitSelectFiltre(T: TOB);
var
  Lines: HTStrings;
  i: integer;
  stChamp, stVal: string;
begin
  if V_PGI.AGLDesigning then exit;
  if T <> nil then else exit;
  Lines := HTStringList.Create;
  for i := 0 to T.Detail.Count - 1 do
  begin
    stChamp := T.Detail[i].GetValue('N');
    stVal := T.Detail[i].GetValue('V');
    Lines.Add(stChamp + ';' + stVal);
  end;
  VideFiltre(ComboFiltre, PanelCriteres, false);
  ChargeCritMemoire(Lines, PanelCriteres);
  Lines.Free;
end;

procedure TOF_PG_SAISPRIM.ParseParamsFiltre(Params: HTStrings);
var
  T: TOB;
begin
  if V_PGI.AGLDesigning then exit;
  LesFiltres.AddVersion;
  T := LesFiltres.Add;
  //en position 0 de Params se trouve le nom du filtre
  T.PutValue('NAME', XMLDecodeSt(Params[0]));
  Params.Delete(0);
  T.PutValue('USER', '---');
  LesFiltres.AffecteTOBFiltreMemoire(T, Params);
end;

procedure TOF_PG_SAISPRIM.UpgradeFiltre(T: TOB);
begin
  if V_PGI.AGLDesigning then exit;
end;

procedure TOF_PG_SAISPRIM.bEffaceAvanceClick(Sender: TObject);
begin
  THValComboBox(GetControl('Z_C1')).ItemIndex := -1;
  THValComboBox(GetControl('ZO1')).ItemIndex := -1;
  TEdit(GetControl('ZV1')).text := '';
  THValComboBox(GetControl('ZG1')).ItemIndex := -1;
  THValComboBox(GetControl('Z_C2')).ItemIndex := -1;
  THValComboBox(GetControl('ZO2')).ItemIndex := -1;
  TEdit(GetControl('ZV2')).text := '';
  THValComboBox(GetControl('ZG2')).ItemIndex := -1;
  THValComboBox(GetControl('Z_C3')).ItemIndex := -1;
  THValComboBox(GetControl('ZO3')).ItemIndex := -1;
  TEdit(GetControl('ZV3')).text := '';
end;

procedure TOF_PG_SAISPRIM.InitChargeChampSQL;
var
  StSQL, StTable, StPref, Fct: string;
  Typs, Vals, Libs: HTStringList;
  AllFields: boolean;

  procedure ParseFieldSelected;
  var
    i: integer;
  begin
    if AllFields then exit;
    i := 0;
    while i < Vals.Count do
    begin
      if (Pos(Vals[i], stSQL) = 0) or (Copy(Libs[i], 1, 2) = '.-') then
      begin
        Vals.Delete(i);
        Libs.Delete(i);
      end
      else Inc(i);
    end;
  end;

begin
  if V_PGI.AGLDesigning then exit;
  bEffaceAvanceClick(nil);
  StSQL := 'SELECT * FROM PGSALRESPONSVAR';
  AllFields := Copy(StSQL, 1, 14) = 'SELECT * FROM ';
  StTable := GetAllTableNamesFromSQL(StSQL, TRUE, FALSE);
  if V_PGI.ExtendedFieldSelection <> '' then Fct := V_PGI.ExtendedFieldSelection else fct := 'L';
  V_PGI.ExtendedFieldSelection := '';
  Typs := HTStringList.Create;
  Libs := HTStringList.Create;
  Vals := HTStringList.Create;
  while StTable <> '' do
  begin
    StPref := ReadTokenSt(StTable);
    if StPref = '' then continue;
    ExtractFields(StPref, Fct, Libs, Vals, Typs, FALSE);
    ParseFieldSelected;
    THValComboBox(GetControl('Z_C1')).Items.AddStrings(Libs);
    THValComboBox(GetControl('Z_C1')).Values.AddStrings(Vals);
  end;
  THValComboBox(GetControl('Z_C2')).Items.Assign(THValComboBox(GetControl('Z_C1')).Items);
  THValComboBox(GetControl('Z_C2')).Values.Assign(THValComboBox(GetControl('Z_C1')).Values);
  THValComboBox(GetControl('Z_C3')).Items.Assign(THValComboBox(GetControl('Z_C1')).Items);
  THValComboBox(GetControl('Z_C3')).Values.Assign(THValComboBox(GetControl('Z_C1')).Values);
  FreeAndNil(Typs);
  FreeAndNil(Libs);
  FreeAndNil(Vals);
  THValComboBox(GetControl('Z_C1')).Reload;
  THValComboBox(GetControl('Z_C2')).Reload;
  THValComboBox(GetControl('Z_C3')).Reload;
end;


procedure TOF_PG_SAISPRIM.OnArgument(Arguments: string);
var
  F: TFVierge;
  st, LeType, LaRubrique: string;
  Q: TQuery;
  NumMsq, Mois: string;
  NbD, i, j: Integer;
  Abandon: Boolean;
  T1, TS: TOB;
  LaRem, CodeC: string;
  NbC, zz, IndLign, k: Integer;
  BtnValid, BExport, BtnFerme, BtnImp, BtnSal, BtnHistSal, BtnHistoSal, BTnTotal, BtnAideCal: TToolbarButton97;
  HMTrad: THSystemMenu;
  LeWhere: string;
  Mt1, Mt: Double;
  //  NewValue,Table, Crit, tri, Champ, Titre, Larg, Align, Params, LeTitre, NomCol, Perso: string ;
  //  OkTri, OkNumCol: boolean ;
begin
  inherited;
  SetControlProperty('GRILLE', 'SortEnabled', True);
  TFVierge(Ecran).WindowState := wsMaximized;
  TFVierge(Ecran).Align := AlClient;
  if Ecran.Name = 'A' then
  begin
    BoutonFiltre := TToolbarButton97(GetControl('BFiltre'));
    ComboFiltre := THValComboBox(GetControl('FFiltres'));
    PanelCriteres := TPageControl(GetControl('Pages'));
    BCherche := TToolbarButton97(GetControl('Bcherche'));
    BCherche.OnClick := BChercherClick;
    bEffaceAvance := TToolbarButton97(GetControl('bEffaceAvance'));
    bEffaceAvance.OnClick := bEffaceAvanceClick;
    HMTrad := THSystemMenu.Create(TFVierge(Ecran));
    HMTrad.ActiveResize := true;
    HMTrad.ResizeDBGrid := true;

    LesFiltres := TListByUser.Create(ComboFiltre, BoutonFiltre, toFiltre, false {FiltreDisabled});
    LesFiltres.LoadDB(FFicheName);
    with LesFiltres do
    begin
      OnSelect := InitSelectFiltre;
      OnInitGet := InitGetFiltre;
      OnInitAdd := InitAddFiltre;
      OnUpgrade := UpgradeFiltre;
      OnParams := ParseParamsFiltre;
      //@@      OnItemNouveau := bNouvRechClick;
    end;

    QZ := THQuery.Create(Application);
    QZ.PageCriteres := PanelCriteres;
    InitChargeChampSQL;
    //@@    QZ.RechargerLaListe := false;
    TFVierge(Ecran).OnAfterFormShow := AfterShow;
  end;


  TSal := nil;
  // NewValue := 'PGSALRESPONSVAR' ;
  //  ChargeHListe (NewValue, Table, Crit, tri, Champ, Titre, Larg, Align,Params, LeTitre, NomCol, Perso, OkTri, OkNumCol);
  BtnImp := TToolbarButton97(GetControl('Bimprimer'));
  if BtnImp <> nil then BtnImp.OnClick := ImpClik;
  BTnTotal := TToolbarButton97(GetControl('BTNCALCUL'));
  if BTnTotal <> nil then BTnTotal.OnClick := TotClik;
  BExport := TToolBarButton97(GetControl('BEXPORT'));
  if BExport <> nil then BExport.OnClick := ExporterGrille;
  LaGrille := THGrid(Getcontrol('GRILLE')); //PT2
  if ConsultP then
  begin
    SetControlVisible('BTNAIDECALCUL', FALSE);
    SetControlVisible('BVALIDER', FALSE);
    //PT1 - Début réactivation
    {$IFDEF EPRIMES}
    // Désactivation en consultation
{PT2
    SetControlVisible('BTNHISTSAL', FALSE);
    SetControlEnabled ('GRILLE', FALSE);
}
    LaGrille.Options:= LaGrille.Options-[goEditing]+[goRowSelect];
//FIN PT2
    SetControlVisible('BTNSAL', FALSE);
    {$ENDIF}
    //PT1 - Fin
  end;
  BtnAideCal := TToolbarButton97(GetControl('BTNAIDECALCUL'));
  if BtnAideCal <> nil then BtnAideCal.OnClick := AideCalcul;
  if Arguments <> '' then
  begin
    st := Trim(Arguments);
    NumMsq := ReadTokenSt(st); // Recup des masques
    NumMsq := StringReplace(NumMsq, '+', ';', [rfReplaceAll]);
    DateDebut := StrToDate(ReadTokenSt(st)); // Recup de la date de debut
    DateFin := StrToDate(ReadTokenSt(st)); // Recup Date fin
    if st <> '' then LeWhere := ReadTokenSt(st)
    else LeWhere := '';
  end
  else
  begin
    NumMsq := '2';
    DateDebut := StrToDate('01/01/2004');
    Datefin := StrToDate('31/01/2004');
    LeWhere := 'WHERE (((PSA_SUSPENSIONPAIE <> "X") AND (PSA_DATEENTREE <="01/31/2004") AND (((PSA_DATESORTIE >="01/01/2004")) OR (PSA_DATESORTIE IS NULL) OR (PSA_DATESORTIE <= "01/01/1900")) AND PSE_RESPONSVAR = "0000000001"))';
    QMUL := OpenSQL('SELECT * FROM PGSALRESPONSVAR ' + LeWhere, TRUE);
  end;

  DebutM1 := DebutDeMois(PLUSMOIS(DateDebut, -1));
  FinM1 := FINDEMOIS(PLUSMOIS(DateFin, -1));
  DecodeDate(DateDebut, AA, MM, JJ);
  // recuperation de la query du multicritere
  if not (Ecran is TFVierge) then exit;
{PT2
  LaGrille := THGrid(Getcontrol('GRILLE'));
}
  AfficheTitre;
  LaGrille.Align := AlClient;
  SetcontrolVisible('GRILLE1', False);
  SetcontrolVisible('LECOMMENTAIRE', False);
  if LaGrille = nil then exit;
  F := TFVierge(Ecran);
  if (F <> nil) and (Ecran.Name <> 'A') then
  begin
{$IFDEF EAGLCLIENT}
    QMUL := THQuery(F.FMULQ).TQ;
{$ELSE}
    QMUL := F.FMULQ;
{$ENDIF}
  end;
  if QMUL = nil then exit;
  Mois := IntToStr(MM);
  if Length(Mois) = 1 then Mois := '0' + Mois;
  if ConsultP then Ecran.Caption := 'Consultation '
  else Ecran.Caption := 'Saisie ';
  Ecran.Caption := Ecran.Caption + VH_Paie.PGLibSaisPrim + ' du mois de ' + RechDom('PGMOIS', Mois, FALSE) + '/' + IntToStr(AA) + ' responsable : ' + RechDom('TTUTILISATEUR',
    V_PGI.User, FALSE);
  UpdateCaption(Ecran);
  //if ConsultP then LaGrille.Options := LaGrille.Options-[goEditing];
  DecodeDate(DebutM1, AA, MM, JJ);
  BtnValid := TToolbarButton97(GetControl('BVALIDER'));
  if BtnValid <> nil then BtnValid.OnClick := ValiderClick;
  BtnFerme := TToolbarButton97(GetControl('BFERME'));
  if BtnFerme <> nil then BtnFerme.OnClick := FermeClick;
  BtnSal := TToolbarButton97(GetControl('BTNSAL'));
  if BtnSal <> nil then Btnsal.OnClick := AccesSal;
  BtnHistSal := TToolbarButton97(GetControl('BTNHISTSAL'));
  if BtnHistSal <> nil then BtnHistsal.OnClick := HistSal;
  BtnHistoSal := TToolbarButton97(GetControl('BTNHISTOSAL'));
  if BtnHistoSal <> nil then BtnHistosal.OnClick := HistoSalarie;
{$IFDEF aucasou}
  if TOB_Rem <> nil then
  begin
    TOB_Rem.free;
    TOB_Rem := nil;
  end;
  TOB_Rem := TOB.Create('les remunerations', nil, -1);
  TOB_Rem.LoadDetailDB('REMUNERATION', '', '', nil, FALSE, False);
  TOB_Rem.detail.Sort('PRM_RUBRIQUE');
{$ENDIF}
{$IFNDEF EPRIMES}
  InitTOB_Rem();
{$ENDIF EPRIMES}
  // design des grille et definition des tailles des colonnes
{  for i := 3 to 37 do
  begin
    LaGrille.ColLengths[i] := 17;
  end;}
  LaGrille.ColLengths[1] := 1;
  for i := 0 to 37 do
  begin
    if i < 3 then LaGrille.ColAligns[i] := taLeftJustify
    else LaGrille.ColAligns[i] := taCenter;
  end;
  for i := 1 to 5 do
  begin
    LesMasq[i] := '';
    LesDecal[i] := 0;
    NbCol[i] := 0;
  end;
{$IFNDEF EPRIMES}
  TOB_Masque := TOB.Create('Les caractéristiques du masque', nil, -1);
  st := 'SELECT * FROM MASQUESAISRUB ORDER BY PMR_ORDRE';
  Q := OpenSql(st, TRUE);
  TOB_Masque.LoadDetailDB('MASQUESAISRUB', '', '', Q, FALSE, FALSE);
  Ferme(Q);
  T_RemMere := TOB.create('Les Rem des masques', nil, -1);
{$ENDIF EPRIMES}
  for i := 1 to 4 do
  begin
    LesMasq[i] := ReadTokenSt(NumMsq); // codes des masques traités
    if LesMasq[i] = '' then break;
    LeMasque := TOB_Masque.FindFirst(['PMR_ORDRE'], [LesMasq[i]], FALSE);
    // Les décalages sont pour le traitement des colonnes suivantes donc indice +1 par rapport au masque en cours de traitement
    if LeMasque <> nil then
    begin
      NbC := LeMasque.getvalue('PMR_NBRECOL');
      LesDecal[i + 1] := LesDecal[i] + NbC;
      NbCol[i + 1] := NbCol[i + 1] + NbC;
{$IFNDEF EPRIMES}
      for zz := 1 to NbC do
      begin // recherche de la liste des remunerations contenues dans les masques
        LaRem := LeMasque.GetValue('PMR_COL' + IntToStr(zz));
        if LaRem <> '' then
        begin
          T_RemFille := T_RemMere.FindFirst(['CODEBASE'], [LaRem], FALSE);
          if T_RemFille = nil then
          begin
            T_RemFille := TOB.create('Une Rem', T_RemMere, -1);
            T_RemFille.AddChampSup('CODEBASE', False);
            T_RemFille.AddChampSup('LIBELBASE', False);
            T_RemFille.PutValue('CODEBASE', LaRem);
            T_RemFille.PutValue('LIBELBASE', LeMasque.GetValue('PMR_LIBCOL' + IntToStr(zz)));
          end;
        end;
      end;
{$ENDIF EPRIMES}
    end;
  end;
  if LesMasq[1] = '' then
  begin
    PgiBox('Attention, aucun masque séléctionné', Ecran.caption);
    exit;
  end;
  // recup des remunerations ayant au moins un champ à saisir
  // pour connaitre les nbre de decimales à saisir dans la colonne en fonction du type de champ saisi
{$IFNDEF EPRIMES}
  LesRem := TOB.Create('Les Remunérations', nil, -1);
  Q := OpenSql('SELECT * FROM REMUNERATION WHERE ##PRM_PREDEFINI## PRM_TYPEBASE="00" OR PRM_TYPETAUX="00" OR PRM_TYPECOEFF="00" OR PRM_TYPEMONTANT="00"', TRUE);
  LesRem.LoadDetailDB('REMUNERATION', '', '', Q, FALSE, FALSE);
  Ferme(Q);
{$ENDIF EPRIMES}
  TSal := TOB.Create('Les Salaries', nil, -1);
  St := 'SELECT PSA_SALARIE,PSA_CONFIDENTIEL,PSA_SALAIREMOIS1,PSA_SALAIREMOIS2,PSA_SALAIREMOIS3,PSA_SALAIREMOIS4,PSA_SALAIREMOIS5,PSA_BOOLLIBRE2,PSA_SALAIRANN1 FROM SALARIES ';
  st := st + 'LEFT JOIN DEPORTSAL ON PSE_SALARIE = PSA_SALARIE ';
  if LeWhere <> '' then st := st + LeWhere;
{$IFDEF EAGLCLIENT}
  if not ConsultP then
  begin
    if LeSalarie <> '' then st := st + ' AND PSE_RESPONSVAR = "' + LeSalarie + '"'
  end
  else
  begin
    if LeSalarie <> '' then st := st + ' AND PSE_RESPONSABS = "' + LeSalarie + '"';
  end;
{$ENDIF}
  St := St + ' ORDER BY PSA_SALARIE';
  Q := OpenSql(st, TRUE);
  TSal.LoadDetailDB('SALARIES', '', '', Q, FALSE, FALSE);
  Ferme(Q);

  for i := 1 to 35 do LesAReporter[i] := 'NON';
  LaGrille.CacheEdit;
  LaGrille.SynEnabled := False;
  LaGrille.Cells[0, 0] := 'Collaborateur';
  LaGrille.Cells[1, 0] := 'Nom';
  LaGrille.Cells[2, 0] := 'Emploi';
  PgNbSalLib := VH_Paie.PgNbSalLib;
  NbDecal := 0;
  for i := 1 to 5 do PrimAffic[i] := FALSE;

  for i := 1 to PgNbSalLib do
  begin
    if i = 1 then
      if VH_Paie.PgPrimAffichSal1 then PrimAffic[i] := TRUE;
    if i = 2 then
      if VH_Paie.PgPrimAffichSal2 then PrimAffic[i] := TRUE;
    if i = 3 then
      if VH_Paie.PgPrimAffichSal3 then PrimAffic[i] := TRUE;
    if i = 4 then
      if VH_Paie.PgPrimAffichSal4 then PrimAffic[i] := TRUE;
    if i = 5 then
      if VH_Paie.PgPrimAffichSal5 then PrimAffic[i] := TRUE;
  end;

  st := 'SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_PGPRIMAFFICHSAL6"';
  Q := OpenSql(St, TRUE);
  if not Q.EOF then SalAnn1 := Q.FindField('SOC_DATA').AsString
  else SalAnn1 := '-';
  FERME(Q);
  if (SalAnn1 = 'X') then PrimAffic[5] := TRUE; // On force sur le 5eme élement
  for i := 1 to 5 do
  begin
    if PrimAffic[i] then NbDecal := NbDecal + 1;
  end;
  LaGrille.FixedCols := 3 + NbDecal;
  LaGrille.OnCellExit := GrilleCellexit;
  LaGrille.OnRowEnter := GrilleRowEnter;
  LaGrille.OnCellEnter := GrilleCellEnter;
  LaGrille.PostDrawCell := PostDrawCell;
  LaGrille.OnDblClick := GrilleDblClick;
  laGrille.OnKeyDown := KeyDown;
  //  LaGrille.ColWidths  [0]:=30;
  //  LaGrille.ColWidths [1]:=35;
  //  LaGrille.ColWidths [2]:=22;
  //  for i := 3 to LaGrille.ColCount-1 do begin  LaGrille.ColWidths [i]:=22; end;
  //  for i := 0 to LaGrille.ColCount-1  do LaGrille.ColAligns[i]:=taleftJustify ;
  //  for i := 3 to LaGrille.ColCount-1 do LaGrille.ColAligns[i]:=taRightJustify ;


  for i := 1 to PgNbSalLib do
  begin
    if i = 1 then
    begin
      if VH_Paie.PgPrimAffichSal1 then
        LibColRem[i] := VH_Paie.PgSalLib1
    end
    else if i = 2 then
    begin
      if VH_Paie.PgPrimAffichSal2 then
        LibColRem[i] := VH_Paie.PgSalLib2
    end
    else if i = 3 then
    begin
      if VH_Paie.PgPrimAffichSal3 then
        LibColRem[i] := VH_Paie.PgSalLib3
    end
    else if i = 4 then
    begin
      if VH_Paie.PgPrimAffichSal4 then
        LibColRem[i] := VH_Paie.PgSalLib4;
    end
    else if i = 5 then
    begin
      if VH_Paie.PgPrimAffichSal5 then
        LibColRem[i] := VH_Paie.PgSalLib5;
      if (SalAnn1 = 'X') then LibColRem[i] := 'Variable';
    end;
  end;

  for j := 1 to 4 do
  begin
    if LesMasq[j] = '' then break;
    LeMasque := TOB_Masque.FindFirst(['PMR_ORDRE'], [LesMasq[j]], FALSE);
    if LeMasque <> nil then
    begin
      for i := 1 to 7 do
      begin
        LaRubrique := LeMasque.GetValue('PMR_COL' + IntToStr(i));
        T1 := LesRem.FindFirst(['PRM_RUBRIQUE'], [LaRubrique], FALSE);
        if T1 <> nil then
        begin
          LeType := LeMasque.GetValue('PMR_TYPECOL' + IntToStr(i));
          if LeType = 'BAS' then NbD := T1.GetValue('PRM_DECBASE')
          else if LeType = 'COE' then NbD := T1.GetValue('PRM_DECCOEFF')
          else if LeType = 'TAU' then NbD := T1.GetValue('PRM_DECTAUX')
          else if LeType = 'MON' then NbD := T1.GetValue('PRM_DECMONTANT');
          LesRub[i + LesDecal[j]] := LeMasque.GetValue('PMR_COL' + IntToStr(i));
          LesTyp[i + LesDecal[j]] := LeMasque.GetValue('PMR_TYPECOL' + IntToStr(i));
          LesDecimales[i + LesDecal[j]] := NbD;
          LesAReporter[i + LesDecal[j]] := LeMasque.GetValue('PMR_REPORTCOL' + IntToStr(i)); // Stockage des colonnes à reporter
          LaGrille.Cells[i + 2 + NbDecal + LesDecal[j], 0] := LeMasque.GetValue('PMR_LIBCOL' + IntToStr(i));
        end; // Fin si Rem connue
      end; // Fin Boucle sur nbre de colonnes gérées
    end; // si masque existe
  end; // boucle sur les masques
  j := 1;
  for i := 1 to PgNbSalLib do
  begin
    if LibColRem[i] <> '' then
    begin
      LaGrille.Cells[2 + j, 0] := LibColRem[i];
      j := j + 1;
    end;
  end;
  // Boucle de chargement de la Grille
  j := 0;
  for i := 1 to 5 do j := j + NbCol[i]; // nbre de colonnes + identification salarie
  LaGrille.ColCount := j + 3 + NbDecal;
  j := 0;
  QMul.First;
  LaGrille.RowCount := 2;
  Ligne := 1; // Ligne courante
  Abandon := FALSE;
  LaGrille.ColCount := LaGrille.ColCount + 1; // Zone de commentaire en fin de saisie
  LaGrille.Cells[LaGrille.ColCount - 1, 0] := 'Commentaire';
  LaGrille.Colwidths[LaGrille.ColCount - 1] := 200;
  LaGrille.Colwidths[5] := 120;
  LaGrille.ColLengths[LaGrille.ColCount - 1] := 35;

  if VH_Paie.PgPrimMoisPrec then IndLign := 1
  else IndLign := 0;
  IdentCol;
  // Boucle pour identifier les colonnes avec des % soit de façon à dimensionner la colonne
  k := 1;
  for i := 3 + NbDecal to LaGrille.ColCount - 3 - NbDecal do
  begin // Boucle pour calculer à l'origine les valeurs en fct des taux ou coeff
    LeCode := CodeCal[k];
    if LeCode = '' then continue;
    CodeC := Readtokenst(LeCode);
    LeCode := Readtokenst(LeCode);
    //    if LeCode <> '0' then LaGrille.ColWidths[i] := LaGrille.ColWidths[i] + 15;
    k := k + 1;
  end;
  Ceg := GetParamSocSecur('SO_IFDEFCEGID', TRUE);
  if Ecran.Name <> 'A' then RempliGrille;
  SetControlProperty('LECOMMENTAIRE', 'CAPTION', 'Commentaire : ' + LaGrille.Cells[LaGrille.ColCount - 1, 1]);
  for i := 3 to Lagrille.ColCount - 2 do
  begin
    LaGrille.ColAligns[i] := taRightJustify;
  end;
{$IFDEF EPRIMES}
  LaGrille.RowCount := 2;
  LaGrille.FixedRows := 1;
{$ENDIF EPRIMES}

  St := '';
  for i := 1 to LaGrille.ColCount - 4 do
  begin
    St := St + ';CHAMP' + IntToStr(i);
  end;
{$IFDEF EPRIMES}
//  LaGrille.ColWidths[2] := 0;
  TobSaiPrim.PutGridDetail(LaGrille, False, False, 'SALARIE;NOM;EMPLOI' + St + ';COMMENTAIRE', False);
  k := 1;
  for i := 3 + NbDecal to LaGrille.ColCount - 1 do
  begin // Boucle pour calculer à l'origine les valeurs en fct des taux ou coeff
    LeCode := CodeCal[k];
    if LeCode = '' then continue;
    CodeC := Readtokenst(LeCode);
    LeCode := Readtokenst(LeCode);
    if LeCode <> '0' then
    begin
      for zz := 1 to LaGrille.RowCount do
      begin
        Mt1 := Valeur(LaGrille.Cells[i, zz]);
        if Copy(LaGrille.Cells[0, zz], 1, 3) = '-->' then continue;
        TS := TSal.FindFirst(['PSA_SALARIE'], [LaGrille.Cells[0, zz]], FALSE);
        if TS <> nil then
        begin
          if StrToInt(LeCode) < 6 then Mt := TS.GetValue('PSA_SALAIREMOIS' + LeCode)
          else Mt := TS.GetValue('PSA_SALAIRANN1');
          if CodeC = '05' then
          begin
            //            if mt <> 1 then
            Mt1 := Mt * (Mt1 / 100);
            //             else mt := mt * (MT1 * 100) ;
          end
          else Mt1 := Mt * Mt1;
          Mt1 := ARRONDI(Mt1, 2);
          if mt <> 1 then LaGrille.cells[i, zz] := LaGrille.cells[i, zz] + '% soit ' + DoubleToCell(Mt1, 2)
          else LaGrille.cells[i, zz] := DoubleToCell(Mt1, 2);
          ;
        end;
      end;
    end;
    k := k + 1;
  end;
{$ENDIF EPRIMES}
  HMTrad.ResizeGridColumns(THGrid(GetControl('GRILLE')));
end;

procedure TOF_PG_SAISPRIM.RempliGrille;
var
  okok, Abandon, Areprendre: Boolean;
  i, j, k: Integer;
  Salarie, CodeC, Libelle: string;
  Mt, Mt1, Montt: Double;
  TS: TOB;
  zz, IndLign: Integer;
begin
  i := QMUL.RecordCount;
{$IFNDEF EPRIMES}
  InitMoveProgressForm(nil, 'Chargement des données de la saisie', 'Veuillez patienter SVP ...', i, TRUE, TRUE);
{$ENDIF EPRIMES}

  if VH_Paie.PgPrimMoisPrec then IndLign := 1
  else IndLign := 0;

{$IFNDEF EPRIMES}

  while not QMul.EOF do
  begin
    okok := MoveCurProgressForm(IntToStr(Ligne));
    if not OkOk then
    begin
      Abandon := TRUE;
      break;
    end;
    Salarie := QMul.FindField('PSA_SALARIE').AsString;
    if VH_Paie.PgPrimMoisPrec then
    begin
      LaGrille.Cells[0, Ligne] := '-->Mois ' + IntToStr(MM) + '/' + IntToStr(AA);
      LaGrille.Cells[1, Ligne] := QMul.FindField('PSA_LIBELLE').AsString + ' ' + QMul.FindField('PSA_PRENOM').AsString;
      LaGrille.Cells[2, Ligne] := RechDom('PGLIBEMPLOI', QMul.FindField('PSA_LIBELLEEMPLOI').AsString, FALSE);
    end;
    LaGrille.Cells[0, Ligne + IndLign] := Salarie;
    LaGrille.Cells[1, Ligne + IndLign] := QMul.FindField('PSA_LIBELLE').AsString + ' ' + QMul.FindField('PSA_PRENOM').AsString;
    ;
    LaGrille.Cells[2, Ligne + IndLign] := RechDom('PGLIBEMPLOI', QMul.FindField('PSA_LIBELLEEMPLOI').AsString, FALSE);
    j := 1;
    for i := 1 to 5 do
    begin
      if LibColRem[i] <> '' then
      begin
        if (i = 1) and Ceg then Montt := QMul.FindField('PSA_SALAIREMOIS1').AsFloat + QMul.FindField('PSA_SALAIREMOIS2').AsFloat + QMul.FindField('PSA_SALAIREMOIS4').AsFloat
        else Montt := QMul.FindField('PSA_SALAIREMOIS' + IntToStr(i)).AsFloat;
        LaGrille.Cells[2 + j, Ligne + IndLign] := StrfMontant(Montt, 7, 2, '', TRUE);
        if (i = 5) and (SalAnn1 = 'X') then
        begin
          Montt := QMul.FindField('PSA_SALAIRANN1').AsFloat;
          if Montt = 1 then Montt := 0; // Pour forcer affichage
          LaGrille.Cells[2 + j, Ligne + IndLign] := StrfMontant(Montt, 7, 2, '', TRUE);
        end;
        j := j + 1;
      end;
    end;
    for i := 1 to 35 do LesMontants[i] := 0;

    for j := 1 to 4 do // boucle sur les masques
    begin
      if LesMasq[j] = '' then break;
      // REcupération des infos saisies le mois précédent
      if VH_Paie.PgPrimMoisPrec then
      begin
        RendSaisieRub(Salarie, DebutM1, FinM1, LesDecal[j], NbCol[j + 1], LeMasque, Libelle); // Recup de la saisie déjà effectuée le mois précédent
        LeMasque := TOB_Masque.FindFirst(['PMR_ORDRE'], [LesMasq[j]], FALSE);
        for i := 1 to LesDecal[j + 1] do
        begin
          LaGrille.Cells[i + 2 + NbDecal + LesDecal[j], Ligne] := DoubleToCell(LesMontants[i + LesDecal[j]], LesDecimales[i + LesDecal[j]]); //
        end;
      end;
      for i := 1 to 35 do LesMontants[i] := 0; // RAZ pour traiter le mois en cours
      // Fin traitement du mois précédent

      AReprendre := RendSaisieRub(Salarie, DateDebut, DateFin, LesDecal[j], NbCol[j + 1], LeMasque, Libelle); // Recup de la saisie déjà effectuée
      if AReprendre = FALSE then RendSaisiePrec(Salarie, DateDebut, DateFin, LesDecal[j], NbCol[j + 1], LeMasque); // Recup de la saisie précédente
      for i := 1 to LesDecal[j + 1] do
      begin
        if i + 2 + NbDecal + LesDecal[j] > LaGrille.ColCount then exit;
        if AReprendre = TRUE then LaGrille.Cells[i + 2 + NbDecal + LesDecal[j], Ligne + IndLign] := DoubleToCell(LesMontants[i + LesDecal[j]], LesDecimales[i + LesDecal[j]])
        else // Saisie à reprendre depuis la derniere saisie
        begin
          if LesAReporter[i + LesDecal[j]] = 'OUI' then
          begin // montant à reporter du mois précédent
            LaGrille.Cells[i + 2 + NbDecal + LesDecal[j], Ligne + IndLign] := DoubleToCell(LesMontants[i + LesDecal[j]], LesDecimales[i + LesDecal[j]]); //
          end
          else LaGrille.Cells[i + 2 + NbDecal + LesDecal[j], Ligne + IndLign] := '';
        end;
      end; // fin boucle sur les colonnes
    end; // fin boucle sur les masques
    if libelle <> '' then LaGrille.Cells[LaGrille.ColCount - 1, Ligne + IndLign] := Libelle; // Affichage du commentaire
    QMul.NEXT;
    Ligne := Ligne + IndLign + 1;
    okok := MoveCurProgressForm(IntToStr(Ligne));
    if not OkOk then break;
    LaGrille.RowCount := LaGrille.RowCount + IndLign + 1;

  end; // si query non nulle
  FiniMoveProgressForm();

  k := 1;
  for i := 3 + NbDecal to LaGrille.ColCount - 1 do
  begin // Boucle pour calculer à l'origine les valeurs en fct des taux ou coeff
    LeCode := CodeCal[k];
    if LeCode = '' then continue;
    CodeC := Readtokenst(LeCode);
    LeCode := Readtokenst(LeCode);
    if LeCode <> '0' then
    begin
      for zz := 1 to LaGrille.RowCount do
      begin
        Mt1 := Valeur(LaGrille.Cells[i, zz]);
        if Copy(LaGrille.Cells[0, zz], 1, 3) = '-->' then continue;
        TS := TSal.FindFirst(['PSA_SALARIE'], [LaGrille.Cells[0, zz]], FALSE);
        if TS <> nil then
        begin
          if StrToInt(LeCode) < 6 then Mt := TS.GetValue('PSA_SALAIREMOIS' + LeCode)
          else Mt := TS.GetValue('PSA_SALAIRANN1');
          if CodeC = '05' then
          begin
            //            if mt <> 1 then
            Mt1 := Mt * (Mt1 / 100);
            //             else mt := mt * (MT1 * 100) ;
          end
          else Mt1 := Mt * Mt1;
          Mt1 := ARRONDI(Mt1, 2);
          if mt <> 1 then LaGrille.cells[i, zz] := LaGrille.cells[i, zz] + '% soit ' + DoubleToCell(Mt1, 2)
          else LaGrille.cells[i, zz] := DoubleToCell(Mt1, 2);
          ;
        end;
      end;
    end;
    k := k + 1;
  end;
{$ENDIF EPRIMES}
  if Abandon = TRUE then Close; // Abandon de la saisie par rubrique
  LaGrille.MontreEdit;
  LaGrille.SynEnabled := True;
  LaGrille.RowCount := LaGrille.RowCount - 1;
  if SalAdm = 'X' then
  begin
    LaGrille1 := THGrid(Getcontrol('GRILLE1'));
    if LaGrille1 <> nil then
    begin
      LaGrille1.ColCount := LaGrille.ColCount - 1;
      for i := 0 to LaGrille1.colcount - 1 do
      begin
        LaGrille1.ColWidths[i] := LaGrille.ColWidths[i];
        LaGrille1.ColAligns[i] := LaGrille.ColAligns[i];
        if i > 2 then LaGrille1.Cells[i, 0] := LaGrille.Cells[i, 0];
      end;
    end;
    CalculTotaux;
  end
  else SetControlVisible('GRILLE1', FALSE);
  //  HMTrad.ResizeGridColumns(LaGrille);
  //  HMTrad.ResizeGridColumns(LaGrille1);

end;


// Fonction de validation de la saisie

procedure TOF_PG_SAISPRIM.ValiderClick(Sender: TObject);
var
  rep: Integer;
begin
  inherited;
  if ConsultP then exit;
  Rep := PGIAsk('Voulez vous sauvegarder votre saisie ?', Ecran.Caption);
  if rep = mrNo then exit;
  if rep = mrCancel then exit;
  if rep = mryes then OnSauve;
end;
// fermeture de la forme

procedure TOF_PG_SAISPRIM.FermeClick(Sender: TObject);
begin
  ValiderClick(Sender);
  Close;
end;

// fonction d'ecriture des elements saisis

function TOF_PG_SAISPRIM.OnSauve: boolean;
var
  LeType, LaVal, CodeC: string;
  i, j, z: Integer;
  T_Sais, LeSal, T: TOB;
  Mt, Mt1: Double;
begin
  result := TRUE;
  if ConsultP then exit;
  if LaGrille.RowCount - 1 = 1 then Modifier := TRUE;
  if Modifier = FALSE then exit;
  BeginTrans;
  try
    Modifier := FALSE;
    T_Sais := TOB.Create('Les lignes de saisie', nil, -1);
    try
    for i := 1 to LaGrille.RowCount - 1 do
    begin
      if (Copy(LaGrille.Cells[0, i], 1, 7)) = '-->Mois' then continue; // on ne prend pas les lignes d'info
      LeSal := TSal.FindFirst(['PSA_SALARIE'], [LaGrille.Cells[0, i]], FALSE);
      for j := 1 to LaGrille.colCount - 3 - NbDecal - 1 do
      begin
        T := TOB.Create('HISTOSAISPRIM', T_Sais, -1);
        if T <> nil then
        begin
          if LeSal <> nil then T.PutValue('PSP_CONFIDENTIEL', LeSal.GetValue('PSA_CONFIDENTIEL'));
          T.PutValue('PSP_ORIGINEMVT', 'SRB');
          T.PutValue('PSP_SALARIE', LaGrille.Cells[0, i]);
          T.PutValue('PSP_DATEDEBUT', DateDebut);
          T.PutValue('PSP_DATEFIN', DateFin);
          T.PutValue('PSP_RUBRIQUE', LesRub[j]);
          T.PutValue('PSP_TYPALIMPAIE', LesTyp[j]);
          LeType := LesTyp[j];
          LaVal := LaGrille.Cells[J + 2 + NbDecal, i];
          z := Pos('%', LaVal);
          if (IsNumeric(LaVal) or (z > 0)) then
          begin
            if z > 0 then LaVal := Copy(LaVal, 1, z - 1);
          end;
          LeCode := CodeCal[j];
          CodeC := Readtokenst(LeCode);
          LeCode := Readtokenst(LeCode);
          mt := 0;
          if LeCode <> '' then
          begin
            mt := 0;
            LeSal := TSal.FindFirst(['PSA_SALARIE'], [LaGrille.Cells[0, i]], FALSE);
            if LeSal <> nil then
            begin
              if StrToInt(LeCode) < 6 then Mt := LeSal.GetValue('PSA_SALAIREMOIS' + LeCode)
              else Mt := LeSal.GetValue('PSA_SALAIRANN1');
            end;
          end;
          if mt = 1 then Mt1 := Valeur(LaVal) * 100
          else Mt1 := Valeur(LaVal);
          if LeType = 'MON' then T.PutValue('PSP_MONTANT', Mt1);
          if LeType = 'BAS' then T.PutValue('PSP_BASE', Mt1);
          if LeType = 'TAU' then T.PutValue('PSP_TAUX', Mt1);
          if LeType = 'COE' then T.PutValue('PSP_COEFF', Mt1);

          T.PutValue('PSP_ORDRE', j);
          T.PutValue('PSP_DATEINTEGRAT', IDate1900);
        end;
      end; // boucle sur les colonnes
      if T <> nil then T.Putvalue('PSP_LIBELLE', LaGrille.Cells[LaGrille.ColCount - 1, i]);
    end; // Boucle sur les lignes de la grille
    T_Sais.SetAllModifie(TRUE);
    T_Sais.InsertOrUpdateDB(TRUE);
    finally
      T_Sais.free;
    end;
    CommitTrans;
  except
    on e: exception do
    begin
      result := FALSE;
      Rollback;
      PGIBox('Une erreur est survenue lors de la validation de la saisie', E.Message);
    end;
  end;
  Modifier := FALSE;
  SetControlEnabled('BCherche', TRUE);
end;


function TOF_PG_SAISPRIM.NbreDecimale(NbD: Integer): string;
var
  i: Integer;
begin
  result := '';
  if NbD = 0 then exit;
  result := '.';
  for i := 1 to NbD do
  begin
    result := result + '0';
  end;
end;

function TOF_PG_SAISPRIM.RendSaisieRub(Salarie: string; DateDebut, DateFin: TDateTime; Decalage, NbreCols: Integer; LeMasque: TOB; var libelle: string): Boolean;
var
  Q: TQuery;
  St: string;
  i: Integer;
  Montant: Double;
begin
  result := FALSE;
  Libelle := ''; // #############
  st := 'SELECT * FROM HISTOSAISPRIM WHERE PSP_SALARIE="' + Salarie + '" AND PSP_DATEDEBUT ="' + USDateTime(DateDebut)
    + '" AND PSP_ORDRE >"' + InttoStr(decalage) + '" AND PSP_ORDRE <="' + IntToStr(Decalage + NbreCols)
    + '" AND PSP_DATEFIN ="' + USDateTime(DateFin) + '" AND PSP_ORIGINEMVT="SRB" ORDER BY PSP_SALARIE,PSP_DATEDEBUT,PSP_DATEFIN,PSP_ORDRE';
  Q := OpenSql(St, TRUE);
  while not Q.EOF do
  begin
    i := Q.FindField('PSP_ORDRE').AsInteger;
    Montant := 0;
    if Q.FindField('PSP_TYPALIMPAIE').AsString = 'BAS' then Montant := Q.FindField('PSP_BASE').AsFloat
    else if Q.FindField('PSP_TYPALIMPAIE').AsString = 'MON' then Montant := Q.FindField('PSP_MONTANT').AsFloat
    else if Q.FindField('PSP_TYPALIMPAIE').AsString = 'COE' then Montant := Q.FindField('PSP_COEFF').AsFloat
    else if Q.FindField('PSP_TYPALIMPAIE').AsString = 'TAU' then Montant := Q.FindField('PSP_TAUX').AsFloat;
    if (i > decalage) and (i <= decalage + Nbrecols) then
    begin
      if LesRub[i] = Q.FindField('PSP_RUBRIQUE').AsString then
        LesMontants[i] := Montant;
    end;
    result := TRUE;
    Libelle := Q.FindField('PSP_LIBELLE').AsString;
    Q.next;
  end;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Philippe Dumet
Créé le ...... : 30/05/2001
Modifié le ... :   /  /
Description .. : Saisie par rubriques des éléménts variables
Suite ........ : Utilise le CTRL_C et CTRL_V entre une feuille EXCEL
Suite ........ : et une grille
Suite ........ : Analyse de la tob et rempli la grille en remettant dans les
Suite ........ : bonnes lignes et les bonnes colonnes
Suite ........ : Indique si erreur identification
Mots clefs ... : PAIE;PGSAISRUB
*****************************************************************}

function TOF_PG_SAISPRIM.RendSaisiePrec(Salarie: string; DateDebut, DateFin: TDateTime; Decalage, NbreCols: Integer; LeMasque: TOB): Boolean;
var
  Q: TQuery;
  St: string;
  i: Integer;
  Deb, Fin: TDateTime;
  Montant: Double;
begin
  result := FALSE;
  Deb := DebutDeMois(DateDebut) - 1;
  Deb := DebutDeMois(Deb);
  Fin := FindeMois(Deb);
  st := 'SELECT * FROM HISTOSAISPRIM WHERE PSP_SALARIE="' + Salarie + '" AND PSP_DATEFIN ="' + USDateTime(FIN)
    + '" AND PSP_ORDRE >"' + InttoStr(decalage) + '" AND PSP_ORDRE <="' + IntToStr(Decalage + NbreCols)
    + '" AND PSP_DATEDEBUT ="' + USDateTime(Deb) + '" AND PSP_ORIGINEMVT="SRB" ORDER BY PSP_SALARIE,PSP_DATEDEBUT,PSP_DATEFIN,PSP_ORDRE';
  Q := OpenSql(St, TRUE);
  while not Q.EOF do
  begin
    i := Q.FindField('PSP_ORDRE').AsInteger;
    Montant := 0;
    if (LesAReporter[i] = 'OUI') and (i >= decalage) and (i <= decalage + nbrecols) then
    begin
      if Q.FindField('PSP_TYPALIMPAIE').AsString = 'BAS' then Montant := Q.FindField('PSP_BASE').AsFloat
      else if Q.FindField('PSP_TYPALIMPAIE').AsString = 'MON' then Montant := Q.FindField('PSP_MONTANT').AsFloat
      else if Q.FindField('PSP_TYPALIMPAIE').AsString = 'COE' then Montant := Q.FindField('PSP_COEFF').AsFloat
      else if Q.FindField('PSP_TYPALIMPAIE').AsString = 'TAU' then Montant := Q.FindField('PSP_TAUX').AsFloat;
      if i <> 0 then LesMontants[i] := Montant;
      if LesRub[i] = Q.FindField('PSP_RUBRIQUE').AsString then
        result := TRUE;
    end;
    Q.next;
  end;
  Ferme(Q);
end;

procedure TOF_PG_SAISPRIM.GrilleCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  i, s: Integer;
  St, CodeCalcul: string;
  Mt, Mt1, Nbre: Double;
  TS: TOB;
begin
  Modifier := TRUE;
  if ACol = LaGrille.ColCount - 1 then
  begin
//  SetControlProperty ('LECOMMENTAIRE', 'Caption', 'Commentaire : '+LaGrille.Cells[Acol, ARow]);
    exit; // Saisie du commentaire donc pas de controle
  end;
  if (LaGrille.Row - 1 = 0) and (ARow = 1) and VH_Paie.PgPrimMoisPrec then
  begin
    LaGrille.Cells[Acol, LaGrille.Row] := ' ';
    exit;
  end;
  if (LaGrille.Cells[Acol, Arow] = '') and (ARow <> LaGrille.RowCount - 1) then exit;
  i := Pos('%', LaGrille.Cells[Acol, ARow]);
  if (IsNumeric(LaGrille.Cells[Acol, ARow]) or (i > 0)) then
  begin
    if i > 0 then
      LaGrille.Cells[ACol, ARow] := Copy(LaGrille.Cells[ACol, ARow], 1, i - 1);
    if ARow = LaGrille.RowCount - 1 then
    begin
      if IsNumeric(LaGrille.Cells[Acol, ARow]) then
      begin
        if LaGrille.Cells[Acol, ARow] <> '' then Nbre := Valeur(LaGrille.Cells[Acol, ARow])
        else Nbre := 0;
        if ARow <> 0 then LaGrille.Cells[Acol, ARow] := DoubleToCell(Nbre, LesDecimales[ACol + 1 - LaGrille.FixedCols]);
      end
      else if ARow <> 0 then LaGrille.Cells[Acol, ARow] := '';
    end
    else
    begin
      Nbre := Valeur(LaGrille.Cells[Acol, Arow]);
      if Arow <> 0 then LaGrille.Cells[Acol, ARow] := DoubleToCell(Nbre, LesDecimales[ACol + 1 - LaGrille.FixedCols]);
    end;
  end
  else
  begin
    if ARow <> 0 then LaGrille.Cells[Acol, ARow] := DoubleToCell(0, LesDecimales[ACol + 1 - LaGrille.FixedCols]);
    if ARow = LaGrille.RowCount - 1 then
    begin
      if IsNumeric(LaGrille.Cells[Acol, ARow]) then
      begin
        Nbre := Valeur(LaGrille.Cells[Acol, ARow]);
        if ARow <> 0 then LaGrille.Cells[Acol, ARow] := DoubleToCell(Nbre, LesDecimales[ACol + 1 - LaGrille.FixedCols]);
      end
      else if ARow <> 0 then LaGrille.Cells[Acol, ARow] := '';
    end;
  end;
  i := ACol - 2 - NbDecal;
  if LaGrille.Cells[Acol, ARow] <> '' then
  begin
    if (LesTyp[i] = 'TAU') or (LesTyp[i] = 'COE') then
    begin
      if CodeCal[i] <> '' then
      begin
        St := CodeCal[i];
        CodeCalcul := ReadTokenSt(St);
        s := StrToInt(ReadTokenSt(St));
        if s > 0 then
        begin
          TS := TSal.findFirst(['PSA_SALARIE'], [LaGrille.Cells[0, ARow]], FALSE);
          Mt := 0;
          if TS <> nil then
          begin
            if s < 6 then Mt := TS.GetValue('PSA_SALAIREMOIS' + IntToStr(s))
            else Mt := TS.GetValue('PSA_SALAIRANN1');
          end;
          Mt1 := Valeur(LaGrille.Cells[ACol, ARow]);
          if CodeCalcul = '05' then
          begin
            if Mt <> 1 then Mt1 := Mt * (Mt1 / 100);
          end
          else Mt1 := Mt * Mt1;
          Mt1 := ARRONDI(Mt1, 2);
          if IsNumeric(LaGrille.Cells[ACol, ARow]) and
            (Copy(LaGrille.Cells[0, ARow], 1, 7) <> '-->Mois') then
          begin
            if Mt <> 1 then LaGrille.cells[ACol, ARow] := LaGrille.cells[ACol, ARow] + '% soit ' + DoubleToCell(Mt1, 2); // LesDecimales[i]
          end;
        end;
      end;
    end; // Fin si Taux ou Coeff
  end;
  if SalAdm = 'X' then CalculTotaux; // on recalcule systematiquement les totaux
end;


procedure TOF_PG_SAISPRIM.OnClose;
begin
  If Assigned(HMTrad) Then FreeAndNil(HMTrad); //PT1
  If Assigned(QZ) Then FreeAndNil(QZ); //PT1
  If Assigned(LesFiltres) Then FreeAndNil(LesFiltres); //PT1
  if Ecran.Name = 'A' then
  begin
    if QMUL <> nil then Ferme(QMUL);
  end;

  if TSal <> nil then
  begin
    TSal.Free;
    TSal := nil;
  end;
{$IFNDEF EPRIMES}
  if TOB_Masque <> nil then
  begin
    TOB_Masque.free;
    TOB_Masque := nil;
  end;
  if LesRem <> nil then
  begin
    LesRem.free;
    LesRem := nil;
  end;
  if T_RemMere <> nil then
  begin
    T_RemMere.free;
    T_RemMere := nil;
  end;
{$ENDIF EPRIMES}
{$IFDEF aucasou}
  if TOB_Rem <> nil then
  begin
    TOB_Rem.free;
    TOB_Rem := nil;
  end;
{$ENDIF}
  inherited;

end;

procedure TOF_PG_SAISPRIM.KeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  if (Key = 86) and (ssCtrl in Shift) then
    GrilleCopierColler;
end;

procedure TOF_PG_SAISPRIM.GrilleCopierColler;
var
  T, TT: TOB;
  i, j, k, NbreCol, LstErrorL, LstErrorP: integer;
  St, Salarie: string;
  LeVar: Variant;
  TitreCol: array[1..35] of string;
begin
  LstErrorL := 0;
  LstErrorP := 0;
  T := TOBLoadFromClipBoard;
  T.Detail[0].Free;

  NbreCol := LaGrille.ColCount - 2 - NbDecal; // Car 3 colonnes figees Matricule,Nom,Prenom
  for i := 1 to 35 do TitreCol[i] := '';
  for i := 1 to 35 do
  begin
    if i > NbreCol then break;
    TitreCol[i] := LaGrille.Cells[i + 2 + NbDecal, 0];
  end;
  for i := 0 to T.Detail.Count - 1 do
  begin
    TT := T.Detail[i];
    LeVar := TT.GetValue('Matricule');
    Salarie := VarToStr(LeVar);
    k := IdentSalarie(Salarie);
    if k = -1 then LstErrorL := LstErrorL + 1;
    if (k > 0) and (k < LaGrille.RowCount) then
    begin // Boucle sur les lignes de la TOB
      for j := 1 to 35 do
      begin // Boucle sur les colonnes de la grille qui correspondent aux colonnes de la feuilles EXCEL
        st := TitreCol[j];
        LeVar := TT.GetValue(TitreCol[j]);
        if LeVar <> '' then // colonne identifiée
          if IsNumeric(VarToStr(LeVar)) then LaGrille.Cells[j + 2 + NbDecal, k] := DoubleToCell(Valeur(VarToStr(LeVar)), 2)
          else LstErrorP := LstErrorP + 1;
      end;
    end;
  end;
  T.Free;
  St := 'Attention, ';
  if LstErrorL > 0 then St := St + IntTOstr(LstErrorL) + ' ligne(s) n''ont pas été intégrées';
  if LstErrorP > 0 then St := St + IntTOstr(LstErrorP) + ' cellule(s) colonne n''ont pas été intégrées';
  if (LstErrorL > 0) then PGIBox(St, Ecran.Caption);
end;

function TOF_PG_SAISPRIM.IdentSalarie(Salarie: string): Integer;
var
  i: Integer;
begin
  result := -1;
  for i := 1 to LaGrille.RowCount - 1 do
  begin
    if VH_PAie.PgTypeNumSal = 'NUM' then
    begin
      if Valeur(Salarie) = Valeur(LaGrille.Cells[0, i]) then
      begin
        Result := i;
        break;
      end;
    end
    else
    begin
      if LaGrille.Cells[0, i] = Salarie then
      begin
        Result := i;
        break;
      end;
    end;
  end;
end;

procedure TOF_PG_SAISPRIM.ImpClik(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  PrintDBGrid(TCustomGrid(LaGrille), nil, Ecran.Caption, '');
{$ELSE}
  //  PrintDBGrid(LAGRILLE, '', Ecran.Caption, '');
//  PrintGrid([LaGille],Ecran.Caption);
{$ENDIF}
end;


procedure TOF_PG_SAISPRIM.PostDrawCell(ACol, ARow: Integer;
  Canvas: TCanvas; AState: TGridDrawState);
begin
  if ((Copy(LaGrille.Cells[0, ARow], 1, 7) = '-->Mois') and (ARow > 0)) then
    GridGriseCell(LaGrille, Acol, Arow, Canvas);
end;

procedure TOF_PG_SAISPRIM.GrilleCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var
  i: Integer;
begin
  if ConsultP then
  begin
    cancel := FALSE;
    Exit;
  end;
  Modifier := TRUE;
  if (Copy(LaGrille.Cells[0, LaGrille.Row], 1, 7) = '-->Mois') then
  begin
    if LaGrille.Row <= 2 then
    begin
      {      Cancel := TRUE;
            PGIBox ('Vous ne pouvez pas modifier une ligne concernant le mois précédent', Ecran.caption);}
      LaGrille.Row := LaGrille.Row + 1;
      exit;
    end;
  end;
  i := Pos('%', LaGrille.Cells[LaGrille.Col, LaGrille.Row]);
  if i > 0 then
    LaGrille.Cells[LaGrille.Col, LaGrille.Row] := Copy(LaGrille.Cells[LaGrille.Col, LaGrille.Row], 1, i - 1);
  ACol := LaGrille.Col;
  if (Copy(LaGrille.Cells[0, LaGrille.Row], 1, 7) = '-->Mois') then
  begin
    if LaGrille.Row < Arow then LaGrille.Row := LaGrille.Row - 1
    else
    begin
      if LaGrille.Row + 1 < LaGrille.RowCount then LaGrille.Row := LaGrille.Row + 1
      else LaGrille.Row := LaGrille.Row - 1;
    end;
  end;
end;

procedure TOF_PG_SAISPRIM.AccesSal(Sender: TObject);
var
  CodeSalarie: string;
begin
  CodeSalarie := LaGrille.Cells[0, Lagrille.Row];
  if (CodeSalarie <> '') and (CodeSalarie[1] <> '-') then AglLanceFiche('PAY', 'SALARIE_PRIM', '', CodeSalarie, 'ACTION=CONSULTATION');
end;

procedure TOF_PG_SAISPRIM.HistSal(Sender: TObject);
var
  CodeSalarie, St1, Nom, Decal: string;
  DD, DF: TDateTime;
  TS: TOB;
  YY, JJ, MM: WORD;
  Ceg: Boolean;
begin
  CodeSalarie := LaGrille.Cells[0, Lagrille.Row];
  Nom := LaGrille.Cells[1, Lagrille.Row] + ' ' + LaGrille.Cells[2, Lagrille.Row];
  TheTob := T_RemMere;
  DecodeDate(DateFin, YY, MM, JJ);
  Ceg := GetParamSocSecur('SO_IFDEFCEGID', TRUE);
  TS := TSal.FindFirst(['PSA_SALARIE'], [CodeSalarie], FALSE);
  // Specifique CEGID Boolean libre 2 ==> indique le décaalage ou non du calcul des primes
  if TS <> nil then Decal := TS.GetValue('PSA_BOOLLIBRE2');
  if Ceg then
  begin
    if (Decal = 'X') and (MM = 1) then YY := YY - 1;
    if Decal = 'X' then DD := EncodeDate(YY, 02, 01)
    else DD := EncodeDate(YY, 01, 01);
    DF := FindeMois(Plusmois(DD, 11));
  end
  else
  begin
    DD := PLUSMOIS(DateDebut, -12);
    DF := PLUSMOIS(DateFin, -1);
  end;
  DF := FINDEMOIS(DF);
  St1 := CodeSalarie + ';' + ' ' + ';' + DateToStr(DD) + ';' + DateToStr(DF) + ';' + Nom + ';' + LaGrille.Cells[LaGrille.Col, 0];
  AglLanceFiche('PAY', 'SAISPRIM_REM', '', '', St1);
  TheTOB := nil;
end;
{ fonction de calcul des totaux de grille en fonction du type de la zone
saisie (Montant ou Taux ou Coeff)
si taux ou coeff alors on regarde si le calcul est fait à partir des elts de
salaires
}

procedure TOF_PG_SAISPRIM.CalculTotaux;
var
  i, j, zz, nb: Integer;
  ColFix: array[1..5] of double;
  ColVar: array[1..35] of double;
  TS: TOB;
  St: string;
  s, k: Integer;
  Mt, Mt1: Double;
  CodeCalCul: string;
begin
  j := 1;
  for i := 1 to 5 do ColFix[i] := 0;
  if NbDecal > 0 then
  begin
    for i := 1 to NbDecal do
    begin // calcul de la somme des elts de paie
      for zz := 1 to LaGrille.RowCount do
      begin
        if IsNumeric(LaGrille.Cells[i + 2, zz]) and
          (Copy(LaGrille.Cells[0, zz], 1, 7) <> '-->Mois') then
          colFix[i] := ColFix[i] + Valeur(LaGrille.Cells[i + 2, zz]);
      end;
    end;
  end;
  LaGrille1.cells[2, 1] := 'Total';
  nb := 2 + NbDecal; // indice de déplacement dans la grille pour connaitre la pemiere zone saisissable
  for i := 1 to 5 do
  begin // affichage totalisation colonnes elt de salaire du salarié
    if ColFix[i] <> 0 then LaGrille1.Cells[i + 2, 1] := DoubleToCell(ColFix[i], 2);
  end;
  // on va cumuler toutes les autres colonnes
  for i := 1 to 35 do ColVar[i] := 0;

  for i := nb + 1 to LaGrille.ColCount - 1 do
  begin
    if (LesTyp[i - nb] = 'MON') or (LesTyp[i - nb] = 'BAS') then
    begin // cas ou montant donc on somme simplement
      for zz := 1 to LaGrille.rowCount do
      begin
        if IsNumeric(LaGrille.Cells[i, zz]) and
          (Copy(LaGrille.Cells[0, zz], 1, 7) <> '-->Mois') then
          ColVar[i - Nb] := ColVar[i - Nb] + Valeur(LaGrille.Cells[i, zz]);
      end;
      LaGrille1.cells[i, 1] := DoubleToCell(ColVar[i - Nb], LesDecimales[i - nb]);
    end;
    if (LesTyp[i - nb] = 'TAU') or (LesTyp[i - nb] = 'COE') then
    begin
      if CodeCal[i - nb] <> '' then
      begin
        St := CodeCal[i - nb];
        CodeCalcul := ReadTokenSt(St);
        s := StrToInt(ReadTokenSt(St));
        if s > 0 then
        begin
          for zz := 1 to LaGrille.rowCount do
          begin
            TS := TSal.findFirst(['PSA_SALARIE'], [LaGrille.Cells[0, zz]], FALSE);
            Mt := 0;
            k := Pos('%', LaGrille.Cells[i, zz]);
            if k > 0 then Mt1 := Valeur(Copy(LaGrille.Cells[i, zz], 1, k - 1))
            else Mt1 := Valeur(LaGrille.Cells[i, zz]);
            if TS <> nil then
            begin
              if s < 6 then Mt := TS.GetValue('PSA_SALAIREMOIS' + IntToStr(s))
              else Mt := TS.GetValue('PSA_SALAIRANN1');
            end;
            if CodeCalcul = '05' then
            begin
              if Mt <> 1 then Mt1 := Mt * (Mt1 / 100);
            end
            else Mt1 := Mt * Mt1;
            Mt1 := ARRONDI(Mt1, 2);
            if ((IsNumeric(LaGrille.Cells[i, zz])) or (k > 0)) and
              (Copy(LaGrille.Cells[0, zz], 1, 7) <> '-->Mois') then
              ColVar[i - Nb] := ColVar[i - Nb] + Mt1;
          end;
          LaGrille1.cells[i, 1] := DoubleToCell(ColVar[i - Nb], LesDecimales[i - nb]);
        end;
      end;
    end; // Fin si Taux ou Coeff
  end;
end;

procedure TOF_PG_SAISPRIM.TotClik(Sender: TObject);
begin
  if SalAdm = 'X' then CalculTotaux;
end;

procedure TOF_PG_SAISPRIM.IdentCol;
var
  i, s, Nb: Integer;
  CodeCalcul, LaBase: string;
  T1: TOB;
begin
  nb := 2 + NbDecal;
  for i := 1 to 35 do CodeCal[i] := '';
  for i := nb to LaGrille.ColCount - 1 do
  begin
    if (LesTyp[i + 1 - nb] = 'TAU') or (LesTyp[i + 1 - nb] = 'COE') then
    begin
      T1 := LesRem.FindFirst(['PRM_RUBRIQUE'], [LesRub[i + 1 - nb]], FALSE);
      if T1 <> nil then
      begin
        s := 0;
        CodeCalcul := T1.GetValue('PRM_CODECALCUL');
        if (CodeCalcul = '04') or (CodeCalcul = '05') or (CodeCalcul = '08') then
        begin
          if T1.GetValue('PRM_TYPEBASE') = '03' then //Alors une variable
          begin
            LaBase := T1.GetValue('PRM_BASEREM');
            if LaBase = '0170' then s := 1
            else if LaBase = '0171' then s := 2
            else if LaBase = '0172' then s := 3
            else if LaBase = '0173' then s := 4
            else if LaBase = '0178' then s := 5
            else if LaBase = '0174' then s := 6;
          end;
          if s > 0 then
          begin
            CodeCal[i + 1 - nb] := CodeCalcul + ';' + IntToStr(s);
          end;
        end;
      end;
    end;
  end;
end;
{ Fonction aide au calcul d'un pourcentage ou coeff ou taux en fonction de la rubrique,du salarie,
du montant de la base
On a besoin de passer en paramètre : Nom+Prenom salarie,
valeur de la base,
codeCalcul,
libelle de la remunération,
nbre de decimales de la zone saisie

et retour si valide on recupere la valeur dans la zone taux de la fiche aide
}

procedure TOF_PG_SAISPRIM.AideCalcul(Sender: TObject);
var
  CodeSalarie, St, Ret, CodeCalcul, NumCh, Champ: string;
  Base: DOUBLE;
  Col, s: Integer;
  TS: TOB;
begin
  if ConsultP then exit;
  if CodeCal[LaGrille.Col - 2 - NbDecal] = '' then
  begin
    PgiBox('La colonne n''est pas calculable', Ecran.Caption);
    exit;
  end;
  CodeSalarie := LaGrille.Cells[0, Lagrille.Row];
  if (CodeSalarie = '') or (CodeSalarie[1] = '-') then exit;
  CodeCalcul := CodeCal[LaGrille.Col - 2 - NbDecal];
  Col := POS(';', CodeCalcul);
  NumCh := Copy(CodeCalcul, Col + 1, 1);
  TS := TSal.findFirst(['PSA_SALARIE'], [CodeSalarie], FALSE);
  if TS = nil then exit;
  s := StrToInt(NumCh);
  if s < 6 then Base := TS.GetValue('PSA_SALAIREMOIS' + NumCh)
  else Base := TS.GetValue('PSA_SALAIRANN1');
  if Base = 1 then exit;
  CodeCalcul := ReadTokenst(CodeCalcul);
  Col := LaGrille.Col;
  Champ := LaGrille.Cells[LaGrille.Col, LaGrille.Row]; // Recup du taux ou coeff
  Col := Pos('%', Champ);
  if Col > 0 then Champ := Copy(Champ, 1, Col - 1);
  St := DoubleToCell(Base, 2) + ';' + Champ + ';';
  St := St + LaGrille.Cells[1, Lagrille.Row] + ' ' + LaGrille.Cells[2, Lagrille.Row] + ';';
  st := st + CodeCalcul + ';' + IntToStr(LesDecimales[LaGrille.Col - 2 - NbDecal]) + ';';
  St := St + LaGrille.cells[LaGrille.Col, 0]; // entete de la colonne
  if (CodeSalarie <> '') and (CodeSalarie[1] <> '-') then
  begin
    ret := AglLanceFiche('PAY', 'AIDE_CALPRIME', '', '', St);
    if Ret <> '' then LaGrille.Cells[LaGrille.Col, LaGrille.Row] := Ret;
  end;
end;

procedure TOF_PG_SAISPRIM.GrilleDblClick(Sender: TObject);
begin
  AideCalcul(Sender);
end;

procedure TOF_PG_SAISPRIM.HistoSalarie(Sender: TObject);
begin
  AglLanceFiche('PAY', 'PGMULHISTOSAL', '', '', LaGrille.Cells[0, Lagrille.Row]);
end;


procedure TOF_PG_SAISPRIM.GrilleRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  SetControlProperty('LECOMMENTAIRE', 'CAPTION', 'Commentaire : ' + LaGrille.Cells[LaGrille.ColCount - 1, THGrid(Sender).Row]);
end;

procedure TOF_PG_SAISPRIM.AfficheTitre;
var i, j: Integer;
  LaRubrique, LeType: string;
  T1: Tob;
  NbD: Integer;
begin
  for i := 0 to LaGrille.ColCount - 1 do
  begin
    LaGrille.Cells[i, 0] := '                                          ';
    LaGrille.ColWidths[i] := 500;
  end;
  HMTrad.ResizeGridColumns(THGrid(GetControl('GRILLE')));
  lAGrille.Refresh;
end;

procedure TOF_PG_SAISPRIM.ExporterGrille(Sender: TObject);
var Fichier, StChampGrid: string;
  TobTemp, T, TobAug: Tob;
  i: Integer;
begin
  NextPrevControl(TFVierge(Ecran));
  Fichier := AGLLanceFiche('PAY', 'AUGM_XLS', '', '', '');
  if Fichier <> '' then
  begin
    TobAug := Tob.create('_temp_', nil, -1);
    TobAug.GetGridDetail(LaGrille, LaGrille.RowCount - 1, '', 'INDICATEUR;' + StChampGrid);
    TobTemp := Tob.Create('Temporaire', nil, -1);
    for i := 0 to TobAug.Detail.Count - 1 do
    begin
      T := Tob.Create('Fille_temp', TobTemp, -1);
      T.AddChampSupValeur('SALARIE', TobAug.Detail[i].GetValue('PSA_LIBELLE'), False);
      T.AddChampSupValeur('EMPLOI', TobAug.Detail[i].GetValue('PSA_LIBELLEEMPLOI'), False);
      T.AddChampSupValeur('DATE_ENTREE', TobAug.Detail[i].GetValue('PSA_DATEENTREE'), False);
      T.AddChampSupValeur('ETAT', RechDom('PGAUGMETATVALID', TobAug.Detail[i].GetValue('PBG_ETATINTAUGM'), False), False);
      T.AddChampSupValeur('FIXE', TobAug.Detail[i].GetValue('PBG_FIXEAV'), False);
      T.AddChampSupValeur('POURCENTAGE_FIXE', TobAug.Detail[i].GetValue('PBG_PCTFIXE'), False);
      T.AddChampSupValeur('FIXE_AUGMENTE', TobAug.Detail[i].GetValue('PBG_FIXEAP'), False);
      T.AddChampSupValeur('VARIABLE', TobAug.Detail[i].GetValue('PBG_VARIABLEAV'), False);
      T.AddChampSupValeur('POURCENTAGE_VARIABLE', TobAug.Detail[i].GetValue('PBG_PCTVARIABLE'), False);
      T.AddChampSupValeur('VARIABLE_AUGMENTE', TobAug.Detail[i].GetValue('PBG_VARIABLEAP'), False);
      T.AddChampSupValeur('BRUT', TobAug.Detail[i].GetValue('TOTALAV'), False);
      T.AddChampSupValeur('POURCENTAGE_BRUT', TobAug.Detail[i].GetValue('PCTTOTAL'), False);
      T.AddChampSupValeur('BRUT_AUGMENTE', TobAug.Detail[i].GetValue('TOTALAP'), False);
      T.AddChampSupValeur('PBG_COMMENTAIREABR', TobAug.Detail[i].GetValue('TOTALAP'), False);
    end;
    TobTemp.SaveToExcelFile(Fichier);
    FreeAndNil(tobTemp);
    FreeAndNil(TobAug);
  end;
end;

initialization
  registerclasses([TOF_PG_SAISPRIM]);
end.

