{***********UNITE*************************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 21/05/2002
Modifié le ... : 05/06/2003
Description .. : Statistiques des stocks :
Suite ........ :  -par article
Suite ........ :  -à la dimension
Mots clefs ... : BOS5;STATISTIQUES;DIMENSION;
*****************************************************************}
unit StatDim_Tof;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, M3FP, Stat, HDimension,
  {$IFDEF EAGLCLIENT}
  emul, MaineAGL,
  {$ELSE}
  dbTables, db, DBGrids, mul, Fe_main,
  {$ENDIF}
  UTobView, TVProp, UTob, UtilArticle, EntGC, UtilGC, UtilPGI, HQry;

type
  TOF_StatDim = class(TOF)
  private
    Dim: boolean;
    TobViewer1: TTobViewer;
    stChampsCompl, sType: string; // Champs complémentaires de la requête
    procedure TVOnDblClickCell(Sender: TObject);
    procedure SetLastError(Num: integer; ou: string);
    function Formate_StWhere(NaturesDoc, ListeNatPiece, DatePieceDeb, DatePieceFin: string): string;
  public
    StatVente, StatAchat: string; // Statistique portant sur les ventes d'une période
    stDimensions, stSelectDim, stGroupByDim: string; // Champs sélectionnés et jointures pour les dimensions
    ListeNatPieceVen, DatePieceDebVen, DatePieceFinVen: string;
    ListeNatPieceAch, DatePieceDebAch, DatePieceFinAch: string;
    procedure OnArgument(Arguments: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  end;

const
  // libellés des messages
  TexteMessage: array[1..4] of string = (
    {1}'Cet article est défini en taille unique'
    {2}, 'Opération impossible'
    {3}, 'Attention, vous n''avez pas sélectionné de nature de document pour les ventes'
    {4}, 'Attention, vous n''avez pas sélectionné de nature de document pour les achats'
    );

procedure LanceStat(st: string);

implementation

procedure TOF_StatDim.SetLastError(Num: integer; ou: string);
begin
  LastError := Num;
  LastErrorMsg := TexteMessage[LastError];
  if ou <> '' then SetFocusControl(ou);
end;

procedure TOF_StatDim.OnArgument(Arguments: string);
var Nbr, iCol: integer;
  TCB: TCheckBox;
  stIndice: string;
  F: TFStat;
begin
  inherited;
  Dim := (Arguments = 'DIM');
  sType := 'CODEARTICLE';
  //if Dim then sType := 'ARTICLE' else sType := 'CODEARTICLE';
  // Paramétrage des libellés des familles, stat. article et de la collection
  ChangeLibre2('TGA_COLLECTION', Ecran);
  for iCol := 1 to 3 do
  begin
    stIndice := IntToStr(iCol);
    ChangeLibre2('TGA_FAMILLENIV' + stIndice, Ecran);
  end;
  stChampsCompl := '';
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation = ART_ORLI) then
  begin
    for iCol := 4 to 8 do
    begin
      stIndice := IntToStr(iCol);
      if ChangeLibre2('TGA2_FAMILLENIV' + stIndice, Ecran) then
        stChampsCompl := stChampsCompl + ',max(GA2_FAMILLENIV' + stIndice + ') as GA2_FAMILLENIV' + stIndice;
    end;
    for iCol := 1 to 2 do
    begin
      stIndice := IntToStr(iCol);
      if ChangeLibre2('TGA2_STATART' + stIndice, Ecran) then stChampsCompl := stChampsCompl + ',max(GA2_STATART' + stIndice + ') as GA2_STATART' + stIndice;
    end;
  end;

  F := TFStat(Ecran);
  // Si on travaille à la dimension : mise en forme des libellés des dimensions
  if Dim then
  begin
    SetControlVisible('GBDIM', true);
    F.Caption := TraduireMemoire('Tableau de bord des stocks à la dimension');
    for iCol := 1 to MaxDimension do
    begin
      TCB := TCheckBox(GetControl('CB' + InttoStr(iCol)));
      TCB.Caption := RechDom('GCCATEGORIEDIM', 'DI' + InttoStr(iCol), False);
      if TCB.Caption = '.-' then TCB.Visible := False;
    end;
  end else
    F.Caption := TraduireMemoire('Tableau de bord des stocks par article');
  UpdateCaption(F);

  TobViewer1 := TTobViewer(getcontrol('TV'));
  TobViewer1.OnDblClick := TVOnDblClickCell;

  if (ctxMode in V_PGI.PGIContexte) then SetControlProperty('GQ_DEPOT', 'Plus', 'GDE_SURSITE="X"');

  // Paramètrage des libellés des onglets des zones et tables libres du dépôt
  if not VH_GC.GCMultiDepots then
  begin
    SetControlCaption('PTABLESLIBRESDEP', 'Tables libres étab.');
    SetControlCaption('PZONESLIBRESDEP', 'Zones libres étab.');
  end;

  // Paramétrage des libellés des tables libres article et dépôt
  if (GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'GA_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False);
  if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GDE_LIBREDEP', 10, '') = 0) then SetControlVisible('PTABLESLIBRESDEP', False);

  // Mise en forme des libellés des dates, booléans libres et montants libres article et dépôt
  Nbr := 0;
  if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else inc(Nbr);
  if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else inc(Nbr);
  if (GCMAJChampLibre(TForm(Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else inc(Nbr);
  {$IFNDEF CCS3}
  if (Nbr = 0) then
    {$ENDIF}
    SetControlVisible('PZONESLIBRES', False);

  Nbr := 0;
  if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GDE_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VALDEP', False) else inc(Nbr);
  if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GDE_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATEDEP', False) else inc(Nbr);
  if (GCMAJChampLibre(TForm(Ecran), False, 'BOOL', 'GDE_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOLDEP', False) else inc(Nbr);
  {$IFNDEF CCS3}
  if (Nbr = 0) then
    {$ENDIF}
    SetControlVisible('PZONESLIBRESDEP', False);

  // Initialisation de l'onglet : Compléments
  SetControlChecked('VAL_DPA', False);
  SetControlChecked('VAL_DPR', False);
  SetControlChecked('VAL_PMAP', True);
  SetControlChecked('VAL_PMRP', False);
  SetControlChecked('STAT_VENTE', False);
  if (VH_GC.GCSeria = True) or (V_PGI.VersionDemo = True) then
  begin
    SetControlText('GL_NATUREPIECEG', 'FAC;AVC;FFO');
    SetControlText('GL_NATUREPIECEG1', 'ALF;CF;BLF');
  end else
  begin
    SetControlText('GL_NATUREPIECEG', 'FFO');
    SetControlText('GL_NATUREPIECEG1', 'BLF');
  end;

  StatVente := '-';
  StatAchat := '-';
  ListeNatPieceVen := '';
  ListeNatPieceAch := '';
  DatePieceDebVen := '';
  DatePieceDebAch := '';
  DatePieceFinVen := '';
  DatePieceFinAch := '';
end;

procedure TOF_StatDim.OnLoad;
var xx_where: string;
begin
  inherited;
  xx_where := '';

  // Gestion des checkBox : booléens libres article
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'BOOL', 'GA_BOOLLIBRE', 3, '');

  // Gestion des dates libres article
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'DATE', 'GA_DATELIBRE', 3, '_');

  // Gestion des montants libres article
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'EDIT', 'GA_VALLIBRE', 3, '_');

  // Gestion des checkBox : booléens libres  dépôts
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'BOOL', 'GDE_BOOLLIBRE', 3, '');

  // Gestion des dates libres dépôts
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'DATE', 'GDE_DATELIBRE', 3, '_');

  // Gestion des montants libres dépôts
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'EDIT', 'GDE_VALLIBRE', 3, '_');

  SetControlText('XX_WHERE', xx_where);
end;

function TOF_StatDim.Formate_StWhere(NaturesDoc, ListeNatPiece, DatePieceDeb, DatePieceFin: string): string;
var StWhere: string; //  , CodeNature
begin
  StWhere := 'left join PIECE on GP_NATUREPIECEG=GL_NATUREPIECEG and GP_SOUCHE=GL_SOUCHE ' +
    'and GP_NUMERO=GL_NUMERO and GP_INDICEG=GL_INDICEG Where ';
  ListeNatPiece := '';
  if NaturesDoc <> TraduireMemoire('<<Tous>>') then
    {  begin
        While NaturesDoc <> '' do
        begin
          CodeNature := ReadTokenSt (NaturesDoc) ;
          if ListeNatPiece <> '' then ListeNatPiece := ListeNatPiece+',' ;
          ListeNatPiece := ListeNatPiece + '"' + CodeNature + '"' ;
        end;
      end;   }
    ListeNatPiece := MultiComboInSQL(NaturesDoc);
  if ListeNatPiece <> '' then
  begin
    StWhere := StWhere + 'GP_NATUREPIECEG IN (' + ListeNatPiece + ') and ';
    StWhere := StWhere + 'GP_DATEPIECE>="' + UsDateTime(StrToDate(DatePieceDeb)) + '" and ' +
      'GP_DATEPIECE<="' + UsDateTime(StrToDate(DatePieceFin)) + '" and ' +
      'GP_DEPOT=GQ_DEPOT and GL_' + sType + '=ART.GA_' + sType + ' and ' +
      'GL_TYPELIGNE="ART"';
  end;
  result := StWhere;
end;

procedure TOF_StatDim.OnUpdate;
var StSQL, StWhere, stSelect, stSelectTablesLibres, stGroupBy, suffixe : string ;
    stCaption, iDim, NaturesDoc, StWhereVen, StWhereAch, stOrderByDim, sep: string;
  stDim: array of string;
  iCol,i: integer;
  CalcCoeff: TCheckBox;
begin
  inherited;
  StatVente := GetControlText('STAT_VENTE');
  StatAchat := GetControlText('STAT_ACHAT');
  CalcCoeff := TcheckBox(GetControl('COEFF'));

  // Contrôle sur l'onglet : Compléments
  if StatVente = 'X' then
  begin
    if (GetControlText('GL_NATUREPIECEG') = '') or (GetControlText('GL_NATUREPIECEG') = TraduireMemoire('<<Tous>>')) then
    begin
      SetLastError(3, 'GL_NATUREPIECEG');
      exit;
    end;
  end;
  if StatAchat = 'X'
    then
  begin
    if (GetControlText('GL_NATUREPIECEG1') = '') or (GetControlText('GL_NATUREPIECEG1') = TraduireMemoire('<<Tous>>')) then
    begin
      SetLastError(4, 'GL_NATUREPIECEG1');
      exit;
    end;
  end;
  stSQL := '';
  stSelect := '';
  stGroupBy := '';
  stSelectDim := '';
  stDimensions := ' ';
  stGroupByDim := ' ';
  stOrderByDim := '';
  stSelectTablesLibres := '' ;
  // Récupération des dimensions
  if Dim then
  begin
    SetLength(stDim, MaxDimension); // allocation de chaines autant que de dimensions possibles
    for iCol := 1 to MaxDimension do
    begin
      iDim := IntToStr(iCol);
      stDim[iCol - 1] := 'left join DIMENSION GDI' + iDim + ' on GDI' + iDim + '.GDI_TYPEDIM="DI' + iDim +
        '" AND GA_GRILLEDIM' + iDim + '=GDI' + iDim + '.GDI_GRILLEDIM AND GA_CODEDIM' + iDim + '=GDI' + iDim + '.GDI_CODEDIM ';
      if TCheckBox(GetControl('CB' + iDim)).Checked then
      begin
        stCaption := Trim(TCheckBox(GetControl('CB' + iDim)).Caption);
        stSelectDim := stSelectDim + 'GDI' + iDim + '.GDI_RANG,' + 'GDI' + iDim + '.GDI_LIBELLE as ' + StrToAlias(stCaption) + ',ART.GA_GRILLEDIM' + iDim +
          ',';
        stDimensions := stDimensions + stDim[iCol - 1];
        stGroupByDim := stGroupByDim + 'GDI' + iDim + '.GDI_RANG,' + 'GDI' + iDim + '.GDI_LIBELLE ' + ',ART.GA_GRILLEDIM' + iDim + ',';
        if stOrderByDim <> '' then sep := ',';
        stOrderByDim := stOrderByDim + sep + 'GDI' + iDim + '.GDI_RANG ';
      end;
    end;

    // Si aucune dimension n'est sélectionné alors on travaille par défaut sur la 1ère
    if stSelectDim = '' then
    begin
      stSelectDim := 'GDI1.GDI_RANG,GDI1.GDI_LIBELLE as ' + StrToAlias(Trim(TCheckBox(GetControl('CB1')).Caption)) + ',GA_GRILLEDIM1,';
      stDimensions := ' left join DIMENSION GDI1 on GDI1.GDI_TYPEDIM="DI1" AND GA_GRILLEDIM1=GDI1.GDI_GRILLEDIM AND GA_CODEDIM1=GDI1.GDI_CODEDIM ';
      stGroupByDim := 'GDI1.GDI_RANG,GDI1.GDI_LIBELLE,GA_GRILLEDIM1,';
      stOrderByDim := ' ORDER BY GDI1.GDI_RANG';
    end else
      stOrderByDim := ' ORDER BY ' + stOrderByDim;
  end;

  StGroupBy := Copy(StGroupBy, 0, Length(StGroupBy) - 1);
  stSelect := Copy(stSelect, 0, Length(stSelect) - 1);
  if StatVente = 'X' then
  begin
    // Statistique portant sur les ventes d'une période
    NaturesDoc := GetControlText('GL_NATUREPIECEG');
    DatePieceDebVen := GetControlText('GL_DATEPIECEV');
    DatePieceFinVen := GetControlText('GL_DATEPIECEV_');
    StWhereVen := Formate_StWhere(NaturesDoc, ListeNatPieceVen, DatePieceDebVen, DatePieceFinVen);
  end;
  if StatAchat = 'X' then
  begin
    // Statistique portant sur les achats d'une période
    NaturesDoc := GetControlText('GL_NATUREPIECEG1');
    DatePieceDebAch := GetControlText('GL_DATEPIECEA');
    DatePieceFinAch := GetControlText('GL_DATEPIECEA_');
    StWhereAch := Formate_StWhere(NaturesDoc, ListeNatPieceAch, DatePieceDebAch, DatePieceFinAch);
  end;

  // Select des tables libres

  For i:=1 to 10 do
  begin
    if i<10 then
      suffixe:=IntToStr(i) else
      suffixe:='A';
  if THValComboBox(GetControl('GA_LIBREART'+suffixe)).Visible then
    stSelectTablesLibres := stSelectTablesLibres +  ',max(ART.GA_LIBREART' + suffixe  + ') AS GA_LIBREART' + suffixe ;
  end ;


  // Constitution de la requête SQL

 { ExecuteSQL ('@@') ;
  ExecuteSQL ('set quoted_identifier off') ;
  ExecuteSQL ('Set nocount on') ;

  ExecuteSQL ('@@IF (object_id(''tempdb..#ART'') IS not Null) Drop Table #ART') ;
  ExecuteSQL ('@@IF (object_id(''tempdb..#DISP'') IS not Null) Drop Table #DISP') ;
  ExecuteSQL ('@@IF (object_id(''tempdb..#ACHAT'') IS not Null) Drop Table #ACHAT') ;
  ExecuteSQL ('@@IF (object_id(''tempdb..#VENTE'') IS not Null) Drop Table #VENTE') ;

  exit ;

  ExecuteSQL ('SELECT gq_article,gq_depot into #DISP FROM DISPO WHERE GQ_DEPOT IN() AND GQ_EMPLACEMENT="' + GetControlText ('GQ_EMPLACEMENT') + '"') ;}

  StSQL := 'SELECT ' + stSelectDim + ' GQ_DEPOT,ART.GA_STATUTART,ART.GA_' + sType + ',ART.GA_LIBELLE,ART.GA_LIBCOMPL,ART.GA_TYPEARTICLE,ART.GA_COLLECTION,' +
    'ART.GA_FOURNPRINC,FOU.T_LIBELLE as NOM_FOURNISSEUR,GQ_EMPLACEMENT,GQ_STOCKMIN,GQ_STOCKMAX,' +
    'max(ART.GA_FAMILLENIV1) as GA_FAMILLENIV1, max(ART.GA_FAMILLENIV2) as GA_FAMILLENIV2, max(ART.GA_FAMILLENIV3) as GA_FAMILLENIV3' +
    stSelectTablesLibres ;
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation = ART_ORLI)
    then StSQL := StSQL + stChampsCompl;

  StSQL := StSQL + ',(SUM (GQ_LIVREFOU)) QTE_RECU_FOURN,' +
    '(SUM (GQ_TRANSFERT)) QTE_TRANSFERE,' +
    '(SUM (GQ_ENTREESORTIES)) QTE_ENTREE_SORTIE,' +
    '(SUM (GQ_PHYSIQUE)) QTE_STOCK,' +
    '(SUM (GQ_RESERVECLI)) QTE_RESA_CLIENT,' +
    '(SUM (GQ_RESERVEFOU)) QTE_CDE_FOURN,' +
    '(SUM (GQ_PREPACLI)) QTE_PREPA_CLIENT,' +
    '(SUM (GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU-GQ_PREPACLI)) QTE_STOCK_NET';
  if (StatVente = '-') and (StatAchat = '-') then
    StSQL := StSQL + ',(SUM (GQ_LIVRECLIENT)) QTE_LIVRE_CLIENT,' +
      '(SUM (GQ_VENTEFFO)) QTE_VENTE_FO,' +
      '(SUM (GQ_LIVRECLIENT+GQ_VENTEFFO)) QTE_VENDU';

  if (GetControlText('VAL_DPA') <> '-') then
  begin
    if (StatVente = '-') and (StatAchat = '-') then
      StSQL := StSQL + ',(SUM ((GQ_LIVRECLIENT+GQ_VENTEFFO)*GQ_DPA)) DPA_VENDU';
    StSQL := StSQL + ',(SUM (GQ_LIVREFOU*GQ_DPA)) DPA_RECU_FOURN,' +
      '(SUM (GQ_TRANSFERT*GQ_DPA)) DPA_TRANSFERE,' +
      '(SUM (GQ_ENTREESORTIES*GQ_DPA)) DPA_ENTREE_SORTIE,' +
      '(SUM (GQ_PHYSIQUE*GQ_DPA)) DPA_STOCK,' +
      '(SUM (GQ_RESERVECLI*GQ_DPA)) DPA_RESA_CLIENT,' +
      '(SUM (GQ_RESERVEFOU*GQ_DPA)) DPA_CDE_FOURN,' +
      '(SUM (GQ_PREPACLI)) DPA_PREPA_CLIENT,' +
      '(SUM ((GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU-GQ_PREPACLI)*GQ_DPA)) DPA_STOCK_NET';
  end;

  if (GetControlText('VAL_PMAP') <> '-') then
  begin
    if (StatVente = '-') and (StatAchat = '-') then
      StSQL := StSQL + ',(SUM ((GQ_LIVRECLIENT+GQ_VENTEFFO)*GQ_PMAP)) PMAP_VENDU';
    StSQL := StSQL + ',(SUM (GQ_LIVREFOU*GQ_PMAP)) PMAP_RECU_FOURN,' +
      '(SUM (GQ_TRANSFERT*GQ_PMAP)) PMAP_TRANSFERE,' +
      '(SUM (GQ_ENTREESORTIES*GQ_PMAP)) PMAP_ENTREE_SORTIE,' +
      '(SUM (GQ_PHYSIQUE*GQ_PMAP)) PMAP_STOCK,' +
      '(SUM (GQ_RESERVECLI*GQ_PMAP)) PMAP_RESA_CLIENT,' +
      '(SUM (GQ_RESERVEFOU*GQ_PMAP)) PMAP_CDE_FOURN,' +
      '(SUM (GQ_PREPACLI)) PMAP_PREPA_CLIENT,' +
      '(SUM ((GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU-GQ_PREPACLI)*GQ_PMAP)) PMAP_STOCK_NET';
  end;

  if (GetControlText('VAL_DPR') <> '-') then
  begin
    if (StatVente = '-') and (StatAchat = '-') then
      StSQL := StSQL + ',(SUM ((GQ_LIVRECLIENT+GQ_VENTEFFO)*GQ_DPR)) DPR_VENDU';
    StSQL := StSQL + ',(SUM (GQ_LIVREFOU*GQ_DPR)) DPR_RECU_FOURN,' +
      '(SUM (GQ_TRANSFERT*GQ_DPR)) DPR_TRANSFERE,' +
      '(SUM (GQ_ENTREESORTIES*GQ_DPR)) DPR_ENTREE_SORTIE,' +
      '(SUM (GQ_PHYSIQUE*GQ_DPR)) DPR_STOCK,' +
      '(SUM (GQ_RESERVECLI*GQ_DPR)) DPR_RESA_CLIENT,' +
      '(SUM (GQ_RESERVEFOU*GQ_DPR)) DPR_CDE_FOURN,' +
      '(SUM (GQ_PREPACLI)) DPR_PREPA_CLIENT,' +
      '(SUM ((GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU-GQ_PREPACLI)*GQ_DPR)) DPR_STOCK_NET';
  end;

  if (GetControlText('VAL_PMRP') <> '-') then
  begin
    if (StatVente = '-') and (StatAchat = '-') then
      StSQL := StSQL + ',(SUM ((GQ_LIVRECLIENT+GQ_VENTEFFO)*GQ_PMRP)) PMRP_VENDU';
    StSQL := StSQL + ',(SUM (GQ_LIVREFOU*GQ_PMRP)) PMRP_RECU_FOURN,' +
      '(SUM (GQ_TRANSFERT*GQ_PMRP)) PMRP_TRANSFERE,' +
      '(SUM (GQ_ENTREESORTIES*GQ_PMRP)) PMRP_ENTREE_SORTIE,' +
      '(SUM (GQ_PHYSIQUE*GQ_PMRP)) PMRP_STOCK,' +
      '(SUM (GQ_RESERVECLI*GQ_PMRP)) PMRP_RESA_CLIENT,' +
      '(SUM (GQ_RESERVEFOU*GQ_PMRP)) PMRP_CDE_FOURN,' +
      '(SUM (GQ_PREPACLI)) PMRP_PREPA_CLIENT,' +
      '(SUM ((GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU-GQ_PREPACLI)*GQ_PMRP)) PMRP_STOCK_NET';
  end;

  //
  // les champs calculés des achats
  //
  if (StatAchat = 'X') then
  begin
    StSQL := StSQL + ',(Select SUM(GL_QTEFACT) from LIGNE ' + StWhereAch + ') ACHAT_QTE,' +
      '(Select SUM(GL_TOTALHT) from LIGNE ' + StWhereAch + ') ACHAT_MTHT,' +
      '(Select SUM(GL_TOTALTTC) from LIGNE ' + StWhereAch + ') ACHAT_MTTC';
    if (GetControlText('VAL_DPA') <> '-') then
      StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_DPA) from LIGNE ' + StWhereAch + ') ACHAT_DPA';
    if (GetControlText('VAL_PMAP') <> '-') then
      StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_PMAP) from LIGNE ' + StWhereAch + ') ACHAT_PMAP';
    if (GetControlText('VAL_DPR') <> '-') then
      StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_DPR) from LIGNE ' + StWhereAch + ') ACHAT_DPR';
    if (GetControlText('VAL_PMRP') <> '-') then
      StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_PMRP) from LIGNE ' + StWhereAch + ') ACHAT_PMRP';
  end;

  //
  // les champs calculés des ventes
  //
  if (StatVente = 'X') then
  begin
    StSQL := StSQL + ',(Select SUM(GL_QTEFACT) from LIGNE ' + StWhereVen + ') VENTE_QTE,' +
      '(Select SUM(GL_TOTALHT) from LIGNE ' + StWhereVen + ') VENTE_MTHT,' +
      '(Select SUM(GL_TOTALTTC) from LIGNE ' + StWhereVen + ') VENTE_MTTC';
    if (GetControlText('VAL_DPA') <> '-') then
      StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_DPA) from LIGNE ' + StWhereVen + ') VENTE_DPA';
    if (GetControlText('VAL_PMAP') <> '-') then
      StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_PMAP) from LIGNE ' + StWhereVen + ') VENTE_PMAP';
    if (GetControlText('VAL_DPR') <> '-') then
      StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_DPR) from LIGNE ' + StWhereVen + ') VENTE_DPR';
    if (GetControlText('VAL_PMRP') <> '-') then
      StSQL := StSQL + ',(Select SUM(GL_QTEFACT*GL_PMRP) from LIGNE ' + StWhereVen + ') VENTE_PMRP';
  end;

  //
  // les coefficients
  //
  if (CalcCoeff.Enabled and CalcCoeff.Checked) then
  begin
    StSQL := StSQL +
      ',((Select SUM(GL_TOTALTTC) from LIGNE ' + StWhereVen + ')*(Select SUM(GL_QTEFACT) from LIGNE ' + StWhereAch + '))/' +
      '((Select SUM(GL_QTEFACT) from LIGNE ' + StWhereVen + ')*(Select SUM(GL_TOTALHT) from LIGNE ' + StWhereAch + ')) COEF_ACHAT_VENTE,' +
      '((Select SUM(GL_QTEFACT) from LIGNE ' + StWhereVen + ') / (Select SUM(GL_QTEFACT) from LIGNE ' + StWhereAch + ')*100) COEF_VENDU_RECU,' +
      '((SUM (GQ_PHYSIQUE)) / (Select SUM(GL_QTEFACT) from LIGNE ' + StWhereAch + ')*100) COEF_STOCK_RECU';
  end;

  //
  // les jointures
  //
  StSQL := StSQL + ' FROM DISPO' +
    ' LEFT JOIN DEPOTS ON GDE_DEPOT=GQ_DEPOT' +
    ' LEFT JOIN ARTICLE ART ON GA_ARTICLE=GQ_ARTICLE';
  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation = ART_ORLI)
    then StSQL := StSQL + ' LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=GQ_ARTICLE';
  StSQL := StSQL + ' LEFT JOIN TIERS FOU ON FOU.T_NATUREAUXI="FOU" and FOU.T_TIERS=ART.GA_FOURNPRINC';
  StSQL := StSQL + stDimensions;

  StSQL := StSQL + ' WHERE GQ_CLOTURE="-" AND ART.GA_STATUTART IN ("DIM","UNI") '; //AND ART.GA_TYPEARTICLE="MAR" ';

  //
  // le group by
  //
  StSQL := StSQL + ' GROUP BY ' + stGroupByDim + ' GQ_DEPOT,ART.GA_STATUTART,ART.GA_' + sType + ',ART.GA_LIBELLE,ART.GA_LIBCOMPL,' +
    'ART.GA_TYPEARTICLE,ART.GA_COLLECTION,ART.GA_FOURNPRINC,FOU.T_LIBELLE,GQ_EMPLACEMENT,GQ_STOCKMIN,GQ_STOCKMAX';

  stSQL := StSQL + stOrderByDim;
  {  if (ctxMode in V_PGI.PGIContexte) and (GetPresentation=ART_ORLI)
then StSQL := StSQL + stChampsCompl ;    }

// Ajout du complément de la clause WHERE, pour prendre en compte les critères
// de sélection renseignés par l'utilisateur (collection, fournisseur, ...)
  stWhere := RecupWhereCritere(TPageControl(GetControl('Pages')));
  if Uppercase(copy(stWhere, 1, 6)) = 'WHERE ' then
  begin
    StWhere := Copy(StWhere, 7, length(StWhere) - 6);
    StSQL := InsertSQLWhere(StSQL, StWhere);
  end;
  TFstat(Ecran).StSQL := StSQL;

end;

procedure TOF_StatDim.TVOnDblClickCell(Sender: TObject);
var //bCloture : boolean;
  st1, st2, depot, cloture, statutart, args: string;
begin
  inherited;
  with TTobViewer(sender) do
  begin
    st1 := AsString[ColIndex('GA_' + sType), CurrentRow];
    st2 := CodeArticleUnique2(st1, '');
    depot := AsString[ColIndex('GQ_DEPOT'), CurrentRow];
    //bCloture := AsBoolean[ColIndex('GQ_CLOTURE'), CurrentRow] ;
    //cloture := CheckToString(bCloture) ;
    statutart := AsString[ColIndex('GA_STATUTART'), CurrentRow];
    if (statutart <> 'GEN') and (statutart <> 'DIM') then
    begin
      HShowMessage('0;' + TexteMessage[2] + ';' + TexteMessage[1] + ';W;O;O;O', '', '');
      exit;
    end else
      if (StatAchat = 'X') then
      args := 'ACTION=CONSULTATION;STOCKVTE=X;DEPOT=' + depot +
        ';CLOTURE=' + cloture + ';NATPIECE=' + ListeNatPieceAch + ';DATEPIECEDEB=' + DatePieceDebAch +
        ';DATEPIECEFIN=' + DatePieceFinAch
    else
      if (StatVente = 'X') then
      args := 'ACTION=CONSULTATION;STOCKVTE=X;DEPOT=' + depot +
        ';CLOTURE=' + cloture + ';NATPIECE=' + ListeNatPieceVen + ';DATEPIECEDEB=' + DatePieceDebVen +
        ';DATEPIECEFIN=' + DatePieceFinVen
    else args := 'ACTION=CONSULTATION;STOCK=X;DEPOT=' + depot; // + ';CLOTURE=' + cloture ;
    V_PGI.DispatchTT(7, taConsult, st2, args, '');
  end;
end;

procedure LanceStat(st: string);
begin
  DispatchArtMode(13, '', '', st);
end;

procedure TOF_StatDim.OnClose;
begin
  inherited;
  //ExecuteSQL ('Set arithabort on') ;
end;

initialization
  registerclasses([TOF_StatDim]);

end.
