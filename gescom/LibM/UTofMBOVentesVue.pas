{***********UNITE*************************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 05/06/2003
Modifié le ... : 05/06/2003
Description .. : TobViewer des achats et des ventes
Mots clefs ... : TOBVIEWER;STAT;DIMENSION
*****************************************************************}
unit UTofMBOVentesVue;

interface

uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
  HCtrls, HEnt1, HMsgBox, UTOF, M3FP, Stat, Hdimension, HQry,
  {$IFDEF EAGLCLIENT}
  emul, MaineAGL,
  {$ELSE}
  dbTables, db, DBGrids, mul, Fe_main,
  {$ENDIF}
  UTobView, TVProp, UTob, UtilArticle, AGLInit, EntGC, UtilGC, utilPGI, FactUtil;

type
  TOF_MBOVentesVue = class(TOF)
  private
    TobViewer1: TTobViewer;
    Dim: boolean;
    procedure TVOnDblClickCell(Sender: TObject);
    procedure RemplirTOBDim(naturepieceg, souche: string; numero: Integer; codeart: string; TOBDim: TOB);
    procedure SetLastError(Num: integer; ou: string);
  public
    procedure OnArgument(Arguments: string); override;
    procedure OnUpdate; override;
  end;

const
  // libellés des messages
  TexteMessage: array[1..2] of string = (
    {1}'Cet article est défini en taille unique'
    {2}, 'Opération impossible');

implementation

procedure TOF_MBOVentesVue.SetLastError(Num: integer; ou: string);
begin
  if ou <> '' then SetFocusControl(ou);
  LastError := Num;
  LastErrorMsg := TexteMessage[LastError];
  PGIBox(LastErrorMsg, TexteMessage[2]);
end;

procedure TOF_MBOVentesVue.OnArgument(Arguments: string);
var Critere, NomCritere, ValCritere, LibType, LibDim: string;
  x, iCol: integer;
  TCB: TCheckBox;
  F: TFStat;
  {$IFDEF NOMADE}
  stPlus, stNatDef: string;
  {$ENDIF}
begin
  inherited;
  repeat
    Critere := Trim(ReadTokenSt(Arguments));
    if Critere <> '' then
    begin
      x := pos('=', Critere);
      if x <> 0 then
      begin
        NomCritere := uppercase(copy(Critere, 1, x - 1));
        ValCritere := copy(Critere, x + 1, length(Critere));
        if NomCritere = 'TYPESTA' then LibType := ValCritere
        else
          if NomCritere = 'TYPEART' then LibDim := ValCritere;
      end;
    end;
  until Critere = '';
  F := TFStat(Ecran);

  // Si on travaille à la dimension : mise en forme des libellés des dimensions
  if LibDim = 'DIM' then Dim := true;
  if Dim then
  begin
    SetControlVisible('GBDIM', true);
    if F.Name = 'VENTEVUE' then
      F.Caption := TraduireMemoire('Tableau de bord des ventes à la dimension') else
      F.Caption := TraduireMemoire('Tableau de bord des achats à la dimension');
    for iCol := 1 to MaxDimension do
    begin
      TCB := TCheckBox(GetControl('CB' + InttoStr(iCol)));
      TCB.Caption := RechDom('GCCATEGORIEDIM', 'DI' + InttoStr(iCol), False);
      if TCB.Caption = '.-' then TCB.Visible := False;
    end;
  end else
  begin
    if F.Name = 'VENTEVUE' then
      F.Caption := TraduireMemoire('Tableau de bord des ventes par article') else
      F.Caption := TraduireMemoire('Tableau de bord des achats par article');
  end;
  UpdateCaption(F);

  TobViewer1 := TTobViewer(getcontrol('TV'));
  TobViewer1.OnDblClick := TVOnDblClickCell;

  if (ctxMode in V_PGI.PGIContexte) then
  begin
    SetControlProperty('GL_DEPOT', 'Plus', 'GDE_SURSITE="X"');
    SetControlProperty('GL_ETABLISSEMENT', 'Plus', 'ET_SURSITE="X"');
  end;
  if not VH_GC.GCMultiDepots then
  begin
    SetControlVisible('GL_DEPOT', False);
    SetControlVisible('TGL_DEPOT', False);
  end;

  {$IFDEF NOMADE}
  stPlus := ' AND GPP_VENTEACHAT="' + LibType + '"';
  if LibType = 'VEN' then
  begin
    stPlus := stPlus + ' AND (GPP_NATUREPIECEG IN ("CC","DE") OR GPP_TRSFVENTE="X")';
    stNatDef := 'CC';
  end else
  begin
    stPlus := stPlus + ' AND (GPP_NATUREPIECEG="CF" OR GPP_TRSFACHAT="X")';
    stNatDef := 'CF';
  end;
  SetControlProperty('GL_NATUREPIECEG', 'PLUS', stPlus);
  SetControlText('GL_NATUREPIECEG', stNatDef);
  {$ELSE}
  if LibType = 'VEN' then
  begin
    if (VH_GC.GCSeria = True) or (V_PGI.VersionDemo = True)
      then SetControlText('GL_NATUREPIECEG', 'FAC;AVC;FFO')
    else SetControlText('GL_NATUREPIECEG', 'FFO');
  end else
  begin
    SetControlText('GL_NATUREPIECEG', 'BLF');
  end;
  {$ENDIF}

  // Paramétrage des libellés des tables libres article
  if (GCMAJChampLibre(F, False, 'COMBO', 'GL_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESLIBRES', False);
  x := 0;
  if (GCMAJChampLibre(F, False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else inc(x);
  if (GCMAJChampLibre(F, False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else inc(x);
  if (GCMAJChampLibre(F, False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else inc(x);
  {$IFNDEF CCS3}
  if (x = 0) then
    {$ENDIF}
    SetControlVisible('PZONESLIBRES', False);

  // Paramétrage des libellés des tables libres tiers
  if (GCMAJChampLibre(F, False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '') = 0) then SetControlVisible('PTABLESLIBRESTIERS', False);
  x := 0;
  if (GCMAJChampLibre(F, False, 'EDIT', 'YTC_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL1', False) else inc(x);
  if (GCMAJChampLibre(F, False, 'EDIT', 'YTC_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE1', False) else inc(x);
  if (GCMAJChampLibre(F, False, 'BOOL', 'YTC_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL1', False) else inc(x);
  if (GCMAJChampLibre(F, False, 'EDIT', 'YTC_TEXTELIBRE', 3, '') = 0) then SetControlVisible('GB_TEX1', False) else inc(x);
  {$IFNDEF CCS3}
  if (x = 0) then
    {$ENDIF}
    SetControlVisible('PZONESLIBRESTIERS', False);
end;

procedure TOF_MBOVentesVue.TVOnDblClickCell(Sender: TObject);
var codearticle, depot: string;
  naturepieceg, souche: string;
  numero: integer;
  TOBDim: TOB;
begin
  // Affichage de la fenêtre Dimension si Article générique.

  with TTobViewer(sender) do
  begin
    codearticle := AsString[ColIndex('GL_CODEARTICLE'), CurrentRow];

    if not ExisteSQL('Select GA_ARTICLE from ARTICLE where GA_CODEARTICLE="' + codearticle + '"' +
      ' and GA_STATUTART="GEN"') then
    begin
      SetLastError(1, '');
      exit;
    end;

    naturepieceg := AsString[ColIndex('GL_NATUREPIECEG'), CurrentRow];
    souche := AsString[ColIndex('GL_SOUCHE'), CurrentRow];
    numero := AsInteger[ColIndex('GL_NUMERO'), CurrentRow];
    depot := AsString[ColIndex('GL_DEPOT'), CurrentRow];

    TOBDim := TOB.Create('', nil, -1);
    RemplirTOBDim(naturepieceg, souche, numero, codearticle, TOBDim);

    if TOBDim.Detail.Count > 0 then
    begin
      TheTob := TOBDim;
      //AglLanceFiche ('GC','GCMSELECTDIMDOC','','', 'GA_CODEARTICLE='+codearticle+';ACTION=CONSULT;CHAMP= ;TOP=90;LEFT=80;OU= ;TYPEPARAM='+TheTob.GetValue('GP_VENTEACHAT')+';CODEPARAM='+TheTob.GetValue('GP_NATUREPIECEG')+';DEPOT='+Depot+';HEIGTH=400;WIDTH=650') ;
      AglLanceFiche('GC', 'GCMSELECTDIMDOC', '', '', 'GA_CODEARTICLE=' + codearticle +
        ';ACTION=CONSULT;CHAMP= ;TOP=90;LEFT=80;OU= ;TYPEPARAM=TVV;CODEPARAM=TVV;DEPOT=' + Depot + ';HEIGTH=400;WIDTH=650');
      TheTob.Free;
      TheTob := nil;
    end else
    begin
      TOBDim.Free;
    end;
  end;
end;

procedure TOF_MBOVentesVue.RemplirTOBDim(naturepieceg, souche: string; numero: Integer; codeart: string; TOBDim: TOB);
var jDim: integer;
  TOBPiece: TOB;
  TOBL, TOBD: TOB;
  Q: TQuery;
begin
  Q := OpenSQL('Select GL_ARTICLE,GL_QTEFACT,GL_QTERELIQUAT,GL_PUHTDEV,GL_PUTTCDEV,GL_REMISELIGNE ' +
    'from LIGNE where GL_NATUREPIECEG="' + naturepieceg + '" AND ' +
    'GL_SOUCHE="' + souche + '" AND GL_NUMERO="' + IntToStr(numero) + '" AND ' +
    'GL_CODEARTICLE="' + codeart + '" AND GL_TYPELIGNE="ART" ' +
    'order by GL_ARTICLE', True);
  if Q.EOF then
  begin
    // Pas de ligne pour cet article
    Ferme(Q);
    exit;
  end;
  TOBPiece := TOB.Create('', nil, -1);
  TOBPiece.LoadDetailDB('LIGNE', '', '', Q, False, True);
  Ferme(Q);

  Q := OpenSQL('Select GA_ARTICLE from ARTICLE where GA_CODEARTICLE="' + CodeArt + '" AND GA_STATUTART="DIM" order by GA_ARTICLE', True);
  while not Q.EOF do
  begin
    TOBD := TOB.Create('', TOBDim, -1);
    TOBD.AddChampSup('GA_ARTICLE', False);
    TOBD.PutValue('GA_ARTICLE', Q.FindField('GA_ARTICLE').AsString);
    TOBD.AddChampSup('GL_QTEFACT', False);
    TOBD.PutValue('GL_QTEFACT', 0);
    TOBD.AddChampSup('GL_QTERELIQUAT', False);
    TOBD.PutValue('GL_QTERELIQUAT', 0);
    TOBD.AddChampSup('GL_PUHTDEV', False);
    TOBD.PutValue('GL_PUHTDEV', 0);
    TOBD.AddChampSup('GL_PUTTCDEV', False);
    TOBD.PutValue('GL_PUTTCDEV', 0);
    TOBD.AddChampSup('GL_REMISELIGNE', False);
    TOBD.PutValue('GL_REMISELIGNE', 0);
    TOBD.AddChampSup('_QteDejaSaisi', False);
    TOBD.PutValue('_QteDejaSaisi', 0);
    TOBD.AddChampSup('_ReliqDejaSaisi', False);
    TOBD.PutValue('_ReliqDejaSaisi', 0);
    Q.Next;
  end;
  ferme(Q);

  if TOBDim.Detail.Count > 0 then
  begin
    for jDim := 0 to TOBDim.Detail.Count - 1 do
    begin
      TOBD := TOBDim.Detail[jDim];
      TOBL := TOBPiece.findfirst(['GL_ARTICLE'], [TOBD.GetValue('GA_ARTICLE')], false);
      while TOBL <> nil do
      begin
        TOBD.PutValue('GL_QTEFACT', TOBD.GetValue('GL_QTEFACT') + TOBL.GetValue('GL_QTEFACT'));
        TOBD.PutValue('GL_QTERELIQUAT', TOBD.GetValue('GL_QTERELIQUAT') + TOBL.GetValue('GL_QTERELIQUAT'));
        TOBD.PutValue('GL_PUHTDEV', TOBL.GetValue('GL_PUHTDEV'));
        TOBD.PutValue('GL_PUTTCDEV', TOBL.GetValue('GL_PUTTCDEV'));
        TOBD.PutValue('GL_REMISELIGNE', TOBD.GetValue('GL_REMISELIGNE') + TOBL.GetValue('GL_REMISELIGNE'));
        TOBD.PutValue('_QteDejaSaisi', TOBD.GetValue('GL_QTEFACT'));
        TOBL := TOBPiece.findnext(['GL_ARTICLE'], [TOBD.GetValue('GA_ARTICLE')], false);
      end;
    end;

    TOBDim.AddChampSup('GP_DEVISE', False);
    TOBDim.AddChampSup('GP_TAUXDEV', False);
    TOBDim.AddChampSup('GP_SAISIECONTRE', False);
    TOBDim.AddChampSup('GP_FACTUREHT', False);
    TOBDim.AddChampSup('GP_VENTEACHAT', False);
    TOBDim.AddChampSup('GP_NATUREPIECEG', False);
    TOBDim.AddChampSup('GP_DATEPIECE', False); // recup tarif
    TOBDim.AddChampSup('GP_TIERS', False); // recup tarif achat
    TOBDim.AddChampSup('GP_TARIFTIERS', False); // recup tarif tiers
    TOBDim.PutValue('GP_SAISIECONTRE', '-');
    TOBDim.PutValue('GP_FACTUREHT', '-');

    Q := OpenSQL('Select GP_DEVISE,GP_TAUXDEV,GP_SAISIECONTRE,GP_FACTUREHT,GP_VENTEACHAT,' +
      'GP_NATUREPIECEG,GP_DATEPIECE,GP_TIERS,GP_TARIFTIERS ' +
      'from PIECE where GP_NATUREPIECEG="' + naturepieceg + '" AND ' +
      'GP_SOUCHE="' + souche + '" AND GP_NUMERO="' + IntToStr(numero) + '"', True);
    if not Q.EOF then
    begin
      TOBDim.PutValue('GP_DEVISE', Q.FindField('GP_DEVISE').AsString);
      TOBDim.PutValue('GP_TAUXDEV', Q.FindField('GP_TAUXDEV').AsFloat);
      TOBDim.PutValue('GP_SAISIECONTRE', Q.FindField('GP_SAISIECONTRE').AsString);
      TOBDim.PutValue('GP_FACTUREHT', Q.FindField('GP_FACTUREHT').AsString);
      TOBDim.PutValue('GP_VENTEACHAT', Q.FindField('GP_VENTEACHAT').AsString);
      TOBDim.PutValue('GP_NATUREPIECEG', Q.FindField('GP_NATUREPIECEG').AsString);
      TOBDim.PutValue('GP_DATEPIECE', Q.FindField('GP_DATEPIECE').AsDateTime);
      TOBDim.PutValue('GP_TIERS', Q.FindField('GP_TIERS').AsString);
      TOBDim.PutValue('GP_TARIFTIERS', Q.FindField('GP_TARIFTIERS').AsString);
    end;
    ferme(Q);
  end;

  TOBPiece.Free;
end;

procedure TOF_listeDepotStat(parms: array of variant; nb: integer);
var F: TTobViewer;
  listeetab, etab, ListeDepot, ListeDepotsav, NomDepot: string;
begin
  F := TTobViewer(Longint(Parms[0]));
  if (F.Name <> 'ACHATVUE') and (F.Name <> 'VENTEVUE') and
    (F.Name <> 'ACHATCUB') and (F.Name <> 'VENTECUB') then exit;
  listeetab := Parms[1];
  NomDepot := Parms[2];
  ListeDepotsav := '';
  // Comme c'est un multi combo il faut traiter tous les établissements, si au moins un des établissements
  // choisi est lié à tous les dépôts alors on a toute la liste des dépôts, sinon on a que les dépôts
  // présent dans l'ensemble des dépôtlié des établissements.
  etab := readTokenSt(listeetab);
  while etab <> '' do
  begin
    // Appel de la fonction renvoyant la liste des dépôts liés à l'établissement
    ListeDepot := ListeDepotParEtablissement(etab);
    if ListeDepot = '' then
    begin
      etab := '';
      ListeDepotsav := '';
    end
    else
    begin
      if ListeDepotsav <> '' then ListeDepotsav := ListeDepotsav + ',';
      ListeDepotsav := ListeDepotsav + ListeDepot;
      etab := readTokenSt(listeetab);
    end;
  end;
  ListeDepot := ListeDepotsav;
  if ListeDepot <> '' then
  begin
    if THMultiValComboBox(F.FindComponent(NomDepot)).Plus <> '' then
      THMultiValComboBox(F.FindComponent(NomDepot)).Plus := THMultiValComboBox(F.FindComponent(NomDepot)).Plus + ' AND ';
    THMultiValComboBox(F.FindComponent(NomDepot)).Plus := THMultiValComboBox(F.FindComponent(NomDepot)).Plus + 'GDE_DEPOT in (' + ListeDepot + ')';
  end
  else
  begin
    if (ctxMode in V_PGI.PGIContexte) then THMultiValComboBox(F.FindComponent(NomDepot)).Plus := 'GDE_SURSITE="X"' // On prend tous les dépôts gérer sur site
    else THMultiValComboBox(F.FindComponent(NomDepot)).Plus := '';
  end;
end;

procedure TOF_MBOVentesVue.OnUpdate;
var StSQL, StWhere, xx_where, stCaption, iDim, sep: string;
  stSelectDim, stDimensions, stGroupByDim, stSelectStock, stGroupByStock, stJointureStock, stOrderByDim: string;
  stDim: array of string;
  iCol: integer;
  F: TFStat;
begin
  inherited;
  F := TFStat(Ecran);
  stSelectStock := '';
  stGroupByStock := '';
  stJointureStock := '';
  stSQL := '';
  sep := '';
  stSelectDim := '';
  stDimensions := ' ';
  stGroupByDim := ' ';
  stOrderByDim := '';
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
        stSelectDim := stSelectDim + 'GDI' + iDim + '.GDI_RANG,' + 'GDI' + iDim + '.GDI_LIBELLE as ' + StrToAlias(stCaption) + ',GA_GRILLEDIM' + iDim + ',';
        stDimensions := stDimensions + stDim[iCol - 1];
        stGroupByDim := stGroupByDim + 'GDI' + iDim + '.GDI_RANG,' + 'GDI' + iDim + '.GDI_LIBELLE ' + ',GA_GRILLEDIM' + iDim + ',';
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

  // si consultation des stocks
  if TCheckBox(GetControl('CBSTOCKS')).Checked then
  begin
    stSelectStock := 'SUM(GQ_PHYSIQUE) STOCK_PHY,SUM(GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU-GQ_PREPACLI) STOCK_NET,SUM(GQ_STOCKMIN) STOCK_MIN,SUM(GQ_STOCKMAX) STOCK_MAX,';
    //stGroupByStock := ',GQ_PHYSIQUE,GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU-GQ_PREPACLI,GQ_STOCKMIN,GQ_STOCKMAX';
    stJointureStock := 'LEFT JOIN DISPO ON GQ_ARTICLE=GL_ARTICLE AND GQ_DEPOT=GL_DEPOT AND GQ_CLOTURE="-" ';
  end;
  StSql := 'SELECT ' + stSelectDim + 'GL_NATUREPIECEG,GL_DATEPIECE,GL_SOUCHE,' + //GL_NUMERO,' +
    'GL_DEPOT,GL_DEVISE,GL_ETABLISSEMENT,GL_FOURNISSEUR,GL_CAISSE,GL_CODEARTICLE,GA_LIBELLE,GL_COLLECTION,' +
    'GL_FAMILLENIV1,GL_FAMILLENIV2,GL_FAMILLENIV3,GL_FAMILLETAXE1,GL_FAMILLETAXE2,' +
    'GL_DATELIVRAISON,GL_REPRESENTANT,GL_APPORTEUR,GL_TIERS,GL_TIERSFACTURE,GL_TIERSLIVRE,GL_TIERSPAYEUR,' +
    'T_AUXILIAIRE,T_LIBELLE as NOM_TIERS,T_PAYS,GL_SOCIETE,GL_TYPEREMISE,' +
    'MAX(GL_LIBREART1) GL_LIBREART1,MAX(GL_LIBREART2) GL_LIBREART2,MAX(GL_LIBREART3) GL_LIBREART3,MAX(GL_LIBREART4) GL_LIBREART4,MAX(GL_LIBREART5) GL_LIBREART5,' +
    'MAX(GL_LIBREART6) GL_LIBREART6,MAX(GL_LIBREART7) GL_LIBREART7,MAX(GL_LIBREART8) GL_LIBREART8,MAX(GL_LIBREART9) GL_LIBREART9,MAX(GL_LIBREARTA) GL_LIBREARTA,' +
    'GL_TYPEARTICLE,GL_TYPEDIM,GL_TYPELIGNE,' + stSelectStock +
    'SUM(GL_QTEFACT) Quantite, SUM(GL_QTERESTE) QteReste, SUM(GL_ESCOMPTE)/COUNT(*) Escompte,' +
    'SUM(GL_REMISELIGNE)/COUNT(*) Remise,' +
    'SUM(GL_QTERESTE * GL_PUHT * (100-GL_REMISELIGNE)/100) MONTANTRESTEAVECREMISE,' +
    'SUM(GL_QTERESTE * GL_PUHT) MONTANTRESTE,SUM (GL_MONTANTHT) MONTANTHT,' +
    'SUM (GL_MONTANTHTDEV) MONTANTHTDEV,SUM (GL_MONTANTTTC) MONTANTTTC,' +
    'SUM (GL_MONTANTTTCDEV) MONTANTTTCDEV,SUM (GL_TOTALHT) TOTALHT,' +
    'SUM (GL_TOTALHTDEV) TOTALHTDEV,SUM (GL_TOTALTTC) TOTALTTC,SUM (GL_TOTALTTCDEV) TOTALTTCDEV,' +
    'SUM (GL_TOTALTAXE1) TOTALTAXE1,SUM (GL_TOTALTAXE2) TOTALTAXE2,SUM (GL_TOTALTAXEDEV1) TOTALTAXEDEV1,' +
    'SUM (GL_TOTALTAXEDEV2) TOTALTAXEDEV2,SUM (GL_TOTESCLIGNE) TOTESCLIGNE,SUM (GL_TOTESCLIGNEDEV) TOTESCLIGNEDEV,' +
    'SUM (GL_TOTREMLIGNE) TOTREMLIGNE,SUM (GL_TOTREMLIGNEDEV) TOTREMLIGNEDEV,SUM (GL_TOTREMPIED) TOTREMPIED,' +
    'SUM (GL_TOTREMPIEDDEV) TOTREMPIEDDEV,SUM (GL_VALEURREMDEV) VALEURREMDEV FROM LIGNE ' +
    'LEFT JOIN ARTICLE ON GA_ARTICLE=GL_ARTICLE ' +
    stJointureStock +
    'LEFT JOIN TIERS ON T_TIERS=GL_TIERS ' +
    'LEFT JOIN TIERSCOMPL ON YTC_AUXILIAIRE=GL_TIERS' +
    stDimensions +
    ' WHERE GL_TYPELIGNE="ART" ' +
    'GROUP BY ' + stGroupByDim + ' GL_NATUREPIECEG,GL_DATEPIECE,GL_SOUCHE,' + //GL_NUMERO,' +
    'GL_DEPOT,GL_DEVISE,GL_ETABLISSEMENT,GL_FOURNISSEUR,GL_CAISSE,' +
    'GL_CODEARTICLE,GA_LIBELLE,GL_COLLECTION,' +
    'GL_FAMILLENIV1,GL_FAMILLENIV2,GL_FAMILLENIV3,' +
    'GL_FAMILLETAXE1,GL_FAMILLETAXE2,' +
    'GL_DATELIVRAISON,GL_REPRESENTANT,GL_APPORTEUR,' +
    'GL_TIERS,GL_TIERSFACTURE,GL_TIERSLIVRE,GL_TIERSPAYEUR,' +
    'T_AUXILIAIRE,T_LIBELLE,T_PAYS,GL_SOCIETE,GL_TYPEREMISE,' +
    'GL_TYPEARTICLE,GL_TYPEDIM,GL_TYPELIGNE' + stOrderByDim;

  // Récupération des zones libres articles
  xx_where := '';
  xx_where := GCXXWhereChampLibre(F, xx_where, 'BOOL', 'GA_BOOLLIBRE', 3, '');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'DATE', 'GA_DATELIBRE', 3, '_');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'EDIT', 'GA_VALLIBRE', 3, '_');
  // Récupération des zones libres tiers
  xx_where := GCXXWhereChampLibre(F, xx_where, 'BOOL', 'YTC_BOOLLIBRE', 3, '');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'DATE', 'YTC_DATELIBRE', 3, '_');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'EDIT', 'YTC_VALLIBRE', 3, '_');
  xx_where := GCXXWhereChampLibre(F, xx_where, 'STRING', 'YTC_TEXTELIBRE', 3, '');
  SetControlText('XX_WHERE', xx_where);

  // Récupération des autres critères
  stWhere := RecupWhereCritere(TPageControl(GetControl('Pages')));
  if Uppercase(copy(stWhere, 1, 6)) = 'WHERE ' then
  begin
    StWhere := Copy(StWhere, 7, length(StWhere) - 6);
    StSQL := InsertSQLWhere(StSQL, StWhere);
  end;
  F.StSQL := StSQL;

end;

initialization
  registerclasses([TOF_MBOVentesVue]);
  RegisterAglProc('ListeDepotStat', True, 3, TOF_listeDepotStat);
end.
