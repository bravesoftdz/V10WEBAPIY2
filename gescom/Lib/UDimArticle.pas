unit UDimArticle;

interface
uses {$IFDEF VER150} variants,{$ENDIF} StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, HStatus,
  HCtrls, HEnt1, HMsgBox, Hdimension, UTOB, UTOM, AGLInit, ParamSoc, EntGC,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, DBGrids, Fe_Main,
  {$ENDIF}
  UtilArticle, UtilGC, TarifUtil,
  math;

const MaxDimChamp = 10;
  NAT_ARTICLE = 'ART'; // Articles
  NAT_STOCK = 'STO'; // Stocks
  NAT_ACHAT = 'ACH'; // Achats
  NAT_VENTE = 'VEN'; // Ventes
  NAT_TARIF = 'TAF'; // Tarifs
  NAT_TRANSFERT = 'TRF'; // Transferts
  NAT_INVENTAIRE = 'INV'; // Inventaires
  NAT_ETIQART = 'ETA'; // Etiquettes article
  NAT_PROPTRF = 'PTR'; // Propositions de transfert : paramétrage stock mini / maxi
  NAT_PROPSAI = 'PSA'; // Propositions de transfert : saisie
  NAT_PROPDEP = 'PDE'; // Propositions de transfert : consultation dépôt
  NAT_ENTREXC = 'EEX'; // Entrée de stock exceptionnelle
  NAT_SORTEXC = 'SEX'; // Sortie de stock exceptionnelle
  NAT_STOVTE = 'STV'; // Stocks et ventes
  NAT_TOBVIEW = 'TVV'; // TobViewer Ventes et Achats
  NAT_ARRETESTO = 'ARS'; // Arreté de stock
  NAT_AFFCDEAFF = 'ACF'; // Affectation des commandes clients
  NAT_AFFCDERES = 'ACR'; // Reservation des commandes clients
  NAT_VENTESFFO = 'FFO'; // Ventes FFO

  NATPIE_ARTICLE = 'art';
  NATPIE_STOCK = 'sto';

type TODimArticle = class
    Dim: THDimension;
    TableDim: THDimensionItemList;
    TOBArticleDim: TOB;
    TOBArtGen: TOB; // TOB temporaire pour traitement de maj des prix des articles dimensionnés
    CodeArticle: string;
    MasqueDim: string;
    Depot, Cloture: string;
    ListeNatPiece, DatePieceDeb, DatePieceFin: string;
    DimDepot: string; // Champ affiché dans le THDIM pour la dimension DEPOT : tablette GCDIMDEPOT
    NatureDoc: string; // DC 25/01/01
    CodeSerie: string; // Propositions de transferts
    PrixUnique, bMasqueDepot: boolean;
    MasquePosDepot: string; // Position du dépôt dans un masque multi-dépôts
    //pas utilise StockNull : boolean ; //Utilisé pour les TRF
    IsArticle: Boolean;
    OldDimChamp: array[1..MaxDimChamp] of string;
    NewDimChamp: array[1..MaxDimChamp] of string; // param USERPREFDIM
    HideUnUsed: boolean; // param USERPREFDIM
    SauveParamGrille: boolean; // param USERPREFDIM
    TypeMasqueDim: string; // param USERPREFDIM
    {$IFDEF HDIM}
    AffDimDet, AffDimTotDeb, AffDimTotFin: array[1..MaxDimension] of boolean;
    AffTotVal: array[1..MaxDimChamp] of boolean;
    TobParamDim: TOB; // Tob des paramètres du THDIM (définis dans la table USERPREFDIM).
    {$ELSE}
    {$ENDIF}
    DimPlus: array[1..MaxDimChamp] of string;
    TobPrefEtab: TOB; // Tob des établissements définis dans la table USERPREFDIM.
    TobDimensions: TOB; // Tob pré-chargée avant appel de la fonction.
    // Utilisée par : proposition de transfert, Arrêté de stock.
    DepotEmet: string; // Proposition de transfert : code dépôt émetteur
    FAction: TActionFiche;
    SupItemDim: Boolean; //Variable temporaire
    TarifChampsModifiables: string; // Tarif : Champ modifiable -> 1 seul pour l'instant ...
    ConserveTobArtDim: boolean; // Defaut=False ; si True, TobArticleDim n'est pas détruit lors du destroy (récup. par TheTob)
    bChargtOptimise: boolean; // Chargement optimisé = dimensions existantes uniqt ?
    DimDecQte, DimDecPrix : string; // Nb de décimales des qté et prix

  private
    LibChampsDim, ChampsDim: TStringList;
    procedure UpdateArticleDim(ArtDim: Tob);
    procedure InitListeChampDim(TableChampsDim, TableChampsDimPlus: string);

    procedure OnVoir(Sender: TObject);
    procedure SetAction(const Value: TActionFiche);
    procedure LibereDim;
  public
    constructor create(FDim: THDimension; stArticle, stMasque, stChampAffiche, TableChampsDim, stNatureDoc, naturePiece, stDepot, stCloture: string;
      bPrixUnique: boolean;
      stListeNatPiece: string = ''; stDatePieceDeb: string = ''; stDatePieceFin: string = ''; bZeroDim: boolean = False);
    property Action: TActionFiche read FAction write SetAction;
    destructor Destroy; override;
    procedure SetTypeDonnee(NomChamp: string; NumChamp: integer);
    procedure SetReadOnly(NomChamp: string; NumChamp: integer; prixUnique: Boolean);
    procedure ChangeChampDimMul(Chargement: boolean);
    procedure MiseEnForme;
    procedure RefreshGrid(HideUnUsed: boolean);
    procedure LoadUsersPref(naturePiece: string);
    procedure LoadDimensions(LoadInitial: boolean);
    function UpdateDB(ArtGen, ArtCompl: TOB; var ArtError: string): integer;
    function isModified: boolean;
    function CombienModifie: integer;
    procedure refresh(ItemDim: THDimensionItem);
    procedure CreateItemDim(ItemDim: THDimensionItem; TobArtGen: TOB);
    function  DeleteItemDim(ItemDim: THDimensionItem; Article: string) : boolean;
    procedure MajPrixUniques(TobArt: TOB);
  end;

function FormateChamp(CodeArticle: string; PCB: Double): string;

implementation

uses
  FactUtil,
  UtilDimArticle,
  {$IFDEF GPAO}
    wGammeTet,
  {$ENDIF}
  UTomArticle
  ;

procedure TODimArticle.InitListeChampDim(TableChampsDim, TableChampsDimPlus: string);
begin
  if LibChampsDim <> nil then
  begin
    LibChampsDim.free;
    LibChampsDim := nil;
  end;
  if ChampsDim <> nil then
  begin
    ChampsDim.free;
    ChampsDim := nil;
  end;
  LibChampsDim := TStringList.create;
  ChampsDim := TStringList.create;
  RemplirValCombo(TableChampsDim, TableChampsDimPlus, '', LibChampsDim, ChampsDim, False, False);
end;

{ Paramètres à fournir :
 NAT_PROPTRF : - stNatureDoc = NAT_PROPTRF
               - stArticle = code serie
 NAT_ETIQART : - stNatureDoc = NAT_ETIQART
               - stArticle = code article
               - stMasque = masque de l'article
 NAT_PROPSAI,
 NAT_PROPDEP :
               - stNatureDoc = NAT_PROPSAI, NAT_PROPDEP
               - stArticle = code article
               - stChampAffiche = 1 si multi-dépôt
                                = 2 si mono-dépôt -> stNatureDoc devient NAT_PROPDEP
                                                      pour sauvegarde du paramétrage
                                                      et champs "DEP_" disponibles
               - stDepot = dépôt sélectionné -> visu si masque mono-dépôt
               - TheTob contient : -> les articles dimensionnés
                                   -> les qtés de l'établissement émetteur
                                   -> les qtés de l'établissement récepteur si mono-dépôt
                                      ou les qtés des établissements récepteurs si multi-dépôt
                 La fonction ajoute les codes et grilles dimensions et pcb aux tobs filles.
 NAT_TARIF   : - stNatureDoc = NAT_TARIF
               - stChampAffiche = nom du champ en saisie possible, read only pour les autres.
 NAT_ARRETESTO : - stNatureDoc = NAT_ARRETESTO
               - stArticle = code article
               - TheTob contient : -> les articles dimensionnés
                                   -> les qtés de la table DISPO recalculées à la date d'arrêté.
                 La tob contient les codes et grilles dimensions.
}

constructor TODimArticle.create(FDim: THDimension; stArticle, stMasque, stChampAffiche, TableChampsDim, stNatureDoc, naturePiece, stDepot, stCloture: string;
  bPrixUnique: boolean;
  stListeNatPiece: string = ''; stDatePieceDeb: string = ''; stDatePieceFin: string = ''; bZeroDim: boolean = False);

var iChamp, i : integer;
  TobTransf: TOB;
begin
  inherited create;
  Depot := stDepot;
  Cloture := stCloture;
  ListeNatPiece := stListeNatPiece;
  DatePieceDeb := stDatePieceDeb;
  DatePieceFin := stDatePieceFin;
  MasqueDim := stMasque;
  bMasqueDepot := False;
  MasquePosDepot := '';
  ConserveTobArtDim := False;
  NatureDoc := stNatureDoc;
  //if (NatureDoc=NAT_PROPSAI) and (stChampAffiche='2') then BEGIN NatureDoc:=NAT_PROPDEP ; NaturePiece:=NatureDoc END ;
  if NaturePiece = '' then NaturePiece := NatureDoc; // Affectation valeur par défaut à NaturePiece
  Dim := Fdim;
  // DCA - FQ MODE 11033 - Formatage des qtés et montants dans l'objet dimension
  DimDecQte := '';
  for i := 1 to GetParamSoc('SO_DECQTE') do DimDecQte := DimDecQte + '0';
  DimDecPrix := '';
  for i := 1 to GetParamSoc('SO_DECPRIX') do DimDecPrix := DimDecPrix + '0';
  Dim.Formats.ReelDisplayFormat := '# ##0.' + DimDecPrix + ';;# ##0.' + DimDecPrix;
  PrixUnique := bPrixUnique;
  IsArticle := (natureDoc = NAT_ARTICLE);
  InitListeChampDim(TableChampsDim, natureDoc);
  Dim.OnVoir := OnVoir; // par defaut
  bChargtOptimise := (GetParamSoc('SO_CHARGEDIMDEGRADE') and not bZeroDim);
  // DCA - FQ MODE 10138 - NAT_ETIQART On passe dans le cas standard pour charger les preferences utilisateur
  if (NatureDoc = NAT_PROPTRF) then
  begin
    if NatureDoc = NAT_PROPTRF then CodeSerie := stArticle
    else if NatureDoc = NAT_ETIQART then CodeArticle := stArticle;
    TypeMasqueDim := VH_GC.BOTypeMasque_Defaut;
    HideUnUsed := (NatureDoc = NAT_ETIQART);
    {$IFDEF HDIM}
    TobParamDim := nil;
    {$ELSE}
    {$ENDIF}
    TobPrefEtab := nil;
    // Chargement des champs affichés dans le THDIM
    for iChamp := 1 to MaxDimChamp do NewDimChamp[iChamp] := '';
    Dim.TitreValeur := '';
    ControlePresenceChampsObligatoires(Self, natureDoc);
  end else
  begin
    CodeArticle := stArticle;
    if natureDoc = NAT_TARIF then TarifChampsModifiables := stChampAffiche
    else if (natureDoc = NAT_PROPSAI) or (natureDoc = NAT_PROPDEP) then
    begin
      TobDimensions := TheTob;
      // Recherche du dépôt émetteur dans la tob : présence obligatoire
      TobTransf := TobDimensions.FindFirst(['GTL_DEPOTDEST'], ['...'], False);
      if TobTransf <> nil then DepotEmet := TobTransf.GetValue('GDE_DEPOT') else DepotEmet := '';
    end
    else if (natureDoc = NAT_ARRETESTO) or (natureDoc = NAT_AFFCDEAFF) or (natureDoc = NAT_AFFCDERES) then TobDimensions := TheTob;
    LoadUsersPref(naturePiece); // Chargement de TobParamDim
  end;
  LoadDimensions(True);
end;

procedure TODimArticle.LoadUsersPref(naturePiece: string);
var QQ, QM: TQuery;
  TobFille: TOB;
  iChamp, position: integer;
  bEtabOk: boolean;
  stSql, LibDepot: string;
begin
  // Définition des valeurs par défaut.
  {$IFDEF HDIM}
  for iChamp := 1 to MaxDimChamp do
  begin
    NewDimChamp[iChamp] := '';
    AffTotVal[iChamp] := False;
  end;
  for iChamp := 1 to MaxDimension do
  begin
    AffDimTotDeb[iChamp] := False;
    AffDimTotFin[iChamp] := False;
    AffDimDet[iChamp] := True;
  end;
  {$ELSE}
  for iChamp := 1 to MaxDimChamp do NewDimChamp[iChamp] := '';
  {$ENDIF}
  Dim.TitreValeur := '';
  HideUnUsed := Boolean(natureDoc <> NAT_ARTICLE);
  SauveParamGrille := True;
  if GetParamSoc('SO_GCARTTYPEMASQUE') then TypeMasqueDim := VH_GC.BOTypeMasque_Defaut else TypeMasqueDim := '';
  {$IFDEF HDIM}
  TobParamDim := nil; // Chargement de tous les établissements (en fin de fct).
  {$ELSE}
  {$ENDIF}
  TobPrefEtab := nil; // Par défaut, chargement de tous les établissements (en fin de fct).

  // Paramétrage des dimensions par nature de pièce.
  // DCA - FQ MODE 10138 - NAT_ETIQART On passe dans le cas standard pour charger les preferences utilisateur
  if (NatureDoc = NAT_ETIQART) or NaturePieceGeree(naturePiece) then
  begin
    {Lecture des préférences utilisateurs}
    stSql := 'select GUD_POSITION,GUD_CHAMP,GUD_LIBELLE,GUD_DETAIL,GUD_DEFAUT,' +
      ' GUD_TYPEMASQUE,GUD_ETABLISSEMENT,GDE_LIBELLE,GDE_ABREGE' +
      ' from USERPREFDIM' +
      ' left join DEPOTS on GDE_DEPOT=GUD_ETABLISSEMENT' +
      ' where GUD_NATUREPIECE="' + naturePiece + '" and GUD_UTILISATEUR="' + V_PGI.User + '"' +
      ' order by GUD_POSITION';
    QQ := OpenSQL(stSql, True,-1,'',true);
    if not QQ.Eof then while not QQ.EOF do
      begin
        position := QQ.FindField('GUD_POSITION').AsInteger;
        if position = 0 then
        begin
          DimDepot := QQ.FindField('GUD_CHAMP').AsString;
          if DimDepot = '' then DimDepot := 'GDE_ABREGE';
          HideUnUsed := Boolean(QQ.FindField('GUD_DETAIL').AsString <> 'X');
          SauveParamGrille := Boolean(QQ.FindField('GUD_DEFAUT').AsString = 'X');
          if QQ.FindField('GUD_TYPEMASQUE').AsString <> '' then TypeMasqueDim := QQ.FindField('GUD_TYPEMASQUE').AsString;
        end
        else if (position > 0) and (position < 11) then
        begin
          NewDimChamp[position] := QQ.FindField('GUD_CHAMP').AsString;
          if Dim.TitreValeur <> '' then Dim.TitreValeur := Dim.TitreValeur + ';';
          Dim.TitreValeur := Dim.TitreValeur + QQ.FindField('GUD_LIBELLE').AsString;
          {$IFDEF HDIM}
          // Affichage du total des 1..10 valeurs
          AffTotVal[position - 10] := Boolean(QQ.FindField('GUD_DETAIL').AsString = 'X');
        end
        else if (position > 10) and (position < 16) then
        begin // AffDimTotDeb[1..5]
          AffDimTotDeb[position - 10] := Boolean(QQ.FindField('GUD_DETAIL').AsString = 'X');
        end
        else if (position > 15) and (position < 21) then
        begin // AffDimTotFin[1..5]
          AffDimTotFin[position - 15] := Boolean(QQ.FindField('GUD_DETAIL').AsString = 'X');
        end
        else if (position > 20) and (position < 26) then
        begin // AffDimDet[1..5]
          AffDimDet[position - 20] := Boolean(QQ.FindField('GUD_DETAIL').AsString = 'X');
          {$ENDIF}
        end
        else if position > 100 then
        begin
          if TobPrefEtab = nil then TobPrefEtab := TOB.Create('Liste établissements', nil, -1);
          TobFille := TOB.Create('', TobPrefEtab, -1);
          TobFille.AddChampSup('GUD_ETABLISSEMENT', False);
          TobFille.AddChampSup('GDE_DEPOT', False);
          TobFille.AddChampSup('GDE_LIBELLE', False);
          TobFille.AddChampSup('GDE_ABREGE', False);
          TobFille.PutValue('GUD_ETABLISSEMENT', QQ.FindField('GUD_ETABLISSEMENT').AsString);
          TobFille.PutValue('GDE_DEPOT', QQ.FindField('GUD_ETABLISSEMENT').AsString);
          TobFille.PutValue('GDE_LIBELLE', QQ.FindField('GDE_LIBELLE').AsString);
          TobFille.PutValue('GDE_ABREGE', QQ.FindField('GDE_ABREGE').AsString);
        end;
        QQ.next;
      end
    else if natureDoc = NAT_PROPSAI then
      begin // Définition des options par défaut
      // Utilisation du premier masque multi-dépôt trouvé (si existe)
      stSql := 'select GMQ_TYPEMASQUE from TYPEMASQUE' +
        ' where GMQ_MULTIETAB="X"' +
        ' order by GMQ_TYPEMASQUE';
      QM := OpenSQL(stSql, True,-1,'',true);
      if not QM.Eof then TypeMasqueDim := QM.FindField('GMQ_TYPEMASQUE').AsString;
      Ferme(QM);
      if DimDepot = '' then DimDepot := 'GDE_ABREGE';
      end;
    Ferme(QQ);
  end;
  ControlePresenceChampsObligatoires(Self, natureDoc);

  // Préférences utilisateurs non définies -> défaut = tous les établissements.
  // NAT_PROPSAI -> Chargement des dépôts concernés en tob pour maj combo dimension
  if (TobPrefEtab = nil) and
    ((natureDoc = NAT_ARTICLE) or (natureDoc = NAT_STOCK) or
    (natureDoc = NAT_PROPSAI)) then
  begin
    bEtabOk := True;
    stSql := 'select GDE_DEPOT,GDE_LIBELLE,GDE_ABREGE' +
      ' from DEPOTS' +
      ' order by GDE_DEPOT';
    QQ := OpenSQL(stSql, True,-1,'',true);
    while not QQ.EOF do
    begin
      if natureDoc = NAT_PROPSAI then
      begin
        // Sélection des établissements présents dans la tob des propositions
        bEtabOk := Boolean(TobDimensions.FindFirst(['GDE_DEPOT'], [QQ.FindField('GDE_DEPOT').AsString], False) <> nil);
      end;
      if bEtabOk then
      begin
        if TobPrefEtab = nil then TobPrefEtab := TOB.Create('Liste établissements', nil, -1);
        TobFille := TOB.Create('', TobPrefEtab, -1);
        TobFille.AddChampSup('GUD_ETABLISSEMENT', False);
        TobFille.AddChampSup('GDE_DEPOT', False);
        TobFille.AddChampSup('GDE_LIBELLE', False);
        TobFille.AddChampSup('GDE_ABREGE', False);
        TobFille.PutValue('GUD_ETABLISSEMENT', QQ.FindField('GDE_DEPOT').AsString);
        TobFille.PutValue('GDE_DEPOT', QQ.FindField('GDE_DEPOT').AsString);
        TobFille.PutValue('GDE_LIBELLE', QQ.FindField('GDE_LIBELLE').AsString);
        TobFille.PutValue('GDE_ABREGE', QQ.FindField('GDE_ABREGE').AsString);
      end;
      QQ.next;
    end;
    ferme(QQ);
    if natureDoc = NAT_PROPSAI then
    begin
      // Positionnement de l'établissement émetteur en 1ère position dans la tob -> 1ère position dans le THDIM
      TobFille := TobPrefEtab.FindFirst(['GUD_ETABLISSEMENT'], [DepotEmet], False);
      if TobFille <> nil then
      begin
        LibDepot := TraduireMemoire('Dépôt:');
        TobFille.PutValue('GUD_ETABLISSEMENT', TobFille.GetValue('GUD_ETABLISSEMENT'));
        TobFille.PutValue('GDE_DEPOT', LibDepot + TobFille.GetValue('GDE_DEPOT'));
        TobFille.PutValue('GDE_LIBELLE', LibDepot + TobFille.GetValue('GDE_LIBELLE'));
        TobFille.PutValue('GDE_ABREGE', LibDepot + TobFille.GetValue('GDE_ABREGE'));
        TobFille.ChangeParent(TobPrefEtab, 0);
      end;
    end;
  end;
end;

procedure TODimArticle.LoadDimensions(LoadInitial: boolean);
var QQ, QMASQUE, QTM: TQuery;
  Position: array[1..7] of string;
  SQL, NomChamp, stTmp, stDim: string;
  ItemDim: THDimensionItem;
  I, iTob, iChamp, iEtab, iDim: Integer;
  MasquePosition, ArtDimValue, MasqueMultiEtab: string;
  ArtDim, TobDim: TOB;
  AddChampSup, NouveauChamp, SupChampSup, PrefixeCodeDim: string; //AC
  jointure_stock, jointure_ligne, jointure_artcompl: boolean;
  AddGlPiecePrecedente: Boolean;
begin
  LibereDim; // Free de TableDim et TobArticleDim
  AddGlPiecePrecedente := False; { NEWPIECE }
  if (MasqueDim <> '') or (NatureDoc = NAT_PROPSAI) or (NatureDoc = NAT_PROPDEP) or
    (NatureDoc = NAT_ARRETESTO) or (NatureDoc = NAT_AFFCDEAFF) or (NatureDoc = NAT_AFFCDERES) then
  begin
    if Dim = nil then exit;
    for iChamp := 1 to MaxDimChamp do
    begin
      OldDimChamp[iChamp] := NewDimChamp[iChamp];
      SetTypeDonnee(NewDimChamp[iChamp], iChamp);
    end;

    if ((NatureDoc = NAT_PROPSAI) or (NatureDoc = NAT_PROPDEP) or (NatureDoc = NAT_AFFCDEAFF) or (NatureDoc = NAT_AFFCDERES)) and (TobDimensions <> nil) then
    begin // Ajout des dimensions à la tob passée en paramètre
      if LoadInitial then
      begin
        if TobDimensions.Detail.Count > 0 then
        begin
          for i := 1 to 5 do
          begin
            TobDimensions.Detail[0].AddChampSup('GA_CODEDIM' + IntToStr(i), True);
            TobDimensions.Detail[0].AddChampSup('GA_GRILLEDIM' + IntToStr(i), True);
          end;
          TobDimensions.Detail[0].AddChampSup('GA_PCB', True);
        end;
        SQL := 'select GA_ARTICLE,GA_CODEARTICLE,GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5'
          + ',GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5'
          + ',GA_PCB,GA_DIMMASQUE'
          + ' from ARTICLE where GA_CODEARTICLE="' + CodeArticle + '" and GA_STATUTART="DIM"';
        QQ := OpenSQL(SQL, True,-1,'',true);
        while not QQ.EOF do
        begin
          MasqueDim := QQ.FindField('GA_DIMMASQUE').AsString;
          if (NatureDoc = NAT_AFFCDEAFF) or (NatureDoc = NAT_AFFCDERES) then TobDim := TobDimensions.FindFirst(['GA_ARTICLE'], [QQ.FindField('GA_ARTICLE').AsString], False)
          else TobDim := TobDimensions.FindFirst(['GTL_ARTICLE'], [QQ.FindField('GA_ARTICLE').AsString], False);
          while TobDim <> nil do
          begin
            for i := 1 to 5 do
            begin
              TobDim.PutValue('GA_CODEDIM' + IntToStr(i), QQ.FindField('GA_CODEDIM' + IntToStr(i)).AsString);
              TobDim.PutValue('GA_GRILLEDIM' + IntToStr(i), QQ.FindField('GA_GRILLEDIM' + IntToStr(i)).AsString);
            end;
            TobDim.PutValue('GA_PCB', QQ.FindField('GA_PCB').AsString);
            if (NatureDoc = NAT_AFFCDEAFF) or (NatureDoc = NAT_AFFCDERES) then TobDim := TobDimensions.FindNext(['GA_ARTICLE'], [QQ.FindField('GA_ARTICLE').AsString], False)
            else TobDim := TobDimensions.FindNext(['GTL_ARTICLE'], [QQ.FindField('GA_ARTICLE').AsString], False);
          end;
          QQ.next;
        end;
        Ferme(QQ);
        TobDimensions.SetAllModifie(False);
      end; // if LoadInitial then
    end; // if ((NatureDoc=NAT_PROPSAI) or (NatureDoc=NAT_PROPDEP)) and (TobDimensions<>nil) then
    if GetParamSoc('SO_GCARTTYPEMASQUE') then
    begin
      // Gestion du type de masque
      // Si le type de masque n'est pas celui par défaut on vérifie s'il est multi-dépôt
      // Si c'est le cas, on vérifie qu'il existe du stock.
      // Si pas de stock, on force le type de masque à celui par défaut (NON multi-dépôt)
      if TypeMasqueDim <> VH_GC.BOTypeMasque_Defaut then
      begin
        QTM := OpenSQL('select GMQ_MULTIETAB from TYPEMASQUE where GMQ_TYPEMASQUE="' + TypeMasqueDim + '"', True,-1,'',true);
        if not QTM.Eof then MasqueMultiEtab := QTM.FindField('GMQ_MULTIETAB').AsString else MasqueMultiEtab := '';
        Ferme(QTM);
        if MasqueMultiEtab = 'X' then
        begin
          // Cas d'un masque multi-dépôt : contrôle existence d'infos dans le stock
          if not ExisteSQL('select GQ_ARTICLE from DISPO where GQ_ARTICLE in (select GA_ARTICLE from ARTICLE where GA_STATUTART="DIM" and GA_CODEARTICLE="' +
            CodeArticle + '")')
            then TypeMasqueDim := VH_GC.BOTypeMasque_Defaut;
        end;
      end;

      QMasque := OpenSQL('select * from DIMMASQUE where GDM_MASQUE="' + MasqueDim + '" and GDM_TYPEMASQUE="' + TypeMasqueDim + '"', TRUE,-1,'',true);
      if (QMasque.EOF) and (TypeMasqueDim <> VH_GC.BOTypeMasque_Defaut) then
      begin
        Ferme(QMasque);
        QMasque := OpenSQL('select * from DIMMASQUE where GDM_MASQUE="' + MasqueDim + '" and GDM_TYPEMASQUE="' + VH_GC.BOTypeMasque_Defaut + '"', TRUE,-1,'',true);
      end;
    end
    else QMasque := OpenSQL('select * from DIMMASQUE where GDM_MASQUE="' + MasqueDim + '"', TRUE,-1,'',true);
    if not QMasque.EOF then
    begin
      for i := 1 to MaxDimension + 1 do Position[i] := QMasque.Findfield('GDM_POSITION' + IntToStr(i)).AsString;

      for iDim := 1 to MaxDimension do DimPlus[iDim] := '';
      if bChargtOptimise then // ModeDegrade
      begin
        // Fiche article + chargt optimisé -> détail décoché
        if NatureDoc = NAT_ARTICLE then HideUnUsed := True;

        // Chargement des clauses Plus des combos des grilles de dimensions
        // pour limiter l'utilisation des seuls codes utilisés par l'article.
        for iDim := 1 to MaxDimension do if position[iDim] <> '' then
          begin
            stTmp := IntToStr(iDim);
            SQL := 'select distinct GDI_CODEDIM from ARTICLE ' +
              'left join DIMENSION on GDI_TYPEDIM="DI' + stTmp +
              '" and GDI_GRILLEDIM=GA_GRILLEDIM' + stTmp + ' and GDI_CODEDIM=GA_CODEDIM' + stTmp +
              ' where GA_STATUTART="DIM" and GA_CODEARTICLE="' + CodeArticle + '"';
            QQ := OpenSQL(SQL, True,-1,'',true);
            if QQ.EOF then
            begin // Aucune dimension existante pour cet article -> on force min
              Ferme(QQ);
              SQL := 'select min(GDI_CODEDIM) as GDI_CODEDIM from DIMENSION ' +
                ' where GDI_TYPEDIM="DI' + stTmp +
                '" and GDI_GRILLEDIM="' + QMasque.Findfield('GDM_TYPE' + stTmp).AsString + '"';
              QQ := OpenSQL(SQL, True,-1,'',true);
            end;
            if not QQ.EOF then
            begin
              stTmp := '';
              while not QQ.EOF do
              begin
                if stTmp <> '' then stTmp := stTmp + ',';
                stTmp := stTmp + '"' + QQ.FindField('GDI_CODEDIM').AsString + '"';
                QQ.Next;
              end;
              DimPlus[iDim] := 'AND GDI_CODEDIM IN (' + stTmp + ') ';
            end;
            Ferme(QQ);
          end;
      end;
      Ferme(QMasque);

      TableDim := THDimensionItemList.create;
      AddChampSup := '';
      bMasqueDepot := Boolean(Position[6] <> '');
      jointure_stock := False;
      jointure_ligne := False;
      jointure_artcompl := False;
      PrefixeCodeDim := 'GA_CODEDIM';
      if (NatureDoc = NAT_PROPSAI) or (NatureDoc = NAT_PROPDEP) or (NatureDoc = NAT_ARRETESTO) or (NatureDoc = NAT_AFFCDEAFF) or (NatureDoc = NAT_AFFCDERES) then
      begin // TobArticleDim déjà chargée
        if (NatureDoc = NAT_PROPSAI) and (not bMasqueDepot) then
        begin
          // Si masque mono-dépôt, on ne reprend dans la tob que les articles du dépôt concerné
          TOBArticleDim := TOB.Create('prop transf mono depot', nil, -1);
          TobDim := TobDimensions.FindFirst(['GDE_DEPOT'], [Depot], True);
          while TobDim <> nil do
          begin
            TobDim.ChangeParent(TOBArticleDim, -1);
            TobDim := TobDimensions.FindNext(['GDE_DEPOT'], [Depot], True);
          end;
        end else
        begin
          // Sinon on reprend la tob telle quelle
          TOBArticleDim := TobDimensions;
        end;
      end
      else if NatureDoc = NAT_PROPTRF then
      begin
        // Jointure DIMMASQUE - DIMENSION pour proposer toutes
        {SQL:='select GDM_MASQUE as GAM_DIMMASQUE,'+
             'GDI1.GDI_CODEDIM as GAM_CODEDIM1,GDI1.GDI_GRILLEDIM as GAM_GRILLEDIM1,'+
             'GDI2.GDI_CODEDIM as GAM_CODEDIM2,GDI2.GDI_GRILLEDIM as GAM_GRILLEDIM2,'+
             'GDI3.GDI_CODEDIM as GAM_CODEDIM3,GDI3.GDI_GRILLEDIM as GAM_GRILLEDIM3,'+
             'GDI4.GDI_CODEDIM as GAM_CODEDIM4,GDI4.GDI_GRILLEDIM as GAM_GRILLEDIM4,'+
             'GDI5.GDI_CODEDIM as GAM_CODEDIM5,GDI5.GDI_GRILLEDIM as GAM_GRILLEDIM5,'+
             '"'+CodeSerie+'" as GAM_CODESERIE,GAM_QTEDISPOMINI,GAM_QTEDISPOMAXI'+
             ' from DIMMASQUE '+
             'left join DIMENSION GDI1 on GDM_TYPE1=GDI1.GDI_GRILLEDIM and GDI1.GDI_TYPEDIM="DI1" '+
             'left join DIMENSION GDI2 on GDM_TYPE2=GDI2.GDI_GRILLEDIM and GDI2.GDI_TYPEDIM="DI2" '+
             'left join DIMENSION GDI3 on GDM_TYPE3=GDI3.GDI_GRILLEDIM and GDI3.GDI_TYPEDIM="DI3" '+
             'left join DIMENSION GDI4 on GDM_TYPE4=GDI4.GDI_GRILLEDIM and GDI4.GDI_TYPEDIM="DI4" '+
             'left join DIMENSION GDI5 on GDM_TYPE5=GDI5.GDI_GRILLEDIM and GDI5.GDI_TYPEDIM="DI5" '+
             'left join ARTICLESERIE on GDM_MASQUE=GAM_DIMMASQUE and GAM_CODESERIE="'+CodeSerie+'" and '+
             '(GDI1.GDI_CODEDIM is null or GDI1.GDI_CODEDIM=GAM_CODEDIM1) and'+
             '(GDI2.GDI_CODEDIM is null or GDI2.GDI_CODEDIM=GAM_CODEDIM2) and'+
             '(GDI3.GDI_CODEDIM is null or GDI3.GDI_CODEDIM=GAM_CODEDIM3) and'+
             '(GDI4.GDI_CODEDIM is null or GDI4.GDI_CODEDIM=GAM_CODEDIM4) and'+
             '(GDI5.GDI_CODEDIM is null or GDI5.GDI_CODEDIM=GAM_CODEDIM5) '+
             'where GDM_MASQUE="'+MasqueDim+'"' ;}

        SQL := 'select GDM_MASQUE as GAM_DIMMASQUE,' +
          'GDI1.GDI_CODEDIM as GAM_CODEDIM1,GDI1.GDI_GRILLEDIM as GAM_GRILLEDIM1,' +
          'GDI2.GDI_CODEDIM as GAM_CODEDIM2,GDI2.GDI_GRILLEDIM as GAM_GRILLEDIM2,' +
          'GDI3.GDI_CODEDIM as GAM_CODEDIM3,GDI3.GDI_GRILLEDIM as GAM_GRILLEDIM3,' +
          'GDI4.GDI_CODEDIM as GAM_CODEDIM4,GDI4.GDI_GRILLEDIM as GAM_GRILLEDIM4,' +
          'GDI5.GDI_CODEDIM as GAM_CODEDIM5,GDI5.GDI_GRILLEDIM as GAM_GRILLEDIM5,' +
          '"' + CodeSerie + '" as GAM_CODESERIE,0.0 as GAM_QTEDISPOMINI,0.0 as GAM_QTEDISPOMAXI ' +
          'from DIMMASQUE ' +
          'left join DIMENSION GDI1 on GDM_TYPE1=GDI1.GDI_GRILLEDIM and GDI1.GDI_TYPEDIM="DI1" ' +
          'left join DIMENSION GDI2 on GDM_TYPE2=GDI2.GDI_GRILLEDIM and GDI2.GDI_TYPEDIM="DI2" ' +
          'left join DIMENSION GDI3 on GDM_TYPE3=GDI3.GDI_GRILLEDIM and GDI3.GDI_TYPEDIM="DI3" ' +
          'left join DIMENSION GDI4 on GDM_TYPE4=GDI4.GDI_GRILLEDIM and GDI4.GDI_TYPEDIM="DI4" ' +
          'left join DIMENSION GDI5 on GDM_TYPE5=GDI5.GDI_GRILLEDIM and GDI5.GDI_TYPEDIM="DI5" ' +
          'where GDM_MASQUE="' + MasqueDim + '"';
        if GetParamSoc('SO_GCARTTYPEMASQUE') then SQL := SQL + ' and GDM_TYPEMASQUE="' + VH_GC.BOTypeMasque_Defaut + '"';
        PrefixeCodeDim := 'GAM_CODEDIM';
      end else
      begin
        // Modif BTP Rajout de GA_TYPEARTICLE pour gestion des articles en prix posés
        SQL := 'select GA_ARTICLE,GA_CODEARTICLE,GA_TYPEARTICLE,GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5'
          + ',GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5'
          + ',GA_STATUTART,GA_QUALIFCODEBARRE,GA_PCB,GA_CODEBARRE as ZGA_CODEBARRE'
          + ',GA_TARIFARTICLE';
        for i := 0 to ChampsDim.count - 1 do
        begin
          if (copy(ChampsDim[i], 1, 3) = 'GQ_') then
          begin
            jointure_stock := True;
            if ChampsDim[i] = 'GQ__STOCKNET' // Champ calculé
            then SQL := SQL + ', GQ_PHYSIQUE+GQ_RESERVEFOU-GQ_RESERVECLI-GQ_PREPACLI as GQ__STOCKNET'
            else SQL := SQL + ',' + ChampsDim[i];
          end
          else if (copy(ChampsDim[i], 1, 4) = 'GA2_') then
          begin
            if GetPresentation = ART_ORLI then
            begin
              jointure_artcompl := True;
              SQL := SQL + ',' + ChampsDim[i];
            end; // Sinon champ ignoré
          end
          else if copy(ChampsDim[i], 1, 3) <> 'GA_' then
          begin
            if ((NatureDoc = NAT_STOVTE) and (copy(ChampsDim[i], 1, 3) = 'GL_')) then
            begin // Cas consultation stocks et lignes de vente
              jointure_ligne := True;
              SQL := SQL + ',' + ChampsDim[i];
            end
            else if (NatureDoc = NAT_ETIQART) and (ChampsDim[i] = 'NBETIQ') then // Champ ignoré cf + bas
            else
            begin
              if AddChampSup <> '' then AddChampSup := AddChampSup + ';' + ChampsDim[i] //AC
              else AddChampSup := ChampsDim[i];
              { NEWPIECE }
              if (copy(ChampsDim[i], 1, 3) = 'GL_') and (not AddGlPiecePrecedente) then
              begin
                 AddChampSup := AddChampSup + ';GL_PIECEPRECEDENTE';
                 AddGlPiecePrecedente := True;
              end;
            end;
          end
          else SQL := SQL + ',' + ChampsDim[i]; //AC
        end;
        if NatureDoc = NAT_ETIQART then
        begin
          SQL := SQL + ',0 as NBETIQ from ARTICLE where GA_CODEARTICLE="' + CodeArticle + '" and GA_STATUTART="DIM"';
        end
        else if bMasqueDepot then
        begin
          SQL := SQL + ',GDE_DEPOT from ARTICLE,DEPOTS,DISPO';
          if jointure_artcompl then SQL := SQL + ',ARTICLECOMPL';
          SQL := SQL + ' where GDE_DEPOT in (';
          // Ajout clause where sur la liste des dépôts de USERPREFDIM
          if TobPrefEtab <> nil then
          begin
            for iEtab := 0 to TobPrefEtab.Detail.count - 1 do
            begin
              SQL := SQL + '"' + TobPrefEtab.Detail[iEtab].GetValue('GUD_ETABLISSEMENT') + '"';
              if iEtab < TobPrefEtab.Detail.count - 1 then SQL := SQL + ',';
            end;
          end else SQL := SQL + '"' + VH_GC.GCDepotDefaut + '"';
          SQL := SQL + ') and GA_ARTICLE=GQ_ARTICLE and GQ_CLOTURE="-" and GQ_DEPOT=GDE_DEPOT and GA_CODEARTICLE="' + CodeArticle +
            '" and GA_STATUTART="DIM" ';
          if jointure_artcompl then SQL := SQL + 'and GA2_ARTICLE=GA_ARTICLE ';
        end
        else if jointure_stock or jointure_ligne then
        begin
          if Depot = '' then Depot := VH_GC.GCDepotDefaut;
          SQL := SQL + ' from ARTICLE';
          if jointure_artcompl then
          begin
            SQL := SQL + ' left join ARTICLECOMPL on GA2_ARTICLE=GA_ARTICLE';
          end;
          if jointure_stock then
          begin
            SQL := SQL + ' left join DISPO on GQ_ARTICLE=GA_ARTICLE and GQ_DEPOT="' + Depot + '"';
            if Cloture <> '' then SQL := SQL + ' and GQ_CLOTURE="' + Cloture + '"';
          end;
          if jointure_ligne then
          begin
            SQL := SQL + ' left join LIGNE on GL_ARTICLE=GA_ARTICLE and GL_DEPOT="' + Depot + '"';
            if ListeNatPiece <> '' then SQL := SQL + ' and GL_NATUREPIECEG in (' + ListeNatPiece + ')';
            if DatePieceDeb <> '' then SQL := SQL + ' and GL_DATEPIECE>="' + USDateTime(StrToDate(DatePieceDeb)) + '"';
            if DatePieceFin <> '' then SQL := SQL + ' and GL_DATEPIECE<="' + USDateTime(StrToDate(DatePieceFin)) + '"';
          end;
          SQL := SQL + ' where GA_CODEARTICLE="' + CodeArticle + '" and GA_STATUTART="DIM"'
        end
        else SQL := SQL + ' from ARTICLE Where GA_CODEARTICLE="' + CodeArticle + '" AND GA_STATUTART="DIM"';
      end;
      if (NatureDoc <> NAT_PROPSAI) and (NatureDoc <> NAT_PROPDEP) and (NatureDoc <> NAT_ARRETESTO) and (NatureDoc <> NAT_AFFCDEAFF) and (NatureDoc <>
        NAT_AFFCDERES) then
      begin
        QQ := OpenSQL(SQL, True,-1,'',true);
        if QQ.EOF then
        begin
          if IsArticle then
          begin
            if HideUnused then HideUnused := False; // Pas de données -> forcer affichage du détail
            {$IFDEF HDIM}
            ShowDimension(MasqueDim, TypeMasqueDim, Dim, TobPrefEtab, DimDepot, nil, AffDimDet, AffDimTotFin); //vide
            {$ELSE}
            ShowDimension(MasqueDim, TypeMasqueDim, Dim, TobPrefEtab, DimDepot, nil, DimPlus); //vide
            {$ENDIF}
          end;
          Ferme(QQ);
          exit; //PGIBox('Cet article ne possède pas de dimensions, vous ne pouvez pas le sélectionner','Attention')
        end;
        TOBArticleDim := TOB.CreateDB('_ARTICLESDIM', nil, -1, QQ);
        TOBArticleDim.LoadDetailDB('_ARTICLESDIM', '', '', QQ, False);
        Ferme(QQ);
        //
        // Cas des stocks minimums : Recherche dans la table ARTICLESERIE
        //
        if NatureDoc = NAT_PROPTRF then
        begin
          {$IFDEF MAUVAISE_SOLUTION}
          for iTob := 0 to TOBArticleDim.Detail.Count - 1 do
          begin
            SQL := 'select GAM_QTEDISPOMINI,GAM_QTEDISPOMAXI ' +
              'from ARTICLESERIE ' +
              'where GAM_DIMMASQUE="' + TOBArticleDim.Detail[iTob].GetValue('GAM_DIMMASQUE') +
              '" and GAM_CODESERIE="' + TOBArticleDim.Detail[iTob].GetValue('GAM_CODESERIE') + '"';
            for iDim := 1 to 5 do
            begin
              stDim := IntToStr(iDim);
              // Attention, si null !?
              if (TOBArticleDim.Detail[iTob].GetValue('GAM_CODEDIM' + stDim) <> null) and
                (TOBArticleDim.Detail[iTob].GetValue('GAM_CODEDIM' + stDim) <> '') then
              begin
                SQL := SQL + ' and GAM_CODEDIM' + stDim + '="' + TOBArticleDim.Detail[iTob].GetValue('GAM_CODEDIM' + stDim) + '"';
              end;
            end;
            QQ := OpenSQL(SQL, True,-1,'',true);
            if not QQ.EOF then
            begin
              TOBArticleDim.Detail[iTob].PutValue('GAM_QTEDISPOMINI', QQ.FindField('GAM_QTEDISPOMINI').AsFloat);
              TOBArticleDim.Detail[iTob].PutValue('GAM_QTEDISPOMAXI', QQ.FindField('GAM_QTEDISPOMAXI').AsFloat);
            end;
            Ferme(QQ);
          end;
          {$ENDIF}
          if TOBArticleDim.Detail.Count > 0 then
          begin
            // Maj des GAM_CODEDIM1..5 pour remplacer les null par des '' !
            for iTob := 0 to TOBArticleDim.Detail.Count - 1 do
              for iDim := 1 to 5 do
              begin
                stDim := IntToStr(iDim);
                if TOBArticleDim.Detail[iTob].GetValue('GAM_CODEDIM' + stDim) = null then
                  TOBArticleDim.Detail[iTob].PutValue('GAM_CODEDIM' + stDim, '');
              end;

            SQL := 'select GAM_CODEDIM1,GAM_CODEDIM2,GAM_CODEDIM3,GAM_CODEDIM4,GAM_CODEDIM5,' +
              'GAM_QTEDISPOMINI,GAM_QTEDISPOMAXI ' +
              'from ARTICLESERIE ' +
              'where GAM_DIMMASQUE="' + MasqueDim +
              '" and GAM_CODESERIE="' + CodeSerie + '"';
            QQ := OpenSQL(SQL, True,-1,'',true);
            while not QQ.EOF do
            begin
              TobDim := TOBArticleDim.FindFirst(['GAM_CODEDIM1', 'GAM_CODEDIM2', 'GAM_CODEDIM3', 'GAM_CODEDIM4', 'GAM_CODEDIM5'],
                [QQ.FindField('GAM_CODEDIM1').AsString, QQ.FindField('GAM_CODEDIM2').AsString,
                QQ.FindField('GAM_CODEDIM3').AsString, QQ.FindField('GAM_CODEDIM4').AsString,
                  QQ.FindField('GAM_CODEDIM5').AsString], False);
              if TobDim <> nil then
              begin
                TobDim.PutValue('GAM_QTEDISPOMINI', QQ.FindField('GAM_QTEDISPOMINI').AsFloat);
                TobDim.PutValue('GAM_QTEDISPOMAXI', QQ.FindField('GAM_QTEDISPOMAXI').AsFloat);
              end;
              QQ.Next;
            end;
            Ferme(QQ);
          end;
        end;
        //
        // FIN Cas des stocks minimums : Recherche dans la table ARTICLESERIE
        //

      end;
      if TOBArticleDim <> nil then for iTob := 0 to TOBArticleDim.Detail.Count - 1 do
        begin
          ArtDim := TOBArticleDim.Detail[iTob];
          ///////// debut AC
          SupChampSup := AddChampSup;
          while SupChampSup <> '' do
          begin
            NouveauChamp := uppercase(Trim(ReadTokenSt(SupChampSup)));
            ArtDim.AddChampSup(NouveauChamp, False);
            ArtDim.PutValue(NouveauChamp, SetInitTypeDonnee(NouveauChamp));
          end;
          //////// fin AC
          ItemDim := THDimensionItem.create;
          if GetParamSoc('SO_GCARTTYPEMASQUE') then
          begin
            for i := 1 to MaxDimension + 1 do
            begin
              MasquePosition := Position[i];
              if (i < MaxDimension + 1) or (bMasqueDepot) then
              begin
                if i <> MaxDimension + 1 then
                begin
                  NomChamp := PrefixeCodeDim + IntToStr(I);
                  if ArtDim.GetValue(NomChamp) <> Null
                    then ArtDimValue := ArtDim.GetValue(NomChamp)
                  else ArtDimValue := SetInitTypeDonnee(NomChamp);
                end
                else
                begin
                  ArtDimValue := ArtDim.GetValue('GDE_DEPOT');
                  MasquePosDepot := MasquePosition;
                end;
                if MasquePosition = 'ON1' then ItemDim.Ong := ArtDimValue
                else if MasquePosition = 'LI1' then ItemDim.Lig1 := ArtDimValue
                else if MasquePosition = 'LI2' then ItemDim.Lig2 := ArtDimValue
                else if MasquePosition = 'CO1' then ItemDim.Col1 := ArtDimValue
                else if MasquePosition = 'CO2' then ItemDim.Col2 := ArtDimValue
              end;
            end;
          end else
          begin
            for i := 1 to MaxDimension do
            begin
              MasquePosition := Position[i];
              if ArtDim.GetValue(PrefixeCodeDim + IntToStr(I)) <> Null
                then ArtDimValue := ArtDim.GetValue(PrefixeCodeDim + IntToStr(I))
              else ArtDimValue := SetInitTypeDonnee(PrefixeCodeDim + IntToStr(I));
              if MasquePosition = 'ON1' then ItemDim.Ong := ArtDimValue
              else if MasquePosition = 'LI1' then ItemDim.Lig1 := ArtDimValue
              else if MasquePosition = 'LI2' then ItemDim.Lig2 := ArtDimValue
              else if MasquePosition = 'CO1' then ItemDim.Col1 := ArtDimValue
              else if MasquePosition = 'CO2' then ItemDim.Col2 := ArtDimValue
            end;
          end;

          // Ajout champ sup pour gestion mise à jour de la table DISPO
          if not ArtDim.FieldExists('GDE_DEPOT') then ArtDim.AddChampSupValeur('GDE_DEPOT', Depot);

          // Ajout champ sup pour gestion mise à jour des tables ARTICLE, ARTICLECOMPL, DISPO
          ArtDim.AddChampSupValeur('_MAJARTICLE', False);
          ArtDim.AddChampSupValeur('_MAJARTICLECOMPL', False);
          ArtDim.AddChampSupValeur('_MAJDISPO', False);

          ItemDim.data := ArtDim;
          //JD Pour éviter que l'objet dimension ai l'état modifié au chargement
          TOB(ItemDim.data).modifie := False;
          TableDim.Add(ItemDim);
        end;
      ChangeChampDimMul(TRUE); // positionne les valeurs pour le(s) champ(s) à afficher
      //RefreshGrid(HideUnUsed) ;
    end
    else Ferme(Qmasque);
  end else
  begin // dimmasque non renseigné
    {$IFDEF HDIM}
    ShowDimension('', '', Dim, TobPrefEtab, DimDepot, nil, AffDimDet, AffDimTotFin); // pour effacer en attendant mieux
    {$ELSE}
    ShowDimension('', '', Dim, TobPrefEtab, DimDepot, nil, DimPlus); // pour effacer en attendant mieux
    {$ENDIF}
  end;
end;

procedure TODimArticle.MiseEnForme;
var iChamp: integer;
  //pcb: double;
  DimDec, stDisplayFormat: string ;
begin
  // DCA - FQ MODE 11033 - Formatage des qtés et montants dans l'objet dimension
  // Obsolète
  //if (NatureDoc <> NAT_PROPTRF) and (TOBARticleDim <> nil) and (TOBARticleDim.Detail.Count > 0)
  //  then pcb := Tob(TOBARticleDim.Detail[0]).GetValue('GA_PCB')
  //else pcb := 1.0;
  Dim.DisplayFormat := '';
  for iChamp := 1 to MaxDimChamp do
  begin
    OldDimChamp[iChamp] := NewDimChamp[iChamp];
    SetTypedonnee(NewDimChamp[iChamp], iChamp);
    if NewDimChamp[iChamp] <> '' then SetReadOnly(NewDimChamp[iChamp], iChamp, PrixUnique);
    stDisplayFormat := '';
    // Inclus GL_QTEFACT, GL_QTERESTE, GL_QTERELIQUAT
    if (copy(NewDimChamp[iChamp],1,6) = 'GL_QTE') then stDisplayFormat := '# ##0.' + DimDecQte  + ';'
    else
      case Dim.TypeDonnee[iChamp] of
        dotReel:
          begin
            if (NatureDoc = NAT_PROPTRF) or (NatureDoc = NAT_ETIQART) or
              (NatureDoc = NAT_PROPSAI) or (NatureDoc = NAT_PROPDEP) or
              (NatureDoc = NAT_AFFCDEAFF) or (NatureDoc = NAT_AFFCDERES) then
            begin
              if (NatureDoc = NAT_ETIQART) and (NewDimChamp[iChamp] <> 'NBETIQ')
                then stDisplayFormat := '# ##0.' + DimDecPrix + ';'
                else stDisplayFormat := '# ##0.' + DimDecQte  + ';';
            end else
            begin
              // ParCombien = 1 par défaut.
              //if (TOBARticleDim <> nil) and (TOBARticleDim.Detail.Count > 0)
              //  then pcb := Tob(TOBARticleDim.Detail[0]).GetValue('GA_PCB')
              //else pcb := 1.0;
              DimDec := nbDecimales(NewDimChamp[iChamp], 'DOUBLE', DimDecQte, DimDecPrix);
              stDisplayFormat := '# ##0.' + DimDec + ';';
            end;
          end;
        dotDate: stDisplayFormat := 'dd/mm/yy;';
        //dotString: Dim.DisplayFormat := Dim.DisplayFormat + ';';
      //else Dim.DisplayFormat := Dim.DisplayFormat + ';';
      end; // case Dim.TypeDonnee[iChamp] of
      // DCA - Mise en place ou suppression d'un format spécifique
      Dim.FormatDonnees.Delete(iChamp);
      if stDisplayFormat <> ''
        then Dim.FormatDonnees.Add(iChamp).ReelDisplayFormat := stDisplayFormat;
  end;
end;

procedure TODimArticle.ChangeChampDimMul(Chargement: boolean);
var i, iChamp: integer;
  ArtDim: Tob;
  ItemDim: THDimensionItem;
  ForceArtDimModifieToOldValue, OldValue: Boolean;
begin
  if tableDim = nil then exit;
  OldValue := False;
  for i := 0 to TableDim.count - 1 do
  begin
    if TableDim.Items[i] <> nil then
    begin
      ItemDim := THDimensionItem(TableDim.Items[i]);
      ArtDim := TOB(ItemDim.data);
      if ArtDim <> nil then
      begin
        for iChamp := 1 to MaxDimChamp do
        begin
          if (not Chargement) and (OldDimChamp[iChamp] <> '') and (ItemDim.valeur[iChamp] <> ArtDim.GetValue(OldDimChamp[iChamp])) then
          begin
            // ATTENTION, ArtDim.IsFieldModified() retourne False !!!
            ForceArtDimModifieToOldValue := (ArtDim.GetValue(OldDimChamp[iChamp]) = Null);
            if ForceArtDimModifieToOldValue then OldValue := ArtDim.Modifie;
            ArtDim.PutValue(OldDimChamp[iChamp], ItemDim.valeur[iChamp]);
            if ForceArtDimModifieToOldValue then ArtDim.Modifie := OldValue;
            // Gestion mise à jour des tables ARTICLE, ARTICLECOMPL, DISPO
            if Copy(OldDimChamp[iChamp], 1, 3) = 'GA_' then ArtDim.PutValue('_MAJARTICLE', True)
            else if Copy(OldDimChamp[iChamp], 1, 3) = 'GQ_' then ArtDim.PutValue('_MAJDISPO', True)
            else if Copy(OldDimChamp[iChamp], 1, 4) = 'GA2_' then ArtDim.PutValue('_MAJARTICLECOMPL', True);
          end;
          if (NewDimChamp[iChamp] <> '') then
          begin
            if ArtDim.GetValue(NewDimChamp[iChamp]) = Null then
            begin
              ItemDim.Valeur[iChamp] := SetInitTypeDonnee(NewDimChamp[iChamp]);
              OldValue := ArtDim.Modifie;
              ArtDim.PutValue(NewDimChamp[iChamp], ItemDim.valeur[iChamp]);
              ArtDim.Modifie := OldValue;
            end
            else ItemDim.Valeur[iChamp] := ArtDim.GetValue(NewDimChamp[iChamp])
          end
          else
            if not ItemDim.ToCreate then ItemDim.Valeur[iChamp] := 'X'
        end;
      end;
    end;
  end;

  MiseEnForme;
end;

procedure TODimArticle.RefreshGrid(HideUnUsed: boolean);
begin
  if MasqueDim = '' then exit;
  Dim.active := True;
  {$IFDEF HDIM}
  ShowDimension(MasqueDim, TypeMasqueDim, Dim, TobPrefEtab, DimDepot, TableDim,
    AffDimDet, AffDimTotFin);
  {$ELSE}
  ShowDimension(MasqueDim, TypeMasqueDim, Dim, TobPrefEtab, DimDepot, TableDim, DimPlus);
  {$ENDIF}
  Dim.HideUnUsed := HideUnUsed; // ShowDimension réinitialise HideUnUsed d'où maj ici.
end;

function TODimArticle.isModified: boolean;
var ItemDim: THDimensionItem;
  i: integer;
begin
  Result := False;
  ChangeChampDimMul(False);
  for i := 0 to TableDim.count - 1 do
  begin
    if TableDim.Items[i] <> nil then
    begin
      ItemDim := THDimensionItem(TableDim.Items[i]);
      if (ItemDim.ToCreate = TRUE) or (ItemDim.ToDelete = TRUE) or ((ItemDim.data <> nil) and (TOB(ItemDim.data).modifie)) then
      begin
        Result := TRUE; // pour passer DS.state en mode dsEdit
        exit;
      end;
    end;
  end;

end;

function TODimArticle.CombienModifie: integer;
var ItemDim: THDimensionItem;
  i: integer;
begin
  Result := 0;
  if TableDim = nil then exit;
  for i := 0 to TableDim.count - 1 do
  begin
    if TableDim.Items[i] <> nil then
    begin
      ItemDim := THDimensionItem(TableDim.Items[i]);
      if (ItemDim.ToCreate = TRUE) or (ItemDim.ToDelete = TRUE) or ((ItemDim.data <> nil) and (TOB(ItemDim.data).modifie)) then
        Result := Result + 1;
    end;
  end;

end;

procedure TODimArticle.SetTypeDonnee(NomChamp: string; NumChamp: integer);
var Typ: string;
begin
  if NomChamp <> '' then
  begin
    if NomChamp = 'NBETIQ' then Typ := 'INTEGER' // Etiquettes articles
    else if (copy(NomChamp, 1, 4) = 'GTL_') or (copy(NomChamp, 1, 4) = 'DEP_') or
      (NomChamp = 'GQ__STOCKNET') or // Champ calculé
      (NomChamp = 'STOCKFINAL') then Typ := 'DOUBLE' // Saisie propositions de transferts
    else Typ := ChampToType(NomChamp);
    if (Typ = 'DOUBLE') or (Typ = 'RATE') then
    begin
      Dim.TypeDonnee[NumChamp] := dotReel;
      Dim.Alignment[NumChamp] := taRightJustify
    end
    else if (Typ = 'INTEGER') or (Typ = 'SMALLINT') then
    begin
      Dim.TypeDonnee[NumChamp] := dotReel;
      Dim.Alignment[NumChamp] := taRightJustify
    end
    else if Typ = 'DATE' then
    begin
      Dim.TypeDonnee[NumChamp] := dotDate;
      Dim.Alignment[NumChamp] := taRightJustify
    end
    else
    begin
      Dim.TypeDonnee[NumChamp] := dotString;
      Dim.Alignment[NumChamp] := taLeftJustify
    end
  end
  else
  begin
    Dim.TypeDonnee[NumChamp] := dotString;
    Dim.Alignment[NumChamp] := taLeftJustify
  end;
end;

procedure TODimArticle.SetReadOnly(NomChamp: string; NumChamp: integer; PrixUnique: Boolean);
begin
  if IsArticle
    // DCA - FQ MODE 10813
    then Dim.ReadOnly[NumChamp] := not ChampModif(NomChamp, PrixUnique, True)
  else Dim.ReadOnly[NumChamp] := False;
  if (natureDoc <> NAT_ARTICLE) and
    ((Copy(NomChamp, 1, 3) = 'GA_') or (Copy(NomChamp, 1, 4) = 'GA2_') or (Copy(NomChamp, 1, 5) = 'GL_DP')) then Dim.ReadOnly[numChamp] := True;
  if (Copy(NomChamp, 1, 3) = 'GQ_') then Dim.ReadOnly[numChamp] := True;
  // DCA - FQ MODE 10813
  // Saisie des stocks mini/maxi depuis la fiche ARTICLE
  if ((natureDoc = NAT_ARTICLE) and ((NomChamp = 'GQ_STOCKMIN') or (NomChamp = 'GQ_STOCKMAX')))
    then Dim.ReadOnly[numChamp] := False
  else if (natureDoc = NAT_STOCK) or (natureDoc = NAT_STOVTE) or (natureDoc = NAT_TOBVIEW) then Dim.ReadOnly[numChamp] := True
  else if (natureDoc = NAT_PROPSAI) or (natureDoc = NAT_PROPDEP) then Dim.ReadOnly[numChamp] := Boolean((NomChamp <> 'GTL_PROPOSITION') or (Depot = DepotEmet))
  else if (natureDoc = NAT_ENTREXC) or (natureDoc = NAT_SORTEXC) then Dim.ReadOnly[numChamp] := Boolean(NomChamp <> 'GL_QTEFACT')
  else if (natureDoc = NAT_TARIF) then Dim.ReadOnly[numChamp] := Boolean(NomChamp <> TarifChampsModifiables)
  else if (natureDoc = NAT_AFFCDEAFF) then Dim.ReadOnly[numChamp] := Boolean(NomChamp <> 'GEL_QTEAFFECTEE')
  else if (natureDoc = NAT_AFFCDERES) then Dim.ReadOnly[numChamp] := Boolean(NomChamp <> 'GEL_QTERESERVEE');

  // DCA - FQ MODE 10625 et MODE 10955
  if  (NomChamp = 'GA_CODEBARRE')   or
      (NomChamp = 'GL_QTERESTE')    or
      (NomChamp = 'GL_QTERELIQUAT') or
      (NomChamp = 'GL_MTRESTE')     or
      (NomChamp = 'GL_MTRELIQUAT')  then
    begin
    if (NomChamp <> 'GA_CODEBARRE') then Dim.ReadOnly[numChamp] := True;
    Dim.CanCopy[NumChamp] := False;
    end ;
end;

function TODimArticle.UpdateDB(ArtGen, ArtCompl: TOB; var ArtError: string): integer;
var QMasque: TQuery;
  I, J, iChamp, CABok, iCAB: integer;
  b_result, PremDim, CABAuto: boolean;
  TobDim, TobArtCompl, TobDispo: TOB;
  ItemDim: THDimensionItem;
  Article, TypeCAB, CodeCAB, FournPrinc, DepotItem: string;
  {$IFDEF BTP}
  // modif BTP
  LibCodeDim, LIBArt: string;
  {$ENDIF}
  {$IFDEF GPAO}
    ArtNewDim: string;
  {$ENDIF}
begin
  Result := 0;
  CodeCAB := '';
  TypeCAB := '';
  ArtError := '';
  iCAB := -1;
  CABAuto := False;
  ChangeChampDimMul(False); // force la mise à jour de la TOB
  InitMove(CombienModifie, '');
  if GetParamSoc('SO_GCARTTYPEMASQUE')
    then QMasque := OpenSQL('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="' + MasqueDim + '" and GDM_TYPEMASQUE="' + TypeMasqueDim + '"', TRUE,-1,'',true)
  else QMasque := OpenSQL('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="' + MasqueDim + '"', TRUE,-1,'',true);
  if not QMasque.EOF then
  begin
    i := 0;
    PremDim := True;
    while i <= TableDim.count - 1 do
    begin
      if TableDim.Items[i] <> nil then
      begin
        ItemDim := THDimensionItem(TableDim.Items[i]);
        if ItemDim.ToCreate = TRUE then
        begin
          ItemDim.ToCreate := False;
          TobDim := TOB.Create('ARTICLE', nil, -1); // a revoir creer un TOB mère article générique et TOB fille dimensionné
          TobDim.dupliquer(ArtGen, True, True, True);
          TobDim.AddChampSup('ZGA_CODEBARRE', False);
          TobDim.PutValue('ZGA_CODEBARRE', TobDim.GetValue('GA_CODEBARRE'));
          TobDim.PutValue('GA_ARTICLE', CodeArticleUnique(ArtGen.GetValue('GA_CODEARTICLE'), AffecteDim(QMasque, 1, ItemDim), AffecteDim(QMasque, 2, ItemDim),
            AffecteDim(QMasque, 3, ItemDim), AffecteDim(QMasque, 4, ItemDim), AffecteDim(QMasque, 5, ItemDim)));
          {$IFDEF BTP}
          LibArt := '';
          {$ENDIF}
          {$IFDEF GPAO}
            ArtNewDim := CodeArticleUnique(ArtGen.GetValue('GA_CODEARTICLE'), AffecteDim(QMasque, 1, ItemDim), AffecteDim(QMasque, 2, ItemDim),
            AffecteDim(QMasque, 3, ItemDim), AffecteDim(QMasque, 4, ItemDim), AffecteDim(QMasque, 5, ItemDim));
          { Initialisation des gammes opératoires pour la ou les nouvelles dimensions }
          uInitGammesNewDim(ArtNewDim);
          {$ENDIF}
          for j := 1 to MaxDimension do
          begin
            TobDim.PutValue('GA_CODEDIM' + IntToStr(j), AffecteDim(QMasque, j, ItemDim));
            TobDim.PutValue('GA_GRILLEDIM' + IntToStr(j), QMasque.Findfield('GDM_TYPE' + IntToStr(j)).AsString);
            {$IFDEF BTP}
            LibCodeDim := GCGetCodeDim(TobDim.GetValue('GA_GRILLEDIM' + IntToStr(j)), TobDim.GetValue('GA_CODEDIM' + IntToStr(j)), j);
            if LibCodeDim <> '' then LibArt := LibArt + ' ' + LibCodeDim;
            {$ENDIF}
          end;
          {$IFDEF BTP}
          TOBDim.PutValue('GA_LIBELLE', TOBDIM.GetValue('GA_LIBELLE') + LibArt);
          {$ENDIF}
          TobDim.PutValue('GA_STATUTART', 'DIM');
          TobDim.PutValue('GA_DATECREATION', Date);
          TobDim.PutValue('GA_DATEMODIF', NowH);
          TobDim.PutValue('GA_CODEBARRE', '');

          {Maj des infos saisies dans la nouvelle TOB créée}
          for iChamp := 1 to MaxDimChamp do
          begin
            if (NewDimChamp[iChamp] <> '') and (copy(NewDimChamp[iChamp], 1, 3) = 'GA_') then
            begin
              case Dim.TypeDonnee[iChamp] of
                dotDate: TobDim.PutValue(NewDimChamp[iChamp], StrToDateTime(ItemDim.Valeur[iChamp]));
                dotReel: TobDim.PutValue(NewDimChamp[iChamp], StrToFloat(ItemDim.Valeur[iChamp]));
                dotString: TobDim.PutValue(NewDimChamp[iChamp], ItemDim.Valeur[iChamp]);
              else TobDim.PutValue(NewDimChamp[iChamp], ItemDim.Valeur[iChamp]);
              end;
            end;
          end;

          //*********** Attribution automatique des Codes à barres *****************
          if TobDim.GetValue('GA_CODEBARRE') = '' then
          begin
            if PremDim then
            begin
              // Recherche du paramétrage du code à barres
              PremDim := False;
              FournPrinc := TobDim.GetValue('GA_FOURNPRINC');
              CABAuto := RechParamCAB_Auto(FournPrinc, TypeCAB);
              // Recherche position du champ GA_CODEBARRE dans la liste des champs affichés
              iChamp := 1;
              while (iChamp < MaxDimChamp) and (NewDimChamp[iChamp] <> 'GA_CODEBARRE') do inc(iChamp);
              if NewDimChamp[iChamp] = 'GA_CODEBARRE' then iCAB := iChamp;
            end;
            if CABAuto then
            begin
              // Attribution du nouveau code à barres
              CodeCAB := AttribNewCodeCAB(GetParamSoc('SO_GCCABARTICLE'), FournPrinc);
              TobDim.PutValue('GA_CODEBARRE', CodeCAB);
              TobDim.PutValue('GA_QUALIFCODEBARRE', TypeCAB);
              // Maj infos tob
              Tob(ItemDim.data).PutValue('GA_CODEBARRE', CodeCAB);
              Tob(ItemDim.data).PutValue('ZGA_CODEBARRE', CodeCAB);
              if iCAB <> -1 then ItemDim.valeur[iCAB] := CodeCAB;
            end;
          end;

          CABok := ControleTobCAB(TobDim);
          if CABok = 0 then
          begin
            b_result := TobDim.InsertDB(nil);
            if GetPresentation = ART_ORLI then
            begin // Gestion ARTICLECOMPL
              TobArtCompl := TOB.Create('ARTICLECOMPL', nil, -1);
              TobArtCompl.Dupliquer(ArtCompl, False, True, False);
              TobArtCompl.PutValue('GA2_ARTICLE', TobDim.GetValue('GA_ARTICLE'));
              TobArtCompl.PutValue('GA2_CODEARTICLE', TobDim.GetValue('GA_CODEARTICLE'));

              // Maj des éventuels champs saisis à la création de l'article
              for iChamp := 1 to MaxDimChamp do
              begin
                if (NewDimChamp[iChamp] <> '') and (copy(NewDimChamp[iChamp], 1, 4) = 'GA2_') then
                begin
                  case Dim.TypeDonnee[iChamp] of
                    dotDate: TobArtCompl.PutValue(NewDimChamp[iChamp], StrToDateTime(ItemDim.Valeur[iChamp]));
                    dotReel: TobArtCompl.PutValue(NewDimChamp[iChamp], StrToFloat(ItemDim.Valeur[iChamp]));
                    dotString: TobArtCompl.PutValue(NewDimChamp[iChamp], ItemDim.Valeur[iChamp]);
                  else TobArtCompl.PutValue(NewDimChamp[iChamp], ItemDim.Valeur[iChamp]);
                  end;
                end;
              end;
              TobArtCompl.InsertOrUpdateDB(False);
              TobArtCompl.Free;
            end; // Gestion ARTICLECOMPL

            // Insertion enreg dans DISPO ?
            TobDispo := nil;
            iChamp := 1;
            while iChamp <= MaxDimChamp do
            begin
              if (NewDimChamp[iChamp] <> '') and (copy(NewDimChamp[iChamp], 1, 3) = 'GQ_') and
                (not Dim.ReadOnly[iChamp]) then
              begin // Il existe au moins 1 champ de la table DISPO modifiable
                // donc on crée l'enregistrement dans la table DISPO !
                TobDispo := TOB.Create('DISPO', nil, -1);
                with TobDispo do
                begin
                  PutValue('GQ_ARTICLE', TobDim.GetValue('GA_ARTICLE'));
                  if bMasqueDepot then
                  begin
                    if MasquePosDepot = 'LI1' then DepotItem := ItemDim.Lig1
                    else if MasquePosDepot = 'CO1' then DepotItem := ItemDim.Col1
                    else if MasquePosDepot = 'LI2' then DepotItem := ItemDim.Lig2
                    else if MasquePosDepot = 'CO2' then DepotItem := ItemDim.Col2
                    else if MasquePosDepot = 'ON1' then DepotItem := ItemDim.Ong;
                  end
                  else DepotItem := Depot;
                  PutValue('GQ_DEPOT', DepotItem);
                  PutValue('GQ_CLOTURE', '-');
                  PutValue('GQ_DPA', TobDim.GetValue('GA_DPA'));
                  PutValue('GQ_PMAP', TobDim.GetValue('GA_PMAP'));
                  PutValue('GQ_DPR', TobDim.GetValue('GA_DPR'));
                  PutValue('GQ_PMRP', TobDim.GetValue('GA_PMRP'));
                  PutValue('GQ_DATECREATION', NowH);
                  PutValue('GQ_DATEMODIF', NowH);
                  PutValue('GQ_UTILISATEUR', V_PGI.User);
                end;

                repeat
                  if (NewDimChamp[iChamp] <> '') and (copy(NewDimChamp[iChamp], 1, 3) = 'GQ_') and
                    (not Dim.ReadOnly[iChamp]) then
                  begin // Il existe au moins 1 champ de la table DISPO modifiable
                    case Dim.TypeDonnee[iChamp] of
                      dotDate: TobDispo.PutValue(NewDimChamp[iChamp], StrToDateTime(ItemDim.Valeur[iChamp]));
                      dotReel: TobDispo.PutValue(NewDimChamp[iChamp], StrToFloat(ItemDim.Valeur[iChamp]));
                      dotString: TobDispo.PutValue(NewDimChamp[iChamp], ItemDim.Valeur[iChamp]);
                    else TobDispo.PutValue(NewDimChamp[iChamp], ItemDim.Valeur[iChamp]);
                    end;

                  end;
                  inc(iChamp);
                until iChamp > MaxDimChamp;
              end
              else inc(iChamp);
            end; // while iChamp <= MaxDimChamp do

            if TobDispo <> nil then
            begin
              // L'enreg DISPO peut exister : avec GQ_PHYSIQUE à 0 !!!
              TobDispo.InsertOrUpdateDB(False);
              TobDispo.Free;
            end;

            if not b_result then Result := 15;
          end else
          begin
            ItemDim.ToCreate := True;
            Result := CABok;
            if ArtError = '' then ArtError := TobDim.GetValue('GA_ARTICLE');
          end;
          if TobDim <> nil then TobDim.Free;
          MoveCur(False);
        end else
          if ItemDim.ToDelete then
        begin
          ItemDim.ToDelete := False;
          Article := CodeArticleUnique(CodeArticle, AffecteDim(QMasque, 1, ItemDim), AffecteDim(QMasque, 2, ItemDim), AffecteDim(QMasque, 3, ItemDim),
            AffecteDim(QMasque, 4, ItemDim), AffecteDim(QMasque, 5, ItemDim));
          if SupItemDim then
          begin
            ExecuteSQL('delete from ARTICLE where GA_ARTICLE="' + Article + '"');
{$IFDEF BTP}
            // ajout BTP pour gestion des articles en prix posés
            if TOB(ItemDim.data).getValue('GA_TYPEARTICLE') = 'ARP' then
            begin
              ExecuteSQL('DELETE FROM NOMENENT WHERE GNE_NOMENCLATURE="' + Article + '"');
              ExecuteSQL('DELETE FROM NOMENLIG WHERE GNL_NOMENCLATURE="' + Article + '"');
            end;
{$ENDIF}
            if GetPresentation = ART_ORLI
              then ExecuteSQL('delete from ARTICLECOMPL where GA2_ARTICLE="' + Article + '"');
          end;
          dec(i);
          MoveCur(False);
        end else
        begin
          TobDim := TOB(ItemDim.data);
          if TobDim.modifie then
          begin
            CABok := ControleTobCAB(TobDim);
            if CABok = 0 then UpdateArticleDim(TobDim)
            else
            begin
              Result := CABok;
              if ArtError = '' then ArtError := TobDim.GetValue('GA_ARTICLE');
            end;
            MoveCur(False);
          end;
        end;
      end;
      inc(i);
    end; // fin while
    //
    // Mémorisation du dernier chrono Code à barres affecté
    //
    //Géré dans AttribNewCodeCAB !! if (CodeCAB<>'') then SauvChronoCAB( FournPrinc ) ;
  end;
  FiniMove;
  ferme(Qmasque);
  if Result = 0 then
  begin
    LoadDimensions(False);
    RefreshGrid(HideUnUsed);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Didier Carret
Créé le ...... : 13/03/2003
Modifié le ... : 29/04/2003
Description .. : Mise à jour des articles dimensionnés.
Suite ........ : Mise à jour des champs modifiés dans l'objet dimension.
Suite ........ : Mise à jour éventuelle des champs de la table
Suite ........ : ARTICLECOMPL ; dans ce cas, si l'article n'a pas été
Suite ........ : modifié dans la table ARTICLE, la modification de
Suite ........ : GA_DATEMODIF est forçée ( TOX ).
Suite ........ : Mise à jour des champs modifiés de la table DISPO
Mots clefs ... : MODIF;DIM;ARTICLE;ARTICLEDIM;ARTICLECOMPL;DISPO
*****************************************************************}

procedure TODimArticle.UpdateArticleDim(ArtDim: Tob);
var i, ii, iGQ: integer;
  ArtModifie, ArtComplModifie, DispoModifie: boolean;
  SQL, SQLArtCompl, SQLDispo, SQLInsertArtComplChp, SQLInsertArtComplVal, stTmp: string;
  TobDispo: Tob;
  ChampGQ, ValeurGQ: array[1..MaxDimChamp] of string; // Utile pour TobDispo.InsertDB
begin
  ArtModifie := False;
  ArtComplModifie := False;
  DispoModifie := False;
  if (ArtDim.modifie) then
  begin
    iGQ := 0;
    for i := 0 to ChampsDim.count - 1 do
    begin //modif RAPIDE du 30/05/01 DC+JCF
      if (ArtDim.GetValue('_MAJARTICLE')) and
        (copy(ChampsDim[i], 1, 3) = 'GA_') and
        (copy(ChampsDim[i], 1, 7) <> 'GA_DATE') then
      begin
        if not ArtModifie then
        begin
          SQL := 'update ARTICLE set ';
          ArtModifie := True;
        end
        else SQL := SQL + ',';
        SQL := SQL + ChampsDim[i] + '=' + ChampTobToSQL(ChampsDim[i], ArtDim);
      end

        // Gestion de la table ARTICLECOMPL
      else
        if (ArtDim.GetValue('_MAJARTICLECOMPL')) and
        (copy(ChampsDim[i], 1, 4) = 'GA2_') then
      begin
        if not ArtComplModifie then
        begin
          SQLArtCompl := 'update ARTICLECOMPL set ';
          SQLInsertArtComplChp := 'insert into ARTICLECOMPL(GA2_ARTICLE,GA2_CODEARTICLE,';
          SQLInsertArtComplVal := ') values ("' + ArtDim.GetValue('GA_ARTICLE') + '","' + ArtDim.GetValue('GA_CODEARTICLE') + '",';
          ArtComplModifie := True;
        end
        else
        begin
          SQLArtCompl := SQLArtCompl + ',';
          SQLInsertArtComplChp := SQLInsertArtComplChp + ',';
          SQLInsertArtComplVal := SQLInsertArtComplVal + ',';
        end;
        stTmp := ChampTobToSQL(ChampsDim[i], ArtDim);
        SQLArtCompl := SQLArtCompl + ChampsDim[i] + '=' + stTmp;
        SQLInsertArtComplChp := SQLInsertArtComplChp + ChampsDim[i];
        SQLInsertArtComplVal := SQLInsertArtComplVal + stTmp;
      end

        // Gestion de la table DISPO
      else if (ArtDim.GetValue('_MAJDISPO')) and
        (copy(ChampsDim[i], 1, 3) = 'GQ_') and
        (copy(ChampsDim[i], 1, 4) <> 'GQ__') then
      begin
        ii := 1;
        while (NewDimChamp[ii] <> ChampsDim[i]) and (ii < 10) do inc(ii);
        if (NewDimChamp[ii] = ChampsDim[i]) and (not Dim.ReadOnly[ii]) then
        begin
          if not DispoModifie then
          begin
            SQLDispo := 'update DISPO set ';
            DispoModifie := True;
          end
          else SQLDispo := SQLDispo + ',';
          inc(iGQ);
          ChampGQ[iGQ] := ChampsDim[i];
          ValeurGQ[iGQ] := ChampTobToSQL(ChampsDim[i], ArtDim);
          SQLDispo := SQLDispo + ChampsDim[i] + '=' + ValeurGQ[iGQ];
        end;
      end;
    end;

    if ArtModifie then
    begin
      SQL := SQL + ',GA_DATEMODIF="' + USTime(NowH);
      SQL := SQL + '",GA_UTILISATEUR="' + V_PGI.User;
      SQL := SQL + '" where GA_ARTICLE="' + ArtDim.GetValue('GA_ARTICLE') + '"';
      ExecuteSQL(SQL);
    end;
    if ArtComplModifie then
    begin
      SQLArtCompl := SQLArtCompl + ' where GA2_ARTICLE="' + ArtDim.GetValue('GA_ARTICLE') + '"';
      // DCA - FQ MODE 10952 - Gestion du cas de la saisie directe dans l'objet dimension
      // de GA2_COLLECTIONBAS pour un GA2_ARTICLE inexistant (normalement, redescend d'orli)
      if ExecuteSQL(SQLArtCompl) = 0
        then ExecuteSQL(SQLInsertArtComplChp+SQLInsertArtComplVal+')');
      // Si l'article n'est pas modifié, on force la mise à jour de GA_DATEMODIF : Tox oblige !
      if not ArtModifie then
      begin
        SQL := 'update ARTICLE set GA_DATEMODIF="' + USTime(NowH);
        SQL := SQL + '",GA_UTILISATEUR="' + V_PGI.User;
        SQL := SQL + '" where GA_ARTICLE="' + ArtDim.GetValue('GA_ARTICLE') + '"';
        ExecuteSQL(SQL);
      end;
    end;
    if DispoModifie then
    begin
      SQLDispo := SQLDispo + ',GQ_DATEMODIF="' + USTime(NowH) +
        '",GQ_UTILISATEUR="' + V_PGI.User +
        '" where GQ_ARTICLE="' + ArtDim.GetValue('GA_ARTICLE') +
        '" and GQ_DEPOT = "' + ArtDim.GetValue('GDE_DEPOT') +
        '" and GQ_CLOTURE = "-"';
      if ExecuteSQL(SQLDispo) = 0 then
      begin // Insertion dans DISPO
        TobDispo := TOB.Create('DISPO', nil, -1);
        with TobDispo do
        begin
          PutValue('GQ_ARTICLE', ArtDim.GetValue('GA_ARTICLE'));
          PutValue('GQ_DEPOT', ArtDim.GetValue('GDE_DEPOT'));
          PutValue('GQ_CLOTURE', '-');
          PutValue('GQ_DATEMODIF', NowH);
          PutValue('GQ_UTILISATEUR', V_PGI.User);
          for ii := 1 to iGQ do PutValue(ChampGQ[ii], ValeurGQ[ii]);
          InsertDB(nil);
          Free;
        end;
      end;
    end;
  end;
end;

procedure TODimArticle.refresh(ItemDim: THDimensionItem);
begin
  Dim.InvalidateDim(ItemDim, true);
end;

procedure TODimArticle.OnVoir(Sender: TObject);
var QMasque: TQuery;
  ArticleDim: string;
  ItemDim: THDimensionItem;
begin
  ItemDim := THDimensionItem(Sender);
  if ItemDim = nil then
  begin
    {HShowMessage( '0;'+'Article'+';Rien à voir;E;O;O;O','','') ;  }
    exit;
  end;
  if GetParamSoc('SO_GCARTTYPEMASQUE')
    then QMasque := OpenSQL('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="' + MasqueDim + '" and GDM_TYPEMASQUE="' + TypeMasqueDim + '"', TRUE,-1,'',true)
  else QMasque := OpenSQL('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="' + MasqueDim + '"', TRUE,-1,'',true);
  if not QMasque.EOF then
  begin
    ArticleDim := CodeArticleUnique(CodeArticle, AffecteDim(QMasque, 1, ItemDim), AffecteDim(QMasque, 2, ItemDim), AffecteDim(QMasque, 3, ItemDim),
      AffecteDim(QMasque, 4, ItemDim), AffecteDim(QMasque, 5, ItemDim));
    {$IFDEF EAGLCLIENT}
    {AFAIREEAGL}
    {$ELSE}
    if TheMulq = nil then TheMulQ := TQuery.Create(Application); // pour passer la clause where à la fiche
    TheMulQ.SQL.Clear;
    TheMulQ.SQL.Add('SELECT * From Article where (GA_STATUTART="DIM") AND GA_CODEARTICLE="' + codeArticle + '"');
    ChangeSQL(TheMulQ);
    {$ENDIF}
    {$IFNDEF GPAO} { GPAO V1.0}
      {$IFDEF BTP}
      V_PGI.DispatchTT(7, taModif, ArticleDim, '', '');
      {$ELSE}
      AGLLanceFiche('GC', 'GCARTICLE', '', ArticleDim, '');
      {$ENDIF}
    {$ELSE}
    V_PGI.DispatchTT(7, taModif, ArticleDim, '', '');
    {$ENDIF}
  end;
  ferme(Qmasque);
end;

procedure TODimArticle.SetAction(const Value: TActionFiche);
begin
  FAction := value;
  Dim.Action := value;
  if FAction = taConsult then Dim.HideUnUsed := True;
end;

procedure TODimArticle.LibereDim;
var iItem: integer;
begin
  if TableDim <> nil then
  begin
    for iItem := 0 to TableDim.count - 1 do
      if TableDim.Items[iItem] <> nil then
      begin
        // Destruction ici car en cas de création de dimensions, les tobs créées ne sont pas rattachées à TobArticleDim !
        if (not ConserveTobArtDim) and (Tob(THDimensionItem(TableDim.Items[iItem]).data) <> nil)
          then Tob(THDimensionItem(TableDim.Items[iItem]).data).free;
        TableDim.Items[iItem].free;
      end;
    TableDim.free;
    TableDim := nil;
  end;
  if (not ConserveTobArtDim) and (TOBArticleDim <> nil) then
  begin
    TOBArticleDim.free;
    TOBArticleDim := nil
  end;
end;

destructor TODimArticle.Destroy;
begin
  LibereDim;
  LibChampsDim.free;
  LibChampsDim := nil;
  ChampsDim.free;
  ChampsDim := nil;
  TobPrefEtab.free;
  TobPrefEtab := nil;
  inherited Destroy;
end;

function FormateChamp(CodeArticle: string; PCB: Double): string;
var //QQ: TQuery ;
  Format: string;
  ParCombien: Double;
  NbPartieEnt, NbPartieDec: Integer;
begin
  //QQ:=OpenSQL('Select GA_PCB from Article where GA_ARTICLE="'+CodeArticleUnique2( codeArticle, '')+'"',True) ;
  //If Not QQ.EOF then
  //Begin
     //ParCombien:=QQ.FindField('GA_PCB').AsFloat ;
  ParCombien := PCB;
  Format := '#,#';
  NbPartieEnt := length(IntToStr(Trunc(ParCombien)));
  Format := Format + stringofchar('0', NbPartieEnt);
  if frac(ParCombien) > 0 then
  begin
    Format := Format + '.';
    // +1 pour la virgule
    NbPartieDec := length(FloatToStr(ParCombien)) - (NbPartieEnt + 1);
    Format := Format + stringofchar('0', NbPartieDec);
  end;
  //end ;
  Result := Format + ';';
  //Ferme(QQ) ;
end;

procedure TODimArticle.CreateItemDim(ItemDim: THDimensionItem; TobArtGen: TOB);
var iChamp: integer;
  QMasque: TQuery;
  TobDim: TOB;
  CodeArtDim: string;
begin
  if GetParamSoc('SO_GCARTTYPEMASQUE')
    then QMasque := OpenSQL('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="' + MasqueDim + '" and GDM_TYPEMASQUE="' + TypeMasqueDim + '"', TRUE,-1,'',true)
  else QMasque := OpenSQL('SELECT * FROM DIMMASQUE WHERE GDM_MASQUE="' + MasqueDim + '"', TRUE,-1,'',true);
  if QMasque.EOF then exit;
  CodeArtDim := CodeArticleUnique(CodeArticle, AffecteDim(QMasque, 1, ItemDim), AffecteDim(QMasque, 2, ItemDim), AffecteDim(QMasque, 3, ItemDim),
    AffecteDim(QMasque, 4, ItemDim), AffecteDim(QMasque, 5, ItemDim));
  Ferme(QMasque);

  TobDim := TOB.Create('ARTICLE', nil, -1);
  TobDim.dupliquer(TobArtGen, True, True, False);
  TobDim.PutValue('GA_ARTICLE', CodeArtDim);
  TobDim.PutValue('GA_STATUTART', 'DIM');
  TobDim.AddChampSup('ZGA_CODEBARRE', False);
  TobDim.PutValue('ZGA_CODEBARRE', TobDim.GetValue('GA_CODEBARRE'));
  TOB(itemDim.data) := TobDim;

  for iChamp := 1 to MaxDimChamp do if (NewDimChamp[iChamp] <> '') then
    begin
      if (copy(NewDimChamp[iChamp], 1, 3) <> 'GA_') then
      begin
        TOB(itemDim.data).AddChampSup(NewDimChamp[iChamp], False);
        TOB(itemDim.data).PutValue(NewDimChamp[iChamp], SetInitTypeDonnee(NewDimChamp[iChamp]));
      end;
      ItemDim.valeur[iChamp] := TOB(itemDim.data).GetValue(NewDimChamp[iChamp]);
    end;

end;

function TODimArticle.DeleteItemDim(ItemDim: THDimensionItem; Article: string) : boolean;
var TobDim: TOB;
  Erreur: string;
  TypeArticle : string;
begin
  Result := True ;
  SupItemDim := False;
  TypeArticle := TOB(ItemDim.data).GetValue('GA_TYPEARTICLE');

  Erreur := ControlArticle(Article, CodeArticle, True,TypeArticle);
  if Erreur <> '' then
    begin
    AGLLanceFiche('GC', 'GCUTILARTDIM', '', '', 'LISTE=' + Erreur + ',ARTICLE=' + Article) ;
    Result := False  // DCA - FQ MODE 10682
    end
  else SupItemDim := True;

  if Result and (ItemDim.ToCreate) then
  begin
    TobDim := Tob(ItemDim.data);
    if (TobDim <> nil) then TobDim.free;
    TableDim.Remove(ItemDim);
  end;
end;

// Report des prix de TobArt sur les Tob des THDimensionItems.

procedure TODimArticle.MajPrixUniques(TobArt: TOB);
{var nbTob : integer;}
var i: integer;
  ArtDimModifie: boolean;
  ArtDim: Tob;
  ItemDim: THDimensionItem;
begin
  if tableDim = nil then exit;
  for i := 0 to TableDim.count - 1 do
  begin
    if TableDim.Items[i] <> nil then
    begin
      ItemDim := THDimensionItem(TableDim.Items[i]);
      ArtDim := TOB(ItemDim.data);
      if ArtDim <> nil then
      begin
        ArtDimModifie := ArtDim.Modifie;
        ArtDim.PutValue('GA_PAHT', TobArt.GetValue('GA_PAHT'));
        ArtDim.PutValue('GA_PRHT', TobArt.GetValue('GA_PRHT'));
        ArtDim.PutValue('GA_PVHT', TobArt.GetValue('GA_PVHT'));
        ArtDim.PutValue('GA_PVTTC', TobArt.GetValue('GA_PVTTC'));
        ArtDim.Modifie := ArtDimModifie; // inchangé : prix uniques maj par 1 requete suivant TobArt
      end;
    end;
  end;
end;

end.
