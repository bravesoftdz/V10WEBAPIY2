{***********UNITE*************************************************
Auteur  ...... : AC
Créé le ...... : 16/12/2002
Modifié le ... : 31/01/2003
Description .. : Source TOF de la FICHE : AFFCDEMODIF ()
Mots clefs ... : TOF;AFFCDEMODIF
*****************************************************************}
unit AFFCDEMODIF_TOF;

interface

uses StdCtrls, Controls, Classes, stat, Windows,
  {$IFDEF EAGLCLIENT}
  eFiche, Maineagl,
  {$ELSE}
  db, dbtables, DBGrids, Fe_Main,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, UTobView, HQry, UTOB, menus,
  UtilArticle, UtilGC, EntGC, AGLInit, UtilDimArticle, UtilAffCde, M3FP, HStatus, ed_tools;

type
  TOF_AFFCDEMODIF = class(TOF)
  private
    TV1: TTobViewer;
    FTV: TFStat;
    TOBDim: TOB; // Articles dimensionnées
    TOBAffect: TOB; // Ligne de commande à traiter
    TOBDispo: TOB; // Stock des articles à traiter
    TOBAffPiece: TOB; // Commande à traiter ;
    Etape: TTypeEtapeAffCde;
    QTobViewer: TQuery;
    StWhere : String ;// Clause where du TobViever
    nbTransacErr,          // Erreur à la maj d'une pièce dans la base par TOBxxx.UpdateDB
    nbTransacTot : integer; // Nb de transactions de maj de la base

    procedure TVOnDblClick(Sender: TObject);
    procedure TraiterArticleDimensionne;
    procedure TraiterArticleUnique;
    procedure RemplirTOBDim(CodeArticle: string; Numero, Indice, NumLigneGen: integer);
    procedure TOBDimToTV1(NaturePiece, souche: string; Numero, Indice, NumLigneGen: Integer);
    procedure ChargeLesTOB(NaturePiece, souche: string; Numero, Indice: Integer; Selection: boolean = False);
    procedure ChargeTobAffect(NaturePiece, souche: string; Numero, Indice: Integer);
    procedure ChargeTobDispo(NaturePiece, souche: string; Numero, Indice: Integer; CodeArticle: string);
    procedure ChargeLaTob(NaturePiece, souche: string; Numero, Indice: Integer);
    procedure ChargeTobAffPiece(NaturePiece, souche: string; Numero, Indice: Integer);
    procedure AlouerTob;
    procedure UpdateTableAffect;
    procedure UpdateTable;
    procedure OnChangeCodeAff;
    procedure OnAffecterSelection;
    procedure OnAffecterCommande;
    procedure OnAffecterLigne;
    procedure AffecterCommande(NaturePiece, Souche: string; Numero, Indice: Integer);
    procedure OnDesAffecterSelection;
    procedure OnDesAffecterCommande;
    procedure OnDesAffecterLigne;
    procedure DesAffecterCommande(NaturePiece, Souche: string; Numero, Indice: Integer);
    procedure MajLigneGen(NaturePiece, Souche: string; Numero, Indice, NumLigneGen: Integer; Qte, PxUnit: DOuble);
    procedure MajPiece(TOBGen: TOB; OldQte, OldMontant: Double);
    procedure MajLaTOB(TOBP: TOB);

  public
    procedure OnArgument(S: string); override;
    procedure OnLoad; override;
    procedure OnUpdate; override;
    procedure OnClose; override;
  end;

function ConstruireInArticle(Const ListeArticles : TStringList; Var Indice : Integer) : string;
procedure AjouteEvent(TypeEvent, LibEvent, MessEvent, EtatEvent: string);

const
  // libellés des messages
  TexteMessage: array[1..2] of string = (
    {1}'Le code de l''affectation est obligatoire'
    {2}, ''
    );

implementation

type
  TTypeJointure = (joinENTETE, joinARTICLE, joinTIERS, joinTIERSCOMPL);
  TTypeJointures = set of TTypeJointure;

const
  NbArtParRequete : integer = 1000;

procedure AjouteEvent(TypeEvent, LibEvent, MessEvent, EtatEvent: string);
var QQ: TQuery;
  MotifPiece: TStrings;
  NumEvent: integer;
begin
  MotifPiece := TStringList.Create;
  MotifPiece.Add(MessEvent);
  NumEvent := 0;
  QQ := OpenSQL('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True);
  if not QQ.EOF then NumEvent := QQ.Fields[0].AsInteger;
  Ferme(QQ);
  Inc(NumEvent);
  QQ := OpenSQL('SELECT * FROM JNALEVENT WHERE GEV_TYPEEVENT="' + TypeEvent + '" AND GEV_NUMEVENT=-1', False);
  QQ.Insert;
  InitNew(QQ);
  QQ.FindField('GEV_NUMEVENT').AsInteger := NumEvent;
  QQ.FindField('GEV_TYPEEVENT').AsString := TypeEvent;
  QQ.FindField('GEV_LIBELLE').AsString := LibEvent;
  QQ.FindField('GEV_DATEEVENT').AsDateTime := Date;
  QQ.FindField('GEV_UTILISATEUR').AsString := V_PGI.User;
  QQ.FindField('GEV_ETATEVENT').AsString := EtatEvent;
  TMemoField(QQ.FindField('GEV_BLOCNOTE')).Assign(MotifPiece);
  QQ.Post;
  Ferme(QQ);
  MotifPiece.Free;
end;

procedure TOF_AFFCDEMODIF.OnLoad;
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
  // Gestion des checkBox : booléens libres  tiers
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'BOOL', 'YTC_BOOLLIBRE', 3, '');
  // Gestion des dates libres tiers
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'DATE', 'YTC_DATELIBRE', 3, '_');
  // Gestion des montants libres tiers
  xx_where := GCXXWhereChampLibre(TForm(Ecran), xx_where, 'EDIT', 'YTC_VALLIBRE', 3, '_');
  SetControlText('XX_WHERE', xx_where);
end;

procedure TOF_AFFCDEMODIF.OnArgument(S: string);
var iCol, Nbr: integer;
  stIndice, stChampsCompl: string;
begin
  inherited;
  TV1 := TTobViewer(getcontrol('TV'));
  TV1.OnDblClick := TVOnDblClick;
  TV1.PopupData := TPopupMenu(getcontrol('POPAFFECT'));
  FTV := TFStat(Ecran);
  FTV.Caption := 'Modification des affectations';
  UpdateCaption(FTV);

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
      if ChangeLibre2('TGA2_FAMILLENIV' + stIndice, Ecran) then stChampsCompl := stChampsCompl + ',GA2_FAMILLENIV' + stIndice;
    end;
    for iCol := 1 to 2 do
    begin
      stIndice := IntToStr(iCol);
      if ChangeLibre2('TGA2_STATART' + stIndice, Ecran) then stChampsCompl := stChampsCompl + ',GA2_STATART' + stIndice;
    end;
  end;

  // Paramétrage des libellés des tables libres article et dépôt
  if (GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'GA_LIBREART', 10, '') = 0) then SetControlVisible('PTABLESARTICLE', False);
  if (GCMAJChampLibre(TForm(Ecran), False, 'COMBO', 'YTC_TABLELIBRETIERS', 10, '') = 0) then SetControlVisible('PTABLESTIERS', False);
  // Mise en forme des libellés des dates, booléans libres et montants libres article et tiers
  Nbr := 0;
  if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GA_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL', False) else inc(Nbr);
  if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'GA_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE', False) else inc(Nbr);
  if (GCMAJChampLibre(TForm(Ecran), False, 'BOOL', 'GA_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL', False) else inc(Nbr);
  {$IFNDEF CCS3}
  if (Abs(Nbr) = 0) then
    {$ENDIF}
    SetControlVisible('PZONESARTICLE', False);

  Nbr := 0;
  if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'YTC_VALLIBRE', 3, '_') = 0) then SetControlVisible('GB_VAL1', False) else inc(Nbr);
  if (GCMAJChampLibre(TForm(Ecran), False, 'EDIT', 'YTC_DATELIBRE', 3, '_') = 0) then SetControlVisible('GB_DATE1', False) else inc(Nbr);
  if (GCMAJChampLibre(TForm(Ecran), False, 'BOOL', 'YTC_BOOLLIBRE', 3, '') = 0) then SetControlVisible('GB_BOOL1', False) else inc(Nbr);
  {$IFNDEF CCS3}
  if (Abs(Nbr) = 0) then
    {$ENDIF}
    SetControlVisible('PZONESTIERS', False);
end;

procedure TOF_AFFCDEMODIF.OnUpdate;
var CodeAff, StSQL: string;
begin
  inherited;
  CodeAff := GetControlText('GEL_CODEAFF');
  if CodeAff = '' then
  begin
    PGIBox(TraduireMemoire(TexteMessage[1]), FTV.Caption);
    SetFocusControl('GEL_CODEAFF');
    FTV.stSQL := '';
    Exit;
  end;
  if not VerifStatutAFFCDEENTETE(CodeAff, afcAffectation) then
  begin
    FTV.stSQL := '';
    Exit;
  end;
  if QTobViewer <> nil then Ferme(QTobViewer);
  // Ajout du complément de la clause WHERE, pour prendre en compte les critères
  // de sélection renseignés par l'utilisateur (collection, fournisseur, ...)
  stWhere := RecupWhereCritere(TPageControl(GetControl('Pages')));
  if Uppercase(copy(stWhere, 1, 6)) = 'WHERE ' then
  begin
    StWhere := Copy(StWhere, 7, length(StWhere) - 6);
    StSQL := InsertSQLWhere(FTV.stSQL, StWhere);
  end;
  FTV.StSQL := StSQL;
  QTobViewer := OpenSQL(FTV.stSQL, true);
  AlouerTOB;
end;

procedure TOF_AFFCDEMODIF.AlouerTOB;
begin
  if TOBDim = nil then TOBDim := TOB.Create('', nil, -1);
  if TOBAffect = nil then TOBAffect := TOB.Create('_Les affectations', nil, -1);
  if TOBAffPiece = nil then TOBAffPiece := TOB.Create('_Les pieces', nil, -1);
  if TOBDispo = nil then TOBDispo := TOB.Create('_Les stocks', nil, -1);
  // TOB du TobViewer
  if FTV.LaTob = nil then FTV.LaTob := TOB.create('_AffectationCDE', nil, -1);
end;

function ConstruireInArticle(Const ListeArticles : TStringList; Var Indice : Integer) : string;
var StIn : string;
    NbArt, i : integer;
begin
  StIn := '';
  NbArt := 0;
  For i := Indice to ListeArticles.Count - 1 do
  begin
    if StIn <> '' then StIn := StIn + ',';
    StIn := StIn + '"' + ListeArticles.Strings[i] + '"';
    Inc(NbArt);
    if NbArt >= NbArtParRequete then Break;
  end;
  if (i = ListeArticles.Count) then Indice := i else Indice := i + 1;
  Result := StIn;
end;

procedure TOF_AFFCDEMODIF.ChargeLesTob(NaturePiece, souche: string; Numero, Indice: Integer; Selection: boolean = False);
var CodeArticle, StInArticle : string;
    ListeArticles : TStringList;
    i, iTob, iArticle : integer;
begin
  ChargeTobAffect(NaturePiece, souche, Numero, Indice);
  ChargeTobAffPiece(NaturePiece, souche, Numero, Indice);
  // Pour pouvoir charger seulement le dispo des articles utilisés
  { DCA - Prise en compte limitation du nb d'elements dans une liste (ORACLE)
    Une liste de valeur ne peut contenir plus de 1000 occurences.

    Chargement des articles en TStringList en éliminant les doublons. }
  ListeArticles := TStringList.Create;
  ListeArticles.Sorted := True;
//  While y < ListeArticles.Count do
  for iTob := 0 to TOBAffect.Detail.Count - 1 do
  begin
    CodeArticle := TOBAffect.Detail[iTob].GetValue('GEL_CODEARTICLE');
    if Not ListeArticles.Find(CodeArticle, i) then ListeArticles.Add(CodeArticle);
    { Ancien code
    if CodeArticle = TOBAffect.Detail[i].GetValue('GEL_CODEARTICLE') then continue
    else CodeArticle := TOBAffect.Detail[i].GetValue('GEL_CODEARTICLE');
    if LesCodeArticles = '' then LesCodeArticles := '"' + TOBAffect.Detail[i].GetValue('GEL_CODEARTICLE') + '" '
    else LesCodeArticles := LesCodeArticles + ', "' + TOBAffect.Detail[i].GetValue('GEL_CODEARTICLE') + '" '
    }
  end;

  { Chargement en tob en limitant le nombre d'article dans la requête }
  TOBDispo.ClearDetail;
  if ListeArticles.Count > 0 then
  begin
    iArticle := 0;
    while iArticle < ListeArticles.Count do
    begin
      StInArticle := ConstruireInArticle (ListeArticles, iArticle);
      if StInArticle <> '' then ChargeTobDispo (NaturePiece, souche, Numero, Indice, StInArticle);
    end;
  end;
  ListeArticles.Free;
  { Ancien code
  ChargeTobDispo(LesCodeArticles);
  }

  if not selection then ChargeLaTob(NaturePiece, souche, Numero, Indice);
end;

procedure TOF_AFFCDEMODIF.ChargeTobAffect(NaturePiece, souche: string; Numero, Indice: Integer);
var SQL, StWhereCde, StWhereForJointure,Groupe,Prfx: string;
  Q: TQuery;
  TypeJoin: TTypeJointures ;
begin
  TobAffect.ClearDetail;
  // Make Jointure
  StWhereForJointure := StWhere ;
  StWhereForJointure := StringReplace (StWhereForJointure,'(','',[rfReplaceAll]) ;
  StWhereForJointure := StringReplace (StWhereForJointure,')','',[rfReplaceAll]) ;
  While StWhereForJointure<>'' do
  begin
    Groupe := ReadTokenPipe (StWhereForJointure,'AND') ;
    Prfx := ExtractPrefixe(Groupe);
    if Prfx = 'YTC' then TypeJoin := TypeJoin + [joinTIERSCOMPL] else
    if Prfx = 'GEA' then TypeJoin := TypeJoin + [joinENTETE] else
    if Prfx = 'GA' then TypeJoin := TypeJoin + [joinARTICLE] else
    if Prfx = 'T' then TypeJoin := TypeJoin + [joinTIERS];
  end ;
  // Make Sql
  SQL := 'SELECT * FROM AFFCDELIGNE' ;
  // Jointures
  if joinTIERS in TypeJoin then
    SQL := SQL + ' LEFT JOIN TIERS ON GEL_TIERS=T_TIERS';
  if joinARTICLE in TypeJoin then
    SQL := SQL + ' LEFT JOIN ARTICLE ON GEL_ARTICLE=GA_ARTICLE';
  if joinTIERSCOMPL in TypeJoin then
    SQL := SQL + ' LEFT JOIN TIERSCOMPL ON GEL_TIERS=YTC_TIERS';
  if joinENTETE in TypeJoin then
    SQL := SQL + ' LEFT JOIN AFFCDEENTETE ON GEL_CODEAFF=GEA_CODEAFF' ;
  // Make clause Where
  StWhereCde := 'AND GEL_NATUREPIECEG="' + NaturePiece + '" AND GEL_SOUCHE="' + souche + '" AND GEL_NUMERO=' + IntToStr(Numero) + ' AND GEL_INDICEG=' +
    intToStr(Indice) + ' ';
  StWhereCde := StWhere + StWhereCde ;
  // SQL terminé
  // Modif DCA - Remplacement stWhere par stWhereCde pour ne charger que la commande concernée
  //SQL := InsertSQLWhere(SQL, StWhere);
  SQL := InsertSQLWhere(SQL, StWhereCde);
  Q := OpenSQL(SQL, True);
  TOBAffect.LoadDetailDB('AFFCDELIGNE', '', '', Q, False);
  Ferme(Q);
end;

procedure TOF_AFFCDEMODIF.ChargeTobAffPiece(NaturePiece, souche: string; Numero, Indice: Integer);
var SQL, StWhere: string;
  Q: TQuery;
begin
  TOBAffPiece.ClearDetail;
  StWhere := 'AND GEP_NATUREPIECEG="' + NaturePiece + '" AND GEP_SOUCHE="' + souche + '" AND GEP_NUMERO=' + IntToStr(Numero) + ' AND GEP_INDICEG=' +
    IntToStr(Indice) + ' ';
  SQL := 'SELECT * FROM AFFCDEPIECE WHERE GEP_CODEAFF="' + GetControlText('GEL_CODEAFF') + '" ' + StWhere;
  Q := OpenSQL(SQL, True);
  TOBAffPiece.LoadDetailDB('AFFCDEPIECE', '', '', Q, False);
  Ferme(Q);
end;

procedure TOF_AFFCDEMODIF.ChargeTobDispo(NaturePiece, souche: string; Numero, Indice: Integer; CodeArticle: string);
var SQL, stWhereCde : string;
  Q: TQuery;
begin
  // Modif DCA - Chargement en plusieurs requêtes si nécessaire
  //TOBDispo.ClearDetail;

  // Modif DCA - Ajout condition sur la pièce traitée (pour utilisation index)
  StWhereCde := 'GEL_CODEAFF="' + GetControlText('GEL_CODEAFF') + '" AND GEL_NATUREPIECEG="' + NaturePiece + '" AND GEL_SOUCHE="' + souche
    + '" AND GEL_NUMERO=' + IntToStr(Numero) + ' AND GEL_INDICEG=' + IntToStr(Indice);

  SQL := 'SELECT * FROM AFFCDEDISPO WHERE GED_CODEAFF="' + GetControlText('GEL_CODEAFF') +
    '" AND GED_ARTICLE in (SELECT GEL_ARTICLE FROM AFFCDELIGNE WHERE ' + stWhereCde + ' AND GEL_CODEARTICLE in (' + CodeArticle + '))';
  Q := OpenSQL(SQL, True);
  // Modif DCA - Chargement en plusieurs requêtes si nécessaire
  TOBDispo.LoadDetailDB('AFFCDEDISPO', '', '', Q, True);
  Ferme(Q);
end;

procedure TOF_AFFCDEMODIF.ChargeLaTob(NaturePiece, souche: string; Numero, Indice: Integer);
var stSQL, stWhere: string;
  Q: TQuery;
begin
  FTV.LaTob.ClearDetail;
  stSQL := FTV.stSQL;
  stWhere := 'AND GEL_NATUREPIECEG="' + NaturePiece + '" AND GEL_SOUCHE="' + souche + '" AND GEL_NUMERO=' + IntToStr(Numero) + ' AND GEL_INDICEG=' +
    IntToStr(Indice) + ' ';
  StSQL := InsertSQLWhere(StSQL, StWhere);
  Q := OpenSQL(stSQL, true);
  FTV.LaTob.LoadDetailDB('_Affect', '', '' Q, False);
  Ferme(Q);
end;

procedure TOF_AFFCDEMODIF.OnClose;
begin
  inherited;
  TOBAffect.Free;
  TOBDispo.Free;
  TOBAffPiece.Free;
  TOBDim.Free;
  FTV.LaTob.Free;
  Ferme(QTobViewer);
end;

procedure TOF_AFFCDEMODIF.TVOnDblClick(Sender: TObject);
{$IFDEF VERSION_AGL2}
var ARow : integer;
{$ENDIF}
begin
{$IFDEF VERSION_AGL2}
  ARow := TV1.GetRowSelect;
{$ENDIF}
  nbTransacErr := 0 ; // Initialisation
  nbTransacTot := 0 ; // Initialisation
  if TV1.AsString[TV1.ColIndex('GEL_TYPEDIM'), TV1.CurrentRow] = 'GEN' then TraiterArticleDimensionne
  else TraiterArticleUnique;
{$IFDEF VERSION_AGL2}
  // DC/EP - Repositionnement sur la ligne
  TV1.SetRowSelect(ARow);
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : A CATHELINEAU
Créé le ...... : 18/12/2002
Modifié le ... :   /  /
Description .. : Affiche l'objet dimensionné
Mots clefs ... :
*****************************************************************}

procedure TOF_AFFCDEMODIF.TraiterArticleDimensionne;
var ValAction, NatureDoc, NaturePiece, Souche: string;
  Numero, Indice, NumLigneGen: Integer;
begin
  if TV1.CurrentCol = TV1.ColIndex('GEL_QTERESERVEE') then
  begin
    Etape := afcReservation;
    NatureDoc := 'ACR';
  end
  else
    if TV1.CurrentCol = TV1.ColIndex('GEL_QTEAFFECTEE') then
  begin
    Etape := afcAffectation;
    NatureDoc := 'ACF';
  end else exit;
  ValAction := 'SAISIE';
  NaturePiece := TV1.AsString[TV1.ColIndex('GEL_NATUREPIECEG'), TV1.CurrentRow];
  Souche := TV1.AsString[TV1.ColIndex('GEL_SOUCHE'), TV1.CurrentRow];
  Numero := TV1.AsInteger[TV1.ColIndex('GEL_NUMERO'), TV1.CurrentRow];
  Indice := TV1.AsInteger[TV1.ColIndex('GEL_INDICEG'), TV1.CurrentRow];
  NumLigneGen := TV1.AsInteger[TV1.ColIndex('GEL_NUMLIGNE'), TV1.CurrentRow];
  // Chargement des Tob affect , piece et dispo
  ChargeLesTob(NaturePiece, Souche, Numero, Indice);
  RemplirTOBDim(TV1.AsString[TV1.ColIndex('GEL_CODEARTICLE'), TV1.CurrentRow], Numero, Indice, NumLigneGen);
  TheTOB := TOBDim;
  AglLanceFiche('MBO', 'AFFCDESAISIEDIM', '', '', 'GA_CODEARTICLE=' + TV1.AsString[TV1.ColIndex('GEL_CODEARTICLE'), TV1.CurrentRow] + ';ACTION=' + ValAction +
    ';CHAMP= ;TYPEPARAM=' + NatureDoc + ';DEPOT=' + TV1.AsString[TV1.ColIndex('GEA_DEPOT'), TV1.CurrentRow] + '');
  if (TheTOB <> nil) and (TheTOB.Detail.Count > 0) then
  begin
  TOBDim := TheTOB ;
  TOBDimToTV1(NaturePiece, Souche, Numero, Indice, NumLigneGen);
  end
  else TOBDim.ClearDetail;
  TheTOB := nil; 
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : A CATHELINEAU
Créé le ...... : 18/12/2002
Modifié le ... :   /  /
Description .. : Chargement de TOBDim utilisée pour l'objet dimensionné
Mots clefs ... :
*****************************************************************}

procedure TOF_AFFCDEMODIF.RemplirTOBDim(CodeArticle: string; Numero, Indice, NumLigneGen: Integer);
var TOBD, TOBA, TOBDis: TOB;
  Q: TQuery;
begin
  TOBDim.ClearDetail;
  Q := OpenSQL('Select GA_Article from Article where GA_CodeArticle="' + CodeArticle + '" AND GA_STATUTART="DIM" order by GA_ARTICLE', True);
  while not Q.EOF do
  begin
    TOBA := TOBAffect.findfirst(['GEL_ARTICLE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNEGEN'], [Q.FindField('GA_ARTICLE').AsString, Numero, Indice,
      NumLigneGen], False);
    if TOBA <> nil then
    begin
      TOBD := TOB.Create('', TOBDim, -1);
      TOBD.AddChampSup('GA_ARTICLE', False);
      TOBD.PutValue('GA_ARTICLE', Q.FindField('GA_ARTICLE').AsString);
      TOBD.AddChampSup('GEL_QTERESERVEE', False);
      TOBD.PutValue('GEL_QTERESERVEE', TOBA.GetValue('GEL_QTERESERVEE'));
      TOBD.AddChampSup('GEL_MODIFIABLE', False);
      TOBD.PutValue('GEL_MODIFIABLE', TOBA.GetValue('GEL_MODIFIABLE'));
      TOBD.AddChampSup('GEL_QTEALIVRER', False);
      TOBD.PutValue('GEL_QTEALIVRER', TOBA.GetValue('GEL_QTEALIVRER'));
      if Etape = afcaffectation then
      begin
        TOBD.AddChampSup('GEL_QTEAFFECTEE', False);
        TOBD.PutValue('GEL_QTEAFFECTEE', TOBA.GetValue('GEL_QTEAFFECTEE'));
      end;
      // Traitement du dispo
      TOBDis := TOBDispo.findfirst(['GED_ARTICLE'], [Q.FindField('GA_ARTICLE').AsString], False);
      TOBD.AddChampSup('GED_QTEPREVI', False);
      if TOBDis <> nil then TOBD.PutValue('GED_QTEPREVI', TOBDis.GetValue('GED_QTEPREVI')) else TOBD.PutValue('GED_QTEPREVI', 0);
      if Etape = afcaffectation then
      begin
        TOBD.AddChampSup('GED_DISPOAFFECTEE', False);
        if TOBDis <> nil then TOBD.PutValue('GED_DISPOAFFECTEE', TOBDis.GetValue('GED_DISPOAFFECTEE')) else TOBD.PutValue('GED_DISPOAFFECTEE', 0);
      end else
      begin
        TOBD.AddChampSup('GED_DISPORESERVEE', False);
        if TOBDis <> nil then TOBD.PutValue('GED_DISPORESERVEE', TOBDis.GetValue('GED_DISPORESERVEE')) else TOBD.PutValue('GED_DISPORESERVEE', 0);
      end;
    end;
    Q.Next;
  end;
  ferme(Q);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : A CATHELINEAU
Créé le ...... : 18/12/2002
Modifié le ... :   /  /
Description .. : Mise à jour des valeurs modifées dans le TobViewer et
Suite ........ : dans la base
Mots clefs ... :
*****************************************************************}

procedure TOF_AFFCDEMODIF.TOBDimToTV1(NaturePiece, Souche: string; Numero, Indice, NumLigneGen: Integer);
var i: integer;
  TOBA, TOBDis: TOB;
  TotQteReservee, TotQteAffectee, AncienneQte, QteLivrer,
    PxUnit, Qte, Montant: Double;
  CodeArticle: string;
begin
  TotQteReservee := 0;
  TotQteAffectee := 0;
  Qte := 0;
  PxUnit := 0;
  CodeArticle := '';
  for i := 0 to TOBDim.Detail.count - 1 do
  begin
    TOBA := TobAffect.FindFirst(['GEL_ARTICLE', 'GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDIGEG', 'GEL_NUMLIGNEGEN'],
      [TOBDim.Detail[i].GetValue('GA_ARTICLE'), NaturePiece, Souche, Numero, Indice, NumLigneGen], False);
    if (TOBA <> nil) and (TOBA.GetValue('GEL_QTEALIVRER') <> 0) then
    begin
      if TobDim.Detail[i].GetValue('MODIFIE') then TOBA.AddChampSup('MODIFIER', False);
      TOBA.PutValue('GEL_MODIFIABLE', TOBDim.Detail[i].GetValue('GEL_MODIFIABLE'));
      if Etape = afcReservation then
      begin
        PxUnit := Arrondi(TOBA.GetValue('GEL_MONTANTHT') / TOBA.GetValue('GEL_QTEALIVRER'), 6);
        AncienneQte := TOBA.GetValue('GEL_QTERESERVEE');
        TOBA.PutValue('GEL_QTERESERVEE', TOBDim.Detail[i].GetValue('GEL_QTERESERVEE'));
        Montant := Arrondi(TOBDim.Detail[i].GetValue('GEL_QTERESERVEE') * PxUnit, 6);
        TOBA.PutValue('GEL_MTHTRESERVEE', Montant);
        TotQteReservee := TotQteReservee + TOBDim.Detail[i].GetValue('GEL_QTERESERVEE');
      end else
      begin
        PxUnit := Arrondi(TOBA.GetValue('GEL_MONTANTHT') / TOBA.GetValue('GEL_QTEALIVRER'), 6);
        AncienneQte := TOBA.GetValue('GEL_QTEAFFECTEE');
        TOBA.PutValue('GEL_QTEAFFECTEE', TOBDim.Detail[i].GetValue('GEL_QTEAFFECTEE'));
        Montant := Arrondi(TOBDim.Detail[i].GetValue('GEL_QTEAFFECTEE') * PxUnit, 6);
        TOBA.PutValue('GEL_MTHTAFFECTEE', Montant);
        TotQteAffectee := TotQteAffectee + TOBDim.Detail[i].GetValue('GEL_QTEAFFECTEE');
      end;
      CodeArticle := Trim(copy(TOBDim.Detail[i].GetValue('GA_ARTICLE'), 1, 18));
      // Mise à jour du dispo
      TOBDis := TobDispo.FindFirst(['GED_ARTICLE'], [TOBDim.Detail[i].GetValue('GA_ARTICLE')], False);
      if TOBDis <> nil then
      begin
        if Etape = afcReservation then QteLivrer := TOBDim.Detail[i].GetValue('GEL_QTERESERVEE')
        else QteLivrer := TOBDim.Detail[i].GetValue('GEL_QTEAFFECTEE');
        MajStockAFFCDE(TOBDis, QteLivrer, Etape, AncienneQte, True);
      end;
    end;
    // Mise à jour ligne générique
    if Etape = afcReservation then Qte := TotQteReservee
    else Qte := TotQteAffectee;
    // MajLigneGen(NaturePiece, Souche, Numero, Indice, NumLigneGen, Qte, PxUnit);
    // Modif tables
    //UpdateTable;
  end;
  MajLigneGen(NaturePiece, Souche, Numero, Indice, NumLigneGen, Qte, PxUnit);
  UpdateTable;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : A CATHELINEAU
Créé le ...... : 18/12/2002
Modifié le ... :   /  /
Description .. : Ouverture d'une fiche pour la saisie des articles uniques
Mots clefs ... :
*****************************************************************}

procedure TOF_AFFCDEMODIF.TraiterArticleUnique;
var QteLivrer, QteReservee, QteAffectee, AncienneQte, QteDispo, PxUnit, Montant,
  OldQte, OldMontant, QteStock: Double;
  Tiers, Article, QteModifie, Modifiable, Modifier, TypeAction,
    NaturePiece, Souche: string;
  Numero, indice: Integer;
  TOBA, TOBDis: TOB;
begin
  OldQte := 0;
  OldMontant := 0;
  QteDispo := 0;
  Montant := 0;
  QteStock := 0;
  if TV1.CurrentCol = TV1.ColIndex('GEL_QTERESERVEE') then
  begin
    Etape := afcReservation;
    TypeAction := 'RESERVATION';
  end
  else
    if TV1.CurrentCol = TV1.ColIndex('GEL_QTEAFFECTEE') then
  begin
    Etape := afcAffectation;
    TypeAction := 'AFFECTATION';
  end else exit;
  QteAffectee := 0;
  Modifiable := 'X';
  NaturePiece := TV1.AsString[TV1.ColIndex('GEL_NATUREPIECEG'), TV1.CurrentRow];
  Souche := TV1.AsString[TV1.ColIndex('GEL_SOUCHE'), TV1.CurrentRow];
  Numero := TV1.AsInteger[TV1.ColIndex('GEL_NUMERO'), TV1.CurrentRow];
  Indice := TV1.AsInteger[TV1.ColIndex('GEL_INDICEG'), TV1.CurrentRow];
  Tiers := TV1.AsString[TV1.ColIndex('GEL_TIERS'), TV1.CurrentRow];
  Article := TV1.AsString[TV1.ColIndex('GEL_CODEARTICLE'), TV1.CurrentRow];
  QteReservee := TV1.AsDouble[TV1.ColIndex('GEL_QTERESERVEE'), TV1.CurrentRow];
  QteLivrer := TV1.AsDouble[TV1.ColIndex('GEL_QTEALIVRER'), TV1.CurrentRow];
  //QteDispo:= TV1.AsDouble[TV1.ColIndex('GED_QTEPREVI'), TV1.CurrentRow] ;
  if Etape = afcAffectation then
  begin
    QteAffectee := TV1.AsDouble[TV1.ColIndex('GEL_QTEAFFECTEE'), TV1.CurrentRow];
    AncienneQte := QteAffectee;
  end else AncienneQte := QteReservee;
  // Chargement des Tob affect , piece et dispo
  ChargeLesTob(NaturePiece, Souche, Numero, Indice);
  TOBA := TobAffect.FindFirst(['GEL_CODEARTICLE', 'GEL_NUMERO', 'GEL_TYPEDIM'], [Article, Numero, 'NOR'], False);
  if TOBA <> nil then
  begin
    Modifiable := TOBA.GetValue('GEL_MODIFIABLE');
    TOBDis := TobDispo.FindFirst(['GED_ARTICLE'], [CodeArticleUnique2(Article, '')], False);
    if TOBDis <> nil then
    begin
      QteStock := TOBDis.GetValue('GED_QTEPREVI');
      QteDispo := TOBDis.GetValue('GED_DISPOAFFECTEE');
    end;
  end;
  QteModifie := AGLLanceFiche('MBO', 'AFFCDESAISIE', '', '', 'QteALivrer=' + FloatToStr(QteLivrer) + ';QteReservee=' + FloatToStr(QteReservee) + ';QteAffectee='
    + FloatToStr(QteAffectee) + ';QteDispo=' + FloatToStr(QteDispo) + ';QteStock=' + FloatToStr(QteStock) + ';Article=' + Article + ';Tiers=' + Tiers +
    ';TypeAction=' + TypeAction + ';Modifiable=' + Modifiable);
  // Mise à jour de la ligne
  QteReservee := StrToFloat(ReadTokenSt(QteModifie));
  QteAffectee := StrToFloat(ReadTokenSt(QteModifie));
  Modifiable := ReadTokenSt(QteModifie);
  Modifier := QteModifie;
  TOBA := TobAffect.FindFirst(['GEL_CODEARTICLE', 'GEL_NUMERO', 'GEL_TYPEDIM'], [Article, Numero, 'NOR'], False);
  if (TOBA <> nil) and (TOBA.GetValue('GEL_QTEALIVRER') <> 0) then
  begin
    if Modifier = 'X' then TOBA.AddChampSup('MODIFIER', False);
    TOBA.PutValue('GEL_MODIFIABLE', Modifiable);
    PxUnit := Arrondi(TOBA.GetValue('GEL_MONTANTHT') / TOBA.GetValue('GEL_QTEALIVRER'), 6);
    if Etape = AfcReservation then
    begin
      OldQte := TOBA.GetValue('GEL_QTERESERVEE');
      OldMontant := TOBA.GetValue('GEL_MTHTRESERVEE');
      TOBA.PutValue('GEL_QTERESERVEE', QteReservee);
      Montant := Arrondi(TOBA.GetValue('GEL_QTERESERVEE') * PxUnit, 6);
      TOBA.PutValue('GEL_MTHTRESERVEE', Montant);
    end
    else
    begin
      OldQte := TOBA.GetValue('GEL_QTEAFFECTEE');
      OldMontant := TOBA.GetValue('GEL_MTHTAFFECTEE');
      TOBA.PutValue('GEL_QTEAFFECTEE', QteAffectee);
      Montant := Arrondi(TOBA.GetValue('GEL_QTEAFFECTEE') * PxUnit, 6);
      TOBA.PutValue('GEL_MTHTAFFECTEE', Montant);
    end;
    // Mise à jour du dispo
    TOBDis := TobDispo.FindFirst(['GED_ARTICLE'], [CodeArticleUnique2(Article, '')], False);
    if TOBDis <> nil then
    begin
      if Etape = AfcReservation then QteLivrer := QteReservee
      else QteLivrer := QteAffectee;
      MajStockAFFCDE(TOBDis, QteLivrer, Etape, AncienneQte, True);
    end;
  end;
  // Mise à jour Piece
  MajPiece(TOBA, OldQte, OldMontant);
  // Mise à Jour TobViewer
  if Etape = AfcReservation then
  begin
    TV1.SetValueCell(TV1.ColIndex('GEL_QTERESERVEE'), TV1.CurrentROW, QteReservee);
    TV1.SetValueCell(TV1.ColIndex('GEL_MTHTRESERVEE'), TV1.CurrentROW, Montant);
  end else
  begin
    TV1.SetValueCell(TV1.ColIndex('GEL_QTEAFFECTEE'), TV1.CurrentROW, QteAffectee);
    TV1.SetValueCell(TV1.ColIndex('GEL_MTHTAFFECTEE'), TV1.CurrentROW, Montant);
  end;
  UpdateTable;
end;

procedure TOF_AFFCDEMODIF.UpdateTableAffect;
var CodeStatut: string;
begin
  TOBAffect.UpdateDB;
  TOBAffPiece.UpdateDB;
  TOBDispo.UpdateDB;

  // Mise à jour Entete
  CodeStatut := DonneCodeStatutAFFCDE(Etape);
  MajStatutAFFCDEENTETE(GetControlText('GEL_CODEAFF'), CodeStatut, True);
end;

procedure TOF_AFFCDEMODIF.UpdateTable;
var stPiece : string;
begin
  inc (nbTransacTot) ;
  if transactions (UpdateTableAffect, 0) <> oeOk then
  begin
    inc (nbTransacErr) ;
    stPiece := TobAffPiece.GetValue ('GEP_NATUREPIECEG') + ';' +
      TobAffPiece.GetValue ('GEP_SOUCHE') + ';' +
      IntToStr(TobAffPiece.GetValue ('GEP_NUMERO')) + ';' +
      IntToStr(TobAffPiece.GetValue ('GEP_INDICEG')) ;
    AjouteEvent ('AFF', 'Erreur AFFECTATION ' + stPiece, 'Erreur AFFECTATION Maj Pièce (TOF_AFFCDEMODIF.UpdateTable) : ' + stPiece, 'ERR');
  end;
end;

procedure TOF_AFFCDEMODIF.MajLigneGen(NaturePiece, Souche: string; Numero, Indice, NumLigneGen: Integer; Qte, PxUnit: Double);
var TOBA: TOB;
  NomChampQte, NomChampMt: string;
  Montant, OldQte, OldMontant: Double;
begin
  Montant := 0;
  OldQte := 0;
  OldMontant := 0;
  TOBA := TobAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE', 'GEL_TYPEDIM'], [NaturePiece, Souche, numero,
    indice, numLigneGen, 'GEN'], False);
  if Etape = AfcReservation then
  begin
    NomChampQte := 'GEL_QTERESERVEE';
    NomChampMt := 'GEL_MTHTRESERVEE'
  end
  else
  begin
    NomChampQte := 'GEL_QTEAFFECTEE';
    NomChampMt := 'GEL_MTHTAFFECTEE'
  end;
  if TOBA <> nil then
  begin
    OldQte := TOBA.GetValue(NomChampQte);
    OldMontant := TOBA.GetValue(NomChampMt);
    TOBA.PutValue(NomChampQte, Qte);
    Montant := Arrondi(Qte * PxUnit, 6);
    TOBA.PutValue(NomChampMt, Montant);
  end;
  TV1.SetValueCell(TV1.ColIndex(NomChampQte), TV1.CurrentROW, Qte);
  TV1.SetValueCell(TV1.ColIndex(NomChampMt), TV1.CurrentROW, Montant);
  MajPiece(TOBA, OldQte, OldMontant);
end;

procedure TOF_AFFCDEMODIF.MajPiece(TOBGen: TOB; OldQte, OldMontant: Double);
var NomChamp, NaturePiece, Souche: string;
  TOBP: TOB;
  Mont, Total, Qte,MontPiece: Double;
  Numero, indice: Integer;
begin
  if TOBGen = nil then Exit;
  if TOBGen.GetValue('GEL_TYPEDIM') = 'DIM' then Exit;
  // recherche de la pièce
  NaturePiece := TOBGen.GetValue('GEL_NATUREPIECEG');
  Souche := TOBGen.GetValue('GEL_SOUCHE');
  Numero := TOBGen.GetValue('GEL_NUMERO');
  Indice := TOBGen.GetValue('GEL_INDICEG');
  TOBP := TOBAffPiece.FindFirst(['GEP_NATUREPIECEG', 'GEP_SOUCHE', 'GEP_NUMERO', 'GEP_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
  if TOBP <> nil then
  begin
    // calcul du total HT
    if Etape = afcReservation then
    begin
      Mont := TOBGen.GetValue('GEL_MTHTRESERVEE');
      NomChamp := 'GEP_TOTHTRESERVEE';
    end else
    begin
      Mont := TOBGen.GetValue('GEL_MTHTAFFECTEE');
      NomChamp := 'GEP_TOTHTAFFECTEE';
    end;
    Total := TOBP.GetValue(NomChamp);
    Total := Arrondi(Total - OldMontant + Mont, 6);
    TOBP.PutValue(NomChamp, Total);
    // calcul du pourcentage du montant
    MontPiece := TOBP.GetValue('GEP_TOTALHT');
    if MontPiece <>0 then
       Mont := Arrondi((Total * 100) / MontPiece, 6)
       else Mont := 0 ;
    if Etape = afcReservation then
      NomChamp := 'GEP_POUMTRES'
    else
      NomChamp := 'GEP_POUMTAFF';
    TOBP.PutValue(NomChamp, Mont);
    // calcul de la quantité totale
    if Etape = afcReservation then
    begin
      Qte := TOBGen.GetValue('GEL_QTERESERVEE');
      NomChamp := 'GEP_TOTALQTERES';
    end else
    begin
      Qte := TOBGen.GetValue('GEL_QTEAFFECTEE');
      NomChamp := 'GEP_TOTALQTEAFF';
    end;
    Total := TOBP.GetValue(NomChamp);
    Total := Arrondi(Total - OldQte + Qte, 6);
    TOBP.PutValue(NomChamp, Total);
    // calcul du pourcentage de la quantite
    Qte := TOBP.GetValue('GEP_TOTALQTE');
    if Qte = 0 then
      Mont := 0
    else Mont := Arrondi((Total * 100) / Qte, 6);
    if Etape = afcReservation then
      NomChamp := 'GEP_POUQTERES'
    else
      NomChamp := 'GEP_POUQTEAFF';
    TOBP.PutValue(NomChamp, Mont);
    if Etape = afcReservation then
    begin
      TV1.SetValueCell(TV1.ColIndex('GEP_TOTHTRESERVEE'), TV1.CurrentROW, TOBP.GetValue('GEP_TOTHTRESERVEE'));
      TV1.SetValueCell(TV1.ColIndex('GEP_TOTALQTERES'), TV1.CurrentROW, TOBP.GetValue('GEP_TOTALQTERES'));
      TV1.SetValueCell(TV1.ColIndex('GEP_POUQTERES'), TV1.CurrentROW, TOBP.GetValue('GEP_POUQTERES'));
      TV1.SetValueCell(TV1.ColIndex('GEP_POUMTRES'), TV1.CurrentROW, TOBP.GetValue('GEP_POUMTRES'));
    end
    else
    begin
      TV1.SetValueCell(TV1.ColIndex('GEP_TOTHTAFFECTEE'), TV1.CurrentROW, TOBP.GetValue('GEP_TOTHTAFFECTEE'));
      TV1.SetValueCell(TV1.ColIndex('GEP_TOTALQTEAFF'), TV1.CurrentROW, TOBP.GetValue('GEP_TOTALQTEAFF'));
      TV1.SetValueCell(TV1.ColIndex('GEP_POUQTEAFF'), TV1.CurrentROW, TOBP.GetValue('GEP_POUQTEAFF'));
      TV1.SetValueCell(TV1.ColIndex('GEP_POUMTAFF'), TV1.CurrentROW, TOBP.GetValue('GEP_POUMTAFF'));
    end;
    MajLaTOB(TOBP);
  end;
end;

procedure TOF_AFFCDEMODIF.MajLaTOB(TOBP: TOB);
var i, numero, Indice, NumLigne, NumRow: Integer;
  NaturePiece, Souche: string;
begin
  for i := 0 to FTV.LaTOB.Detail.count - 1 do
  begin
    NaturePiece := FTV.LaTOB.Detail[i].GetValue('GEL_NATUREPIECEG');
    Souche := FTV.LaTOB.Detail[i].GetValue('GEL_SOUCHE');
    Numero := FTV.LaTOB.Detail[i].GetValue('GEL_NUMERO');
    Indice := FTV.LaTOB.Detail[i].GetValue('GEL_INDICEG');
    NumLigne := FTV.LaTOB.Detail[i].GetValue('GEL_NUMLIGNE');
    NumRow := TV1.FindRowWithKey(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE'], [NaturePiece, Souche, Numero, Indice,
      NumLigne]);
    if Etape = afcReservation then
    begin
      FTV.LaTOB.Detail[i].PutValue('GEP_TOTHTRESERVEE', TOBP.GetValue('GEP_TOTHTRESERVEE'));
      FTV.LaTOB.Detail[i].PutValue('GEP_TOTALQTERES', TOBP.GetValue('GEP_TOTALQTERES'));
      FTV.LaTOB.Detail[i].PutValue('GEP_POUQTERES', TOBP.GetValue('GEP_POUQTERES'));
      FTV.LaTOB.Detail[i].PutValue('GEP_POUMTRES', TOBP.GetValue('GEP_POUMTRES'));
      TV1.SetValueCell(TV1.ColIndex('GEP_TOTHTRESERVEE'), NumRow, TOBP.GetValue('GEP_TOTHTRESERVEE'));
      TV1.SetValueCell(TV1.ColIndex('GEP_TOTALQTERES'), NumRow, TOBP.GetValue('GEP_TOTALQTERES'));
      TV1.SetValueCell(TV1.ColIndex('GEP_POUQTERES'), NumRow, TOBP.GetValue('GEP_POUQTERES'));
      TV1.SetValueCell(TV1.ColIndex('GEP_POUMTRES'), NumRow, TOBP.GetValue('GEP_POUMTRES'));
    end else
    begin
      FTV.LaTOB.Detail[i].PutValue('GEP_TOTHTAFFECTEE', TOBP.GetValue('GEP_TOTHTAFFECTEE'));
      FTV.LaTOB.Detail[i].PutValue('GEP_TOTALQTEAFF', TOBP.GetValue('GEP_TOTALQTEAFF'));
      FTV.LaTOB.Detail[i].PutValue('GEP_POUQTEAFF', TOBP.GetValue('GEP_POUQTEAFF'));
      FTV.LaTOB.Detail[i].PutValue('GEP_POUMTAFF', TOBP.GetValue('GEP_POUMTAFF'));
      TV1.SetValueCell(TV1.ColIndex('GEP_TOTHTAFFECTEE'), NumRow, TOBP.GetValue('GEP_TOTHTAFFECTEE'));
      TV1.SetValueCell(TV1.ColIndex('GEP_TOTALQTEAFF'), NumRow, TOBP.GetValue('GEP_TOTALQTEAFF'));
      TV1.SetValueCell(TV1.ColIndex('GEP_POUQTEAFF'), NumRow, TOBP.GetValue('GEP_POUQTEAFF'));
      TV1.SetValueCell(TV1.ColIndex('GEP_POUMTAFF'), NumRow, TOBP.GetValue('GEP_POUMTAFF'));
    end;
  end;
end;

//************************** Cycle d'affectation *************************************/
//************************************************************************************/
// Affecter la selection
procedure TOF_AFFCDEMODIF.OnAffecterSelection;
var Numero, Indice: Integer;
  NaturePiece, Souche: string;
  Erreur: Boolean;
begin
  Numero := 0;
  Indice := 0;
  Erreur := False;
  nbTransacErr := 0 ; // Initialisation
  nbTransacTot := 0 ; // Initialisation
  Etape := afcAffectation;
  InitMoveProgressForm(FTV, FTV.Caption, 'Affectation en cours', 10, True, True);
  // Traitement des besoin en Query ;
  QTobViewer.First;
  while not QTobViewer.EOF do
  begin
    // Traitement Commande par Commande
    if (QTobViewer.Findfield('GEL_NATUREPIECEG').AsString <> NaturePiece) or
      (QTobViewer.Findfield('GEL_SOUCHE').AsString <> Souche) or
      (QTobViewer.Findfield('GEL_NUMERO').AsInteger <> Numero) or
      (QTobViewer.Findfield('GEL_INDICEG').AsInteger <> Indice) then
    begin
      if not MoveCurProgressForm('Traitement de la commande: "' + IntToStr(QTobViewer.Findfield('GEL_NUMERO').AsInteger) + '"') then
      begin
        PGIError(TraduireMemoire('Traitement annulé par l''utilisateur'), FTV.Caption);
        Erreur := True;
        break;
      end;
      NaturePiece := QTobViewer.Findfield('GEL_NATUREPIECEG').AsString;
      Souche := QTobViewer.Findfield('GEL_SOUCHE').AsString;
      Numero := QTobViewer.Findfield('GEL_NUMERO').AsInteger;
      Indice := QTobViewer.Findfield('GEL_INDICEG').AsInteger;
      // Chargement des Tob affect , piece et dispo
      ChargeLesTob(NaturePiece, Souche, Numero, Indice, True);
      AffecterCommande(NaturePiece, Souche, Numero, Indice);
    end;
    QTobViewer.next;
  end;
  FiniMoveProgressForm;
  if not Erreur then
  begin
    if nbTransacErr > 0
      then PGIInfo('Affectation terminée avec ERREUR, ' + IntToStr (nbTransacErr) + '/' + IntToStr (nbTransacTot) + ' pièces non mises à jour. Veuillez consultez le journal d''évènements', FTV.Caption)
      else PGIInfo('Affectation terminée. ' + IntToStr (nbTransacTot) + ' pièces affectées. Veuillez appliquer les critères', FTV.Caption);
  end;
  //PGIInfo('Affectation totale terminée:'+UsTime(NowH),FTV.Caption) ;
end;

// Affecter la commande
procedure TOF_AFFCDEMODIF.OnAffecterCommande;
var Numero, Indice, NumLigne, NumRow: Integer;
  NumChampArt,NumChampLivrer,NumChampReserve,NumChampQtePrevi,
  NumChampAff,NumChampMon,NumChampMonAff,NumChampType,NumChampLigne,
  NumChampCode,NumChampPQteAff,NumChampPMtAff: Integer;
  TOBA, TOBAffBis, TOBDis, TOBP: TOB;
  Article, CodeArticle, NaturePiece, Souche: string;
  TotQteAffectee, AncienneQte, PxUnit, Montant, OldQte, OldMontant, Qte, Total,
  MontantPiece: Double;
begin
  Etape := afcAffectation;
  NaturePiece := TV1.AsString[TV1.ColIndex('GEL_NATUREPIECEG'), TV1.CurrentRow];
  Souche := TV1.AsString[TV1.ColIndex('GEL_SOUCHE'), TV1.CurrentRow];
  Numero := TV1.AsInteger[TV1.ColIndex('GEL_NUMERO'), TV1.CurrentRow];
  Indice := TV1.AsInteger[TV1.ColIndex('GEL_INDICEG'), TV1.CurrentRow];
  // Chargement des Tob affect , piece et dispo
  ChargeLesTob(NaturePiece, Souche, Numero, Indice);
  // Recupération des numeros de champ
  // TOBAffect
  NumChampArt := TOBAffect.Detail[0].GetNumChamp('GEL_ARTICLE') ;
  NumChampLivrer := TOBAffect.Detail[0].GetNumChamp('GEL_QTEALIVRER') ;
  NumChampReserve := TOBAffect.Detail[0].GetNumChamp('GEL_QTERESERVEE') ;
  NumChampAff := TOBAffect.Detail[0].GetNumChamp('GEL_QTEAFFECTEE') ;
  NumChampMon :=  TOBAffect.Detail[0].GetNumChamp('GEL_MONTANTHT') ;
  NumChampMonAff := TOBAffect.Detail[0].GetNumChamp('GEL_MTHTAFFECTEE');
  NumChampType := TOBAffect.Detail[0].GetNumChamp('GEL_TYPEDIM') ;
  NumChampLigne :=TOBAffect.Detail[0].GetNumChamp('GEL_NUMLIGNE') ;
  NumChampCode := TOBAffect.Detail[0].GetNumChamp('GEL_CODESDIM') ;
  // TOBDispo
  NumChampQtePrevi := TOBDispo.Detail[0].GetNumChamp('GED_QTEPREVI') ;
  // TOBPiece
  NumChampPQteAff := TOBAffPiece.Detail[0].GetNumChamp('GEP_TOTALQTEAFF') ;
  NumChampPMtAff := TOBAffPiece.Detail[0].GetNumChamp('GEP_TOTHTAFFECTEE') ;
  //
  PxUnit := 0;
  TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
  while TOBA <> nil do
  begin
    Article := TOBA.GetValeur(NumChampArt);
    if (TOBA.GetValeur(NumChampType) <> 'GEN') and (TOBA.GetValeur(NumChampLivrer) <> 0) then
    begin
      TOBDis := TOBDispo.FindFirst(['GED_ARTICLE'], [Article], False);
      if (TOBDis <> nil) and (TOBA.GetValeur(NumChampReserve) <= TOBDis.GetValeur(NumChampQtePrevi)) then
      begin
        OldQte := TOBA.GetValeur(NumChampMonAff);
        OldMontant := TOBA.GetValeur(NumChampMonAff);
        if PxUnit <> 0 then
           PxUnit := Arrondi(TOBA.GetValeur(NumChampMon) / TOBA.GetValeur(NumChampLivrer), 6)
           else PxUnit := 0;
        AncienneQte := TOBA.GetValeur(NumChampAff);
        TOBA.PutValeur(NumChampAff, TOBA.GetValeur(NumChampReserve));
        Montant := Arrondi(TOBA.GetValeur(NumChampAff) * PxUnit, 6);
        TOBA.PutValeur(NumChampMonAff, Montant);
        MajStockAFFCDE(TOBDis, TOBA.GetValeur(NumChampAff), Etape, AncienneQte);
        // Mise à jour de la piece pour un article unique
        if TOBA.GetValeur(NumChampType) = 'NOR' then
        begin
          NumLigne := TOBA.GetValeur(NumChampLigne);
          NumRow := TV1.FindRowWithKey(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE'], [NaturePiece, Souche, Numero, Indice,
            NumLigne]);
          if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_QTEAFFECTEE'), NumRow, TOBA.GetValue('GEL_QTERESERVEE'));
          if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_MTHTAFFECTEE'), NumRow, Montant);
          MajPiece(TOBA, OldQte, OldMontant);
        end;
      end;
    end;
    TOBA := TOBAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
  end;
  // Mise à jour ligne générique
  TOBAffBis := TOB.Create('_Les affectations', nil, -1);
  TOBAffBis.Dupliquer(TOBAffect, TRUE, TRUE, TRUE);
  TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, 'GEN'],
    False);
  while TOBA <> nil do
  begin
    if TOBA.GetValeur(NumChampLivrer) <> 0 then
    begin
      CodeArticle := TOBA.GetValeur(NumChampCode);
      NumLigne := TOBA.GetValeur(NumChampLigne);
      NumRow := TV1.FindRowWithKey(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE'], [NaturePiece, Souche, Numero, Indice,
        NumLigne]);
      // Quantité affectee des dimmensions
      TotQteAffectee := TOBAffBis.Somme('GEL_QTEAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNEGEN', 'GEL_CODEARTICLE',
        'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, NumLigne, CodeArticle, 'DIM'], TRUE);
      TOBA.PutValeur(NumChampAff, TotQteAffectee);
      if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_QTEAFFECTEE'), NumRow, TotQteAffectee);
      // Montant affectee
      PxUnit := Arrondi(TOBA.GetValeur(NumChampMon) / TOBA.GetValeur(NumChampLivrer), 6);
      Montant := Arrondi(TotQteAffectee * PxUnit, 6);
      TOBA.PutValeur(NumChampMonAff, Montant);
      if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_MTHTAFFECTEE'), NumRow, Montant);
    end;
    TOBA := TobAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, 'GEN'],
      False);
  end;
  // Mise à jour de la piece
  TOBP := TOBAffPiece.FindFirst(['GEP_NATUREPIECEG', 'GEP_SOUCHE', 'GEP_NUMERO', 'GEP_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
  while TOBP <> nil do
  begin
    // Quantité affectee de la piece
    Qte := TOBAffect.Somme('GEL_QTEAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice], TRUE);
    OldQte := TOBAffBis.Somme('GEL_QTEAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice], TRUE);
    TotQteAffectee := TOBP.GetValeur(NumChampPQteAff);
    TotQteAffectee := Arrondi(TotQteAffectee - OldQte + Qte, 6);
    TOBP.PutValeur(NumChampPQteAff, TotQteAffectee);
    TV1.SetValueCell(TV1.ColIndex('GEP_TOTALQTEAFF'), TV1.CurrentROW, TotQteAffectee);
    // Montant affectee de la piece
    Montant := TOBAffect.Somme('GEL_MTHTAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice],
      TRUE);
    OldMontant := TOBAffBis.Somme('GEL_MTHTAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice],
      TRUE);
    Total := TOBP.GetValeur(NumChampPMtAff);
    Montant := Arrondi(Total - OldMontant + Montant, 6);
    TOBP.PutValeur(NumChampPMtAff, Montant);
    TV1.SetValueCell(TV1.ColIndex('GEP_TOTHTAFFECTEE'), TV1.CurrentROW, Montant);
    // calcul du pourcentage affecté en montant
    MontantPiece:=TOBP.GetValue('GEP_TOTALHT');
    if MontantPiece <> 0 then
       MontantPiece := Arrondi((Montant * 100) / MontantPiece, 6)
       else MontantPiece := 0 ;
    TOBP.PutValue('GEP_POUMTAFF', MontantPiece);
    TV1.SetValueCell(TV1.ColIndex('GEP_POUMTAFF'), TV1.CurrentROW, MontantPiece);
    // calcul du pourcentage affecté en quantite
    Qte := TOBP.GetValue('GEP_TOTALQTE');
    if Qte <> 0 then
       Montant := Arrondi((TotQteAffectee * 100) / Qte, 6)
       else Montant := 0 ;
    TOBP.PutValue('GEP_POUQTEAFF', Montant);
    TV1.SetValueCell(TV1.ColIndex('GEP_POUQTEAFF'), TV1.CurrentROW, Montant);
    TOBP := TobAffPiece.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, 'GEN'],
      False);
  end;
  // Modif tables
  UpdateTable;
  TOBAffBis.Free;
  PGIInfo('Affectation de la commande terminée', FTV.Caption);
end;

// procedure d'affectation de la commande
procedure TOF_AFFCDEMODIF.AffecterCommande(NaturePiece, Souche: string; Numero, Indice: Integer);
var NumLigne,NumChampArt,NumChampLivrer,NumChampReserve,NumChampQtePrevi,
    NumChampAff,NumChampMon,NumChampMonAff,NumChampType,NumChampLigne,
    NumChampCode,NumChampPQteAff,NumChampPMtAff: Integer;
  TOBA, TOBAffBis, TOBDis, TOBP: TOB;
  Article, CodeArticle: string;
  TotQteAffectee, AncienneQte, PxUnit, Montant, OldQte, OldMontant, Qte, Total,MontantPiece: Double;
  LigneTraite: Boolean;
begin
  LigneTraite := False;
  // Recupération des numeros de champ
  // TOBAffect
  NumChampArt := TOBAffect.Detail[0].GetNumChamp('GEL_ARTICLE') ;
  NumChampLivrer := TOBAffect.Detail[0].GetNumChamp('GEL_QTEALIVRER') ;
  NumChampReserve := TOBAffect.Detail[0].GetNumChamp('GEL_QTERESERVEE') ;
  NumChampAff := TOBAffect.Detail[0].GetNumChamp('GEL_QTEAFFECTEE') ;
  NumChampMon :=  TOBAffect.Detail[0].GetNumChamp('GEL_MONTANTHT') ;
  NumChampMonAff := TOBAffect.Detail[0].GetNumChamp('GEL_MTHTAFFECTEE');
  NumChampType := TOBAffect.Detail[0].GetNumChamp('GEL_TYPEDIM') ;
  NumChampLigne :=TOBAffect.Detail[0].GetNumChamp('GEL_NUMLIGNE') ;
  NumChampCode := TOBAffect.Detail[0].GetNumChamp('GEL_CODESDIM') ;
  // TOBDispo
  NumChampQtePrevi := TOBDispo.Detail[0].GetNumChamp('GED_QTEPREVI') ;
  // TOBPiece
  NumChampPQteAff := TOBAffPiece.Detail[0].GetNumChamp('GEP_TOTALQTEAFF') ;
  NumChampPMtAff := TOBAffPiece.Detail[0].GetNumChamp('GEP_TOTHTAFFECTEE') ;
  //
  TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
  while TOBA <> nil do
  begin
    Article := TOBA.GetValeur(NumChampArt) ;
    if (TOBA.GetValeur(NumChampType) <> 'GEN') and (TOBA.GetValeur(NumChampLivrer) <> 0) and (TOBA.GetValeur(NumChampReserve) > 0) then
    begin
      TOBDis := TOBDispo.FindFirst(['GED_ARTICLE'], [Article], False);
      if (TOBDis <> nil) and (TOBA.GetValeur(NumChampReserve) <= TOBDis.GetValeur(NumChampQtePrevi)) then
      begin
        LigneTraite := True;
        OldQte := TOBA.GetValeur(NumChampAff);
        OldMontant := TOBA.GetValeur(NumChampMonAff);
        PxUnit := Arrondi(TOBA.GetValeur(NumChampMon) / TOBA.GetValeur(NumChampLivrer), 6);
        AncienneQte := TOBA.GetValeur(NumChampAff);
        TOBA.PutValeur(NumChampAff, TOBA.GetValeur(NumChampReserve));
        Montant := Arrondi(TOBA.GetValeur(NumChampAff) * PxUnit, 6);
        TOBA.PutValeur(NumChampMonAff, Montant);
        MajStockAFFCDE(TOBDis, TOBA.GetValeur(NumChampAff), Etape, AncienneQte);
        // Mise à jour de la piece pour un article unique
        if TOBA.GetValeur(NumChampType) = 'NOR' then MajPiece(TOBA, OldQte, OldMontant);
      end;
    end;
    TOBA := TOBAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
  end;
  // Si ligne traite faire Mise à jour Ligne gen et Piece
  // Mise à jour ligne générique
  if LigneTraite then
  begin
    TOBAffBis := TOB.Create('_Les affectations', nil, -1);
    TOBAffBis.Dupliquer(TOBAffect, TRUE, TRUE, TRUE);
    TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, 'GEN'],
      False);
    while TOBA <> nil do
    begin
      if TOBA.GetValeur(NumChampLivrer) <> 0 then
      begin
        CodeArticle := TOBA.GetValeur(NumChampCode);
        NumLigne := TOBA.GetValeur(NumChampLigne);
        // Quantité affectee des dimmensions
        TotQteAffectee := TOBAffBis.Somme('GEL_QTEAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNEGEN',
          'GEL_CODEARTICLE', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, NumLigne, CodeArticle, 'DIM'], TRUE);
        TOBA.PutValeur(NumChampAff, TotQteAffectee);
        // Montant affectee
        PxUnit := Arrondi(TOBA.GetValeur(NumChampMon) / TOBA.GetValeur(NumChampLivrer), 6);
        Montant := Arrondi(TotQteAffectee * PxUnit, 6);
        TOBA.PutValeur(NumChampMonAff, Montant);
      end;
      TOBA := TobAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, 'GEN'],
        False);
    end;
    // Mise à jour de la piece
    TOBP := TOBAffPiece.FindFirst(['GEP_NATUREPIECEG', 'GEP_SOUCHE', 'GEP_NUMERO', 'GEP_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
    while TOBP <> nil do
    begin
      // Quantité affectee de la piece
      Qte := TOBAffect.Somme('GEL_QTEAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice], TRUE);
      OldQte := TOBAffBis.Somme('GEL_QTEAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice],
        TRUE);
      TotQteAffectee := TOBP.GetValeur(NumChampPQteAff);
      TotQteAffectee := Arrondi(TotQteAffectee - OldQte + Qte, 6);
      TOBP.PutValeur(NumChampPQteAff, TotQteAffectee);
      // Montant affectee de la piece
      Montant := TOBAffect.Somme('GEL_MTHTAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice],
        TRUE);
      OldMontant := TOBAffBis.Somme('GEL_MTHTAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice],
        TRUE);
      Total := TOBP.GetValeur(NumChampPMtAff);
      Montant := Arrondi(Total - OldMontant + Montant, 6);
      TOBP.PutValeur(NumChampPMtAff, Montant);
      // calcul du pourcentage affecté en montant
     MontantPiece:=TOBP.GetValue('GEP_TOTALHT');
     if MontantPiece <> 0 then
        MontantPiece := Arrondi((Montant * 100) / MontantPiece, 6)
        else MontantPiece := 0 ;
     TOBP.PutValue('GEP_POUMTAFF', MontantPiece);
     TV1.SetValueCell(TV1.ColIndex('GEP_POUMTAFF'), TV1.CurrentROW, MontantPiece);
      // calcul du pourcentage affecté
      Qte := TOBP.GetValue('GEP_TOTALQTE');
      if Qte <> 0 then
         Montant := Arrondi((TotQteAffectee * 100) / Qte, 6)
         else Montant := 0 ;
      TOBP.PutValue('GEP_POUQTEAFF', Montant);
      TOBP := TobAffPiece.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice,
        'GEN'], False);
    end;
    // Modif tables
    UpdateTable;
    TOBAffBis.Free;
  end;
end;

// Affecter la ligne
procedure TOF_AFFCDEMODIF.OnAffecterLigne;
var Numero, Indice, NumLigne, NumRow: Integer;
  TOBA, TOBAffBis, TOBDis: TOB;
  Article, CodeArticle, NaturePiece, Souche: string;
  TotQteAffectee, AncienneQte, PxUnit, Montant, OldQte, OldMontant: Double;
begin
  Etape := afcAffectation;
  NaturePiece := TV1.AsString[TV1.ColIndex('GEL_NATUREPIECEG'), TV1.CurrentRow];
  Souche := TV1.AsString[TV1.ColIndex('GEL_SOUCHE'), TV1.CurrentRow];
  Numero := TV1.AsInteger[TV1.ColIndex('GEL_NUMERO'), TV1.CurrentRow];
  Indice := TV1.AsInteger[TV1.ColIndex('GEL_INDICEG'), TV1.CurrentRow];
  NumLigne := TV1.AsInteger[TV1.ColIndex('GEL_NUMLIGNE'), TV1.CurrentRow];
  // Chargement des Tob affect , piece et dispo
  ChargeLesTob(NaturePiece, Souche, Numero, Indice);
  TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE'], [NaturePiece, Souche, Numero, Indice, NumLigne],
    False);
  while TOBA <> nil do
  begin
    Article := TOBA.GetValue('GEL_ARTICLE');
    if (TOBA.GetValue('GEL_TYPEDIM') <> 'GEN') and (TOBA.GetValue('GEL_QTEALIVRER') <> 0) then
    begin
      TOBDis := TOBDispo.FindFirst(['GED_ARTICLE'], [Article], False);
      if (TOBDis <> nil) and (TOBA.GetValue('GEL_QTERESERVEE') <= TOBDis.GetValue('GED_QTEPREVI')) then
      begin
        OldQte := TOBA.GetValue('GEL_QTEAFFECTEE');
        OldMontant := TOBA.GetValue('GEL_MTHTAFFECTEE');
        PxUnit := Arrondi(TOBA.GetValue('GEL_MONTANTHT') / TOBA.GetValue('GEL_QTEALIVRER'), 6);
        AncienneQte := TOBA.GetValue('GEL_QTEAFFECTEE');
        TOBA.PutValue('GEL_QTEAFFECTEE', TOBA.GetValue('GEL_QTERESERVEE'));
        Montant := Arrondi(TOBA.GetValue('GEL_QTEAFFECTEE') * PxUnit, 6);
        TOBA.PutValue('GEL_MTHTAFFECTEE', Montant);
        MajStockAFFCDE(TOBDis, TOBA.GetValue('GEL_QTEAFFECTEE'), Etape, AncienneQte);
        // Mise à jour de la piece pour un article unique
        if TOBA.GetValue('GEL_TYPEDIM') = 'NOR' then
        begin
          NumLigne := TOBA.GetValue('GEL_NUMLIGNE');
          NumRow := TV1.FindRowWithKey(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE'], [NaturePiece, Souche, Numero, Indice,
            NumLigne]);
          if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_QTEAFFECTEE'), NumRow, TOBA.GetValue('GEL_QTERESERVEE'));
          if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_MTHTAFFECTEE'), NumRow, Montant);
          MajPiece(TOBA, OldQte, OldMontant);
        end;
      end;
    end;
    TOBA := TOBAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNEGEN'], [NaturePiece, Souche, Numero, Indice,
      NumLigne], False);
  end;
  // Mise à jour ligne générique
  TOBAffBis := TOB.Create('_Les affectations', nil, -1);
  TOBAffBis.Dupliquer(TOBAffect, TRUE, TRUE, TRUE);
  TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero,
    Indice, NumLigne, 'GEN'], False);
  while TOBA <> nil do
  begin
    if TOBA.GetValue('GEL_QTEALIVRER') <> 0 then
    begin
      CodeArticle := TOBA.GetValue('GEL_CODESDIM');
      OldQte := TOBA.GetValue('GEL_QTEAFFECTEE');
      OldMontant := TOBA.GetValue('GEL_MTHTAFFECTEE');
      // Quantité affectee des dimmensions
      TotQteAffectee := TOBAffBis.Somme('GEL_QTEAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNEGEN', 'GEL_CODEARTICLE',
        'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, NumLigne, CodeArticle, 'DIM'], TRUE);
      TOBA.PutValue('GEL_QTEAFFECTEE', TotQteAffectee);
      TV1.SetValueCell(TV1.ColIndex('GEL_QTEAFFECTEE'), TV1.CurrentROW, TotQteAffectee);
      // Montant affectee
      PxUnit := Arrondi(TOBA.GetValue('GEL_MONTANTHT') / TOBA.GetValue('GEL_QTEALIVRER'), 6);
      Montant := Arrondi(TotQteAffectee * PxUnit, 6);
      TOBA.PutValue('GEL_MTHTAFFECTEE', Montant);
      TV1.SetValueCell(TV1.ColIndex('GEL_MTHTAFFECTEE'), TV1.CurrentROW, Montant);
      // Mise à jour Piece
      MajPiece(TOBA, OldQte, OldMontant) ;
    end;
    TOBA := TobAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNEGEN', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero,
      Indice, NumLigne, 'GEN'], False);
  end;
  // Modif tables
  UpdateTable;
  TOBAffBis.Free;
end;
//************************** Fin Cycle d'affectation **********************************/
//************************************************************************************/

//************************** Cycle desaffectation ************************************/
//************************************************************************************/
// Désaffecter la selection
procedure TOF_AFFCDEMODIF.OnDesAffecterSelection;
var Numero, Indice: Integer;
  NaturePiece, Souche: string;
  Erreur: Boolean;
begin
  Numero := 0;
  Indice := 0;
  Erreur := False;
  nbTransacErr := 0 ; // Initialisation
  nbTransacTot := 0 ; // Initialisation
  Etape := afcAffectation;
  InitMoveProgressForm(FTV, FTV.Caption, 'Affectation en cours', 10, True, True);
  // Traitement des besoin en Query ;
  QTobViewer.First;
  while not QTobViewer.EOF do
  begin
    // Traitement Commande par Commande
    if (QTobViewer.Findfield('GEL_NATUREPIECEG').AsString <> NaturePiece) or
      (QTobViewer.Findfield('GEL_SOUCHE').AsString <> Souche) or
      (QTobViewer.Findfield('GEL_NUMERO').AsInteger <> Numero) or
      (QTobViewer.Findfield('GEL_INDICEG').AsInteger <> Indice) then
    begin
      if not MoveCurProgressForm('Traitement de la commande: "' + IntToStr(QTobViewer.Findfield('GEL_NUMERO').AsInteger) + '"') then
      begin
        PGIError(TraduireMemoire('Traitement annulé par l''utilisateur'), FTV.Caption);
        Erreur := True;
        break;
      end;
      NaturePiece := QTobViewer.Findfield('GEL_NATUREPIECEG').AsString;
      Souche := QTobViewer.Findfield('GEL_SOUCHE').AsString;
      Numero := QTobViewer.Findfield('GEL_NUMERO').AsInteger;
      Indice := QTobViewer.Findfield('GEL_INDICEG').AsInteger;
      // Chargement des Tob affect , piece et dispo
      ChargeLesTob(NaturePiece, Souche, Numero, Indice, True);
      DesAffecterCommande(NaturePiece, Souche, Numero, Indice);
    end;
    QTobViewer.next;
  end;
  FiniMoveProgressForm;
  if not Erreur then PGIInfo('Desaffectation totale terminée. Veuillez appliquer les critères', FTV.Caption);
  if not Erreur then
  begin
    if nbTransacErr > 0
      then PGIInfo('Désaffectation terminée avec ERREUR, ' + IntToStr (nbTransacErr) + '/' + IntToStr (nbTransacTot) + ' pièces non mises à jour. Veuillez consultez le journal d''évènements', FTV.Caption)
      else PGIInfo('Désaffectation terminée. ' + IntToStr (nbTransacTot) + ' pièces désaffectées. Veuillez appliquer les critères', FTV.Caption);
  end;
end;

// procedure de désaffectation de la commande
procedure TOF_AFFCDEMODIF.DesAffecterCommande(NaturePiece, Souche: string; Numero, Indice: Integer);
var NumLigne: Integer;
  TOBA, TOBAffBis, TOBDis, TOBP: TOB;
  Article, CodeArticle: string;
  TotQteAffectee, AncienneQte, PxUnit, Montant, OldQte, OldMontant : Double;
  LigneTraite: Boolean;
begin
  LigneTraite := False;
  TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
  while TOBA <> nil do
  begin
    Article := TOBA.GetValue('GEL_ARTICLE');
    if (TOBA.GetValue('GEL_TYPEDIM') <> 'GEN') and (TOBA.GetValue('GEL_QTEALIVRER') <> 0) and (TOBA.GetValue('GEL_QTERESERVEE') > 0) then
    begin
      TOBDis := TOBDispo.FindFirst(['GED_ARTICLE'], [Article], False);
      if (TOBDis <> nil) and (TOBA.GetValue('GEL_QTERESERVEE') <= TOBDis.GetValue('GED_QTEPREVI')) then
      begin
        LigneTraite := True;
        OldQte := TOBA.GetValue('GEL_QTEAFFECTEE');
        OldMontant := TOBA.GetValue('GEL_MTHTAFFECTEE');
        AncienneQte := TOBA.GetValue('GEL_QTEAFFECTEE');
        TOBA.PutValue('GEL_QTEAFFECTEE', 0);
        TOBA.PutValue('GEL_MTHTAFFECTEE', 0);
        MajStockAFFCDE(TOBDis, 0, Etape, AncienneQte,True);
        // Mise à jour de la piece pour un article unique
        if TOBA.GetValue('GEL_TYPEDIM') = 'NOR' then
        begin
          MajPiece(TOBA, OldQte, OldMontant);
        end;
      end;
    end;
    TOBA := TOBAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
  end;
  // Si ligne traite faire Mise à jour Ligne gen et Piece
  // Mise à jour ligne générique
  if LigneTraite then
  begin
    TOBAffBis := TOB.Create('_Les affectations', nil, -1);
    TOBAffBis.Dupliquer(TOBAffect, TRUE, TRUE, TRUE);
    TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, 'GEN'],
      False);
    while TOBA <> nil do
    begin
      if TOBA.GetValue('GEL_QTEALIVRER') <> 0 then
      begin
        CodeArticle := TOBA.GetValue('GEL_CODESDIM');
        NumLigne := TOBA.GetValue('GEL_NUMLIGNE');
        // Quantité affectee des dimmensions
        TotQteAffectee := TOBAffBis.Somme('GEL_QTEAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNEGEN',
          'GEL_CODEARTICLE', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, NumLigne, CodeArticle, 'DIM'], TRUE);
        TOBA.PutValue('GEL_QTEAFFECTEE', TotQteAffectee);
        // Montant affectee
        PxUnit := Arrondi(TOBA.GetValue('GEL_MONTANTHT') / TOBA.GetValue('GEL_QTEALIVRER'), 6);
        Montant := Arrondi(TotQteAffectee * PxUnit, 6);
        TOBA.PutValue('GEL_MTHTAFFECTEE', Montant);
      end;
      TOBA := TobAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, 'GEN'],
        False);
    end;
    // Mise à jour de la piece
    TOBP := TOBAffPiece.FindFirst(['GEP_NATUREPIECEG', 'GEP_SOUCHE', 'GEP_NUMERO', 'GEP_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
    while TOBP <> nil do
    begin
      // Quantité affectee de la piece
      TOBP.PutValue('GEP_TOTALQTEAFF', 0);
      // Montant affectee de la piece
      TOBP.PutValue('GEP_TOTHTAFFECTEE', 0);
      // calcul du pourcentage affecté en montant
      TOBP.PutValue('GEP_POUMTAFF', 0);
     // calcul du pourcentage affecté
      TOBP.PutValue('GEP_POUQTEAFF', 0);
      TOBP := TobAffPiece.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice,
        'GEN'], False);
    end;
    // Modif tables
    UpdateTable;
    TOBAffBis.Free;
  end;
end;

// Désaffecter la commande
procedure TOF_AFFCDEMODIF.OnDesAffecterCommande;
var Numero, Indice, NumLigne, NumRow: Integer;
  TOBA, TOBAffBis, TOBDis, TOBP: TOB;
  Article, CodeArticle, NaturePiece, Souche: string;
  TotQteAffectee, AncienneQte, PxUnit, Montant: Double;
begin
  Etape := afcAffectation;
  NaturePiece := TV1.AsString[TV1.ColIndex('GEL_NATUREPIECEG'), TV1.CurrentRow];
  Souche := TV1.AsString[TV1.ColIndex('GEL_SOUCHE'), TV1.CurrentRow];
  Numero := TV1.AsInteger[TV1.ColIndex('GEL_NUMERO'), TV1.CurrentRow];
  Indice := TV1.AsInteger[TV1.ColIndex('GEL_INDICEG'), TV1.CurrentRow];
  // Chargement des Tob affect , piece et dispo
  ChargeLesTob(NaturePiece, Souche, Numero, Indice);
  TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
  while TOBA <> nil do
  begin
    Article := TOBA.GetValue('GEL_ARTICLE');
    if (TOBA.GetValue('GEL_TYPEDIM') <> 'GEN') and (TOBA.GetValue('GEL_QTEALIVRER') <> 0) then
    begin
      TOBDis := TOBDispo.FindFirst(['GED_ARTICLE'], [Article], False);
      if (TOBDis <> nil) and (TOBA.GetValue('GEL_QTERESERVEE') <= TOBDis.GetValue('GED_QTEPREVI')) then
      begin
        AncienneQte := TOBA.GetValue('GEL_QTEAFFECTEE');
        TOBA.PutValue('GEL_QTEAFFECTEE', 0);
        TOBA.PutValue('GEL_MTHTAFFECTEE', 0);
        MajStockAFFCDE(TOBDis, 0, Etape, AncienneQte,True);
        // Mise à jour de la piece pour un article unique
        if TOBA.GetValue('GEL_TYPEDIM') = 'NOR' then
        begin
          NumLigne := TOBA.GetValue('GEL_NUMLIGNE');
          NumRow := TV1.FindRowWithKey(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE'], [NaturePiece, Souche, Numero, Indice,
            NumLigne]);
          if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_QTEAFFECTEE'), NumRow, 0);
          if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_MTHTAFFECTEE'), NumRow, 0);
          MajPiece(TOBA, 0, 0);
        end;
      end;
    end;
    TOBA := TOBAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
  end;
  // Mise à jour ligne générique
  TOBAffBis := TOB.Create('_Les affectations', nil, -1);
  TOBAffBis.Dupliquer(TOBAffect, TRUE, TRUE, TRUE);
  TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, 'GEN'],
    False);
  while TOBA <> nil do
  begin
    if TOBA.GetValue('GEL_QTEALIVRER') <> 0 then
    begin
      CodeArticle := TOBA.GetValue('GEL_CODESDIM');
      NumLigne := TOBA.GetValue('GEL_NUMLIGNE');
      NumRow := TV1.FindRowWithKey(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE'], [NaturePiece, Souche, Numero, Indice,
        NumLigne]);
      // Quantité affectee des dimmensions
      TotQteAffectee := TOBAffBis.Somme('GEL_QTEAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNEGEN', 'GEL_CODEARTICLE',
        'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, NumLigne, CodeArticle, 'DIM'], TRUE);
      TOBA.PutValue('GEL_QTEAFFECTEE', TotQteAffectee);
      if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_QTEAFFECTEE'), NumRow, TotQteAffectee);
      // Montant affectee
      PxUnit := Arrondi(TOBA.GetValue('GEL_MONTANTHT') / TOBA.GetValue('GEL_QTEALIVRER'), 6);
      Montant := Arrondi(TotQteAffectee * PxUnit, 6);
      TOBA.PutValue('GEL_MTHTAFFECTEE', Montant);
      if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_MTHTAFFECTEE'), NumRow, Montant);
    end;
    TOBA := TobAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, 'GEN'],
      False);
  end;
  // Mise à jour de la piece
  TOBP := TOBAffPiece.FindFirst(['GEP_NATUREPIECEG', 'GEP_SOUCHE', 'GEP_NUMERO', 'GEP_INDICEG'], [NaturePiece, Souche, Numero, Indice], False);
  while TOBP <> nil do
  begin
    // Quantité affectee de la piece
    TOBP.PutValue('GEP_TOTALQTEAFF', 0);
    TV1.SetValueCell(TV1.ColIndex('GEP_TOTALQTEAFF'), TV1.CurrentROW, 0);
    // Montant affectee de la piece
    TOBP.PutValue('GEP_TOTHTAFFECTEE',0);
    TV1.SetValueCell(TV1.ColIndex('GEP_TOTHTAFFECTEE'), TV1.CurrentROW, 0);
    // calcul du pourcentage affecté en montant
    TOBP.PutValue('GEP_POUMTAFF', 0);
    TV1.SetValueCell(TV1.ColIndex('GEP_POUMTAFF'), TV1.CurrentROW, 0);
    // calcul du pourcentage affecté en quantite
    TOBP.PutValue('GEP_POUQTEAFF', 0);
    TV1.SetValueCell(TV1.ColIndex('GEP_POUQTEAFF'), TV1.CurrentROW, 0);
    TOBP := TobAffPiece.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, 'GEN'],
      False);
  end;
  // Modif tables
  UpdateTable;
  TOBAffBis.Free;
  PGIInfo('Désaffectation de la commande terminée', FTV.Caption);
end;

// Désaffecter la ligne
procedure TOF_AFFCDEMODIF.OnDesAffecterLigne;
var Numero, Indice, NumLigne, NumRow: Integer;
  TOBA, TOBAffBis, TOBDis: TOB;
  Article, CodeArticle, NaturePiece, Souche: string;
  TotQteAffectee, AncienneQte, PxUnit, Montant, OldQte, OldMontant: Double;
begin
  Etape := afcAffectation;
  NaturePiece := TV1.AsString[TV1.ColIndex('GEL_NATUREPIECEG'), TV1.CurrentRow];
  Souche := TV1.AsString[TV1.ColIndex('GEL_SOUCHE'), TV1.CurrentRow];
  Numero := TV1.AsInteger[TV1.ColIndex('GEL_NUMERO'), TV1.CurrentRow];
  Indice := TV1.AsInteger[TV1.ColIndex('GEL_INDICEG'), TV1.CurrentRow];
  NumLigne := TV1.AsInteger[TV1.ColIndex('GEL_NUMLIGNE'), TV1.CurrentRow];
  // Chargement des Tob affect , piece et dispo
  ChargeLesTob(NaturePiece, Souche, Numero, Indice);
  TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE'], [NaturePiece, Souche, Numero, Indice, NumLigne],
    False);
  while TOBA <> nil do
  begin
    Article := TOBA.GetValue('GEL_ARTICLE');
    if (TOBA.GetValue('GEL_TYPEDIM') <> 'GEN') and (TOBA.GetValue('GEL_QTEALIVRER') <> 0) then
    begin
      TOBDis := TOBDispo.FindFirst(['GED_ARTICLE'], [Article], False);
      if (TOBDis <> nil) and (TOBA.GetValue('GEL_QTERESERVEE') <= TOBDis.GetValue('GED_QTEPREVI')) then
      begin
        OldQte := TOBA.GetValue('GEL_QTEAFFECTEE');
        OldMontant := TOBA.GetValue('GEL_MTHTAFFECTEE');
        AncienneQte := TOBA.GetValue('GEL_QTEAFFECTEE');
        TOBA.PutValue('GEL_QTEAFFECTEE', 0);
        TOBA.PutValue('GEL_MTHTAFFECTEE', 0);
        MajStockAFFCDE(TOBDis, 0, Etape, AncienneQte,True);
        // Mise à jour de la piece pour un article unique
        if TOBA.GetValue('GEL_TYPEDIM') = 'NOR' then
        begin
          NumLigne := TOBA.GetValue('GEL_NUMLIGNE');
          NumRow := TV1.FindRowWithKey(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE'], [NaturePiece, Souche, Numero, Indice,
            NumLigne]);
          if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_QTEAFFECTEE'), NumRow, 0);
          if NumRow >= 0 then TV1.SetValueCell(TV1.ColIndex('GEL_MTHTAFFECTEE'), NumRow,0);
          MajPiece(TOBA, OldQte, OldMontant);
        end;
      end;
    end;
    TOBA := TOBAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNEGEN'], [NaturePiece, Souche, Numero, Indice,
      NumLigne], False);
  end;
  // Mise à jour ligne générique
  TOBAffBis := TOB.Create('_Les affectations', nil, -1);
  TOBAffBis.Dupliquer(TOBAffect, TRUE, TRUE, TRUE);
  TOBA := TOBAffect.FindFirst(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNE', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero,
    Indice, NumLigne, 'GEN'], False);
  while TOBA <> nil do
  begin
    if TOBA.GetValue('GEL_QTEALIVRER') <> 0 then
    begin
      CodeArticle := TOBA.GetValue('GEL_CODESDIM');
      OldQte := TOBA.GetValue('GEL_QTEAFFECTEE');
      OldMontant := TOBA.GetValue('GEL_MTHTAFFECTEE');
      // Quantité affectee des dimmensions
      TotQteAffectee := TOBAffBis.Somme('GEL_QTEAFFECTEE', ['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNEGEN', 'GEL_CODEARTICLE',
        'GEL_TYPEDIM'], [NaturePiece, Souche, Numero, Indice, NumLigne, CodeArticle, 'DIM'], TRUE);
      TOBA.PutValue('GEL_QTEAFFECTEE', TotQteAffectee);
      TV1.SetValueCell(TV1.ColIndex('GEL_QTEAFFECTEE'), TV1.CurrentROW, TotQteAffectee);
      // Montant affectee
      PxUnit := Arrondi(TOBA.GetValue('GEL_MONTANTHT') / TOBA.GetValue('GEL_QTEALIVRER'), 6);
      Montant := Arrondi(TotQteAffectee * PxUnit, 6);
      TOBA.PutValue('GEL_MTHTAFFECTEE', Montant);
      TV1.SetValueCell(TV1.ColIndex('GEL_MTHTAFFECTEE'), TV1.CurrentROW, Montant);
      // Mise à jour Piece
      MajPiece(TOBA, OldQte, OldMontant);
    end;
    TOBA := TobAffect.FindNext(['GEL_NATUREPIECEG', 'GEL_SOUCHE', 'GEL_NUMERO', 'GEL_INDICEG', 'GEL_NUMLIGNEGEN', 'GEL_TYPEDIM'], [NaturePiece, Souche, Numero,
      Indice, NumLigne, 'GEN'], False);
  end;
  // Modif tables
  UpdateTable;
  TOBAffBis.Free;
end;
//************************** Fin Cycle desaffectation ************************************/
//************************************************************************************//


procedure TOF_AFFCDEMODIF.OnChangeCodeAff;
var CodeAff, Stg: string;
  QQ: TQuery;
begin
  CodeAff := GetControlText('GEL_CODEAFF');
  if CodeAff <> '' then
  begin
    Stg := 'SELECT GEA_DEPOT FROM AFFCDEENTETE WHERE GEA_CODEAFF="' + CodeAff + '"';
    QQ := OpenSQl(Stg, True);
    if QQ.Eof then
      Stg := ''
    else
      Stg := QQ.FindField('GEA_DEPOT').asString;
    Ferme(QQ);
    SetControlText('GEA_DEPOT', Stg);
  end;
end;

procedure AGLOnChangeCodeAff(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFStat) then TOTOF := TFStat(F).LaTOF else exit;
  if (TOTOF is TOF_AFFCDEMODIF) then TOF_AFFCDEMODIF(TOTOF).OnChangeCodeAff;
end;

procedure AGLOnAffecterSelection(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFStat) then TOTOF := TFStat(F).LaTOF else exit;
  if (TOTOF is TOF_AFFCDEMODIF) then TOF_AFFCDEMODIF(TOTOF).OnAffecterSelection;
end;

procedure AGLOnAffecterCommande(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFStat) then TOTOF := TFStat(F).LaTOF else exit;
  if (TOTOF is TOF_AFFCDEMODIF) then TOF_AFFCDEMODIF(TOTOF).OnAffecterCommande;
end;

procedure AGLOnAffecterLigne(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFStat) then TOTOF := TFStat(F).LaTOF else exit;
  if (TOTOF is TOF_AFFCDEMODIF) then TOF_AFFCDEMODIF(TOTOF).OnAffecterLigne;
end;

procedure AGLOnDesAffecterSelection(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFStat) then TOTOF := TFStat(F).LaTOF else exit;
  if (TOTOF is TOF_AFFCDEMODIF) then TOF_AFFCDEMODIF(TOTOF).OnDesAffecterSelection;
end;

procedure AGLOnDesAffecterCommande(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFStat) then TOTOF := TFStat(F).LaTOF else exit;
  if (TOTOF is TOF_AFFCDEMODIF) then TOF_AFFCDEMODIF(TOTOF).OnDesAffecterCommande;
end;

procedure AGLOnDesAffecterLigne(parms: array of variant; nb: integer);
var F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFStat) then TOTOF := TFStat(F).LaTOF else exit;
  if (TOTOF is TOF_AFFCDEMODIF) then TOF_AFFCDEMODIF(TOTOF).OnDesAffecterLigne;
end;

initialization
  registerclasses([TOF_AFFCDEMODIF]);
  RegisterAglProc('OnChangeCodeAff', TRUE, 0, AGLOnChangeCodeAff);
  RegisterAglProc('OnAffecterSelection', TRUE, 0, AGLOnAffecterSelection);
  RegisterAglProc('OnAffecterCommande', TRUE, 0, AGLOnAffecterCommande);
  RegisterAglProc('OnAffecterLigne', TRUE, 0, AGLOnAffecterLigne);
  RegisterAglProc('OnDesAffecterSelection', TRUE, 0, AGLOnDesAffecterSelection);
  RegisterAglProc('OnDesAffecterCommande', TRUE, 0, AGLOnDesAffecterCommande);
  RegisterAglProc('OnDesAffecterLigne', TRUE, 0, AGLOnDesAffecterLigne);
end.
