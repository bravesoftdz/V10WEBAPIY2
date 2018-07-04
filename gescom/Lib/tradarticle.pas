unit tradarticle;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, ExtCtrls, HPanel, Grids, Hctrls, StdCtrls, HEnt1, UIUtil, UTOB,
  HSysMenu, SaisUtil,UtilPGI,M3FP,AGLInit,
{$IFDEF EAGLCLIENT}
  MaineAgl,
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_main ,
{$ENDIF}
  Math,AGLInitGC , hmsgbox,Mask
  ,HDimension,UtilArticle, EntGC ,TarifUtil ;

procedure EntreeTradArticle (Action : TActionFiche) ;
procedure SaisieTradArticle (CodeLan : string; Action : TActionFiche) ;

Const NbRowsInit = 50 ;
      NbRowsPlus = 20 ;

Const SG_Art   : integer = -1;
      SG_Lib   : integer = -1;
      SG_Libt  : integer = -1;
      SG_LibC  : integer = -1;
      SG_LibCt : integer = -1;
      SG_Comm  : integer = -1;
      SG_Commt : integer = -1;

type
  TFTradArticle = class(TForm)
    PLANGUE: THPanel;
    Dock971: TDock97;
    Toolbar972: TToolWindow97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    BChercher: TToolbarButton97;
    BInfos: TToolbarButton97;
    PPIED: THPanel;
    PGRID: THPanel;
    GTrad: THGrid;
    GTA_LANGUE: THValComboBox;
    TGTA_LANGUE: THLabel;
    HMTRAD: THSystemMenu;
    MsgBox: THMsgBox;
    FindLigne: TFindDialog;
    HMess: THMsgBox;
    TGA_CODEDIM1: THCritMaskEdit;
    TGA_CODEDIM2: THCritMaskEdit;
    TGA_CODEDIM3: THCritMaskEdit;
    TGA_CODEDIM4: THCritMaskEdit;
    TGA_CODEDIM5: THCritMaskEdit;
    TGA_GRILLEDIM1: THLabel;
    TGA_GRILLEDIM2: THLabel;
    TGA_GRILLEDIM3: THLabel;
    TGA_GRILLEDIM4: THLabel;
    TGA_GRILLEDIM5: THLabel;
    BInsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
                          Shift: TShiftState);
    procedure GTA_LANGUEChange(Sender: TObject);
    procedure GTradEnter(Sender: TObject);
    procedure GTradRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
    procedure GTradRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
    procedure GTradCellEnter(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure GTradCellExit(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure GTradElipsisClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    procedure BInfosClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
  private
    { Déclarations privées }
    TOBTrad, TOBDel, TOBArt : TOB;
    LesColTrad : String ;
    FindDebut,FClosing : Boolean;
    procedure Trad_LoadLesTOB ;
    procedure Trad_ChargeTraduct ;
// Ligne
    Function  Trad_GetTOBLigne ( ARow : integer) : TOB ;
// Actions liées au Grid
    procedure Trad_EtudieColsListe ;
    procedure Trad_InitialiseLigne (ARow : integer) ;
    procedure Trad_AfficheLigne (ARow: integer) ;
    procedure Trad_InitialiseGrille ;
    Procedure Trad_DepileTOBLigne ;
    Procedure Trad_CreerTOBLigne (ARow : integer);
    Function  Trad_LigneVide (ARow : integer) : Boolean;
    procedure Trad_SupprimeTOBTrad (ARow : integer; tout : boolean) ;
    Function  Trad_GrilleModifie : Boolean;
    procedure Trad_InsertLigne (ARow : integer) ;
    procedure Trad_SupprimeLigne (ARow : integer) ;
    Function  Trad_SortDeLaLigne : boolean;
//Manipulation champs
    procedure Trad_TraiterArticle (ACol, ARow : integer; Bouton : boolean);
    procedure Trad_TraiterLibTrad (ACol, ARow : integer);
    procedure Trad_ChercherArticle (ACol, ARow : integer) ;
// Actions liées à l'entête
    procedure Trad_InitialiseEntete ;
    Procedure Trad_TraiterCodeLangue;
// validation
    Procedure Trad_ValideTrad;
    Procedure Trad_VerifLesTOB;
    Function  Trad_QuestionValider : Integer;
// pied
    procedure Trad_AffecteDimensions(TOBL : TOB; ArticleLu : boolean);
    procedure Trad_AffichePied (ARow : integer);
    procedure Trad_RazPied;
  public
    { Déclarations publiques }
    Action      : TActionFiche ;
    CodeLangue  : String ;
  end;

var
  FTradArticle: TFTradArticle;

implementation

{$R *.DFM}

//=================================================================
//                fonctions "point d'entrée"
//=================================================================
procedure EntreeTradArticle (Action : TActionFiche) ;
BEGIN
SaisieTradArticle ('', Action);
END;

// NB si, en l'état, on veut appeler cette fonction avec un code
// langue autre que blanc on a très peu de chances que ça marche !
// cette possibilité n'a été gérée qu'ici

procedure SaisieTradArticle (CodeLan : string; Action : TActionFiche) ;
var FF : TFTradArticle ;
    PPANEL  : THPanel ;
BEGIN
SourisSablier;
FF := TFTradArticle.Create(Application) ;
FF.Action:=Action ;
FF.CodeLangue:= CodeLan ;
PPANEL := FindInsidePanel ;
if PPANEL = Nil then
   BEGIN
    try
      FF.BorderStyle:=bsSingle ;
      FF.ShowModal ;
    finally
      FF.Free ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside (FF, PPANEL) ;
   FF.Show ;
   END ;
END;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}

// chargement des lignes

procedure TFTradArticle.Trad_LoadLesTOB ;
Var i_ind : integer ;
    QQ : TQuery ;
    TOBTmp, TOBL, TOBTL : TOB ;
BEGIN
for i_ind:=0 to TOBTrad.Detail.count - 1 do
    BEGIN
    TOBTrad.Detail[i_ind].Free ;
    END;

// une TOB sans table associée ne peut pas faire de mise à jour de la base
// Une TOB est monotable
// ici, on crée une TOB temporaire sans table associée pour obtenir tous les champs
// de la requête

TOBTmp:= TOB.Create ('', nil, -1) ; if TOBTmp = nil then exit ;
QQ:=OpenSql ('SELECT GTA_ARTICLE, GTA_LANGUE, GTA_LIBELLE, GTA_LIBCOMPL, GTA_COMMENTAIRE, '+
             'GA_CODEARTICLE, GA_LIBELLE, GA_LIBCOMPL, GA_COMMENTAIRE '+
              ' FROM GTRADARTICLE LEFT JOIN ARTICLE ON GTA_ARTICLE=GA_ARTICLE WHERE GTA_LANGUE="' +
              CodeLangue + '" ORDER BY GTA_ARTICLE DESC', true,-1,'',true);
if not QQ.EOF then TOBTmp.LoadDetailDB ('', '', '', QQ, False) ;
Ferme (QQ);

// puis on la déverse dans une TOB liée à la table GTA qu'on veut mettre à jour

for i_ind:=TOBTmp.Detail.count - 1 Downto 0  do
    BEGIN
    TOBTL:=TOBTmp.Detail[i_ind];

    // -1 dans l'indice de la TOB crée une TOB fille "à la suite".
    // remarque : la requète est "DESC" pour que la tob résultante
    // soit "ASC"

    TOBL := TOB.Create ('GTRADARTICLE', TOBTrad, -1);
    TOBL.PutValue ('GTA_ARTICLE', TOBTL.GetValue ('GTA_ARTICLE'));
    TOBL.PutValue ('GTA_LANGUE', TOBTL.GetValue ('GTA_LANGUE'));
    TOBL.PutValue ('GTA_LIBELLE', TOBTL.GetValue ('GTA_LIBELLE'));
    TOBL.PutValue ('GTA_LIBCOMPL', TOBTL.GetValue ('GTA_LIBCOMPL'));
    TOBL.PutValue ('GTA_COMMENTAIRE', TOBTL.GetValue ('GTA_COMMENTAIRE'));
    TOBL.AddChampSup ('GA_CODEARTICLE', False);
    TOBL.AddChampSup ('GA_LIBELLE', False);
    TOBL.AddChampSup ('GA_LIBCOMPL', False);
    TOBL.AddChampSup ('GA_COMMENTAIRE', False);
    TOBL.AddChampSup ('OLD', False);
    TOBL.AddChampSup ('TGA_GRILLEDIM1', False);
    TOBL.AddChampSup ('TGA_GRILLEDIM2', False);
    TOBL.AddChampSup ('TGA_GRILLEDIM3', False);
    TOBL.AddChampSup ('TGA_GRILLEDIM4', False);
    TOBL.AddChampSup ('TGA_GRILLEDIM5', False);
    TOBL.AddChampSup ('TGA_CODEDIM1', False);
    TOBL.AddChampSup ('TGA_CODEDIM2', False);
    TOBL.AddChampSup ('TGA_CODEDIM3', False);
    TOBL.AddChampSup ('TGA_CODEDIM4', False);
    TOBL.AddChampSup ('TGA_CODEDIM5', False);
    TOBL.PutValue ('GA_CODEARTICLE', TOBTL.GetValue ('GA_CODEARTICLE'));
    TOBL.PutValue ('GA_LIBELLE', TOBTL.GetValue ('GA_LIBELLE'));
    TOBL.PutValue ('GA_LIBCOMPL', TOBTL.GetValue ('GA_LIBCOMPL'));
    TOBL.PutValue ('GA_COMMENTAIRE', TOBTL.GetValue ('GA_COMMENTAIRE'));
    TOBL.PutValue ('OLD', 'X');
    Trad_AffecteDimensions(TOBL,False);
    TOBTL.Free ;
    END;
TOBTmp.Free ;
// on va dire que rien n'a été modifié au chargement...
TOBTrad.SetAllModifie(False);
END;

// affichage des lignes

procedure TFTradArticle.Trad_ChargeTraduct ;
var i_ind : integer;
BEGIN
Trad_LoadLesTOB ;
for i_ind:=0 to TOBTrad.Detail.count - 1 do
    BEGIN
    Trad_AfficheLigne (i_ind + 1);
    END;
END;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}

// allocations mémoire

procedure TFTradArticle.FormCreate(Sender: TObject);
begin
GTRAD.RowCount:=NbRowsInit ;
TOBTRAD:=TOB.Create ('', nil, -1) ;
TOBDel:=TOB.Create ('', nil, -1) ;
TOBArt := TOB.Create ('ARTICLE', Nil, -1) ;
FClosing:=False ;
end;

// en entrant dans la fiche...

procedure TFTradArticle.FormShow(Sender: TObject);
begin
  // Appel de la fonction d'empilage dans la liste des fiches
  AglEmpileFiche(Self) ;
GTRAD.ListeParam:='GCTRADARTICLES';
Trad_EtudieColsListe ;
HMTrad.ResizeGridColumns (GTRAD) ;
AffecteGrid (GTRAD,Action) ;
Trad_InitialiseEntete ;
end;

// quand l'utilisateur demande à fermer la fenêtre
// s'il a modifié qq chose on lui demande s'il veut vraiment quitter
// ATTENTION : il faut créer les évènements dans la phase de design de
// la form pour pouvoir les utiliser

procedure TFTradArticle.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
BEGIN
if Trad_GrilleModifie then
   BEGIN
   if MsgBox.Execute(4,Caption,'')<>mrYes then CanClose:=False ;
   END ;
end;

// désallocations mémoire

procedure TFTradArticle.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
GTRAD.VidePile(True) ;
TOBTRAD.Free ; TOBTRAD:=Nil ;
TOBDel.Free ; TOBDel:=Nil ;
TOBArt.Free ; TOBArt:=Nil ;
if IsInside(Self) then Action:=caFree ;
FClosing:=True ;
end;

// c'est la form qui renvoie les touches frappées au clavier
// l'évènement ici est l'appui sur une touche

procedure TFTradArticle.FormKeyDown(Sender: TObject; var Key: Word;
                                     Shift: TShiftState);
BEGIN
Case Key of
    // si return on agit comme pour TAB
    VK_RETURN : Key:=VK_TAB ;
    // on gere ctrl/inser et ctrl/suppr dans la grille uniquement
    VK_INSERT : BEGIN
                if (Screen.ActiveControl = GTrad) then
                    BEGIN
                    Key := 0;
                    Trad_InsertLigne (GTrad.Row);
                    END;
                END;
    VK_DELETE : BEGIN
                if ((Screen.ActiveControl = GTRAD) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    Trad_SupprimeLigne (GTrad.row) ;
                    END ;
                END;
    VK_F10    : BEGIN Key:=0 ; BValiderClick(Nil) ; END ;
    VK_F5     : BEGIN
                if (Screen.ActiveControl=GTRAD) then
                   BEGIN
                   Key:=0 ;
                   if GTRAD.ElipsisButton then Trad_ChercherArticle(GTrad.Col,GTrad.Row) ;
                   END ;
                END;
    END;
END;


{==============================================================================================}
{================================= Manipulation Lignes ========================================}
{==============================================================================================}

// récupération d'une TOB en fonction de l'indice de la ligne
// NB l'indice d'une TOB commence à 0 les lignes de la grid sont numérotées à
// partir de 1 (à cause des titres ?)

Function TFTradArticle.Trad_GetTOBLigne ( ARow : integer) : TOB ;
BEGIN
Result:=Nil ;
if ((ARow<=0) or (ARow>TOBTRAD.Detail.Count)) then Exit ;
Result:=TOBTRAD.Detail[ARow-1] ;
END ;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}

// on récupère les numéros de colonne de chaque zone en fonction
// des titres de la grille

procedure TFTradArticle.Trad_EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol : integer ;
BEGIN
GTRAD.ColWidths[0]:=0;
LesCols:=GTRAD.Titres[0] ; LesColTrad:=LesCols ; icol:=0 ;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        BEGIN
        if NomCol='GA_CODEARTICLE'  then SG_Art:=icol else
        if NomCol='GA_LIBELLE'      then BEGIN SG_Lib:=icol; GTRAD.ColLengths[icol]:=-1; END else
        if NomCol='GA_LIBCOMPL'     then BEGIN SG_LibC:=icol; GTRAD.ColLengths[icol]:=-1; END else
        if NomCol='GA_COMMENTAIRE'  then BEGIN SG_Comm:=icol; GTRAD.ColLengths[icol]:=-1; END else
        if NomCol='GTA_LIBELLE'     then SG_Libt:=icol else
        if NomCol='GTA_LIBCOMPL'    then SG_LibCt:=icol else
        if NomCol='GTA_COMMENTAIRE' then SG_Commt:=icol else
         ;
        END ;
    Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;
END;

// création d'une ligne à blanc

procedure TFTradArticle.Trad_InitialiseLigne (ARow : integer) ;
Var TOBL : TOB ;
    i_ind : integer ;
BEGIN
TOBL:=Trad_GetTOBLigne(ARow) ; if TOBL<>Nil then TOBL.InitValeurs ;
for i_ind := 1 to GTRAD.ColCount-1 do GTRAD.Cells [i_ind, ARow]:='' ;
END;

// affichage d'une ligne

procedure TFTradArticle.Trad_AfficheLigne (ARow: integer) ;
Var TOBL : TOB;
BEGIN
TOBL:=Trad_GetTOBLigne (ARow) ;  if TOBL = Nil then exit;
TOBL.PutLigneGrid (Gtrad, ARow, False, False, LesColTrad) ;
END;

// vidage de la grille et création de 50 lignes écran à blanc

procedure TFTradArticle.Trad_InitialiseGrille ;
BEGIN
Gtrad.VidePile(True) ;
Gtrad.RowCount:= NbRowsInit ;
END;

// désallocation des TOB "filles"
// Ne pas oublier la tobDel... sinon on a de drôles de résultats !

Procedure TFTradArticle.Trad_DepileTOBLigne ;
var i_ind : integer;
BEGIN
for i_ind := TOBTRAD.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBTRAD.Detail[i_ind].Free ;
    END;
for i_ind := TOBDel.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBDel.Detail[i_ind].Free ;
    END;
END;

// ajout d'une TOB vide en fin d'écran

Procedure TFTradArticle.Trad_CreerTOBLigne (ARow : integer);
var TOBL : TOB;
BEGIN
if ARow <> TOBTRAD.Detail.Count + 1 then exit;
TOBL:=TOB.Create ('GTRADARTICLE', TOBTrad, ARow-1) ;
TOBL.AddChampSup ('GA_CODEARTICLE', False);
TOBL.AddChampSup ('GA_LIBELLE', False);
TOBL.AddChampSup ('GA_LIBCOMPL', False);
TOBL.AddChampSup ('GA_COMMENTAIRE', False);
TOBL.AddChampSup ('OLD', False);
TOBL.AddChampSup ('TGA_GRILLEDIM1', False);
TOBL.AddChampSup ('TGA_GRILLEDIM2', False);
TOBL.AddChampSup ('TGA_GRILLEDIM3', False);
TOBL.AddChampSup ('TGA_GRILLEDIM4', False);
TOBL.AddChampSup ('TGA_GRILLEDIM5', False);
TOBL.AddChampSup ('TGA_CODEDIM1', False);
TOBL.AddChampSup ('TGA_CODEDIM2', False);
TOBL.AddChampSup ('TGA_CODEDIM3', False);
TOBL.AddChampSup ('TGA_CODEDIM4', False);
TOBL.AddChampSup ('TGA_CODEDIM5', False);
Trad_InitialiseLigne (ARow) ;
END;

// la ligne est-elle valide ?

Function TFTradArticle.Trad_LigneVide (ARow : integer) : Boolean;
BEGIN
Result := True;
if (Trim (Gtrad.Cells [SG_Libt, ARow]) <> '') and (Trim (Gtrad.Cells [SG_Art, ARow]) <> '') then
    BEGIN
    Result := False;
    Exit;
    END;
END;

// insertion d'une ligne
// TOB.create n'écrase pas la TOB de l'indice indiqué
// si elle existe, il y a insertion (décalage des indices)

procedure TFTradArticle.Trad_InsertLigne (ARow : integer) ;
var TOBL : TOB;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if Trad_LigneVide (ARow) then exit;
if (ARow > TOBTrad.Detail.Count) then Exit;

// pour que l'écran se réaffiche plus "proprement"
GTrad.CacheEdit; GTrad.SynEnabled := False;

TOBL:=TOB.Create ('GTRADARTICLE', TOBTrad, ARow-1) ;
TOBL.AddChampSup ('GA_CODEARTICLE', False);
TOBL.AddChampSup ('GA_LIBELLE', False);
TOBL.AddChampSup ('GA_LIBCOMPL', False);
TOBL.AddChampSup ('GA_COMMENTAIRE', False);
TOBL.AddChampSup ('OLD', False);
TOBL.AddChampSup ('TGA_GRILLEDIM1', False);
TOBL.AddChampSup ('TGA_GRILLEDIM2', False);
TOBL.AddChampSup ('TGA_GRILLEDIM3', False);
TOBL.AddChampSup ('TGA_GRILLEDIM4', False);
TOBL.AddChampSup ('TGA_GRILLEDIM5', False);
TOBL.AddChampSup ('TGA_CODEDIM1', False);
TOBL.AddChampSup ('TGA_CODEDIM2', False);
TOBL.AddChampSup ('TGA_CODEDIM3', False);
TOBL.AddChampSup ('TGA_CODEDIM4', False);
TOBL.AddChampSup ('TGA_CODEDIM5', False);
GTrad.InsertRow (ARow); GTrad.Row := ARow;
Trad_InitialiseLigne (ARow) ;

// fin du "proprement"...
GTrad.MontreEdit; GTrad.SynEnabled := True;

END;

// suppression d'une ligne
// NB tout aussi "proprement" que l'insertion !

procedure TFTradArticle.Trad_SupprimeLigne (ARow : integer) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if (ARow > TOBTrad.Detail.Count) then Exit;
GTrad.CacheEdit; GTrad.SynEnabled := False;
GTrad.DeleteRow (ARow);
if (ARow = TOBTrad.Detail.Count) then Trad_CreerTOBLigne (ARow + 1);
Trad_SupprimeTOBTrad (ARow, True);
if GTrad.RowCount < NbRowsInit then GTrad.RowCount := NbRowsInit;
GTrad.MontreEdit; GTrad.SynEnabled := True;
END;

// suppression de la ligne dans la TOB
// cette fonction est appelée si on efface la ligne et si on change le code article
// de la ligne (==> on supprime la TOB ou non)
// dans les 2 cas, si c'est une ligne issue de la requête SQL on la charge dans
// la TOB des lignes à annuler

procedure TFTradArticle.Trad_SupprimeTOBTrad (ARow : integer; tout : boolean) ;
Var i_ind: integer;
BEGIN
if TOBTrad.Detail[ARow-1].GetValue ('OLD') = 'X' then
    BEGIN
    i_ind := TOBDel.Detail.Count;
    TOB.Create('GTRADARTICLE', TOBDel, i_ind) ;
    TOBDel.Detail[i_ind].Dupliquer (TOBTrad.Detail[ARow-1], False, True);

    // même en création la mise à jour ne se fait que sur les champs modifiés
    // en cas de changement de clé (article) le champ langue qui n'avait pas
    // changé n'était pas rechargé.. on dit donc que TOUT a été modifié...

    TOBTrad.Detail[ARow-1].SetAllModifie(True);

    END;
if (tout) then
    BEGIN
    TOBTrad.Detail[ARow-1].Free;
    END
END;

// est-ce que la TOB a été modifiée ?

Function TFTradArticle.Trad_GrilleModifie : Boolean;
BEGIN
Result:=False ;
if Action=taConsult then Exit ;
Result:=(TOBTrad.IsOneModifie);
END;

// Comme l'utilisateur peut quitter n'importe quand on vérifie
// la validité de la dernière ligne sur laquelle il était

Function TFTradArticle.Trad_SortDeLaLigne : Boolean;
var ACol, ARow : integer;
    Cancel : boolean;
BEGIN
Result := False;
ACol := GTrad.Col;
ARow  := GTrad.Row;
Cancel := False;
GTradCellExit(Nil,ACol,ARow,Cancel);
if Cancel then exit;
GTradRowExit(Nil,ARow,Cancel,False);
if Cancel then exit;
Result := True
END;

{==============================================================================================}
{=============================== Evènements de la Grid ========================================}
{==============================================================================================}

// à l'entrée dans la grille, il faut forcer l'entrée dans la ligne puis
// dans la cellule

procedure TFTradArticle.GTradEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
begin
Cancel := False; Chg := False;
GTradRowEnter (Sender, GTrad.Row, Cancel, Chg);
ACol := GTrad.Col ; ARow := GTrad.Row ;
GTradCellEnter (Sender, ACol, ARow, Cancel);
end;

// Entrée dans la ligne :
// si on est au bout des lignes de la grille, on en crée un nouveau "paquet"
// par contre on ne crée de ligne dans la TOB que si la précédente n'est pas
// vide

procedure TFTradArticle.GTradRowEnter(Sender: TObject; Ou: Integer;
                                      var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
begin
if Ou >= GTrad.RowCount - 1 then GTrad.RowCount := GTrad.RowCount + NbRowsPlus ;;
ARow := Min (Ou, TOBTRAD.detail.count + 1);
if (ARow = TOBTRAD.detail.count + 1) AND (not Trad_LigneVide (ARow - 1)) then
    BEGIN
    Trad_CreerTOBligne (ARow);
    END;
if Ou > TOBTRAD.detail.count then
    BEGIN
    GTrad.Row := TOBTRAD.detail.count;
    END;
Trad_AffichePied(GTrad.row);
end;

// sortie de la ligne : si vide on n'en bouge que vers le haut

procedure TFTradArticle.GTradRowExit(Sender: TObject; Ou: Integer;
                                     var Cancel: Boolean; Chg: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if Trad_LigneVide (Ou) Then GTrad.Row := Min (GTrad.Row,Ou);
end;

// entrée dans la cellule : on refuse de bouger de la cellule article
// tant qu'elle est vide
// on active l'"elipsis Button" seulement sur la cellule article

procedure TFTradArticle.GTradCellEnter(Sender: TObject; var ACol,
                                        ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (GTrad.Col <> SG_Art)  AND (GTrad.Cells [SG_Art,GTrad.Row] = '') then GTrad.Col := SG_Art;
    GTrad.ElipsisButton := (GTrad.Col = SG_Art) ;
    END ;
end;

// en sortie de la cellule : contrôles de saisie

procedure TFTradArticle.GTradCellExit(Sender: TObject; var ACol,
                                     ARow: Integer; var Cancel: Boolean);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if ACol = SG_Art then Trad_TraiterArticle (ACol, ARow,True) else
    if ((ACol=SG_Libt) or (ACol=SG_LibCt) or (ACol=SG_Commt)) then Trad_TraiterLibTrad(ACol,ARow) ;
end;

// si click sur l'"elipsis Button" recherche article

procedure TFTradArticle.GTradElipsisClick(Sender: TObject);
begin
Trad_ChercherArticle(GTrad.Col,GTrad.Row) ;
end;

{==============================================================================================}
{============================== Manipulation des Champs  ======================================}
{==============================================================================================}

// contrôle article

procedure TFTradArticle.Trad_TraiterArticle (ACol, ARow : integer; Bouton : boolean);
var TOBL : TOB;
    RECART : THCritMaskEdit;
    Coord : TRect;
    RechArt : T_RechArt;
    OkArt : Boolean;
    ind : integer;
BEGIN
TOBL := Trad_GetTOBLigne (ARow); if TOBL=nil then exit;
// on sort si code article identique au précédent (on ne refait pas les contrôles)
// si on a blanc, on restitue l'ancienne valeur et on sort
// de toute manière, blanc ne passera pas à cause du Oncellenter

if Gtrad.Cells[ACol,ARow] = trim(TOBL.GetValue('GA_CODEARTICLE')) then exit;
if GTrad.Cells[ACol,ARow] = '' then
    BEGIN
    GTrad.Cells[ACol,ARow] := trim(TOBL.GetValue('GA_CODEARTICLE'));
    exit;
    END;

// donc le cas contrôlé est forcément une modification
// avec saisie de "quelque chose"

// pour ne pas avoir de bouton sur le libelle dans le cas ou on vient
// du OnCellExit et ou on relancerait une recherche
if bouton then GTrad.ElipsisButton := false;

Coord := GTrad.CellRect (ACol, ARow);
RECART := THCritMaskEdit.Create (Self);
RECART.Parent := GTrad;
RECART.Top := Coord.Top;
RECART.Left := Coord.Left;
RECART.Width := 18;
RECART.Visible := False;
RECART.DataType := 'GCARTICLE';
RECART.Text := uppercase(trim(GTrad.Cells[ACol,ARow]));
OkArt := False;
RechArt := TrouverArticle(RECART,TOBArt);
case RechArt of
        traOk : OkArt := True;
     traAucun : BEGIN
                 // Recherche sur code via LookUp ou Recherche avancée
                Trad_ChercherArticle (ACol,ARow);
                exit;
                END;
    traGrille : BEGIN
                if ChoisirDimension(TOBArt.Getvalue('GA_ARTICLE'),TOBArt) then
                    BEGIN
                    OkArt := True;
                    END;
                END;
END; // case....

// vérifier qu'il n'y a pas de doublon
if OkArt = True then
    BEGIN
    for ind :=1 to TOBTrad.detail.count do
        BEGIN
        TOBL := Trad_GetTOBLigne(ind);
        if TOBL <> nil then
            BEGIN
            if TOBArt.GetValue('GA_ARTICLE') = TOBL.GetValue('GTA_ARTICLE') then
                BEGIN
                if (ind <> ARow) then OkArt := False;
                END;
            END;
        END;
    TOBL := Trad_GetTOBLigne (ARow); if TOBL=nil then exit;
    if OkArt = False then MsgBox.execute(3,caption,'');
    END;

// si on modifie le code article sur une ancienne ligne, il faut la stocker
// dans la TOB d'annulation

if (OkArt) then
    BEGIN
    Trad_SupprimeTOBTrad(Arow, False);
    TOBL.PutValue('GTA_LANGUE',CodeLangue);
    TOBL.PutValue('GTA_ARTICLE',TOBArt.GetValue('GA_ARTICLE'));
    TOBL.PutValue('GA_CODEARTICLE',TOBArt.GetValue('GA_CODEARTICLE'));
    TOBL.PutValue('GA_LIBELLE',TOBArt.GetValue('GA_LIBELLE'));
    TOBL.PutValue('GA_LIBCOMPL',TOBArt.GetValue('GA_LIBCOMPL'));
    TOBL.PutValue('GA_COMMENTAIRE',TOBArt.GetValue('GA_COMMENTAIRE'));
    TOBL.PutValue ('OLD',''); // effacer le X
    Trad_AffecteDimensions(TOBL,True);
    END;

// si ce n'était pas OK on récupère les anciennes valeurs

GTrad.Cells[ACol,ARow]:=Trim(TOBL.GetValue('GA_CODEARTICLE')) ;
GTrad.Cells[SG_Lib,ARow]:=TOBL.GetValue('GA_LIBELLE') ;
GTrad.Cells[SG_LibC,ARow]:=TOBL.GetValue('GA_LIBCOMPL') ;
GTrad.Cells[SG_Comm,ARow]:=TOBL.GetValue('GA_COMMENTAIRE') ;
RECART.Destroy ;
Trad_AffichePied(ARow) ;
END;

// on stocke la traduction dans la TOB

procedure TFTradArticle.Trad_TraiterLibTrad (ACol, ARow : integer);
var TOBL : TOB;
BEGIN
TOBL := Trad_GetTOBLigne (ARow); if TOBL=nil then exit;
if Acol=SG_Libt then TOBL.PutValue('GTA_LIBELLE',GTrad.cells[ACol,ARow]) else
 if Acol=SG_LibCt then TOBL.PutValue('GTA_LIBCOMPL',GTrad.cells[ACol,ARow]) else
  if Acol=SG_Commt then TOBL.PutValue('GTA_COMMENTAIRE',GTrad.cells[ACol,ARow]) ;
END;

// recherche article

procedure TFTradArticle.Trad_ChercherArticle (ACol, ARow : integer) ;
var RECART : THCritMaskEdit;
    Coord : TRect;
BEGIN
Coord := GTrad.CellRect (ACol, ARow);
RECART := THCritMaskEdit.Create (Self);
RECART.Parent := GTrad;
RECART.Top := Coord.Top;
RECART.Left := Coord.Left;
RECART.Width := 18;
RECART.Visible := False;
RECART.DataType := 'GCARTICLEGENERIQUE';
RECART.Text := GTrad.Cells[ACol,ARow];
DispatchRechArt(RECART, 1, '','GA_CODEARTICLE=' +
    trim(Copy (RECART.Text,1,18)),'');
if RECART.Text <> '' then
   BEGIN
   RECART.Text := format('%-33.33sX',[RECART.Text]);
   if TOBArt.SelectDB('"' + RECART.Text + '"',Nil) then
       BEGIN
       GTrad.Cells[Acol, Arow] := TOBArt.GetValue('GA_CODEARTICLE');
       END else BEGIN
       GTrad.Cells[Acol,Arow] := '';
       END;
   END;
RECART.destroy;
Trad_TraiterArticle (Acol,Arow, False);
END;

{==============================================================================================}
{============================= Actions liées à l'entête =======================================}
{==============================================================================================}
procedure TFTradArticle.Trad_InitialiseEntete ;
BEGIN
Codelangue := '';
GTA_LANGUE.Text := '';
GTA_LANGUE.Value := '';
GTRAD.Enabled:=False ;
Trad_RazPied;
END ;

Procedure TFTradArticle.Trad_TraiterCodeLangue;
BEGIN
if GTA_LANGUE.Value = CodeLangue then exit;
CodeLangue := GTA_LANGUE.Value;
GTRAD.Enabled:=True ;
Trad_InitialiseGrille;
Trad_DepileTOBLigne;
Trad_ChargeTraduct ;
Trad_RazPied;
END;

{==============================================================================================}
{================================= Evenement Entête ===========================================}
{==============================================================================================}
procedure TFTradArticle.GTA_LANGUEChange(Sender: TObject);
var rep : integer;
    ioerr : TIOErr;
begin
if GTA_LANGUE.Value = Codelangue then exit;
rep := Trad_QuestionValider;
case rep of
    mrYes : BEGIN
            ioerr := Transactions(Trad_ValideTrad, 2);
            case ioerr of
              oeOk      : ;
              oeUnknown : begin MessageAlerte(HMess.Mess[1]); end;
              oeSaisie  : begin MessageAlerte(HMess.Mess[2]); end;
            end;
            END;
    mrNo    : ;
    mrCancel: GTA_LANGUE.Value := Codelangue;
end;
if GTA_LANGUE.Value <> '' then Trad_TraiterCodeLangue
                          else Trad_InitialiseEntete;
end;

{==============================================================================================}
{============================= Evenement liés aux boutons =====================================}
{==============================================================================================}

// pour lancer une recherche dans une grille, il faut y insérer un objet
// Tfinddialog

procedure TFTradArticle.BChercherClick(Sender: TObject);
begin
if GTrad.RowCount < 3 then Exit;
FindDebut:=True ;
FindLigne.Execute ;
end;

procedure TFTradArticle.FindLigneFind(Sender: TObject);
begin
GTrad.ElipsisButton := False;
Rechercher(GTrad, FindLigne, FindDebut);
end;

// appel de la fiche article

procedure TFTradArticle.BInfosClick(Sender: TObject);
var TOBL :TOB;
    st : string;
begin
TOBL := Trad_GetTOBligne(GTrad.row);
if TOBL = Nil then Exit ;
st := TOBL.GetValue('GTA_ARTICLE');
if st <> '' then
BEGIN
{$IFNDEF GPAO}
  if CtxMode in V_PGI.PGIContexte  then
    AGLLanceFiche('MBO','ARTICLE','',st,'ACTION=CONSULTATION')
  else
    AglLanceFiche ('GC', 'GCARTICLE', '', st, 'ACTION=CONSULTATION;TARIF=N');
{$ELSE}
	V_PGI.DispatchTT(7, taConsult, st, 'TARIF=N', '');
{$ENDIF}
END;
end;

{==============================================================================}
{============================= Validation =====================================}
{==============================================================================}

procedure TFTradArticle.BValiderClick(Sender: TObject);
var ioerr : TIOErr;
begin
if Action = taConsult then exit;
ioerr := Transactions (Trad_valideTrad,2);
Case ioerr of
      oeOK : BEGIN Trad_InitialiseEntete; end;
      oeUnknown : BEGIN MessageAlerte(HMess.Mess[0]); END;
      oeSaisie : BEGIN MessageAlerte(HMess.Mess[1]); END;
END;
end;

// fonction appelée par transactions (qui assure le Rollback
// en cas de problème)

procedure TFTradArticle.Trad_ValideTrad;
BEGIN
if not Trad_SortDeLaLigne then exit;
TobDel.DeleteDB (False);
Trad_VerifLesTOB;
TOBTrad.InsertOrUpdateDB(False);
Trad_initialiseGrille;
Trad_DepileTOBLigne;
END;

// on enlève de la TOB toutes les lignes non valides

procedure TFTradArticle.Trad_VerifLesTOB;
var i_ind : integer;
BEGIN
for i_ind := TOBTrad.Detail.count - 1 Downto 0 do
    BEGIN
    if Trad_LigneVide (i_ind + 1) then TOBTrad.Detail[i_ind].Free ;
    END;
END;

// si changement de langue et saisie non sauvegardée
// il faut poser la question de la validation !

Function TFTradArticle.Trad_QuestionValider : Integer;
BEGIN
Result:=mrNo ;
if Action=taConsult then Exit ;
if Trad_GrilleModifie then Result:=MsgBox.Execute(0,Caption,'') ;
END;

{==============================================================================}
{=================== Pied (affichage des dimensions) ==========================}
{==============================================================================}
procedure TFTradArticle.Trad_AffecteDimensions(TOBL : TOB; ArticleLu : boolean);
Var ind :integer;
    GrilleDim,CodeDim,LibDim,LibCode : String ;
BEGIN
if Not ArticleLu then TOBArt.SelectDB ('"' + TOBL.GetValue('GTA_ARTICLE') + '"', nil);
for ind := 1 to MaxDimension do
    BEGIN
    if TOBArt.GetValue ('GA_GRILLEDIM' + IntToStr(ind)) <> '' then
       BEGIN
       GrilleDim:=TOBArt.GetValue('GA_GRILLEDIM'+IntToStr(ind)) ;
       CodeDim:=TOBArt.GetValue('GA_CODEDIM'+IntToStr(ind)) ;
       LibDim:=RechDom('GCGRILLEDIM'+IntToStr(ind),Grilledim,FALSE) ;
       LibCode:=GCGetCodeDim(GrilleDim,CodeDim,Ind) ;
       if LibCode<>'' then
          BEGIN
          TOBL.PutValue('TGA_GRILLEDIM'+IntToStr(ind),LibDim) ;
          TOBL.PutValue('TGA_CODEDIM'+IntToStr(ind),LibCode) ;
          END else
          BEGIN
          TOBL.PutValue('TGA_GRILLEDIM' + IntToStr(ind),'');
          TOBL.PutValue('TGA_CODEDIM' + IntToStr(ind), '');
          END ;
       END else
       BEGIN
       TOBL.PutValue('TGA_GRILLEDIM' + IntToStr(ind),'');
       TOBL.PutValue('TGA_CODEDIM' + IntToStr(ind), '');
       END;
    END;
END;

procedure TFTradArticle.Trad_AffichePied (ARow : integer);
var indT , indE: integer;
    TOBL : TOB;
BEGIN
Trad_Razpied;
indE := 1;
TOBL := Trad_GetTOBLigne (ARow);
if TOBL = nil then exit;
for indT := 1 to MaxDimension do
    BEGIN
    if Trim (TOBL.GetValue('TGA_GRILLEDIM' + IntToStr(indT))) <> '' then
        BEGIN
        if indE = 1 then
            begin
            TGA_GRILLEDIM1.Caption := TOBL.GetValue('TGA_GRILLEDIM' + IntToStr(indT));
            TGA_CODEDIM1.Text := TOBL.GetValue('TGA_CODEDIM' + IntToStr(indT));
            TGA_CODEDIM1.visible := True;
            end;
        if indE = 2 then
            begin
            TGA_GRILLEDIM2.Caption := TOBL.GetValue('TGA_GRILLEDIM' + IntToStr(indT));
            TGA_CODEDIM2.Text := TOBL.GetValue('TGA_CODEDIM' + IntToStr(indT));
            TGA_CODEDIM2.visible := True;
            end;
        if indE = 3 then
            begin
            TGA_GRILLEDIM3.Caption := TOBL.GetValue('TGA_GRILLEDIM' + IntToStr(indT));
            TGA_CODEDIM3.Text := TOBL.GetValue('TGA_CODEDIM' + IntToStr(indT));
            TGA_CODEDIM3.visible := True;
            end;
        if indE = 4 then
            begin
            TGA_GRILLEDIM4.Caption := TOBL.GetValue('TGA_GRILLEDIM' + IntToStr(indT));
            TGA_CODEDIM4.Text := TOBL.GetValue('TGA_CODEDIM' + IntToStr(indT));
            TGA_CODEDIM4.visible := True;
            end;
        if indE = 5 then
            begin
            TGA_GRILLEDIM5.Caption := TOBL.GetValue('TGA_GRILLEDIM' + IntToStr(indT));
            TGA_CODEDIM5.Text := TOBL.GetValue('TGA_CODEDIM' + IntToStr(indT));
            TGA_CODEDIM5.visible := True;
            end;
        Inc(indE);
        END;
    END;
END;

procedure TFTradArticle.Trad_RazPied;
BEGIN
TGA_GRILLEDIM1.Caption := '';
TGA_CODEDIM1.Visible := False;
TGA_GRILLEDIM2.Caption := '';
TGA_CODEDIM2.Visible := False;
TGA_GRILLEDIM3.Caption := '';
TGA_CODEDIM3.Visible := False;
TGA_GRILLEDIM4.Caption := '';
TGA_CODEDIM4.Visible := False;
TGA_GRILLEDIM5.Caption := '';
TGA_CODEDIM5.Visible := False;
END;

procedure TFTradArticle.BAbandonClick(Sender: TObject);
begin
Close ;
if FClosing and IsInside(Self) then THPanel(parent).CloseInside ;
end;

Procedure AGLEntreeTradArticle ( Parms : array of variant ; nb : integer ) ;
Var Action : TActionFiche ;
BEGIN
Action:=StringToAction(String(Parms[1])) ;
EntreeTradArticle(Action) ;
END ;


procedure TFTradArticle.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFTradArticle.BDeleteClick(Sender: TObject);
begin
Trad_SupprimeLigne (GTrad.row) ;
end;

procedure TFTradArticle.BInsertClick(Sender: TObject);
begin
Trad_InsertLigne (GTrad.Row);
end;

procedure TFTradArticle.FormDestroy(Sender: TObject);
begin
  // Appel de la fonction de dépilage dans la liste des fiches
  AglDepileFiche ;
end;

Initialization
RegisterAglProc('EntreeTradArticle',False,1,AGLEntreeTradArticle) ;

END.
