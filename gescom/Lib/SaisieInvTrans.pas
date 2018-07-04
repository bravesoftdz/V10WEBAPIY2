unit SaisieInvTrans;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HSysMenu,HTB97, HPanel, HEnt1, UIUtil, Hctrls,M3FP, ExtCtrls, Grids, StdCtrls,
{$IFDEF EAGLCLIENT}
  UtileAGL, MaineAGL,
{$ELSE}
  dbtables, EdtREtat, Fe_Main,
{$ENDIF}
  ComCtrls, Mask, SaisUtil, UTOB, DetailInv, HMsgBox, EntGC, DimUtil,
  UtilArticle, AGLInit;

procedure EntreeSaisieInvTrans(CodeDepot, CodeTrans : String);
procedure SupprimeInvTrans(CodeDepot, CodeTrans : String);

type
  TFSaisieInvTrans = class(TForm)
    PENTETE: THPanel;
    PPIED: THPanel;
    G_InvTrans: THGrid;
    TGIN_EMPLACEMENT: THLabel;
    GIT_DEPOT: THValComboBox;
    TGIT_DEPOT: THLabel;
    GIT_CODETRANS: TEdit;
    GIT_LIBELLE: TEdit;
    TGIT_LIBELLE: THLabel;
    TGIT_CODETRANS: THLabel;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BChercher: TToolbarButton97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    BAddLigne: TToolbarButton97;
    BDelLigne: TToolbarButton97;
    BDelete: TToolbarButton97;
    PageScroller1: TPageScroller;
    TGA_GRILLEDIM4: THLabel;
    GA_CODEDIM4: THCritMaskEdit;
    TGA_GRILLEDIM5: THLabel;
    GA_CODEDIM5: THCritMaskEdit;
    FindLigne: TFindDialog;
    GIN_EMPLACEMENT: TEdit;
    TOTQTEINV: THNumEdit;
    LDTOTQTEINV: TLabel;
    LGTOTQTEINV: TLabel;
    BImprimer: TToolbarButton97;
    LConsult: TLabel;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GIT_DEPOTChange(Sender: TObject);
    procedure GIN_EMPLACEMENTExit(Sender: TObject);
    procedure G_InvTransEnter(Sender: TObject);
    procedure G_InvTransCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure G_InvTransCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure G_InvTransRowExit(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure G_InvTransRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure G_InvTransKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BDelLigneClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FindLigneFind(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure BImprimerClick(Sender: TObject);
    procedure G_InvTransElipsisClick(Sender: TObject);
  private
    { Déclarations privées }
    FFirstFind : boolean;
    TOBEntete : TOB;
    TOBLignes : TOB;
    FColonneLig,FColonneCB, FColonneArt, FColonneLib, FColonneQte, FColonneEmp : integer;
    FColonneDim1,FColonneDim2,FColonneDim3 : integer;
    NbLigDepart : integer;
    HMTrad : THSystemMenu;

    procedure EtudieColsListe;
    procedure InitEntete;
    procedure ChargeListe;
    procedure LoadLesTOB;
    procedure EcrireEntete;
    procedure EcrireLignes;
    procedure EpureLigneZero;
    procedure InsertLigneG_InvTrans;
    procedure InsertNewArtInG_InvTrans(CodeBarre : string; Q : TQuery ; TOBL : TOB; ARow : integer);
    Procedure ChargeArticleTobLigne (Q : TQuery ; TOBLigne : TOB);
    procedure InsertQteInG_InvTrans(Qte : Double ; ARow : integer);
    procedure InsertEmpInG_InvTrans(Emplacement:string; ARow : integer);

    procedure ValiderSaisie;
    function TransactionValide : boolean;
    function QuestionSaisieEnCours : Integer;
    function ICanContinue : boolean;

  public
    { Déclarations publiques }
    FCodeEmpl : String;
    FCodeTrans : String;
    FDepot : String;
    CodeEmpl_cours : String;
    Action : TActionFiche ;
  end;

var
  FSaisieInvTrans: TFSaisieInvTrans;

implementation

{$R *.DFM}

// Procédure d'ouverture de la fiche de saisie d'inventaire
procedure EntreeSaisieInvTrans(CodeDepot, CodeTrans : String);
var FF : TFSaisieInvTrans;
    PPANEL : THPanel;
begin
SourisSablier;
FF := TFSaisieInvTrans.Create(Application);
if CodeTrans = '' then
   begin
   FF.Action:=taCreat ;
   end else
   begin
   FF.Action:=taModif ;
   FF.Fdepot:=CodeDepot ;
   FF.FCodeTrans:=CodeTrans ;
   end;
PPANEL := FindInsidePanel;
if PPANEL = nil then
  begin
  try
    FF.ShowModal;
  finally
    FF.Free;
  end;
  SourisNormale;
  end else
  begin
  InitInside(FF, PPANEL);
  FF.Show;
  end;
end;

// ------- EVENEMENTS ----------------------------------------------------------

procedure TFSaisieInvTrans.FormCreate(Sender: TObject);
begin
// Création des TOBs
TOBEntete := TOB.Create('Entete d''inventaire', nil, -1);
TOBLignes := TOB.Create('Lignes d''inventaire', nil, -1);
end;

procedure TFSaisieInvTrans.FormDestroy(Sender: TObject);
begin
// Destruction des TOBs
TOBEntete.Free; TOBEntete := nil;
TOBLignes.Free; TOBLignes := nil;
end;

procedure TFSaisieInvTrans.FormShow(Sender: TObject);
begin
// Initialisations diverses de toute la fiche
G_InvTrans.ListeParam := 'GCTRANSINVLIGNE';
EtudieColsListe;
//HMTrad.ResizeGridColumns(G_InvTrans);
InitEntete;
if Action=taConsult then
   begin
   BDelLigne.Enabled:=False;
   BDelete.Enabled:=False;
   LConsult.Visible:=True;
   AffecteGrid(G_InvTrans, Action);
   end
   else AffecteGrid(G_InvTrans, taModif);
ChargeListe;
if TOBLignes.Detail.Count>0 then
   begin
   if Action<>taConsult then G_InvTrans.SetFocus;
   G_InvTrans.Row := 1;
   if TOBLignes.Detail[0].GetValue('GIN_ARTICLE')<>'' then
      begin
      DisplayDimensions(TOBLignes.Detail[0], PPied, 'TGA_GRILLEDIM', 'GA_CODEDIM', 4, false);
      G_InvTrans.Col := FColonneQte;
      end
   else G_InvTrans.Col := FColonneCB;
   end;
HMTrad.ResizeGridColumns(G_InvTrans);
end;

procedure TFSaisieInvTrans.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if IsInside(Self) then Action := caFree;
end;

procedure TFSaisieInvTrans.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var St : string ;
begin
if TOBLignes.IsOneModifie then
    BEGIN
    St:='1;?caption?;Confirmez-vous l''abandon de la saisie ?;Q;YN;Y;N;';
    if HShowMessage(St, Caption, '') <> mrYes then CanClose := False ;
    END;
end;

// ------- METHODES ------------------------------------------------------------
// Repère et mémorise les colonnes importantes sur le grid ; interdit également
//  certaines colonnes au tri par clic sur l'entête
procedure TFSaisieInvTrans.EtudieColsListe;
var NomCol, LesCols : String;
    libDim, FF : String;
    col, i : integer;
begin
LesCols := G_InvTrans.Titres[0]; // Contient les noms des champs séparés par des ;
col := 0;
FColonneDim1 := -1;
FColonneDim2 := -1;
FColonneDim3 := -1;

if V_PGI.OkDecQ > 0 then
begin
  FF := '0.';
  for i := 1 to V_PGI.OkDecQ - 1 do
  begin
    FF := FF + '#';
  end;
  FF := FF + '0';
end;

TOTQTEINV.Decimals := V_PGI.OkDecQ;
TOTQTEINV.Masks.PositiveMask := FF;

repeat
  NomCol := Uppercase(Trim(ReadTokenSt(LesCols)));
  if NomCol = 'GIN_NUMLIGNE' then FColonneLig := col
  else if NomCol = 'GIN_CODEBARRE' then FColonneCB := col
  else if NomCol = 'GIN_CODEARTICLE' then FColonneArt := col
  else if NomCol = 'GA_LIBELLE' then FColonneLib := col
  else if NomCol = 'GIN_QTEINV' then
  begin
    FColonneQte := col;
    G_InvTrans.ColTypes[col] := 'I'; // I = pas pouvoir trier
    G_InvTrans.ColFormats[col] := FF + ';' + FF;
  end
  else if NomCol = 'GIN_EMPLACEMENT' then FColonneEmp := col
  else if NomCol = '(DIM1)' then
  begin
  FColonneDim1 := col;
  LibDim:=RechDom('GCCATEGORIEDIM','DI1',False);
  end
  else if NomCol = '(DIM2)' then
  begin
  FColonneDim2 := col;
  LibDim:=RechDom('GCCATEGORIEDIM','DI2',False);
  end
  else if NomCol = '(DIM3)' then
  begin
    FColonneDim3 := col;
    LibDim:=RechDom('GCCATEGORIEDIM','DI3',False);
  end;

  if (NomCol='(DIM1)') or (NomCol='(DIM2)') or (NomCol='(DIM3)') or (NomCol='(DIM4)') or (NomCol='(DIM5)') then
  begin
    G_InvTrans.ColLengths[col] := -1;
    if (LibDim='') or (LibDim='Error') then G_InvTrans.ColLengths[col] := 0
    else G_InvTrans.cells[col,0]:=LibDim ;
  end;
  inc(col);
until (LesCols = '') or (NomCol = '');
G_InvTrans.ColLengths[FColonneLig] := -1;
G_InvTrans.ColLengths[FColonneArt] := -1;
G_InvTrans.ColLengths[FColonneLib] := -1;
end;

// Initialise l'entête de la fiche (affichage des libellés du dépot et de la liste)
procedure TFSaisieInvTrans.InitEntete;
var Q : TQuery;
begin
NbLigDepart:=0;
if not VH_GC.GCMultiDepots then TGIT_DEPOT.Caption:='&Etablissement' ;

if (ctxMode in V_PGI.PGIContexte) then
  begin
  GIT_DEPOT.Plus:='GDE_SURSITE="X"';
  end;

if Action=taCreat then
   begin
   GIT_DEPOT.Value:='' ;
   GIT_CODETRANS.Enabled:=False ;
   end else
   begin
   GIT_DEPOT.Value:=FDepot ;
   GIT_DEPOT.Enabled:=False ;
   GIT_CODETRANS.Text:=FCodeTrans ;
   GIT_CODETRANS.Enabled:=False ;
   Q := OpenSQL('Select GIT_LIBELLE, GIT_INTEGRATION from TRANSINVENT '+
                'where GIT_DEPOT="'+FDepot+'" and GIT_CODETRANS="'+FCodeTrans+'"',true);
   if Not Q.EOF then
      begin
      GIT_LIBELLE.Text := Q.FindField('GIT_LIBELLE').AsString;
      if (Q.FindField('GIT_INTEGRATION').AsString)='X' then Action:=taConsult;
      end;
   Ferme(Q);
   end;
end;

procedure TFSaisieInvTrans.ChargeListe;
begin
Transactions(LoadLesTOB, 1);
TOBLignes.PutGridDetail(G_InvTrans, false, true, G_InvTrans.Titres[0], true);
with G_InvTrans do if Row >= RowCount then Row := RowCount-1;
end;

// Chargement de la liste des lignes d'inventaire
procedure TFSaisieInvTrans.LoadLesTOB;
var Q : TQuery;
    lig : integer;
    TOBLigne : TOB;
    Emplacement : string;
begin
TOBLignes.ClearDetail;

if FCodeEmpl<>'' then Emplacement := 'AND GIN_EMPLACEMENT = "'+FCodeEmpl+'" '
else Emplacement := '';

Q := OpenSQL('SELECT GIN_DEPOT,GIN_CODETRANS,GIN_NUMLIGNE,GIN_CODEBARRE,GIN_ARTICLE,'+
             'GIN_CODEARTICLE,GIN_EMPLACEMENT,GIN_QTEINV,GA_LIBELLE,GA_CODEDIM1,GA_CODEDIM2,'+
             'GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,'+
             'GA_GRILLEDIM4,GA_GRILLEDIM5 '+
             'FROM TRANSINVLIG LEFT JOIN ARTICLE ON GIN_ARTICLE=GA_ARTICLE '+
             'WHERE GIN_DEPOT="'+FDepot+'" AND GIN_CODETRANS="'+FCodeTrans+'" ' +
             Emplacement + 'ORDER BY GIN_NUMLIGNE', true);
TOBLignes.LoadDetailDB('TRANSINVLIG','','',Q,false,true);
NbLigDepart:=TOBLignes.Detail.Count;
CodeEmpl_cours:=FCodeEmpl;
TOTQTEINV.Value := TOBLignes.Somme('GIN_QTEINV',[''],[''],False);
Ferme(Q);

// Recherche des lignes où seul le code à barres est renseigné (Code articles à blanc)
TOBLigne:=TOBLignes.FindFirst(['GIN_ARTICLE'],[''],False);
While TOBLigne<>Nil do
   begin
   Q:=OpenSQL('Select GA_ARTICLE,GA_CODEARTICLE,GA_LIBELLE,GA_CODEDIM1,GA_CODEDIM2,'+
              'GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,'+
              'GA_GRILLEDIM4,GA_GRILLEDIM5 from ARTICLE '+
              'where GA_CODEBARRE="'+TOBLigne.GetValue('GIN_CODEBARRE')+'"',True) ;
   if Not Q.EOF then
      begin
      ChargeArticleTobLigne (Q, TOBLigne);
      TOBLigne.SetAllModifie(True);
      end;
   Ferme(Q);
   TOBLigne:=TOBLignes.FindNext(['GIN_ARTICLE'],[''],False);
   end;

// Parcourt la liste pour ajouter à chaque ligne le libellé des codes dimentions.
for lig := 0 to TOBLignes.Detail.Count-1 do
  begin
  TOBLigne:= TOBLignes.Detail[lig];
  TOBLigne.AddChampSup('(Dim1)',false);
  TOBLigne.AddChampSup('(Dim2)',false);
  TOBLigne.AddChampSup('(Dim3)',false);
  ChargeTob_LibCodeDim(TOBLigne, 3);
  end;
end;

procedure TFSaisieInvTrans.GIT_DEPOTChange(Sender: TObject);
var FNextCode : Integer;
begin
if Action=taCreat then
   begin
   FDepot:=GIT_DEPOT.Value ;
   if FDepot='' then exit ;
   FNextCode:=1 ;
   // Recherche du premier N° de transmission libre pour le dépôt
   while ExisteSQL('Select GIT_CODETRANS from TRANSINVENT '+
                   'where GIT_DEPOT="'+FDepot+'" and '+
                   'GIT_CODETRANS="'+Format('%.3d',[FNextCode])+'"') do inc(FNextCode);
   FCodeTrans:=Format('%.3d',[FNextCode]) ;
   GIT_CODETRANS.Text:=FCodeTrans ;
   end;
end;

procedure TFSaisieInvTrans.GIN_EMPLACEMENTExit(Sender: TObject);
var NbLignes  : Integer;
begin
if FCodeEmpl = GIN_EMPLACEMENT.Text then exit;

NbLignes := TOBLignes.Detail.Count ;
if ICanContinue then  // Confirmation avant de changer d'emplacement
  begin
  FCodeEmpl := Trim(GIN_EMPLACEMENT.Text);
  if (Action=taCreat) and (NbLignes>0) then
     begin
     Action:=taModif;
     InitEntete;
     end;
  ChargeListe;
  end
  else GIN_EMPLACEMENT.Text := FCodeEmpl;

if (Action=taCreat) or ((Action=taModif) and (FCodeEmpl='')) then
     BDelLigne.Enabled:=True
else BDelLigne.Enabled:=False;
end;

procedure TFSaisieInvTrans.G_InvTransEnter(Sender: TObject);
begin
if (GIT_DEPOT.Value='') then
   BEGIN
   PGIInfo('Vous devez renseigner l''établissement', '');
   if GIT_DEPOT.CanFocus then GIT_DEPOT.SetFocus ;
   Exit ;
   END ;
end;

procedure TFSaisieInvTrans.G_InvTransCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBLigne : TOB;
    Numlig : integer;
begin
if Action=taConsult then Exit ;
if (G_InvTrans.Col=FColonneCB) and (G_InvTrans.Cells[FColonneArt,G_InvTrans.Row]<>'') then
   G_InvTrans.Col:=FColonneQte;

G_InvTrans.ElipsisButton := G_InvTrans.Cells[FColonneArt,G_InvTrans.Row]='';   

if G_InvTrans.Col=FColonneQte then
begin
   if (G_InvTrans.Cells[FColonneArt,G_InvTrans.Row]='') then
     G_InvTrans.Col:=FColonneCB
   else
   begin
      Numlig:=StrToInt(G_InvTrans.Cells[FColonneLig,G_InvTrans.Row]);
      TOBLigne:=TOBLignes.FindFirst(['GIN_NUMLIGNE'],[Numlig],False);
      if TOBLigne<>nil then
         DisplayDimensions(TOBLigne, PPied, 'TGA_GRILLEDIM', 'GA_CODEDIM', 4, false);
   end;
end;

if G_InvTrans.Col=FColonneEmp then
   begin
   if (G_InvTrans.Cells[FColonneArt,G_InvTrans.Row]='') then
      G_InvTrans.Col:=FColonneCB;
   end;
end;

procedure TFSaisieInvTrans.G_InvTransCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var CodeBarre : string;
    TOBL : TOB;
    NewQte : double;
    Q : TQuery;
begin
if Action=taConsult then Exit ;
CodeBarre := Trim(G_InvTrans.Cells[FColonneCB,ARow]);
if (ACol=FColonneCB) and (CodeBarre<>'') and (G_InvTrans.Cells[FColonneArt,ARow]='') then
  begin
  // Recherche dans TOBLignes
  TOBL:=TOBLignes.FindFirst(['GIN_CODEBARRE'],[CodeBarre],False);
  if TOBL<>nil then
    begin
    if (TOBL.GetValue('GIN_ARTICLE')<>'') then
       InsertNewArtInG_InvTrans(CodeBarre,Nil,TOBL,ARow)
    else
       begin
       MessageAlerte('ATTENTION : ce code à barres ne correspond à aucun article.');
       G_InvTrans.SetFocus;
       end;
    exit;
    end;

    // Recherche dans la table Article
    Q:=OpenSQL('Select GA_ARTICLE,GA_CODEARTICLE,GA_LIBELLE,GA_CODEDIM1,GA_CODEDIM2,'+
       'GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,'+
       'GA_GRILLEDIM4,GA_GRILLEDIM5 from ARTICLE where GA_CODEBARRE="'+CodeBarre+'"',True) ;
    Try
      if Not Q.EOF then InsertNewArtInG_InvTrans(CodeBarre,Q,Nil,ARow)
      else
        begin
        MessageAlerte('ATTENTION : ce code à barres ne correspond à aucun article.');
        G_InvTrans.SetFocus;
        Cancel:=True;
        end;
    finally Ferme(Q); end;
  end;

if (ACol=FColonneQte) and (G_InvTrans.Cells[FColonneArt,ARow]<>'') then
  begin
  NewQte := Abs(Valeur(G_InvTrans.Cells[FColonneQte,ARow]));
  G_InvTrans.Cells[FColonneQte,ARow]:=FloatToStr(NewQte);
  InsertQteInG_InvTrans(NewQte,ARow);
  end;

if (ACol=FColonneEmp) and (G_InvTrans.Cells[FColonneArt,ARow]<>'') then
  InsertEmpInG_InvTrans(Trim(G_InvTrans.Cells[FColonneEmp,ARow]),ARow);
end;

procedure TFSaisieInvTrans.G_InvTransRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var TOBLigne : TOB;
    Numlig : integer;
begin
if Action=taConsult then
   begin
   Numlig:=StrToInt(G_InvTrans.Cells[FColonneLig,G_InvTrans.Row]);
   TOBLigne:=TOBLignes.FindFirst(['GIN_NUMLIGNE'],[Numlig],False);
   if TOBLigne<>nil then
      DisplayDimensions(TOBLigne, PPied, 'TGA_GRILLEDIM', 'GA_CODEDIM', 4, false);
   Exit ;
   end;
G_InvTrans.InvalidateRow(Ou);
end;

procedure TFSaisieInvTrans.G_InvTransRowExit(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
if Action=taConsult then Exit ;
if (G_InvTrans.RowCount<>2) and (G_InvTrans.Cells[FColonneCB,Ou]='') then
     G_InvTrans.DeleteRow(Ou)
else CodeEmpl_cours:=Trim(G_InvTrans.Cells[FColonneEmp,Ou]);
if not Cancel then G_InvTrans.InvalidateRow(Ou);
end;

procedure TFSaisieInvTrans.G_InvTransKeyDown(Sender: TObject;
  var Key: Word; Shift: TShiftState);
Var Vide,Cancel : Boolean;
    ARow,ACol : Longint;
begin
Vide:=(Shift=[]); ARow:=G_InvTrans.Row; ACol:=G_InvTrans.Col; Cancel:=False;
Case Key of
   VK_RETURN : if Vide then SendMessage(G_InvTrans.Handle, WM_KeyDown, VK_DOWN, 0);
   VK_TAB    : if Vide and (ACol=FColonneEmp) and (ARow=G_InvTrans.RowCount -1) then
                  SendMessage(G_InvTrans.Handle, WM_KeyDown, VK_DOWN, 0);
   VK_DOWN   : BEGIN
               if Vide and (ARow=G_InvTrans.RowCount-1) then
                 begin
                   G_InvTransCellExit(G_InvTrans,ACol,ARow,Cancel);
                   if Not Cancel then G_InvTransRowExit(G_InvTrans,ARow,Cancel,False);
                   if (Not Cancel) and ((Action=taCreat) or ((Action=taModif) and (FCodeEmpl=''))) then
                      InsertLigneG_InvTrans;
                 end;
               END;
  end ;
end;

procedure TFSaisieInvTrans.InsertLigneG_InvTrans;
var ARow : Longint;
begin
if (G_InvTrans.Row=G_InvTrans.Rowcount-1) and (G_InvTrans.Cells[FColonneCB,G_InvTrans.Rowcount-1]<>'') then
  begin
  G_InvTrans.InsertRow(G_InvTrans.Rowcount);
  ARow:=G_InvTrans.Row+1;
  if G_InvTrans.Cells[FColonneArt,ARow]='' then
     begin
     G_InvTrans.Col:=FColonneCB;
     G_InvTrans.Cells[FColonneLig,ARow]:=IntToStr(ARow);
     G_InvTrans.ElipsisButton := True;
     G_InvTrans.CacheEdit; G_InvTrans.MontreEdit;
     end
     else G_InvTrans.Col:=FColonneQte;
  end;
end;

procedure TFSaisieInvTrans.InsertNewArtInG_InvTrans(CodeBarre : string; Q : TQuery ; TOBL : TOB; ARow : integer);
var TOBLigne : TOB;
    Numlig : integer;
begin
// Ajout de la ligne dans TOBLignes
Numlig := Arow;
TOBLigne:=TOBLignes.FindFirst(['GIN_NUMLIGNE'],[Numlig],False);
if TOBL<>nil then
   begin
   if TOBL.GetValue('GIN_NUMLIGNE')<>Numlig then
      begin
      if TOBLigne<>nil then TobLigne.Free;
      TOBLigne := TOB.Create('TRANSINVLIG',TOBLignes,-1);
      TobLigne.Dupliquer (TOBL, False, True, True) ;
      end;
   TOBLigne.SetAllModifie(True);
   end else
   begin
   if TOBLigne<>nil then TobLigne.Free;
   TOBLigne := TOB.Create('TRANSINVLIG',TOBLignes,-1);
   TOBLigne.PutValue('GIN_CODEBARRE',CodeBarre);
   TOBLigne.AddChampSup('GA_LIBELLE',True);
   TOBLigne.AddChampSup('GA_CODEDIM1',True);
   TOBLigne.AddChampSup('GA_CODEDIM2',True);
   TOBLigne.AddChampSup('GA_CODEDIM3',True);
   TOBLigne.AddChampSup('GA_CODEDIM4',True);
   TOBLigne.AddChampSup('GA_CODEDIM5',True);
   TOBLigne.AddChampSup('GA_GRILLEDIM1',True);
   TOBLigne.AddChampSup('GA_GRILLEDIM2',True);
   TOBLigne.AddChampSup('GA_GRILLEDIM3',True);
   TOBLigne.AddChampSup('GA_GRILLEDIM4',True);
   TOBLigne.AddChampSup('GA_GRILLEDIM5',True);
   ChargeArticleTobLigne (Q, TOBLigne);
   TOBLigne.AddChampSup('(Dim1)',false);
   TOBLigne.AddChampSup('(Dim2)',false);
   TOBLigne.AddChampSup('(Dim3)',false);
   ChargeTob_LibCodeDim(TOBLigne, 3);
   end;

   TOBLigne.PutValue('GIN_NUMLIGNE', Numlig);
   TOBLigne.PutValue('GIN_DEPOT', FDepot);
   TOBLigne.PutValue('GIN_CODETRANS', FCodeTrans);
   TOBLigne.PutValue('GIN_EMPLACEMENT', CodeEmpl_cours);
   TOBLigne.PutValue('GIN_QTEINV', 1);
   TOTQTEINV.Value := (TOTQTEINV.Value)+1;

   G_InvTrans.Cells[FColonneLig,ARow]:=IntToStr(Numlig);
   G_InvTrans.Cells[FColonneArt,ARow]:=TOBLigne.GetValue('GIN_CODEARTICLE');
   G_InvTrans.Cells[FColonneLib,ARow]:=TOBLigne.GetValue('GA_LIBELLE');
   if FColonneDim1>=0 then G_InvTrans.Cells[FColonneDim1,ARow]:=TOBLigne.GetValue('(Dim1)');
   if FColonneDim2>=0 then G_InvTrans.Cells[FColonneDim2,ARow]:=TOBLigne.GetValue('(Dim2)');
   if FColonneDim3>=0 then G_InvTrans.Cells[FColonneDim3,ARow]:=TOBLigne.GetValue('(Dim3)');
   G_InvTrans.Cells[FColonneQte,ARow]:='1';
   G_InvTrans.Cells[FColonneEmp,ARow]:=TOBLigne.GetValue('GIN_EMPLACEMENT');

   if TOBLigne.GetValue('GIN_ARTICLE')<>'' then
      DisplayDimensions(TOBLigne, PPied, 'TGA_GRILLEDIM', 'GA_CODEDIM', 4, false);
end;

Procedure TFSaisieInvTrans.ChargeArticleTobLigne (Q : TQuery ; TOBLigne : TOB);
begin
   TOBLigne.PutValue('GIN_ARTICLE',Q.FindField('GA_ARTICLE').AsString);
   TOBLigne.PutValue('GIN_CODEARTICLE',Q.FindField('GA_CODEARTICLE').AsString);
   TOBLigne.PutValue('GA_LIBELLE',Q.FindField('GA_LIBELLE').AsString);
   TOBLigne.PutValue('GA_CODEDIM1',Q.FindField('GA_CODEDIM1').AsString);
   TOBLigne.PutValue('GA_CODEDIM2',Q.FindField('GA_CODEDIM2').AsString);
   TOBLigne.PutValue('GA_CODEDIM3',Q.FindField('GA_CODEDIM3').AsString);
   TOBLigne.PutValue('GA_CODEDIM4',Q.FindField('GA_CODEDIM4').AsString);
   TOBLigne.PutValue('GA_CODEDIM5',Q.FindField('GA_CODEDIM5').AsString);
   TOBLigne.PutValue('GA_GRILLEDIM1',Q.FindField('GA_GRILLEDIM1').AsString);
   TOBLigne.PutValue('GA_GRILLEDIM2',Q.FindField('GA_GRILLEDIM2').AsString);
   TOBLigne.PutValue('GA_GRILLEDIM3',Q.FindField('GA_GRILLEDIM3').AsString);
   TOBLigne.PutValue('GA_GRILLEDIM4',Q.FindField('GA_GRILLEDIM4').AsString);
   TOBLigne.PutValue('GA_GRILLEDIM5',Q.FindField('GA_GRILLEDIM5').AsString);
end;

procedure TFSaisieInvTrans.InsertQteInG_InvTrans(Qte : Double ; ARow : integer);
var TOBLigne : TOB;
    Numlig : integer;
begin
// Mise à jour de la quantité dans TOBLignes
Numlig:=StrToInt(G_InvTrans.Cells[FColonneLig,ARow]);
TOBLigne:=TOBLignes.FindFirst(['GIN_NUMLIGNE'],[Numlig],False);
if TOBLigne<>nil then
  begin
  TOTQTEINV.Value := (TOTQTEINV.Value)+ Qte - TOBLigne.GetValue('GIN_QTEINV');
  TOBLigne.PutValue('GIN_QTEINV', Qte);
  end;
end;

procedure TFSaisieInvTrans.InsertEmpInG_InvTrans(Emplacement : string ; ARow : integer);
var TOBLigne : TOB;
    Numlig : integer;
begin
// Mise à jour de l'emplacement dans TOBLignes
Numlig:=StrToInt(G_InvTrans.Cells[FColonneLig,ARow]);
TOBLigne:=TOBLignes.FindFirst(['GIN_NUMLIGNE'],[Numlig],False);
if TOBLigne<>nil then
  begin
  TOBLigne.PutValue('GIN_EMPLACEMENT', Emplacement);
  end;
end;

procedure TFSaisieInvTrans.BDelLigneClick(Sender: TObject);
var TOBLigne : TOB;
    ARow,Numlig,Numlig_cour : integer;
begin
ARow:=G_InvTrans.Row;
Numlig_cour:=StrToInt(G_InvTrans.Cells[FColonneLig,ARow]);
if Action=taConsult then Exit ;
if ARow<=0 then Exit ;
if G_InvTrans.RowCount<2 then Exit ;
if G_InvTrans.RowCount=2 then G_InvTrans.RowCount:=3;
G_InvTrans.DeleteRow(ARow);

// Suppression de la ligne dans TOBLignes
Numlig:=Numlig_cour;
TOBLigne:=TOBLignes.FindFirst(['GIN_NUMLIGNE'],[Numlig],False);
if TOBLigne<>nil then
   begin
   TOTQTEINV.Value := (TOTQTEINV.Value) - TOBLigne.GetValue('GIN_QTEINV');
   TOBLigne.Free;
   // Décrémentation de 1 du n° de ligne pour les suivantes
   Inc(Numlig);
   TOBLigne:=TOBLignes.FindFirst(['GIN_NUMLIGNE'],[Numlig],False);
   if (TOBLigne<>Nil) or (G_InvTrans.RowCount=1) then
      begin
      While TOBLigne<>Nil do
         begin
         TOBLigne.PutValue('GIN_NUMLIGNE',Numlig-1);
         TOBLigne.SetAllModifie(True);
         Inc(Numlig);
         TOBLigne:=TOBLignes.FindFirst(['GIN_NUMLIGNE'],[Numlig],False);
         end;
      TOBLignes.PutGridDetail(G_InvTrans, false, true, G_InvTrans.Titres[0], true);
      G_InvTrans.Row:=Numlig_cour;
      end
      //else if Numlig_cour > 1 then G_InvTrans.Row:=Numlig_cour-1 else G_InvTrans.Row:=1;
   end;

if G_InvTrans.Cells[FColonneArt,ARow-1]='' then G_InvTrans.Col:=FColonneCB else G_InvTrans.Col:=FColonneQte;
G_InvTrans.SetFocus;
end;

procedure TFSaisieInvTrans.BDeleteClick(Sender: TObject);
var St : string ;
    i  : integer;
    TOBLigne : TOB;
begin
if TOBLignes.Detail.Count = 0 then exit ;

St:='1;?caption?;Confirmez-vous la remise à zéro de toutes les quantités ?;Q;YN;Y;N;';
if HShowMessage(St, Caption, '') = mrYes then
   begin
   for i:=0 to TOBLignes.Detail.Count-1 do
      begin
      TOBLigne := TOBLignes.Detail[i] ;
      TOBLigne.PutValue('GIN_QTEINV', 0);
      end;
   TOTQTEINV.Value := 0;
   TOBLignes.PutGridDetail(G_InvTrans, false, true, G_InvTrans.Titres[0], true);
   end;
end;

procedure TFSaisieInvTrans.BValiderClick(Sender: TObject);
begin
ValiderSaisie;
end;

procedure TFSaisieInvTrans.ValiderSaisie;
var ACol,ARow : integer;
    Cancel : boolean;
begin
ACol:=G_InvTrans.Col; ARow:=G_InvTrans.Row; Cancel := False;
G_InvTransCellExit(G_InvTrans,ACol,ARow,Cancel);
if Cancel or ((TOBLignes.Detail.Count=0) and (Action=taCreat)) then exit;
if Action<>taConsult then
   begin
     try
      EcrireEntete;
      EcrireLignes;
     finally
     end;
   end;
TOBLignes.ClearDetail;
end;

// Ecriture en Table de l'entête de l'inventaire transmis
procedure TFSaisieInvTrans.EcrireEntete;
var FNextCode : Integer;
    CodeTrans : String;
begin
if Action=taCreat then
   begin
   // Recherche à nouveau du premier N° transmission disponible.
   // Le n° attribué en début de saisie a pu être utilisé par un autre utilisateur.
   FNextCode:=StrToInt(Trim(FCodeTrans)) ;
   while ExisteSQL('Select GIT_CODETRANS from TRANSINVENT '+
                   'where GIT_DEPOT="'+FDepot+'" and '+
                   'GIT_CODETRANS="'+Format('%.3d',[FNextCode])+'"') do inc(FNextCode);
   CodeTrans:=Format('%.3d',[FNextCode]) ;
   if CodeTrans<>FCodeTrans then
      begin
      FCodeTrans:=CodeTrans ;
      GIT_CODETRANS.Text:=FCodeTrans ;
      PGIInfo('ATTENTION, le n° de transmission a changé', '');
      end;
   end;

with TOB.Create('TRANSINVENT', TOBEntete, -1) do
   begin
   PutValue('GIT_DEPOT', FDepot);
   PutValue('GIT_CODETRANS', FCodeTrans);
   PutValue('GIT_LIBELLE', GIT_LIBELLE.Text);
   PutValue('GIT_ETABLISSEMENT', FDepot);
   //PutValue('GIT_DATECREATION', NowH);
   //PutValue('GIT_DATEMODIF', NowH);
   PutValue('GIT_UTILISATEUR', V_PGI.USER);
   PutValue('GIT_INTEGRATION', '-');
   end;

if Action=taCreat then
   begin
   TOBEntete.Detail[0].PutValue('GIT_CREATEUR', V_PGI.USER);
   TOBEntete.InsertDB(nil, False);
   end else
   begin
   TOBEntete.InsertOrUpdateDB(False);
   end;
end;

procedure TFSaisieInvTrans.EcrireLignes;
var i, NbLig : integer ;
    ExisteQteZero : boolean ;
    SQL : string ;
    TOBL : TOB;
begin
  ExisteQteZero := False ;
  for i:=0 to TOBLignes.Detail.Count-1 do
    begin
    TobL := TOBLignes.Detail[i] ;
    TobL.PutValue('GIN_DEPOT', FDepot);
    TobL.PutValue('GIN_CODETRANS',FCodeTrans);
    if TobL.GetValue('GIN_QTEINV')=0 then ExisteQteZero := True ;
    end;

  if Action=taCreat then
     begin
     TOBLignes.InsertDB(nil, True);
     end else
     begin
     TOBLignes.InsertOrUpdateDB(True);

     NbLig:=TOBLignes.Detail.Count;
     if NbLigDepart>NbLig then
        begin
        SQL:='delete from TRANSINVLIG '+
                     'where GIN_DEPOT="'+FDepot+'" and GIN_CODETRANS="'+FCodeTrans+'" and '+
                     'GIN_NUMLIGNE>'+IntToStr(NbLig) ;
        ExecuteSQL(SQL) ;
        end;
     if (NbLig=0) and (FCodeEmpl='') then
        begin
        // Suppression de l'entête de l'inventaire transmis
        SQL:='delete from TRANSINVENT '+
                     'where GIT_DEPOT="'+FDepot+'" and GIT_CODETRANS="'+FCodeTrans+'"';
        ExecuteSQL(SQL) ;
        end;
     end;

  if ExisteQteZero = True then EpureLigneZero ;
end;

procedure TFSaisieInvTrans.EpureLigneZero;
var CodeEmpl_sav : string ;
    i,j : integer;
    Numlig : integer;
    Nblig_Tot : integer;
    TOBLigne : TOB;
begin
// Suppression des lignes avec une quantité à zéro, et renumérotation des lignes restantes sur la totalité de l'inventaire transmis
CodeEmpl_sav := FCodeEmpl ;
FCodeEmpl := '';
LoadLesTOB;
j:=0;
Numlig:=1;
Nblig_Tot:=TOBLignes.Detail.Count;
for i:=0 to Nblig_Tot-1 do
   begin
   TOBLigne := TOBLignes.Detail[j] ;
   if TOBLigne.GetValue('GIN_QTEINV')=0 then TOBLigne.Free
   else
      begin
      if TOBLigne.GetValue('GIN_NUMLIGNE')<>Numlig then
         begin
         TOBLigne.PutValue('GIN_NUMLIGNE', Numlig);
         TOBLigne.SetAllModifie(True);
         end;
      inc(Numlig);
      inc(j);
      end;
   end;
EcrireLignes;
FCodeEmpl := CodeEmpl_sav ;
end;

procedure TFSaisieInvTrans.BAbandonClick(Sender: TObject);
begin

end;

// Transactionne la methode précédente, avec messages d'erreur et tout
function TFSaisieInvTrans.TransactionValide : boolean;
var IOErr : TIOErr;
begin
IOErr := Transactions(ValiderSaisie, 2);
result := false;
 case IOErr of
  oeOK : result := true;
  oeUnknown : PGIBox('Saisie non enregistrée !','');
  oeSaisie : PGIBox('Saisie non enregistrée (en cours de traitement par un autre utilisateur) !','');
 end;
end;

// Si la saisie a été modifiée, demande confirmation pour enregistrer...
function TFSaisieInvTrans.QuestionSaisieEnCours : Integer;
begin
if TOBLignes.IsOneModifie
  then result := HShowMessage('0;?caption?;La saisie a été modifiée, voulez-vous enregistrer?;Q;YNC;Y;C;','','')
  else result := mrNo;
end;

// Je peux continuer ? Seulement après avoir demandé confirmation, et être sur
//  que la sauvegarde s'est bien faite
function TFSaisieInvTrans.ICanContinue : boolean;
begin
result := false;
 case QuestionSaisieEnCours of
    mrYes : if TransactionValide then result := true;
    mrNo : result := true;
 end;
end;


// Recherche
procedure TFSaisieInvTrans.BChercherClick(Sender: TObject);
begin
if G_InvTrans.RowCount < 3 then Exit;
FFirstFind := true;
FindLigne.Execute;
end;

procedure TFSaisieInvTrans.FindLigneFind(Sender: TObject);
begin
Rechercher(G_InvTrans, FindLigne, FFirstFind);
end;


procedure SupprimeInvTrans(CodeDepot, CodeTrans : String);
var SQL : string ;
begin
   if (PGIAsk('Etes-vous sûr de vouloir supprimer cet inventaire ?','') = mrNo) then exit;

   SQL:='delete from TRANSINVENT '+
                     'where GIT_DEPOT="'+CodeDepot+'" and GIT_CODETRANS="'+CodeTrans+'"';
   ExecuteSQL(SQL) ;

   SQL:='delete from TRANSINVLIG '+
                     'where GIN_DEPOT="'+CodeDepot+'" and GIN_CODETRANS="'+CodeTrans+'"';
   ExecuteSQL(SQL) ;
end;

// Procédure pour l'appel de la saisie d'inventaire depuis le script
procedure AGLEntreeSaisieInvTrans(Parms : Array of Variant; Nb : Integer);
begin
EntreeSaisieInvTrans(Parms[0], Parms[1]);
end;

procedure AGLSupprimeInvTrans(Parms : Array of Variant; Nb : Integer);
begin
SupprimeInvTrans(Parms[0], Parms[1]);
end;

procedure TFSaisieInvTrans.BImprimerClick(Sender: TObject);
var stwhere,critere : string ;
    Qnbempl : TQuery ;
    nbempl : integer ;
begin
     stwhere:=' GIN_DEPOT="'+FDepot+'" and GIN_CODETRANS="'+FCodeTrans+'"';
     critere:=';XX_ETABLISSEMENT='+FDepot+'  '+GIT_DEPOT.Text+';XX_EMPLACEMENT='+FCodeEmpl+';XX_TRANSMISSION='+FCodeTrans+';XX_LIBELLE='+GIT_LIBELLE.Text;
     // Si on a choisi un emplacement, j'édite la liste des inventaires pour cet emplacement
     if (GIN_EMPLACEMENT.Text<>'') then
       begin
            stwhere:= stwhere+' AND GIN_EMPLACEMENT = "'+FCodeEmpl+'" ORDER BY gin_numligne,gin_depot,gin_emplacement';
            LanceEtat('E','EIT','EIT',True,False,False,Nil,stWhere,'',False,0,'XX_RUPTURE='+critere);
       end
     // Sinon si dans la liste il y a au moins un emplacement different de vide je pose la question sur le saut de page
     else
       begin
            Qnbempl:=OpenSQL('SELECT count(*) FROM TRANSINVLIG WHERE GIN_CODETRANS="'+FCodeTrans+'" AND GIN_EMPLACEMENT<>""',True);
            nbempl := Qnbempl.Fields[0].AsInteger;
            Ferme(Qnbempl);
            if nbempl<>0 then
              begin
               if (PGIAsk('Voulez vous un saut de page par emplacement ?','') = mrNo) then
                    // si pas de saut de page, j'édite la liste complète dans l'ordre de l'ecran
                 begin
                    stwhere := stwhere + ' ORDER BY gin_numligne,gin_depot,gin_emplacement';
                    LanceEtat('E','EIT','EIT',True,False,False,Nil,stWhere,'',False,0,'XX_RUPTURE='+critere);
                 end
               else // sinon j'édite la liste avec une rupture par emplacement avec un saut de page
                 begin
                    stwhere := stwhere + ' ORDER BY gin_emplacement,gin_numligne,gin_depot';
                    LanceEtat('E','EIT','EIT',True,False,False,Nil,stWhere,'',False,0,'XX_RUPTURE=X'+critere);
                 end;
              end
            else
              begin
                   stwhere := stwhere + ' ORDER BY gin_numligne,gin_depot,gin_emplacement';
                   LanceEtat('E','EIT','EIT',True,False,False,Nil,stWhere,'',False,0,'XX_RUPTURE='+critere);
              end;
       end;
end;

procedure TFSaisieInvTrans.G_InvTransElipsisClick(Sender: TObject);
var Article, CB : String;
    QQ : TQuery;
begin
  CB := G_InvTrans.Cells[FColonneCB,G_InvTrans.Row];
  if ctxMode in V_PGI.PGIContexte then
    Article := DispatchArtMode(1,'','','XX_WHERE=GA_TENUESTOCK="X";GA_CODEBARRE=' + CB)
  else
    Article := AGLLanceFiche('GC','GCARTICLE_RECH','','','XX_WHERE=GA_TENUESTOCK="X";GA_CODEBARRE=' + CB);

  if Article = '' then Exit;

  QQ := OpenSQL('Select GA_CODEBARRE, GA_STATUTART From Article Where GA_ARTICLE="' + Article + '"', True);
  if Not QQ.Eof then
  begin
    if QQ.FindField('GA_STATUTART').AsString = 'GEN' then
    begin
      TheTOB := TOB.Create('',Nil,-1);
      AglLanceFiche ('GC','GCSELECTDIM','','', 'GA_ARTICLE='+Article+';ACTION=SELECT;CHAMP= ');
      if TheTOB <> nil then
      begin
        Article := TheTOB.Detail[0].GetValue('GA_ARTICLE');
        Ferme(QQ);
        QQ := OpenSQL('Select GA_CODEBARRE From Article Where GA_ARTICLE="' + Article + '"', True);
      end;
    end;
  end;
  if Not QQ.Eof then G_InvTrans.Cells[FColonneCB,G_InvTrans.Row] := QQ.FindField('GA_CODEBARRE').AsString;
  Ferme(QQ);
end;

Initialization
RegisterAGLProc('EntreeSaisieInvTrans',false,2,AGLEntreeSaisieInvTrans);
RegisterAGLProc('SupprimeInvTrans',false,2,AGLSupprimeInvTrans);
end.
