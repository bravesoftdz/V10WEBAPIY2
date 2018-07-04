unit ZCentral ;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Grids, Hctrls, Mask, ExtCtrls, Buttons, Menus,
  Hspliter, HEnt1, HMsgBox, HQry, HFLabel, HSysMenu, HPop97, HTB97, HPanel,
  ed_tools, Ent1, UiUtil, SaisUtil, ParamSoc, ImgList, ComCtrls
{$IFDEF EAGLCLIENT}
  ,uTOB
{$ELSE}
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
  ,DB
{$ENDIF}
   ,Saisie
  ,SaisBor, HImgList
;

function Centralisateur(Action: TActionFiche; TypeEcr: string; bANO: Boolean) : Boolean ;
function GetFromCentralisateur(TypeEcr: string; var ParPeriode, ParCodeJal, ParNumFolio : string; CbPeriode: THValComboBox) : Boolean ;
function CWhereConfidentiel : string;

const SC_TYPEJAL  : integer = 0 ;
      SC_NUMPIECE : integer = 1 ;
      SC_LIB      : integer = 2 ;
      SC_DEBIT    : integer = 3 ;
      SC_CREDIT   : integer = 4 ;
      SC_LIGNES   : integer = 5 ;
      SC_FIRST    : integer = 2 ;
      SC_LAST     : integer = 5 ;

const RC_TOUS             =  0 ;
      RC_BORDEREAU        =  1 ;
      RC_PIECE            =  2 ;
      RC_NORECORD         =  3 ;
      RC_TOTAL            =  4 ;
      RC_NOMVT            =  5 ;
      RC_CVISU            =  6 ;
      RC_CMODIF           =  7 ;
      RC_CCOURANTES       =  8 ;
      RC_CSIMULATION      =  9 ;
      RC_CPREVISION       = 10 ;
      RC_CSITUATION       = 11 ;
      RC_CREVISION        = 12 ;
      RC_CANOUVEAU        = 13 ;
      RC_VERYLONGTIME     = 14 ;
      RC_TOTALGENERAL     = 15 ;

const MAX_ROW             = 100 ;

const NRC_VERYLONGTIME    = '14;?caption?;Attention ce traitement peut être long#10#13          Voulez-vous continuer ?;Q;YN;Y;Y;' ;

const ColLibelle          = 300 ;
      ColMouvement        = 100 ;

type
  TModeVisu = (tmvPeriode, tmvJournal, tmvEcriture) ;

type
  TTriTw = ( ttPeriode, ttJal ) ;

  RNAVIG = record
    Exercice, Periode, Journal : Integer ; Code : String ;
  end ;
  PNAVIG = ^RNAVIG ;

  RMVTSTATE = record
    DebPeriode : TDateTime ; Journal : string ;
  end ;
  PMVTSTATE = ^RMVTSTATE ;

type
  TFCentral = class(TForm)
    PEntete: THPanel;
    HCentral: THMsgBox;
    HMTrad: THSystemMenu;
    DockTop: TDock97;
    DockRight: TDock97;
    DockLeft: TDock97;
    DockBottom: TDock97;
    Valide97: TToolbar97;
    BValide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    EXO: THValComboBox;
    Outils: TToolbar97;
    BCherche: TToolbarButton97;
    HLabel3: THLabel;
    BInit: TToolbarButton97;
    BNivPrec: TToolbarButton97;
    BStop: TToolbarButton97;
    Splitter1: TSplitter;
    HLabel2: THLabel;
    PERIODE: THValComboBox;
    HLabel1: THLabel;
    JOURNAUX: THValComboBox;
    ImageList: THImageList;
    BThread: TToolbarButton97;
    HPanel1: THPanel;
    TreeView: TTreeView;
    Splitter2: TSplitter;
    HPanel3: THPanel;
    BTree: TToolbarButton97;
    POPB: TPopupMenu;
    MenuParPeriode: TMenuItem;
    MenuParJournal: TMenuItem;
    GS: THGrid;
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormShow(Sender: TObject);
    procedure EXOChange(Sender: TObject);
    procedure GSDblClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BInitClick(Sender: TObject);
    procedure BNivPrecClick(Sender: TObject);
    procedure BStopClick(Sender: TObject);
    procedure TreeViewChange(Sender: TObject; Node: TTreeNode);
    procedure TreeViewExpanding(Sender: TObject; Node: TTreeNode; var AllowExpansion: Boolean);
    procedure BAideClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Splitter2Moved(Sender: TObject);
    procedure MenuParPeriodeClick(Sender: TObject);
    procedure MenuParJournalClick(Sender: TObject);
    procedure ChangeTri( vBoParPeriode : Boolean );
    
  private
    iLastExo, iLastPeriode, iLastJournaux : Integer ;
    NavPrec      : TList ;
    MvtState     : TList ;
    bStopSearch  : Boolean ;
    FAction      : TActionFiche ;
    FTypeEcr     : string ;
    FCbPeriode   : THValComboBox ;
    FbANO        : Boolean ;
    FbExoClos    : Boolean ;
    memPivot     : RDevise ;
    bExpanding   : Boolean ;
    TCount       : LongInt ;
    TDebit       : Double ;
    TCredit      : Double ;
    FThreadEnded : Boolean ;
    FRow         : LongInt ;

    // GCO - 09/09/2004 - FQ 10049
    FMonTotalExo  : Integer;
    FMonDebitExo  : Double;
    FMonCreditExo : Double;

    FTriTw : TTriTW; // stocke le mode de tri courant pour le TreeView, par defaut TriPeriode

    // Fonctions ressource
    function  GetMessageRC(MessageID : integer) : string ;

    // Fonctions utilitaires
    procedure AfficheSolde(Row : LongInt) ;
    procedure AfficheTotalGeneral ;
    function  ExistPeriode(Per : string) : Boolean ;
    function  ExistJournal(Per, Jal : string) : Boolean ;
    function  ExistJournalSeul( Jal : string ) : Boolean;
    function  GetTotal : Boolean ;
    procedure InitCaption ;
    procedure InitHelp ;
    procedure InitView ;
    procedure InitGrid ;
    procedure GetCellCanvas(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
    function  GetMode : TModeVisu ;
    procedure PutMontant(Montant : Double; Col, Row : LongInt) ;
    function  VerifCloture(sPeriode, sJournal: string) : Boolean ;
    procedure VerifTreeView(sPeriode, sJournal : string) ;
    procedure SyncTreeView(Node : TTreeNode) ;
    procedure FillTreeView ;
    function  PutPeriodeView(var Row : LongInt; Year, Month : Word;
                             sPeriode, CodeJal, LibJal, NumPiece : string;
                             ModeSaisie: string; ModeVisu : TModeVisu) : Boolean ;

    // Click boutons
    function  StopSearch : Boolean ;
    procedure SearchClick(bAddPrec : Boolean=TRUE) ;
    function  ValClick : Boolean ;
    procedure ByeClick ;
    procedure InitClick ;
  public
    bFromFolio, bByParams : Boolean ;
    ParPeriode, ParCodeJal, ParNumFolio : string ;
  end;

implementation

{$R *.DFM}

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/01/2003
Modifié le ... : 20/01/2003
Description .. : Retourne la condition qui gère le E_CONFIDENTIEL 
Mots clefs ... : E_CONFIDENTIEL
*****************************************************************}
function CWhereConfidentiel : string;
begin
  Result := '';
  if V_PGI.Confidentiel = '0' then
    Result := ' AND (E_CONFIDENTIEL="0" OR E_CONFIDENTIEL="-") '
  else
    Result := ' AND (E_CONFIDENTIEL<="' + V_PGI.Confidentiel + '" OR E_CONFIDENTIEL="X" OR E_CONFIDENTIEL="-") ';
  (*
  if V_PGI.Confidentiel = '1' then
    Result := ' AND (E_CONFIDENTIEL<="' + V_PGI.Confidentiel + '" OR E_CONFIDENTIEL="X" OR E_CONFIDENTIEL="-") '
  else
    Result := ' AND (E_CONFIDENTIEL<="' + V_PGI.Confidentiel + '" OR E_CONFIDENTIEL="X" OR E_CONFIDENTIEL="-") ';
  *)
end;

//=======================================================
//======== Point d'entrée dans la saisie balance ========
//=======================================================
function CanOpenCentralisateur : Boolean ;
begin
Result:=TRUE ;
if not GetParamSoc('SO_ZJALCENTRAL') then Exit ;
if (HShowMessage(NRC_VERYLONGTIME, '', '')<>mrYes) then Result:=FALSE ;
end ;

function Centralisateur(Action: TActionFiche; TypeEcr: string; bANO: Boolean) : Boolean ;
var Central : TFCentral; PP : THPanel ;
begin
Result:=TRUE ;
if not CanOpenCentralisateur then begin Result:=FALSE ; Exit ; end ;
Central:=TFCentral.Create(Application) ;
Central.bFromFolio:=FALSE ;
Central.FAction:=Action ;
Central.FTypeEcr:=TypeEcr ;
Central.FbANO:=bANO ;
PP:=FindInsidePanel ;
if PP=nil then
   begin
   try
     Central.ShowModal ;
   finally
     Central.Free ;
   end ;
  Screen.Cursor:=SyncrDefault ;
  end else
  begin
  InitInside(Central, PP) ;
  Central.Show ;
  end ;
end ;

function GetFromCentralisateur(TypeEcr: string; var ParPeriode, ParCodeJal, ParNumFolio : string; CbPeriode: THValComboBox) : Boolean ;
var Central : TFCentral;
begin
if not CanOpenCentralisateur then begin Result:=FALSE ; Exit ; end ;
Central:=TFCentral.Create(Application) ;
try
  Central.bFromFolio:=TRUE ;
  Central.FTypeEcr:=TypeEcr ;
  Central.FCbPeriode:=CbPeriode ;
  Central.ShowModal ;
  Result:=Central.bByParams ;
  if Result then
    begin
    ParPeriode:=Central.ParPeriode ;
    ParCodeJal:=Central.ParCodeJal ;
    ParNumFolio:=Central.ParNumFolio ;
    end ;
finally
  Central.Free ;
end ;
Screen.Cursor:=SyncrDefault ;
end ;

//=======================================================
//================= Fonctions Ressource =================
//=======================================================
function TFCentral.GetMessageRC(MessageID : Integer) : string ;
begin
Result:=HCentral.Mess[MessageID] ;
end ;

//=======================================================
//================ Evénements de la Form ================
//=======================================================
procedure TFCentral.FormShow(Sender: TObject);
var QuelExo : string ;
    i,j : integer;
begin
FMonTotalExo  := 0;
FMonDebitExo  := 0;
FMonCreditExo := 0;

// GCO - 18/07/2002
for i := 0 to Journaux.Items.Count - 1 do
begin
  for j := 0 to Journaux.Items.Count - 1 do
  begin
    if ( Journaux.Items.Strings[i] = Journaux.Items.Strings[j] ) and ( Journaux.Values[i] <> Journaux.Values[j] ) then
    begin
      Journaux.Items.Strings[i] := Journaux.Items.Strings[i] + ' (' + Journaux.Values[i] + ')';
      Journaux.Items.Strings[j] := Journaux.Items.Strings[j] + ' (' + Journaux.Values[j] + ')';
    end;
  end;
end;
// FIN GCO

bExpanding:=FALSE ;
BValide.Visible:=FALSE ; BNivPrec.Enabled:=FALSE ;
LookLesDocks(Self) ;
ParPeriode:='' ; ParCodeJal:='' ; ParNumFolio:='' ; bByParams:=FALSE ;
iLastExo:=-1 ; iLastPeriode:=-1 ; iLastJournaux:=-1 ; bStopSearch:=FALSE ;
memPivot.Code:=V_PGI.DevisePivot ; if memPivot.Code<>'' then GetInfosDevise(memPivot) ;
// Thread
TCount:=0 ; TDebit:=0 ; TCredit:=0 ; FThreadEnded:=TRUE ; FRow:=0 ;
InitCaption ;
InitHelp ;
MvtState:=TList.Create ;
//NavPrec:=TList.Create ;
if VH^.CPExoRef.Code<>'' then QuelExo:=VH^.CPExoRef.Code else
   BEGIN
   if (VH^.Entree.Code=VH^.Encours.Code) or (VH^.Entree.Code=VH^.Suivant.Code) then QuelExo:=VH^.Entree.Code else QuelExo:=VH^.Encours.Code ;
   END ;
EXO.ItemIndex:=EXO.Values.IndexOf(QuelExo) ;

if GetParamSocSecur('SO_CPJALCENTRALSURPERIODE', True) then
  FTriTw := TTPeriode
else
  FTriTw := TTJal;
ChangeTri(FTriTw = TTPeriode);

InitGrid ;
EXOChange(nil) ;
SearchClick ;
TreeView.SetFocus ;
if bFromFolio then EXO.Enabled:=FALSE ;
HMTrad.ResizeGridColumns(GS) ;
end;

procedure TFCentral.FormClose(Sender: TObject;
  var Action: TCloseAction);
//var Nav : PNAVIG ;
var Mvt : PMVTSTATE ;
begin
(*
if NavPrec<>nil then
  begin
  while NavPrec.Count>0 do
    begin Nav:=NavPrec.Items[0] ; Dispose(Nav) ; NavPrec.Delete(0) ; end ;
  NavPrec.Free ;
  end ;
*)
if MvtState<>nil then
  begin
  while MvtState.Count>0 do
    begin Mvt:=MvtState.Items[0] ; Dispose(Mvt) ; MvtState.Delete(0) ; end ;
  MvtState.Free ;
  end ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFCentral.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
var Vide : Boolean ;
begin
if Key=VK_ESCAPE then begin Key:=0 ; ByeClick ; end ;
if not GS.SynEnabled then begin Key:=0 ; Exit ; end ;
if not (Screen.ActiveControl=GS) then Exit ;
Vide:=(Shift=[]) ;
case Key of
  VK_RETURN : if (Vide) then Key:=VK_TAB ;
  VK_F5     : if (Vide) then begin Key:=0 ; GSDblClick(nil); end ;
  end ;
end;

procedure TFCentral.EXOChange(Sender: TObject);
var i : Integer ;
begin
// Exercice clos ?
if (EXO.Values[EXO.ItemIndex]=VH^.Encours.Code) or
   (EXO.Values[EXO.ItemIndex]=VH^.Suivant.Code) then FbExoClos:=FALSE else FbExoClos:=TRUE ;
PERIODE.Clear ; PERIODE.Values.Clear ;
InitGrid ;
// Gestion du TreeView
if bFromFolio then
  begin
  for i:=0 to FCbPeriode.Items.Count-1 do
    begin
    PERIODE.Items.Add(FCbPeriode.Items[i]) ;
    PERIODE.Values.Add(FCbPeriode.Values[i]) ;
    end ;
  end else ListePeriode(EXO.Values[EXO.ItemIndex], PERIODE.Items, PERIODE.Values, TRUE) ;
PERIODE.Items.Insert(0, GetMessageRC(RC_TOUS)) ;
PERIODE.Values.Insert(0, GetMessageRC(RC_TOUS)) ;
PERIODE.ItemIndex:=1 ;
JOURNAUX.ItemIndex:=0 ;
TreeView.Items.Clear ;
InitView ;
FillTreeView ;
if TreeView.Items.Count>0 then TreeView.Selected:=TreeView.Items[0] else SearchClick ;

// GCO - 31/10/2006 - FQ 18996
GetTotal ;
end;

//=======================================================
//================ Fonctions utilitaires ================
//=======================================================
function TFCentral.GetTotal : Boolean ;
var Q : TQuery ; stSQL: string ;
begin
  stSQL := 'SELECT COUNT(*) LETOTAL, SUM(E_DEBIT) LEDEBIT, SUM(E_CREDIT) LECREDIT FROM ECRITURE WHERE E_EXERCICE="'+EXO.Values[EXO.ItemIndex]+'" AND E_QUALIFPIECE="'+FTypeEcr+'"' ;
  try
    try
      Q := OpenSql( StSql, True );
      if not Q.EOF then
      begin
        TCount := Q.FindField('LETOTAL').AsInteger;
        TDebit := Q.FindField('LEDEBIT').AsFloat;
        TCredit:= Q.FindField('LECREDIT').AsFloat;
      end ;
    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : GetTotal');
    end;
  finally
    Ferme( Q );
  end;
  AfficheTotalGeneral ;
  Result := TRUE ;
end ;

procedure TFCentral.InitCaption ;
begin
Exit; // GCO 05/02/2003
case FAction of
  taModif: Caption:=GetMessageRC(RC_CMODIF) ;
  else     Caption:=GetMessageRC(RC_CMODIF) ;
end ;
if FbANO then Caption:=Caption+' '+GetMessageRC(RC_CANOUVEAU)
else if FTypeEcr='N' then Caption:=Caption+' '+GetMessageRC(RC_CCOURANTES)
else if FTypeEcr='S' then Caption:=Caption+' '+GetMessageRC(RC_CSIMULATION)
else if FTypeEcr='P' then Caption:=Caption+' '+GetMessageRC(RC_CPREVISION)
else if FTypeEcr='U' then Caption:=Caption+' '+GetMessageRC(RC_CSITUATION)
else if FTypeEcr='R' then Caption:=Caption+' '+GetMessageRC(RC_CREVISION) ;
UpdateCaption(Self) ;
end ;

procedure TFCentral.InitHelp ;
begin
if FTypeEcr='N' then HelpContext:=7261000 else
  if FTypeEcr='S' then HelpContext:=7278000 else
    if FTypeEcr='P' then HelpContext:=7318000 else
      if FTypeEcr='U' then HelpContext:=7302000 ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 03/04/2003
Modifié le ... :   /  /    
Description .. : - 03/04/2003 - FB 12134 - on redimmensionne la grille qd 
Suite ........ : elle pas en mode inside
Mots clefs ... : 
*****************************************************************}
procedure TFCentral.InitGrid ;
begin
  GS.VidePile(FALSE) ;
  GS.RowCount:=GS.FixedRows+1 ;
  GS.Refresh ;
  GS.ColWidths[SC_TYPEJAL] := -1;
  GS.ColWidths[SC_NUMPIECE]:= -1;
  if not IsInside(self) then
   begin
    GS.ColWidths[SC_LIB]     := colLibelle - 50 ;
    GS.ColWidths[SC_DEBIT]   := 110 ;
    GS.ColWidths[SC_CREDIT]  := 110 ;
    GS.ColWidths[SC_LIGNES]  := 70 ;
    GS.Cells[SC_LIGNES,0]    := 'Mvt.' ;
   end
   else
    begin
     GS.ColWidths[SC_LIB]     := colLibelle;
     GS.ColWidths[SC_LIGNES]  := ColMouvement ;
    end ;
  GS.ColAligns[SC_DEBIT]:=taRightJustify ;
  GS.ColAligns[SC_CREDIT]:=taRightJustify ;
  GS.ColAligns[SC_LIGNES]:=taRightJustify ;
  GS.GetCellCanvas:=GetCellCanvas ;
  HMTrad.ResizeGridColumns(GS);
//  GS.ColWidths[SC_LIGNES]:= ColMouvement - 13;
end ;

procedure TFCentral.GetCellCanvas(ACol, ARow : LongInt; Canvas : TCanvas; AState : TGridDrawState) ;
begin
  if Arow = 0 then Exit;
  if (GS.RowCount - GS.FixedRows - ARow) < 2 then
    Canvas.Font.Style :=[fsBold];
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 03/04/2003
Modifié le ... :   /  /    
Description .. : - 30/09/2004 - FB 14631 - rajout du code exo pour l'ouverture de la saisie bordereau
Mots clefs ... : 
*****************************************************************}
procedure TFCentral.GSDblClick(Sender: TObject);
var Q : TQuery ; Cols, Where, OrderBy : string ;
    Year, Month, Day : Word ;
    lQuery : TQuery;
    lSt    : string;
begin
  if (GS.RowCount - GS.FixedRows - GS.Row) < 2 then Exit;

case GetMode of
  tmvPeriode  :
    begin
    PERIODE.ItemIndex:=PERIODE.Items.IndexOf(GS.Cells[SC_LIB, GS.Row]) ;
    BChercheClick(nil) ;
    end ;
  tmvJournal  :
    begin
    JOURNAUX.ItemIndex:=JOURNAUX.Items.IndexOf(GS.Cells[SC_LIB, GS.Row]) ;
    BChercheClick(nil) ;
    end ;
  tmvEcriture :
    begin
    if ValClick then begin ByeClick ; Exit ; end ;
    DecodeDate(StrToDate(PERIODE.Values[PERIODE.ItemIndex]), Year, Month, Day) ;
    Cols := 'E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE' ;

    Where := 'E_DATECOMPTABLE>="'+USDateTime(EncodeDate(Year, Month, 1))+'" '
            + ' AND E_EXERCICE = "' + EXO.Values[EXO.ItemIndex] + '" '
            +' AND E_DATECOMPTABLE<="'+USDateTime(FinDeMois(EncodeDate(Year, Month, 1)))+'"'
            +' AND E_JOURNAL="'+JOURNAUX.Values[JOURNAUX.ItemIndex]+'"'
            +' AND E_NUMEROPIECE='+GS.Cells[SC_NUMPIECE, GS.Row]
            +' AND E_NUMLIGNE=1 AND E_NUMECHE<=1'
            +' AND E_QUALIFPIECE="'+FTypeEcr+'"' ;

    OrderBy := 'E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE';

    Q :=OpenSQL('SELECT '+Cols+' FROM ECRITURE WHERE '+Where+' ORDER BY '+OrderBy, TRUE) ;

    // GCO - 10/09/2004 - FQ 10049
    lSt := 'SELECT COUNT(*) LETOTAL, SUM(E_DEBIT) LEDEBIT, SUM(E_CREDIT) LECREDIT FROM ECRITURE WHERE ' +
           ' E_EXERCICE = "' + EXO.Values[EXO.ItemIndex] + '" ' +
           ' AND E_DATECOMPTABLE >= " ' + USDateTime(EncodeDate(Year, Month, 1)) + '" ' +
           ' AND E_DATECOMPTABLE <= "' + USDateTime(FinDeMois(EncodeDate(Year, Month, 1))) + '" ' +
           ' AND E_JOURNAL = "' + JOURNAUX.Values[JOURNAUX.ItemIndex] + '" ' +
           ' AND E_QUALIFPIECE = "' + FTypeEcr + '"' ;

    lQuery := OpenSql(lSt, True);
    FMonTotalExo  := lQuery.FindField('LETOTAL').AsInteger;
    FMonDebitExo  := lQuery.FindField('LEDEBIT').AsFloat;
    FMonCreditExo := lQuery.FindField('LECREDIT').AsFloat;
    Ferme(lQuery);

    if (GS.Cells[SC_TYPEJAL, GS.Row]<>'-') and (GS.Cells[SC_TYPEJAL, GS.Row]<>'') then
      LanceSaisieFolio(Q, taModif)
    else
      TrouveEtLanceSaisie(Q, taModif, FTypeEcr) ;

    Ferme(Q) ;

    lQuery := OpenSql(lSt, True);
    FMonTotalExo  := FMonTotalExo  - lQuery.FindField('LETOTAL').AsInteger;
    FMonDebitExo  := FMonDebitExo  - lQuery.FindField('LEDEBIT').AsFloat;
    FMonCreditExo := FMonCreditExo - lQuery.FindField('LECREDIT').AsFloat;
    Ferme(lQuery);

    BChercheClick(nil) ;
    end ;
  end ;
end;

function TFCentral.GetMode : TModeVisu ;
var bAllMonth, bAllJal : Boolean ;
begin
  bAllMonth := (PERIODE.ItemIndex=0) ;
  bAllJal := (JOURNAUX.ItemIndex=0) ;
  if bAllMonth then
    Result:=tmvPeriode
  else
   if bAllJal then
     Result:=tmvJournal
   else
     Result:=tmvEcriture ;
end ;

procedure TFCentral.PutMontant(Montant : Double; Col, Row : LongInt) ;
begin
if Montant<>0 then GS.Cells[Col, Row]:=StrFMontant(Montant, 15, V_PGI.OkDecV,'',TRUE)
              else GS.Cells[Col, Row]:='' ;
end ;

function TFCentral.VerifCloture(sPeriode, sJournal: string) : Boolean ;
var Q : TQuery ;
    bExoN : Boolean ;
//    DateDeb : TDateTime ;
    Col, st : string ;
    Per : Integer ;
begin
Result:=FALSE ;
bExoN:=(EXO.Values[EXO.ItemIndex]=VH^.EnCours.Code) ;
if (not bExoN) and (EXO.Values[EXO.ItemIndex]<>VH^.Suivant.Code) then Exit ;
if bExoN then
  begin
  Col:='J_VALIDEEN' ;
//  DateDeb:=VH^.EnCours.Deb ;
  end else
  begin
  Col:='J_VALIDEEN1' ;
//  DateDeb:=VH^.Suivant.Deb ;
  end ;

  // GCO - 07/09/2004 - FQ 14339 et 14448 et 14383
  Per := PERIODE.Items.IndexOf(sPeriode);

Q:=OpenSQL('SELECT '+Col+' FROM JOURNAL WHERE J_JOURNAL="'+JOURNAUX.Values[JOURNAUX.Items.IndexOf(sJournal)]+'"', TRUE) ;
if not Q.EOF then
  begin
  st:=Q.Fields[0].AsString ;
  if (Length(St)<24) then St := St +  StringOfChar('-', 24-Length(St)); // 14710
  if st[Per]='X' then
  Result:=TRUE ;
end ;
Ferme(Q) ;
end ;

//=======================================================
//================= Fonctions TreeView ==================
//=======================================================
procedure TFCentral.VerifTreeView(sPeriode, sJournal : string) ;
var i, j : Integer ;
    bFind : Boolean ;
    pNode, Node : TTreeNode ;
begin
  bFind:=FALSE ;
  pNode:=nil ;
  for i:=0 to TreeView.Items.Count-1 do
  begin
    if FTriTW = TTPeriode then
    begin
      if TreeView.Items[i].Text=sPeriode then
      begin
        pNode:=TreeView.Items[i] ;
        for j:=0 to pNode.Count-1 do
          if pNode.Item[j].Text=sJournal then bFind:=TRUE ;
      end ;
    end
    else
    begin
      if TreeView.Items[i].Text=sJournal then
      begin
        pNode:=TreeView.Items[i] ;
        for j:=0 to pNode.Count-1 do
          if pNode.Item[j].Text=sPeriode then bFind:=TRUE ;
      end ;
    end;
  end;

  if (not bFind) and (pNode<>nil) then
  begin
    for i:=0 to pNode.Count-1 do
    begin
      if pNode.Item[i].Text=GetMessageRC(RC_NORECORD) then
        TreeView.Items.Delete(pNode.Item[i]) ;
    end;

    if FTriTw = TTPeriode then
      Node:=TreeView.Items.AddChild(pNode, sJournal)
    else
      Node:=TreeView.Items.AddChild(pNode, sPeriode);
    
    if VerifCloture(sPeriode, sJournal) then
      Node.ImageIndex := 0
    else
      Node.ImageIndex := pNode.ImageIndex ;

    Node.SelectedIndex:=Node.ImageIndex ;
  end ;
end ;

procedure TFCentral.SyncTreeView(Node : TTreeNode) ;
var NodeRef : TTreeNode ; i : Integer ;
begin
  NodeRef := nil ;

  if FTriTw = TTPeriode then
  begin
    //Période
    if (Node.Level=0) then
    begin
      NodeRef := Node ;
      JOURNAUX.ItemIndex := 0;
    end ;
  end
  else
  begin
    //Journal
    if (Node.Level=0) then
    begin
      NodeRef := Node ;
      PERIODE.ItemIndex := 0;
    end ;
  end;

  if (Node.Level=1) then
    NodeRef:=Node.Parent ;

  if FTriTw = TTPeriode then
  begin
    // Période
    for i:=0 to PERIODE.Items.Count-1 do
     begin
     if PERIODE.Items[i]=NodeRef.Text then
      begin
        PERIODE.ItemIndex:=i ;
        Break ;
      end ;
    end;

    // Journal
    if (Node.Level=1) then
    begin
      for i:=0 to JOURNAUX.Items.Count-1 do
      begin
        if JOURNAUX.Items[i]=Node.Text then
        begin
          JOURNAUX.ItemIndex:=i ;
          Break ;
        end ;
      end;
    end;
  end
  else
  begin
    // Période
    for i:=0 to JOURNAUX.Items.Count-1 do
    begin
      if JOURNAUX.Items[i]=NodeRef.Text then
      begin
        JOURNAUX.ItemIndex:=i ;
        Break ;
      end ;
    end;

    // Journal
    if (Node.Level=1) then
    begin
      for i:=0 to PERIODE.Items.Count-1 do
      begin
        if PERIODE.Items[i]=Node.Text then
        begin
          PERIODE.ItemIndex:=i ;
          Break ;
        end ;
      end;
    end;
  end;
end ;

function TFCentral.ExistPeriode(Per : string) : Boolean ;
var Mvt : PMVTSTATE ; i : Integer ;
begin
Result:=FALSE ;
for i:=0 to MvtState.Count-1 do
  begin
  Mvt:=MvtState.Items[i] ;
  if Mvt^.DebPeriode=StrToDate(Per) then begin Result:=TRUE ; Exit ; end ;
  end ;
end ;

function TFCentral.ExistJournal(Per, Jal : string) : Boolean ;
var Mvt : PMVTSTATE ; i : Integer ;
begin
Result:=FALSE ;
for i:=0 to MvtState.Count-1 do
  begin
  Mvt:=MvtState.Items[i] ;
  if (Mvt^.DebPeriode=StrToDate(Per)) and (Mvt^.Journal=Jal) then
    begin Result:=TRUE ; Exit ; end ;
  end ;
end ;

function TFCentral.ExistJournalSeul( Jal : string ) : Boolean;
var Mvt : PMVTSTATE ; i : Integer ;
begin
  Result := False;
  for i := 0 to MvtState.Count - 1 do
  begin
    Mvt := MvtState.Items[i] ;
    if Mvt^.Journal = Jal then
    begin
      Result:=TRUE ;
      Exit ;
    end ;
  end ;

end;



procedure TFCentral.InitView ;
var Q : TQuery ; Mvt : PMVTSTATE ; y, m : Word ; p : Integer ; Cols, Where : string ;
begin
if MvtState<>nil then
  begin
  while MvtState.Count>0 do
    begin Mvt:=MvtState.Items[0] ; Dispose(Mvt) ; MvtState.Delete(0) ; end ;
  MvtState.Free ;
  MvtState:=TList.Create ;
  end ;
Cols:='E_PERIODE, E_JOURNAL' ;
Where:=' E_EXERCICE="'+EXO.Values[EXO.ItemIndex]+'" AND E_QUALIFPIECE="'+FTypeEcr+'"' ;
if bFromFolio then Where:=Where+' AND E_MODESAISIE<>"-" AND E_MODESAISIE<>""' ;
Q:=OpenSQL('SELECT DISTINCT '+Cols+' FROM ECRITURE WHERE'+Where, TRUE) ;
while not Q.EOF do
  begin
  if (VH^.JalAutorises<>'') then
    if not (Pos(';'+Q.Fields[1].AsString+';', VH^.JalAutorises)>0) then
      begin Q.Next ; Continue ; end ;
  New(Mvt) ;
  p:=Q.Fields[0].AsInteger ; if p=0 then begin Q.Next ; Continue ; end ;
  y:=Trunc(Int(p/100)) ; m:=p-(y*100) ;
  Mvt^.DebPeriode:=EncodeDate(y, m, 1) ;
  Mvt^.Journal:=Q.Fields[1].AsString ;
  MvtState.Add(Mvt) ;
  Q.Next ;
  end ;
Ferme(Q) ;
end ;

procedure TFCentral.FillTreeView ;
var
  Node : TTreeNode ;
  i : Integer ;
begin
//VH^.DateCloturePer>0) and (FinDeMois(VH^.Entree.Fin)<=VH^.DateCloturePer

  if FTriTw = TTJal then
  begin
    for i:=1 to JOURNAUX.Items.Count-1 do
    begin
      if not ExistJournalSeul(JOURNAUX.Values[i]) then Continue ;
      Node:=TreeView.Items.Add(nil, JOURNAUX.Items[i]) ;

      //if (VH^.DateCloturePer>0) and
      //   (FinDeMois(StrToDate(PERIODE.Values[i]))<=VH^.DateCloturePer) then
        Node.ImageIndex := 1;
      //else
      //  Node.ImageIndex:=1 ;


      if FbExoClos then
        Node.ImageIndex := 3;

      Node.SelectedIndex:=Node.ImageIndex ;
      Node:=TreeView.Items.AddChild(Node, GetMessageRC(RC_NORECORD)) ;

      if FbExoClos then
        Node.ImageIndex:=3
      else
        Node.ImageIndex:=1;

      //Node.SelectedIndex := Node.ImageIndex;
      //for j:=0 to JOURNAUX.Items.Count-1 do
      //TreeView.Items.AddChild(Node, JOURNAUX.Items[j]) ;
    end ;
  end
  else
  begin
    for i:=1 to PERIODE.Items.Count-1 do
    begin
      if not ExistPeriode(PERIODE.Values[i]) then Continue ;
      Node:=TreeView.Items.Add(nil, PERIODE.Items[i]) ;
      if (VH^.DateCloturePer>0) and
         (FinDeMois(StrToDate(PERIODE.Values[i]))<=VH^.DateCloturePer) then
        Node.ImageIndex:=0
      else
        Node.ImageIndex:=1 ;

      if FbExoClos then Node.ImageIndex:=3 ;
      Node.SelectedIndex:=Node.ImageIndex ;
      Node:=TreeView.Items.AddChild(Node, GetMessageRC(RC_NORECORD)) ;
      if FbExoClos then Node.ImageIndex:=3 else Node.ImageIndex:=1 ;
      Node.SelectedIndex:=Node.ImageIndex ;
      //for j:=0 to JOURNAUX.Items.Count-1 do
      //TreeView.Items.AddChild(Node, JOURNAUX.Items[j]) ;
    end ;
  end;
end ;

(*
procedure TFCentral.FillTreeView(ModeVisu : TModeVisu; sPeriode, sJournal : string) ;
var i : Integer ; Node : TTreeNode ;
begin
case ModeVisu of
  tmvPeriode:
    begin
    Node:=TreeView.Items.Add(nil, sItem) ;
    TreeView.Items.AddChild(Node, '¿PFU?') ;
    end ;
  tmvJournal:
    begin
    for i:=0 to TreeView.Items.Count-1 do
      if TreeView.Items[i].Text=sPeriode then
        begin
        Node:=TreeView.Items[i] ;
        if Node.GetFirstNode.Text='¿PFU?' then
          begin
          Delete(Node.GetFirstNode) ;

          end ;
        end ;

    Node:=TreeView.Items.Add(nil, sItem) ;
    TreeView.Items.AddChild(Node, '¿PFU?') ;
    end ;
  end ;
end ;
*)

function TFCentral.PutPeriodeView(var Row : LongInt; Year, Month : Word;
                                  sPeriode, CodeJal, LibJal, NumPiece : string;
                                  ModeSaisie : string; ModeVisu : TModeVisu) : Boolean ;
var Q : TQuery ; Cols, Where, OrderBy, GroupBy : string ;
begin
OrderBy:='' ; GroupBy:='' ;
case ModeVisu of
  tmvPeriode:
    begin
    Cols:='COUNT(*), SUM(E_DEBIT), SUM(E_CREDIT)' ;
    Where:='E_DATECOMPTABLE>="'+USDateTime(EncodeDate(Year, Month, 1))+'"'
          +' AND E_DATECOMPTABLE<="'+USDateTime(FinDeMois(EncodeDate(Year, Month, 1)))+'"'
          +' AND E_QUALIFPIECE="'+FTypeEcr+'"' ;
    if CodeJal<>'' then Where:=Where+' AND E_JOURNAL="'+CodeJal+'"' ;
    // Modif GCO - 20/01/2003
    //if V_PGI.Confidentiel<>'1' then Where:=Where+' AND E_CONFIDENTIEL="0"' ;
    Where := Where + CWhereConfidentiel;

    if bFromFolio then Where:=Where+' AND E_MODESAISIE<>"-" AND E_MODESAISIE<>""' ;
    end ;
  tmvJournal:
    begin
    Cols:='COUNT(*), SUM(E_DEBIT), SUM(E_CREDIT)' ;
    Where:='E_JOURNAL="'+CodeJal+'"' ;
    if (Year<>0) and (Month<>0) then
    // GCO - 06/09/2004 - FQ 14361 -> Ajout du code Exo dans la requete, même si la
    // borne de date n'est pas bonne, l'ajout du code EXO suffit
      Where := Where + ' AND E_EXERCICE = "' + EXO.Values[EXO.ItemIndex] + '"'
                     + ' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(Year, Month, 1))+'"'
                     + ' AND E_DATECOMPTABLE<="'+USDateTime(FinDeMois(EncodeDate(Year, Month, 1)))+'"'
                     + ' AND E_QUALIFPIECE="'+FTypeEcr+'"' ;
    // Modif GCO - 20/01/2003
    //if V_PGI.Confidentiel<>'1' then Where:=Where+' AND E_CONFIDENTIEL="0"' ;
    Where := Where + CWhereConfidentiel;

    if bFromFolio then Where:=Where+' AND E_MODESAISIE<>"-" AND E_MODESAISIE<>""' ;
    end ;
  tmvEcriture:
    begin
    Cols:='COUNT(*), SUM(E_DEBIT), SUM(E_CREDIT)' ;
    Where := 'E_EXERCICE = "' + EXO.Values[EXO.ItemIndex] + '"' 
           + ' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(Year, Month, 1))+'"'
           + ' AND E_DATECOMPTABLE<="'+USDateTime(FinDeMois(EncodeDate(Year, Month, 1)))+'"'
           + ' AND E_JOURNAL="'+CodeJal+'"'
           + ' AND E_NUMEROPIECE='+NumPiece
           + ' AND E_QUALIFPIECE="'+FTypeEcr+'"' + cWhereConfidentiel;
    end ;
end ;
Result:=FALSE ;
Q:=OpenSQL('SELECT '+Cols+' FROM ECRITURE WHERE '+Where, TRUE) ;
while not Q.EOF do
  begin
  if Q.Fields[0].AsInteger=0 then begin Q.Next ; Continue ; end ;
  Result:=TRUE ;
  GS.RowCount:=GS.RowCount+1 ;
  case ModeVisu of
    tmvPeriode  : GS.Cells[SC_LIB, Row]:=sPeriode ;
    tmvJournal  : GS.Cells[SC_LIB, Row]:=LibJal ;
    tmvEcriture :
      begin
      if (ModeSaisie<>'-') and (ModeSaisie<>'') then
       begin
       GS.Cells[SC_LIB, Row]:=GetMessageRC(RC_BORDEREAU)+' '+NumPiece ;
       GS.Cells[SC_TYPEJAL, Row]:=ModeSaisie ;
       end else GS.Cells[SC_LIB, Row]:=GetMessageRC(RC_PIECE)+' '+NumPiece ;
      GS.Cells[SC_NUMPIECE, Row]:=NumPiece ;
      end ;
  end ;
  PutMontant(Q.Fields[1].AsFloat, SC_DEBIT,  Row) ;
  PutMontant(Q.Fields[2].AsFloat, SC_CREDIT, Row) ;
  GS.Cells[SC_LIGNES, Row]:=IntToStr(Q.Fields[0].AsInteger) ;
  Q.Next ; Row:=Row+1 ;
  end ;
Ferme(Q) ;
end ;

function TFCentral.StopSearch : Boolean ;
begin
Application.ProcessMessages ;
Result:=bStopSearch ;
end ;

procedure TFCentral.AfficheSolde(Row : LongInt) ;
var i : LongInt ; TotalDebit, TotalCredit : Double ; TotalLignes : LongInt ;
begin
TotalDebit:=0 ; TotalCredit:=0 ; TotalLignes:=0 ;
for i:=GS.FixedRows to GS.RowCount-1 do
  begin
  TotalDebit:=TotalDebit+Valeur(GS.Cells[SC_DEBIT, i]) ;
  TotalCredit:=TotalCredit+Valeur(GS.Cells[SC_CREDIT, i]) ;
  TotalLignes:=TotalLignes+StrToInt(GS.Cells[SC_LIGNES, i]) ;
  end ;

GS.RowCount:=GS.RowCount+1 ;

// GCO 18/07/2002
if Journaux.ItemIndex = 0 then
  GS.Cells[SC_LIB,    Row]:= '     ' + GetMessageRC(RC_TOTAL) + ' ' + Periode.Text
else
  GS.Cells[SC_LIB,    Row]:= '     ' + GetMessageRC(RC_TOTAL) + ' ' + Journaux.Text + '  ' + Periode.Text;
// FIN GCO
GS.Cells[SC_DEBIT,  Row]:=StrFMontant(TotalDebit,15,memPivot.Decimale,'',TRUE) ;
GS.Cells[SC_CREDIT, Row]:=StrFMontant(TotalCredit,15,memPivot.Decimale,'',TRUE) ;
GS.Cells[SC_LIGNES, Row]:=IntToStr(TotalLignes) ;
GS.RowCount:=GS.RowCount+1 ;
FRow:=Row ;
end ;

procedure TFCentral.AfficheTotalGeneral ;
begin
  if FRow=0 then Exit ;

  TCount  := TCount  - FMonTotalExo;
  TDebit  := TDebit  - FMonDebitExo;
  TCredit := TCredit - FMonCreditExo;

  GS.Cells[SC_LIB,    FRow+1]:= '     ' +GetMessageRC(RC_TOTAL) + ' ' + Exo.Text;
  GS.Cells[SC_DEBIT,  FRow+1]:=StrFMontant(TDebit,15,memPivot.Decimale,'',TRUE) ;
  GS.Cells[SC_CREDIT, FRow+1]:=StrFMontant(TCredit,15,memPivot.Decimale,'',TRUE) ;
  GS.Cells[SC_LIGNES, FRow+1]:=IntToStr(TCount) ;

  FMonTotalExo  := 0;
  FMonDebitExo  := 0;
  FMonCreditExo := 0;
end ;

// Fonctions des Click boutons
procedure TFCentral.SearchClick(bAddPrec : Boolean) ;
var
  Q: TQuery ;
  ModeVisu : TModeVisu ;
  //Nav : PNAVIG ;
  Year, Month, Day : Word ;
  i, Row : Integer ;
  Where : string ;
begin
bStopSearch:=FALSE ; FRow:=0 ;
(*
if (iLastExo=EXO.ItemIndex) and
   (iLastPeriode=PERIODE.ItemIndex) and
   (iLastJournaux=JOURNAUX.ItemIndex) then Exit ;
if (bAddPrec) and (iLastExo<>-1) and (iLastPeriode<>-1) and (iLastJournaux<>-1) then
  begin
  New(Nav) ;
  Nav^.Exercice:=iLastExo ;
  Nav^.Periode:=iLastPeriode ;
  Nav^.Journal:=iLastJournaux ;
  NavPrec.Add(Nav) ;
  BNivPrec.Enabled:=TRUE ;
  end ;
iLastExo:=EXO.ItemIndex ;
iLastPeriode:=PERIODE.ItemIndex ;
iLastJournaux:=JOURNAUX.ItemIndex ;
*)
  ModeVisu:=GetMode ;
  InitGrid ;
  Row:=GS.FixedRows ;

  if FTriTW = TTPeriode then
  begin
    case ModeVisu of
      tmvPeriode:
      begin
        for i:=1 to PERIODE.Values.Count-1 do
        begin
          DecodeDate(StrToDate(PERIODE.Values[i]), Year, Month, Day) ;
          if JOURNAUX.ItemIndex<>0
            then PutPeriodeView(Row, Year, Month, PERIODE.Items[i], JOURNAUX.Values[JOURNAUX.ItemIndex], JOURNAUX.Items[JOURNAUX.ItemIndex], '', '-', ModeVisu)
            else PutPeriodeView(Row, Year, Month, PERIODE.Items[i], '', '', '', '-', ModeVisu) ;
          if StopSearch then Break ;
        end ;

        if GS.Cells[SC_LIB, 1] = '' then
          GS.Cells[SC_LIB, Row]:= GetMessageRC(RC_NORECORD)
        else
          GS.RowCount := GS.RowCount - 1;
      end;

      tmvJournal:
      begin
        if PERIODE.ItemIndex<>0 then
          DecodeDate(StrToDate(PERIODE.Values[PERIODE.ItemIndex]), Year, Month, Day) ;

        for i:=1 to JOURNAUX.Values.Count-1 do
        begin
          if PERIODE.ItemIndex<>0 then
          begin
            if not ExistJournal(PERIODE.Values[PERIODE.ItemIndex], JOURNAUX.Values[i]) then Continue ;

            DecodeDate(StrToDate(PERIODE.Values[PERIODE.ItemIndex]), Year, Month, Day) ;

            if (PutPeriodeView(Row, Year, Month, '', JOURNAUX.Values[i], JOURNAUX.Items[i], '', '-', ModeVisu)) then
              VerifTreeView(PERIODE.Items[PERIODE.ItemIndex], JOURNAUX.Items[i]) ;
          end
          else
           PutPeriodeView(Row, 0, 0, '', JOURNAUX.Values[i], JOURNAUX.Items[i], '', '-', ModeVisu) ;

          if bStopSearch then Break ;
        end ;

        if GS.Cells[SC_LIB, 1] = '' then
          GS.Cells[SC_LIB, Row] := GetMessageRC(RC_NORECORD)
        else
          GS.RowCount := GS.RowCount-1 ;
      end ;

    tmvEcriture :
    begin
      if PERIODE.ItemIndex<>0 then
        DecodeDate(StrToDate(PERIODE.Values[PERIODE.ItemIndex]), Year, Month, Day) ;
      Where:='E_JOURNAL="'+JOURNAUX.Values[JOURNAUX.ItemIndex]+'"' ;
      Where:=Where+' AND E_EXERCICE = "' + EXO.Values[EXO.ItemIndex] + '"';
      Where:=Where+' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(Year, Month, 1))+'"' ;
      Where:=Where+' AND E_DATECOMPTABLE<="'+USDateTime(FinDeMois(EncodeDate(Year, Month, 1)))+'"' ;
      Where:=Where+' AND E_NUMLIGNE=1 AND E_NUMECHE<=1' ;
      Where:=Where+' AND E_QUALIFPIECE="'+FTypeEcr+'"' ;
      // Modif GCO - 20/01/2003
      //if V_PGI.Confidentiel<>'1' then Where:=Where+' AND E_CONFIDENTIEL="0"' ;
      Where := Where + CWhereConfidentiel;

      if bFromFolio then Where:=Where+' AND E_MODESAISIE<>"-" AND E_MODESAISIE<>""' ;
      Q:=OpenSQL('SELECT E_NUMEROPIECE, E_MODESAISIE FROM ECRITURE WHERE '+Where+' ORDER BY E_NUMEROPIECE', TRUE) ;
      while not Q.EOF do
        begin
        PutPeriodeView(Row, Year, Month, '', JOURNAUX.Values[JOURNAUX.ItemIndex], '',
                       Q.Fields[0].AsString, Q.Fields[1].AsString, ModeVisu) ;
        Q.Next ;
        if bStopSearch then Break ;
        end ;
      Ferme(Q) ;
      if GS.Cells[SC_LIB, 1]='' then GS.Cells[SC_LIB, Row]:=GetMessageRC(RC_NORECORD)
                                else GS.RowCount:=GS.RowCount-1 ;
    end;
  end;
  end
  else
  begin
    // TRI PAR JOURNAL
    case ModeVisu of
      tmvPeriode:
      begin
        for i := 1 to PERIODE.Values.Count - 1 do
        begin
          if not ExistPeriode(PERIODE.Values[i]) then Continue ;

          DecodeDate(StrToDate(PERIODE.Values[i]), Year, Month, Day) ;

          if (PutPeriodeView(Row, Year, Month, PERIODE.Items[i], JOURNAUX.Values[JOURNAUX.ItemIndex], JOURNAUX.Items[JOURNAUX.ItemIndex], '', '-', ModeVisu)) then
            VerifTreeView(PERIODE.Items[i], JOURNAUX.Items[JOURNAUX.ItemIndex]) ;

          if bStopSearch then Break ;
        end ;

        if GS.Cells[SC_LIB, 1]='' then
          GS.Cells[SC_LIB, Row] := GetMessageRC(RC_NORECORD)
        else
          GS.RowCount := GS.RowCount-1;
      end;

      tmvJournal:
      begin
        for i:=1 to JOURNAUX.Values.Count-1 do
        begin
          DecodeDate(StrToDate(PERIODE.Values[PERIODE.ItemIndex]), Year, Month, Day) ;
          if PERIODE.ItemIndex <> 0 then
            PutPeriodeView(Row, Year, Month, PERIODE.Items[i], JOURNAUX.Values[JOURNAUX.ItemIndex], JOURNAUX.Items[JOURNAUX.ItemIndex], '', '-', ModeVisu)
          else
            PutPeriodeView(Row, Year, Month, PERIODE.Items[i], '', '', '', '-', ModeVisu);

          if bStopSearch then Break ;
        end ;

        if GS.Cells[SC_LIB, 1] = '' then
          GS.Cells[SC_LIB, Row] := GetMessageRC(RC_NORECORD)
        else
          GS.RowCount := GS.RowCount-1 ;
      end ;

      tmvEcriture :
      begin
        if PERIODE.ItemIndex<>0 then
          DecodeDate(StrToDate(PERIODE.Values[PERIODE.ItemIndex]), Year, Month, Day) ;
        Where:='E_JOURNAL="'+JOURNAUX.Values[JOURNAUX.ItemIndex]+'"' ;
        Where:=Where+' AND E_EXERCICE = "' + EXO.Values[EXO.ItemIndex] + '"';
        Where:=Where+' AND E_DATECOMPTABLE>="'+USDateTime(EncodeDate(Year, Month, 1))+'"' ;
        Where:=Where+' AND E_DATECOMPTABLE<="'+USDateTime(FinDeMois(EncodeDate(Year, Month, 1)))+'"' ;
        Where:=Where+' AND E_NUMLIGNE=1 AND E_NUMECHE<=1' ;
        Where:=Where+' AND E_QUALIFPIECE="'+FTypeEcr+'"' ;
        // Modif GCO - 20/01/2003
        //if V_PGI.Confidentiel<>'1' then Where:=Where+' AND E_CONFIDENTIEL="0"' ;
        Where := Where + CWhereConfidentiel;

        if bFromFolio then Where:=Where+' AND E_MODESAISIE<>"-" AND E_MODESAISIE<>""' ;
        Q:=OpenSQL('SELECT E_NUMEROPIECE, E_MODESAISIE FROM ECRITURE WHERE '+Where+' ORDER BY E_NUMEROPIECE', TRUE) ;
        while not Q.EOF do
        begin
          PutPeriodeView(Row, Year, Month, '', JOURNAUX.Values[JOURNAUX.ItemIndex], '',
                       Q.Fields[0].AsString, Q.Fields[1].AsString, ModeVisu) ;
          Q.Next ;
          if bStopSearch then Break ;
        end ;
        Ferme(Q) ;
        if GS.Cells[SC_LIB, 1] = '' then
          GS.Cells[SC_LIB, Row]:=GetMessageRC(RC_NORECORD)
        else
          GS.RowCount:=GS.RowCount-1 ;
      end;
    end;
  end;

  bStopSearch:=FALSE ;
  if GS.Cells[SC_LIB, 1] <> GetMessageRC(RC_NORECORD) then
  begin
    AfficheSolde(Row) ;
    AfficheTotalGeneral ;
  end ;
end;

function TFCentral.ValClick : Boolean ;
var Year, Month, Day : Word ;
begin
Result:=FALSE ;
DecodeDate(StrToDate(PERIODE.Values[PERIODE.ItemIndex]), Year, Month, Day) ;
if bFromFolio then
  begin
  ParPeriode:=DateToStr(EncodeDate(Year, Month, 1)) ;
  ParCodeJal:=JOURNAUX.Values[JOURNAUX.ItemIndex] ;
  ParNumFolio:=GS.Cells[SC_NUMPIECE, GS.Row] ;
  bByParams:=TRUE ;
  Result:=TRUE ;
  end ;
end ;

procedure TFCentral.ByeClick ;
begin
  // GCO - 13/07/2007 - FQ 20782
  Close;
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
    THPanel(Self.parent).InsideForm := nil;
    THPanel(Self.parent).VideToolBar;
  end;
end ;

procedure TFCentral.InitClick ;
begin PERIODE.ItemIndex:=1 ; JOURNAUX.ItemIndex:=0 ; SearchClick ; end ;

// Click boutons
procedure TFCentral.BChercheClick(Sender: TObject);
begin
//EXOChange(nil) ;
SearchClick ;
end ;

procedure TFCentral.BValideClick(Sender: TObject);
begin if GetMode=tmvEcriture then if ValClick then ByeClick ; end ;

procedure TFCentral.BAbandonClick(Sender: TObject);
begin ByeClick ; end ;

procedure TFCentral.BInitClick(Sender: TObject);
begin InitClick ; end ;

procedure TFCentral.BNivPrecClick(Sender: TObject);
var Nav : PNAVIG ;
begin
if NavPrec.Count>0 then
  begin
  Nav:=NavPrec.Items[NavPrec.Count-1] ;
  EXO.ItemIndex:=Nav^.Exercice ;
  EXOChange(nil) ;
  PERIODE.ItemIndex:=Nav^.Periode ;
  JOURNAUX.ItemIndex:=Nav^.Journal ;
  Dispose(Nav) ; NavPrec.Delete(NavPrec.Count-1) ;
  if NavPrec.Count<=0 then BNivPrec.Enabled:=FALSE ;
  SearchClick(FALSE) ;
  end ;
end;

procedure TFCentral.BStopClick(Sender: TObject);
begin
bStopSearch:=TRUE ;
end;

procedure TFCentral.TreeViewChange(Sender: TObject; Node: TTreeNode);
begin
if bExpanding then Exit ;
SyncTreeView(Node) ;
SearchClick ;
end;

procedure TFCentral.TreeViewExpanding(Sender: TObject; Node: TTreeNode; var AllowExpansion: Boolean);
begin
if Node=TreeView.Selected then Exit ;
bExpanding:=TRUE ;
SyncTreeView(Node) ;
TreeView.Selected:=Node ;
SearchClick ;
bExpanding:=FALSE ;
end;

procedure TFCentral.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end ;

procedure TFCentral.Splitter2Moved(Sender: TObject);
begin
  HMTrad.ResizeGridColumns(GS);
end;



////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TFCentral.MenuParPeriodeClick(Sender: TObject);
begin
  ChangeTri(True);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/07/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCentral.MenuParJournalClick(Sender: TObject);
begin
  ChangeTri(False);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/07/2005
Modifié le ... : 07/09/2005
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFCentral.ChangeTri( vBoParPeriode : Boolean );
begin
  if vBoParPeriode then
    FTriTw := TTPeriode
  else
    FTriTw := TTJal;

  SetParamSoc('SO_CPJALCENTRALSURPERIODE', FTriTw = TTPeriode);

  MenuParPeriode.Checked := vBoParPeriode;
  MenuParJournal.Checked := not vBoParPeriode;
  ExoChange(nil);
end;

////////////////////////////////////////////////////////////////////////////////

end.

