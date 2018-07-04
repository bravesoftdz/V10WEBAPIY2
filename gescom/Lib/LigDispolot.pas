unit LigDispolot;


interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, Grids, Hctrls, ComCtrls, ExtCtrls, HPanel, StdCtrls, Mask, UIUtil,
  Hent1, hmsgbox, HSysMenu, TarifUtil, UTOB, math, AglInit, EntGC,SaisieSerie_TOF,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
{$ENDIF}
  Menus, Buttons, HPop97, SaisUtil, UtilPGI, StockUtil;
function Entree_LigDispolot(TobDispoLot, TobLigLot, TobLigne, TobSerie : TOB ; OkSerie : boolean ; FAction   : TActionFiche) : boolean;
function EntreeACH_LigDispolot(TobDispoLot, TobLigLot, TobLigne, TobSerie : TOB ; QteLigne : double ; bSerie : boolean ; FAction   : TActionFiche) : boolean;
procedure SupprimeLot(TobDispoLot, TobLigneLot : TOB ; Nature : string);   


Type R_IdentifieCol = RECORD
                  ColName : String ;
                  END ;

type
  TFLigDispolot = class(TForm)
    PTETE: THPanel;
    PPIED: THPanel;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    bAbandon: TToolbarButton97;
    bValider: TToolbarButton97;
    bAide: TToolbarButton97;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    TNReste: THLabel;
    TNQTESEL: THLabel;
    bRepartition: TToolbarButton97;
    bsolde: TToolbarButton97;
    GQL: THGrid;
    TARTICLE: THLabel;
    TDEPOT: THLabel;
    EArticle: THLabel;
    EDepot: THLabel;
    bsupprimer: TToolbarButton97;
    NTotalDoc: THNumEdit;
    NReste: THNumEdit;
    PCumul: THPanel;
    TCumul: TLabel;
    NTotDispo: THNumEdit;
    NTotSaisie: THNumEdit;
    MBTitres: THMsgBox;
    ///Gestion de la fiche
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    ///Evénements du grid
    procedure GQLRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GQLRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GQLCellEnter(Sender: TObject; var ACol, ARow: Integer;
          var Cancel: Boolean);
    procedure GQLCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GQLElipsisClick(Sender: TObject);
    procedure GQLDblClick(Sender: TObject);
    procedure GQLColumnWidthsChanged(Sender: TObject);
    Procedure GQLGetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState);
    ///Gestion des boutons
    procedure bValiderClick(Sender: TObject);
    procedure bsoldeClick(Sender: TObject);
    procedure bRepartitionClick(Sender: TObject);
    procedure bsupprimerClick(Sender: TObject);

  private
    { Déclarations privées }

    LesColGQL, ColsReadOnly : string ;
    stArt, stDep : string;
    iSensPiece : integer;
    IdentCols : Array[0..10] of R_IdentifieCol ;
    //Action   : TActionFiche ;

    //Initialisation de la form
    procedure EtudieColsListe ;
    procedure MAJPied ;
    function  PlusMoins(Text_Cell : string) : double;
    procedure InitGrid;

    //Validation des lignes
    function  ValideLaLigne(ARow : integer) : Boolean;
    function  CalculColQteDispo(ARow : integer ; var QteSaisie : double) : double;
    procedure CellExit(ACol,ARow : integer);
    procedure FormateZone(ACol ,ARow : integer);

    //Traitements sur les lignes
    procedure InitialiseLaLigne(ARow : integer);
    function  GQLSommeLot(Col : integer) : double;
    function  LigneVide (ARow : integer) : Boolean;
    procedure VideLaLigne (ARow : integer);
    procedure ZoomDateLot( ACol,ARow : integer ) ;

    //Traitements sur les tob
    procedure InitTobDesLots;
    procedure TobGetCellValues(ARow : integer);
    procedure EpureTobDesLots;
    procedure DispoLotToLigneLot;
    procedure RecupInfoPiece(TOBL : TOB);

    //Fonctions liées aux boutons
    procedure Repartition(Remise : double);
    procedure SupprimeLigne(ARow : integer);
    procedure GereElipsisSerie;
    procedure SupprimeLesSeries(Numerolot : string);

  public
    { Déclarations publiques }

  end;

var
  FLigDispolot: TFLigDispolot;
  TOBDesLots, TobLigneLot, TobLig, TobSer : TOB;
  stopTT, Abandon, GereSerie : boolean;
  ActionPiece : TActionFiche ;

  GQL_MOV  : integer ;
  GQL_NUM  : integer ;
  GQL_DAT  : integer ;
  GQL_QTE  : integer ;
  GQL_SAI  : integer ;

Const NbRowsPlus : integer = 20;

implementation

{$R *.DFM}

//=============================================================================
//============================== Points d'entrée ==============================
//=============================================================================

function Entree_LigDispolot(TobDispoLot, TobLigLot, TobLigne, TobSerie : TOB ; OkSerie : boolean ; FAction   : TActionFiche) : boolean;
var FoLigDispolot : TFLigDispolot;
    TobDuplicDispoLot,TobDuplicSerie: TOB;
begin
FoLigDispolot := TFLigDispolot.Create(Application);
Result := true; ActionPiece := FAction;
GereSerie := OkSerie;
TobDuplicSerie := nil;
if OkSerie then
   begin
   TobSer := TobSerie ;
   TobDuplicSerie := TOB.Create('',nil,-1);
   TobDuplicSerie.Dupliquer(TobSerie,true,true);
   end;
TobLig := TobLigne ;
TOBDesLots := TobDispoLot ; TobLigneLot := TobLigLot;
TobDuplicDispoLot:=TOB.Create('',nil,-1) ;
TobDuplicDispoLot.Dupliquer(TOBDesLots,true,true);
Try
   FoLigDispolot.ShowModal;
   Finally
   FoLigDispolot.free;
   end;
if  Abandon then
    begin
    TobDispoLot.Dupliquer(TobDuplicDispoLot,true,true);
    if OkSerie then TobSerie.Dupliquer(TobDuplicSerie,true,true);
    end;
TobDuplicDispoLot.Free ;
if OkSerie then TobDuplicSerie.Free;
if Valeur(TobLigneLot.GetValue('QUANTITE')) < TobLigne.GetValue('GL_QTEFACT') then Result := false;
end;

function EntreeACH_LigDispolot(TobDispoLot, TobLigLot, TobLigne, TobSerie : TOB ;
           QteLigne : double ; bSerie : boolean ; FAction   : TActionFiche) : boolean;
//var FoLigDispolot : TFLigDispolot;
//    TobDuplicDispoLot,TobDuplicSerie: TOB;
begin
	result := true;
end;

procedure SupprimeLot(TobDispoLot, TobLigneLot : TOB ; Nature : string);
var i_ind, iSens : integer;
    TobD, TobLL  : TOB;
begin
if TobLigneLot.Detail.Count <= 0 then exit;
if GetInfoParPiece(TobLigneLot.Detail[0].GetValue('GLL_NATUREPIECEG'),'GPP_ESTAVOIR')='X' then
   iSens := -1 else iSens := 1;
for i_ind:=0 to TobLigneLot.Detail.Count-1 do
  begin
  TobLL := TobLigneLot.Detail[i_ind];
  TobD:=TobDispoLot.FindFirst(['GQL_NUMEROLOT'],[TobLL.GetValue('GLL_NUMEROLOT')],true);
  if TobD = nil then continue;
  if Nature = 'VEN' then TobD.PutValue('GQL_PHYSIQUE',Valeur(TobD.GetValue('GQL_PHYSIQUE'))
                                     + Valeur(TobLL.GetValue('GLL_QUANTITE')*iSens))
  else if Nature = 'ACH' then
          TobD.PutValue('GQL_PHYSIQUE',Valeur(TobD.GetValue('GQL_PHYSIQUE'))
              - Valeur(TobLL.GetValue('GLL_QUANTITE'))*iSens);
  if TOBD.FieldExists('MODIFIABLE') then TobD.Free;
  end;
end;

//=============================================================================
//=========================== Gestion de la Form ==============================
//=============================================================================

procedure TFLigDispolot.FormShow(Sender: TObject);
var Naturepiece,VenteAchat : string;
begin
GQL.ListeParam:='GCDISPOLOT';
FillChar(IdentCols,Sizeof(IdentCols),#0) ;
EtudieColsListe ;
AffecteGrid (GQL, ActionPiece) ;
ColsReadOnly := 'GQL_NUMEROLOT;GQL_DATELOT';
Naturepiece := TobLig.getvalue('GL_NATUREPIECEG');
VenteAchat := GetInfoParPiece(Naturepiece,'GPP_VENTEACHAT');
if (Naturepiece='TEM') or (VenteAchat='VEN') or (Naturepiece='SEX') then
   begin
   Caption:=MBTitres.Mess[0];
   GQL.Cells[GQL_SAI,0]:=MBTitres.Mess[1];
   iSensPiece := -1;
   end else
if (Naturepiece='TRE') or (VenteAchat='ACH') or (Naturepiece='EEX') then
   begin
   Caption := MBTitres.Mess[2];
   GQL.Cells[GQL_SAI,0]:=MBTitres.Mess[3];
   brepartition.Visible := false;
   ColsReadOnly := ColsReadOnly + ';GQL_PHYSIQUE';
   iSensPiece := 1;
   end;

//Initialisation de l'entête
stArt := TOBDesLots.GetValue('GQ_ARTICLE');
stDep := TOBDesLots.GetValue('GQ_DEPOT');
EArticle.Caption:=Trim(Copy(stArt,1,18)) + '  ' + RechDom('GCARTICLE', stArt ,false);
EDepot.Caption:=RechDom('GCDEPOT', stDep ,false);

InitTobDesLots;
if TOBDesLots.Detail.Count > 0 then
   begin
   TOBDesLots.Detail.Sort('GQL_NUMEROLOT');
   TOBDesLots.PutGridDetail(GQL, false, false, LesColGQL);
   InitGrid;
   end else
   begin
   InitialiseLaLigne(1);
   end;

GQL.DBIndicator:=true;
GQL.OnColumnWidthsChanged := GQLColumnWidthsChanged;
GQL.GetCellCanvas := GQLGetCellCanvas;
GQL.Cells[GQL_MOV,0]:='' ; //GQL.ColWidths[GQL_MOV]:=13;
GQL.ColWidths[GQL_NUM]:=90 ; GQL.ColWidths[GQL_QTE]:=90;
GQL.ColWidths[GQL_DAT]:=90 ; GQL.ColWidths[GQL_SAI]:=110;
GQL.ColAligns[GQL_SAI] := taRightJustify;
//HMTrad.ResizeGridColumns (GQL) ;
GQL.ColLengths[GQL_QTE]:=-1;
GQL.Col := GQL_SAI;
//iSensPiece : incrémente ou décrémente le stock
if GetInfoParPiece(TobLigneLot.GetValue('NATUREPIECEG'),'GPP_ESTAVOIR')='X' then
   iSensPiece := iSensPiece * -1;

//Initialisation du pied
NTotDispo.Ctl3D := false ; NTotSaisie.Ctl3D := false;
NTotDispo.Height := PCumul.Height ; NTotSaisie.Height := PCumul.Height;
NTotDispo.Top := 0 ; NTotSaisie.Top := 0 ;
NTotalDoc.Value := TobLig.GetValue('GL_QTEFACT');
MAJPied;
Abandon := true;
end;

procedure TFLigDispolot.FormClose(Sender: TObject; var Action: TCloseAction);
var i_ind : integer;
    TobLD : TOB;
begin
if ActionPiece=taConsult then exit;
GQL.VidePile(False) ;
for i_ind := TobDesLots.Detail.Count-1 downto 0 do
    begin
    TobLD := TobDesLots.Detail[i_ind];
    if TobLD.GetValue('MODIFIABLE') then
       begin
       if TobLD.GetValue('GQL_PHYSIQUE') = 0 then TobLD.Free;
       end else TOBLD.DelChampSup('MODIFIABLE',false); //nécessaire pour fct InverseStockDesLots
    end;
if TobDesLots.Detail.Count > 0 then
    TOBDesLots.Detail[0].DelChampSup('QTESAISIE',True) ;
end;

procedure TFLigDispolot.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
   VK_RETURN : if Screen.ActiveControl=GQL then Key:=VK_TAB ;
   VK_DELETE :    BEGIN
                  if ((Screen.ActiveControl=GQL) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne (GQL.Row) ;
                    if GQL.Row > TobDesLots.Detail.Count then GQL.Row := GQL.Row - 1;
                    bSupprimer.Enabled := TobDesLots.Detail[GQL.Row-1].GetValue('MODIFIABLE');
                    MAJPied ;
                    END ;
                  END;
   VK_F10    : BEGIN Key:=0 ; BValiderClick(Nil) ; END ;
   VK_F5     : BEGIN
               if (Screen.ActiveControl=GQL) then
                  BEGIN
                  Key:=0 ; if GQL.ElipsisButton then GQLElipsisClick(nil);//ZoomDateLot(GQL.Col,GQL.Row) ;
                  END ;
               END;
   END ;
end;

//=============================================================================
//======================== Evénement sur le Grid ==============================
//=============================================================================

procedure TFLigDispolot.GQLRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
if ActionPiece=taConsult then exit;
if Ou>=GQL.RowCount-1 then GQL.RowCount:=GQL.RowCount+NbRowsPlus ;
if Ou <= TobDesLots.detail.count then
     bSupprimer.Enabled := TobDesLots.Detail[Ou-1].GetValue('MODIFIABLE');
if Ou = TobDesLots.detail.count + 1 then
    begin
    if LigneVide(Ou-1) then
       begin
       GQL.Row := TobDesLots.detail.count;
       GQL.Col := 1; exit;
       end;
    InitialiseLaLigne(Ou);
    end;
if Ou > TobDesLots.detail.count + 1 then
    begin
    GQL.Row := TobDesLots.detail.count + 1 ;
    GQL.Col := 1;
    end;
end;

procedure TFLigDispolot.GQLRowExit(Sender: TObject; Ou: Integer;
                                            var Cancel: Boolean; Chg: Boolean);
begin
if ActionPiece=taConsult then exit;
ValideLaLigne(Ou);
end;

procedure TFLigDispolot.GQLCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
if ActionPiece=taConsult then exit;
if pos(IdentCols[GQL.Col].ColName,ColsReadOnly) > 0 then
   if GQL.Row <= TobDesLots.Detail.Count then
      if not TobDesLots.Detail[GQL.Row-1].GetValue('MODIFIABLE') then
         begin
         GQL.Col := GQL_SAI;
         exit;
         end;
GQL.ElipsisHint := '';
GQL.ElipsisButton := (GQL.Col = GQL_DAT);
if GereSerie and (GQL.Col = GQL_SAI) then
   begin
   GQL.ElipsisButton := true;
   GQL.ElipsisHint := 'Affecter les numéros de série';
   end;
end;

procedure TFLigDispolot.GQLCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
if ActionPiece=taConsult then exit;
FormateZone(ACol,ARow);
CellExit(ACol,ARow);
end;

procedure TFLigDispolot.GQLElipsisClick(Sender: TObject);
begin
if GQL.Col = GQL_SAI then GereElipsisSerie
else if GQL.Col = GQL_DAT then ZoomDateLot(GQL.Col,GQL.Row);
end;

procedure TFLigDispolot.GQLDblClick(Sender: TObject);
begin
if (GQL.Col = GQL_DAT) and (GQL.ElipsisButton) then ZoomDateLot(GQL.Col,GQL.Row);
if (GQL.Col = GQL_SAI) and (GereSerie) then GereElipsisSerie;
end;

procedure TFLigDispolot.GQLColumnWidthsChanged(Sender: TObject);
var Coord : TRect;
begin
Coord:=GQL.CellRect(GQL_QTE,0);
NTotDispo.Left:=Coord.Left + 1;
NTotDispo.Width:=GQL.ColWidths[GQL_QTE] + 2;
Coord:=GQL.CellRect(GQL_SAI,0);
NTotSaisie.Left:=Coord.Left + 1;
NTotSaisie.Width:=GQL.ColWidths[GQL_SAI] + 1;
end;

Procedure TFLigDispolot.GQLGetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
if (ACol <> GQL_SAI) and (ARow > 0) and (ARow <= TobDesLots.detail.count) then
if (not TobDesLots.Detail[ARow-1].GetValue('MODIFIABLE')) and
      (pos(IdentCols[ACol].ColName,ColsReadOnly)>0) then
    Canvas.Font.Style := [fsItalic];
end;

//=============================================================================
//========================== Gestion des boutons ==============================
//=============================================================================

Procedure TFLigDispolot.bValiderClick(Sender: TObject);
var i_ind : integer;
    TobDL,TobSL : TOB;
    Numlot : string;
    QteSai : double;
begin
if ActionPiece=taConsult then exit;
CellExit(GQL.Col,GQL.Row);

if not ValideLaLigne(GQL.Row) then begin ModalResult := 0 ; exit ; end;

if Abs(NTotSaisie.Value) > Abs(NTotalDoc.Value) then
   begin MsgBox.Execute (0,Caption,'') ; ModalResult := 0 ; exit ; end;

if Abs(NTotSaisie.Value) < Abs(NTotalDoc.Value) then
   if MSGBox.Execute(3,'','')=mrno then begin ModalResult := 0 ; exit ; end;

for i_ind:=0 to TobDesLots.Detail.Count -1 do
    begin
    if TobDesLots.Detail[i_ind].GetValue('GQL_PHYSIQUE') < 0  then
       if MSGBox.Execute(10,'','')=mrno then
       begin
       ModalResult := 0 ;
       exit ;
       end else break;
   end;

if GereSerie then
    begin
    TOBSL := nil;
    i_ind := TobLig.GetValue('GL_INDICESERIE');
    if i_ind > 0 then
      if TOBSer.Detail.Count >= i_ind then TOBSL:=TOBSer.Detail[i_ind-1];
    for i_ind:=0 to TobDesLots.Detail.Count -1 do
        begin
        TobDL := TobDesLots.Detail[i_ind];
        Numlot := TobDL.GetValue('GQL_NUMEROLOT');
        if numLot = '' then continue;
        if TobSL <> nil then
             QteSai := TobSL.Somme('GLS_IDSERIE',['GLS_NUMEROLOT'],[NumLot],false,true)
        else QteSai := 0;
        QteSai := GetQteConvert(TobLig,nil,QteSai,true);
        if  QteSai <> TobDL.GetValue('QTESAISIE') then
            begin
            MSGBox.Execute(11,'','');
            ModalResult := 0 ;
            GQL.Row := i_ind + 1;
            exit ;
            end;
        end;
    end;
DispoLotToLigneLot;
EpureTobDesLots;
Abandon := false;
Close;
end;

procedure TFLigDispolot.bsoldeClick(Sender: TObject);
var QteDispo,QteSai : double;
begin
if ActionPiece=taConsult then exit;
if Abs(NTotSaisie.Value) >= Abs(NTotalDoc.Value) then exit;
QteSai := NReste.Value;
QteDispo := CalculColQteDispo(GQL.Row, QteSai);
if QteDispo < 0 then
   begin
   MSGBox.Execute(6,'','');
   GQL.Cells[GQL_SAI,GQL.Row] := StrF00(PlusMoins(GQL.Cells[GQL_QTE,GQL.Row]) +
                                        PlusMoins(GQL.Cells[GQL_SAI,GQL.Row]),V_PGI.OkDecV);
   GQL.Cells[GQL_QTE,GQL.Row] := StrS(0,V_PGI.OkDecV);
   end else
   begin
   GQL.Cells[GQL_SAI,GQL.Row] := StrF00(QteSai,V_PGI.OkDecV);
   GQL.Cells[GQL_QTE,GQL.Row] := StrS(QteDispo,V_PGI.OkDecV);
   end;
TobGetCellValues(GQL.Row);
MAJPied;
end;

procedure TFLigDispolot.bRepartitionClick(Sender: TObject);
begin
if ActionPiece=taConsult then exit;
if Abs(NTotSaisie.Value) >= Abs(NTotalDoc.Value) then exit;
StopTT := false;
Repartition(NReste.Value);
MAJPied;
if NReste.Value > 0 then MSGBox.Execute(5,'','');
end;

procedure TFLigDispolot.bsupprimerClick(Sender: TObject);
begin
SupprimeLigne(GQL.Row);
if GQL.Row > TobDesLots.Detail.Count then GQL.Row := GQL.Row - 1;
bSupprimer.Enabled := TobDesLots.Detail[GQL.Row-1].GetValue('MODIFIABLE');
if (not bsupprimer.Enabled) and
   (pos(IdentCols[GQL.Col].ColName,ColsReadOnly) > 0) then GQL.Col := GQL_SAI;
MAJPied;
end;


//=============================================================================
//====================== Initialisation de la form ============================
//=============================================================================

procedure TFLigDispolot.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp : integer ;
begin
LesCols:=GQL.Titres[0]; LesColGQL := LesCols + 'QTESAISIE';
icol:=0 ;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        begin
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then
           begin
           if NomCol='GQL_NUMEROLOT' then GQL_NUM:=icol else
           if NomCol='GQL_DATELOT' then GQL_DAT:=icol else
           if NomCol='GQL_PHYSIQUE' then GQL_QTE:=icol ;
           IdentCols[icol].ColName:=NomCol ;
           end ;
        end ;
    Inc(icol) ;
    Until ((LesCols='') or (NomCol='')) ;
GQL_MOV := 0;
GQL_SAI := GQL_QTE + 1;
GQL.ColCount := GQL.ColCount+1;
IdentCols[GQL_MOV].ColName:='FIXEDCOL';
IdentCols[GQL_SAI].ColName:='QTESAISIE';
//Formattage des colonnes
GQL.ColFormats[GQL_NUM]:='UPPER';
GQL.ColTypes[GQL_DAT]:='D' ;
GQL.ColFormats[GQL_DAT]:=ShortdateFormat ;
GQL.ColFormats[GQL_SAI]:='##.##';
GQL.ColTypes[GQL_SAI]:='R';
end ;

procedure TFLigDispolot.MAJPied ;
begin
NTotDispo.Value := GQLSommeLot(GQL_QTE);
NTotSaisie.Value := GQLSommeLot(GQL_SAI);
NReste.Value := NTotalDoc.Value - GQLSommeLot(GQL_SAI);
end;

function  TFLigDispolot.PlusMoins(Text_Cell : string) : double;
var iSens : integer;
begin
if NTotalDoc.Value < 0 then iSens := - 1 else iSens := 1;
result := Abs(Valeur(Text_Cell))*iSens;
end;

procedure TFLigDispolot.InitGrid;
var i_row : integer;
begin
for i_row:=1 to TobDesLots.Detail.Count do
    begin
    if GQL.Cells[GQL_DAT,i_row]='' then GQL.Cells[GQL_DAT,i_row]:=DateToStr(2);
    GQL.Cells[GQL_SAI,i_row]:=StrF00(Valeur(GQL.Cells[GQL_SAI,i_row]),V_PGI.OkDecV);
    end;
end;

//=============================================================================
//========================= Validation des lignes =============================
//=============================================================================

function TFLigDispolot.ValideLaLigne(ARow : integer) : Boolean;
var i_lot : integer;
begin
Result := True ;
if LigneVide(Arow) then exit;
if GQL.Cells[GQL_NUM,ARow] = '' then
    begin
    MSGBox.Execute(8,'','');
    Result := False ;
    GQL.Row := ARow ; GQL.Col := GQL_NUM;
    GQL.ElipsisButton := false;
    exit;
    end;

for i_lot := 0 to TobDesLots.Detail.Count -1 do
    begin
    if i_lot = ARow-1 then continue;
    if GQL.Cells[GQL_NUM,ARow] = TOBDesLots.Detail[i_lot].GetValue('GQL_NUMEROLOT') then
       begin
       MSGBox.Execute(9,'','');
       Result := False ;
       GQL.Row := ARow ; GQL.Col := GQL_NUM;
       GQL.ElipsisButton := false;
       exit;
       end;
    end;

TobGetCellValues(Arow);
end;

function TFLigDispolot.CalculColQteDispo(ARow : integer ; var QteSaisie : double) : double;
var QteDispo, NewQteSaisie, OldQteSaisie : double;
begin
OldQteSaisie := Valeur(TobDesLots.Detail[ARow-1].GetValue('QTESAISIE'));
NewQteSaisie := PlusMoins(GQL.Cells[GQL_SAI,ARow]) + QteSaisie;
QteDispo :=  valeur(GQL.Cells[GQL_QTE,ARow]) +
         (NewQteSaisie - OldQteSaisie)*iSensPiece;
Result := QteDispo;
QteSaisie := NewQteSaisie;
end;

procedure TFLigDispolot.CellExit(ACol,ARow : integer);
var QteSai : double;
    TobDL : TOB;
begin
if ACol <> GQL_SAI then exit;
QteSai := 0 ; TobDL := TobDesLots.Detail[ARow-1];
GQL.Cells[GQL_QTE,ARow] := StrS(CalculColQteDispo(ARow,QteSai),V_PGI.OkDecV);
TobDL.PutValue('QTESAISIE',QteSai);
if (QteSai = 0) and GereSerie then
    SupprimeLesSeries(TobDL.GetValue('GQL_NUMEROLOT'));
MAJPied;
end;

procedure TFLigDispolot.FormateZone(ACol ,ARow : integer);
begin
if ACol = GQL_QTE then
  GQL.Cells[ACol,ARow]:=StrS(Valeur(GQL.Cells[ACol,ARow]),V_PGI.OkDecV)
  else if ACol = GQL_SAI then
       GQL.Cells[ACol,ARow]:=StrF00(PlusMoins(GQL.Cells[ACol,ARow]),V_PGI.OkDecV)
       else if ACol = GQL_NUM then
            GQL.Cells[ACol,ARow] := Trim(GQL.Cells[ACol,ARow]);
end;

//=============================================================================
//=================== Traitements sur les lignes ==============================
//=============================================================================

procedure TFLigDispolot.InitialiseLaLigne(ARow : integer);
var TobLD : TOB;
begin
TobLD := TOB.Create('DISPOLOT',TOBDesLots,-1);
TobLD.InitValeurs();
TobLD.AddChampSupValeur('QTESAISIE',0,false);
TobLD.AddChampSupValeur('MODIFIABLE',true,false);
TobLD.PutValue('GQL_ARTICLE', stArt);
TobLD.PutValue('GQL_DEPOT', stDep);
TobLD.PutValue('GQL_NUMEROLOT', GQL.Cells[GQL_NUM,ARow]);
GQL.Cells[GQL_DAT,ARow] := DateToStr(2);
TobLD.PutValue('GQL_DATELOT', StrToDate(GQL.Cells[GQL_DAT,ARow]));

GQL.Cells[GQL_QTE,ARow] := StrS(0,V_PGI.OkDecV);
bSupprimer.Enabled := True;
end;

Function TFLigDispolot.GQLSommeLot(Col : integer) : double;
var i_row : integer;
    dTotal : double;
begin
dTotal := 0;
for i_row := 1 to TobDesLots.Detail.Count do
    dTotal := dTotal + Valeur(GQL.Cells[Col,i_row]);
Result := dToTal;
end;

function TFLigDispolot.LigneVide (ARow : integer) : Boolean;
begin
Result := (GQL.Cells[GQL_NUM, ARow] = '') and (Valeur(GQL.Cells[GQL_SAI, ARow]) = 0);
end;

procedure TFLigDispolot.VideLaLigne (ARow : integer);
begin
GQL.Cells[GQL_NUM, ARow] := '';
GQL.Cells[GQL_QTE, ARow] := StrS(0,V_PGI.OkDecV);
GQL.Cells[GQL_DAT, Arow] := DateToStr(idate1900);
GQL.Cells[GQL_SAI, ARow] := '';
end;

procedure TFLigDispolot.ZoomDateLot( ACol,ARow : integer ) ;
var HDATE : THCritMaskEdit;
    Coord : TRect;
begin
if ACol <> GQL_DAT then Exit ;
Coord:=GQL.CellRect (ACol, ARow);
HDATE := THCritMaskEdit.Create (GQL);
HDATE.Parent := GQL;
HDATE.Top:=Coord.Top ; HDATE.Left:=Coord.Left ; HDATE.Width:=3 ;
HDATE.Visible:=False ; HDATE.OpeType:=otDate ;
GetDateRecherche(TForm(HDATE.Owner),HDATE) ;
if HDATE.Text<>'' then GQL.Cells[GQL_DAT,ARow]:=HDATE.Text ;
HDATE.Free ;
end;

//=============================================================================
//====================== Traitements sur les tob ==============================
//=============================================================================

procedure TFLigDispolot.InitTobDesLots;
var i_ind : integer;
    TobLL, TobDL : TOB;
begin
if TobDesLots.Detail.Count > 0 then
   begin
   TOBDesLots.Detail[0].AddChampSupValeur('QTESAISIE',0,true);
   TOBDesLots.Detail[0].AddChampSupValeur('MODIFIABLE',false,true);
   if Gereserie then
   end;
if (TobLigneLot <> nil) and (TobLigneLot.Detail.Count > 0) then
   for i_ind:=0 to TobLigneLot.Detail.Count-1 do
       begin
       TobLL := TobLigneLot.Detail[i_ind];
       TobDL:=TobDesLots.FindFirst(['GQL_NUMEROLOT'],[TobLL.GetValue('GLL_NUMEROLOT')],true);
       if TobDL <> nil then TobDL.PutValue('QTESAISIE',TobLL.GetValue('GLL_QUANTITE'))
          else begin
               TobDL := TOB.Create('DISPOLOT',TOBDesLots,-1);
               TobDL.InitValeurs();
               TobDL.PutValue('GQL_ARTICLE', stArt); TobDL.PutValue('GQL_DEPOT', stDep);
               TobDL.PutValue('GQL_NUMEROLOT', TobLL.GetValue('GLL_NUMEROLOT'));
               if TobLL.FieldExists('DATELOT') then TobDL.PutValue('GQL_DATELOT', TobLL.GetValue('DATELOT'));
               TobDL.AddChampSupValeur('QTESAISIE',TobLL.GetValue('GLL_QUANTITE'),false);
               TobDL.AddChampSupValeur('MODIFIABLE',true,false);
               TobDL.PutValue('GQL_PHYSIQUE', TobLL.GetValue('GLL_QUANTITE')*iSensPiece);
               end;
       end;
end;

procedure TFLigDispolot.TobGetCellValues(ARow : integer);
var TobLD : TOB;
begin
if ARow > TobDesLots .Detail.Count then exit;
TobLD := TobdesLots.Detail[Arow-1];
TobLD.PutValue('GQL_NUMEROLOT',GQL.Cells[GQL_NUM,ARow]);
TobLD.PutValue('GQL_DATELOT',StrToDate(GQL.Cells[GQL_DAT,ARow]));
TobLD.PutValue('GQL_PHYSIQUE',Valeur(GQL.Cells[GQL_QTE,ARow]));
TobLD.PutValue('QTESAISIE',Valeur(GQL.Cells[GQL_SAI,ARow]));
end;

procedure TFLigDispolot.EpureTobDesLots;
var i_ind : integer;
    TobDL : TOB;
begin
for i_ind := TOBDesLots.Detail.Count-1 downto 0 do
   begin
   TobDL := TOBDesLots.Detail[i_ind];
   if TobDL.GetValue('MODIFIABLE') then TobDL.free;
   (*if VenteAchat = 'VENTE' then
      begin
      if TobDL.GetValue('MODIFIABLE') then TobDL.free;
      end else if VenteAchat = 'ACHAT' then
      begin
      if TobDL.GetValue('MODIFIABLE') and (Valeur(TobDL.GetValue('QTESAISIE'))=0) then
         TobDL.free;
      end;*)
  end;
end;

procedure TFLigDispolot.DispoLotToLigneLot;
var TOBL : TOB;
    i_ind : integer;
begin
TobLigneLot.ClearDetail;
for i_ind:=0 to TOBDesLots.Detail.count-1 do
    begin
    if Valeur(TOBDesLots.Detail[i_ind].GetValue('QTESAISIE')) <> 0 then
       begin
       TOBL:=TOB.Create('LIGNELOT',TobLigneLot,-1);
       RecupInfoPiece(TOBL);
       TOBL.PutValue('GLL_ARTICLE',stArt);
       TOBL.PutValue('GLL_NUMEROLOT',TOBDesLots.Detail[i_ind].GetValue('GQL_NUMEROLOT'));
       TOBL.PutValue('GLL_QUANTITE', TOBDesLots.Detail[i_ind].GetValue('QTESAISIE'));
       TOBL.AddChampSupValeur('DATELOT',TOBDesLots.Detail[i_ind].GetValue('GQL_DATELOT'));
       end;
    end;
TobLigneLot.PutValue('QUANTITE',GQLSommeLot(GQL_SAI));
end;

procedure TFLigDispolot.RecupInfoPiece(TOBL : TOB);
begin
TOBL.PutValue('GLL_NATUREPIECEG',TobLigneLot.GetValue('NATUREPIECEG'));
TOBL.PutValue('GLL_SOUCHE',TobLigneLot.GetValue('SOUCHE'));
TOBL.PutValue('GLL_NUMERO',TobLigneLot.GetValue('NUMERO'));
TOBL.PutValue('GLL_INDICEG',TobLigneLot.GetValue('INDICEG'));
TOBL.PutValue('GLL_NUMLIGNE',TobLigneLot.GetValue('NUMLIGNE'));
TOBL.PutValue('GLL_RANG',TobLigneLot.GetValue('RANG'));
end;

//=============================================================================
//=====================  Fonctions liées aux boutons ==========================
//=============================================================================

procedure TFLigDispolot.Repartition(Remise : double);
var i_ind, iRow, iNbLot : integer;
    Restant, RepartEnt : double;
    QteSai,QteDispo : double;
begin
RepartEnt:=0; iNbLot:=0;

for i_ind := 0 to TobDesLots.Detail.Count -1 do
    if TobDesLots.Detail[i_ind].GetValue('GQL_PHYSIQUE') > 0 then inc(iNbLot);

if iNbLot = 0 then begin StopTT:=true ; exit ; end;

for i_ind := iNbLot downto 1 do
    begin
    RepartEnt := int(Remise/i_ind);
    if Abs(RepartEnt) > 0 then break;
    end;

iNbLot := i_ind ; i_ind := 0;
Restant := Remise - (RepartEnt * iNbLot);
if RepartEnt = 0 then
   if Abs(Restant) > 0 then
      begin
      RepartEnt := Restant;
      Restant:=0;
      end else exit;

for iRow := 1 to TobDesLots.Detail.Count do
  begin
  if Valeur(GQL.Cells[GQL_QTE,iRow]) <= 0 then continue;
  inc(i_ind); //nb de lot concerné par la répartition
  if i_ind>iNbLot then break;
  QteSai := RepartEnt;
  QteDispo := CalculColQteDispo(iRow,QteSai);
  if QteDispo < 0 then
     begin
     Restant := Restant + PlusMoins(FloatToStr((Abs(RepartEnt) - valeur(GQL.Cells[GQL_QTE,iRow]))));
     QteSai := PlusMoins(GQL.Cells[GQL_QTE,iRow]) + PlusMoins(GQL.Cells[GQL_SAI,iRow]);
     QteDispo:=0;
     end;
  GQL.Cells[GQL_SAI,iRow] := StrF00(QteSai,V_PGI.OkDecV);
  GQL.Cells[GQL_QTE,iRow] := StrS(QteDispo,V_PGI.OkDecV);
  TobGetCellValues(irow);
  end;

if Abs(Restant) > 0 then Repartition(Restant);
if StopTT then exit;
end;

procedure TFLigDispolot.SupprimeLigne(ARow : integer);
var TobDL : TOB;
begin
if TobDesLots.Detail.Count = 0 then exit;
if Arow > TobDesLots.Detail.Count then exit;
TobDL := TobDesLots.Detail[Arow-1];
if GereSerie then SupprimeLesSeries(TobDL.GetValue('GQL_NUMEROLOT'));
if TobDesLots.Detail.Count = 1 then begin VideLaLigne(ARow) ; exit ; end;
if not TobDL.GetValue('MODIFIABLE') then exit;
TobDL.free;
GQL.CacheEdit ; GQL.SynEnabled := False;
GQL.DeleteRow (ARow);
GQL.MontreEdit; GQL.SynEnabled := True;
end;

procedure TFLigDispolot.GereElipsisSerie;
Var TobDL : TOB ;
    IndiceSerie : integer ;
    QteLot : Double ;
begin
if not GereSerie then Exit ;
FormateZone(GQL.Col,GQL.Row);
CellExit(GQL.Col,GQL.Row);
if not ValideLaLigne(GQL.Row) then exit;
QteLot := Valeur(GQL.Cells[GQL_SAI,GQL.Row]);
if QteLot = 0 then exit;
QteLot := GetQteConvert(TobLig,nil,QteLot);
IndiceSerie:=TobLig.GetValue('GL_INDICESERIE') ;
if IndiceSerie<=0 then
   BEGIN
   IndiceSerie:=TOBSer.Detail.Count+1 ;
   TobLig.PutValue('GL_INDICESERIE',IndiceSerie) ;
   END ;
TobDL := TobDesLots.Detail[GQL.Row-1];
TobLig.AddChampSupValeur('NUMEROLOT',TobDL.GetValue('GQL_NUMEROLOT'));
TobLig.AddChampSupValeur('QTELOT',QteLot);
Entree_SaisiSerie(TobLig,nil,TOBSer,nil,nil,ActionPiece) ;
TobLig.DelChampSup('NUMEROLOT',false);   TobLig.DelChampSup('QTELOT',false);
end;

procedure TFLigDispolot.SupprimeLesSeries(Numerolot : string);
var TobSerieLig, TobSL : TOB;
    i_ind,IndiceSerie : integer;
begin
TobSerieLig := nil;
IndiceSerie := TobLig.GetValue('GL_INDICESERIE');
if IndiceSerie > 0 then
    if TobSer.Detail.Count >= IndiceSerie then
        TobSerieLig := TobSer.Detail[IndiceSerie-1];
if TobSerieLig = nil then exit;
for i_ind := TobSerieLig.Detail.Count -1 downto 0 do
    begin
    TobSL := TobSerieLig.Detail[i_ind];
    if TOBSL.GetValue('GLS_NUMEROLOT') = Numerolot then TobSL.Free;
    end;
end;

end.

