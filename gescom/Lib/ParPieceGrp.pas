unit ParPieceGrp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFNDEF EAGLCLIENT}
  FichList, HSysMenu, hmsgbox, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Hqry, StdCtrls, HTB97,
  DBCtrls, ExtCtrls, HPanel, Grids, DBGrids, HDB, Hctrls, ComCtrls,
  HRichEdt, HRichOLE, HEnt1, UIUtil, M3FP, Mask, ADODB
{$ELSE}
	eFichList
{$ENDIF}  
  ;


Function EntreeParPieceGrp (NatPiece : string; Action : TActionFiche) : boolean ;


type
  TFParPieceGrp = class(TFFicheListe)
    PENTETE: THPanel;
    TGPR_NATUREPIECEG: THLabel;
    TGPR_NATPIECEGRP: THLabel;
    PGRILLE: THPanel;
    G_COND: THGrid;
    FComboTIE: THValComboBox;
    FComboART: THValComboBox;
    FComboLIG: THValComboBox;
    FComboPIE: THValComboBox;
    GPR_NATPIECEGRP: THDBValComboBox;
    GPR_CONDGRP: THDBRichEditOLE;
    PPIED: THPanel;
    FCTable: THValComboBox;
    FCChamp: THValComboBox;
    FCOpe: THValComboBox;
    GPR_LIBELLE: THDBEdit;
    GPR_NATUREPIECEG: THDBEdit;
    LIBNAT: THLabel;
    FCChp: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure G_CONDRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure G_CONDKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure FCTableChange(Sender: TObject);
    procedure GPR_NATPIECEGRPChange(Sender: TObject);
    procedure ChargeEnreg ; Override ;
    procedure FCChampChange(Sender: TObject);
    procedure FCOpeChange(Sender: TObject);
  private
    { Déclarations privées }
// Initialisation
    procedure InitialiseEntete ;
// Préparation du Grid
    procedure InitComboChamps ;
    procedure RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
    procedure ChargeComboChamps(NomTable : String) ;
    procedure ChargeComboChps(NomTable : String) ;
    procedure GetConditions ;
    function  ValueToItem(CC : THValComboBox ; St : String) : String ;
    procedure PutConditions ;
// Actions liés au Grid
    procedure EffaceGrid ;
    procedure EffaceCombo(CC : THValComboBox) ;
    function  ItemToValue(CC : THValComboBox ; St : String) : String ;
    procedure ChangeLigneGrid(Lig : Integer) ;
    procedure SupprimeLigne;
  public
    { Déclarations publiques }
    Action      : TActionFiche ;
    NaturePiece : string ;
    Chargement  : Boolean ;
  end;

var
  FParPieceGrp: TFParPieceGrp;

implementation

{$R *.DFM}

Function EntreeParPieceGrp (NatPiece : string; Action : TActionFiche) : boolean ;
var FF : TFParPieceGrp ;
    PPANEL  : THPanel ;
begin
SourisSablier;
FF := TFParPieceGrp.Create(Application) ;
FF.Action:=Action ;
FF.NaturePiece:=NatPiece ;
PPANEL := FindInsidePanel ;
if PPANEL = Nil then
   BEGIN
    try
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
Result:=True;
end;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}
procedure TFParPieceGrp.FormShow(Sender: TObject);
begin
TableName:='PARPIECEGRP' ;
UniqueName:= 'GPR_NATUREPIECEG,GPR_NATPIECEGRP' ;
CodeName:='GPR_NATPIECEGRP' ;
LibelleName:='GPR_LIBELLE' ;
NumCle:=0 ;
FRange:= NaturePiece ;
FAGL:=TRUE ;
TypeAction:=Action ;
InitComboChamps ;
InitialiseEntete;
  inherited;
end;

{==============================================================================================}
{=========================      Initialisation    =============================================}
{==============================================================================================}
procedure TFParPieceGrp.InitialiseEntete ;
BEGIN
//GPR_NATUREPIECEG.Text:=NaturePiece;
END;

procedure TFParPieceGrp.GPR_NATPIECEGRPChange(Sender: TObject);
Var F : TField ;
begin
  inherited;
F:=ta.FindField('GPR_LIBELLE');
F.Value:=GPR_NATPIECEGRP.Text;
end;

{==============================================================================================}
{=============================== Préparation du Grid ==========================================}
{==============================================================================================}
procedure TFParPieceGrp.InitComboChamps ;
begin
SourisSablier ;
FCChamp.Items.Clear ; FCChamp.Values.Clear ;
RemplitComboChamps('TIERS',FComboTIE) ;
RemplitComboChamps('ARTICLE',FComboART) ;
RemplitComboChamps('LIGNE',FComboLIG) ;
RemplitComboChamps('PIECE',FComboPIE) ;
SourisNormale ;
end ;

procedure TFParPieceGrp.RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
var Pref : String ;
    Q    : TQuery ;
begin
Pref := TableToPrefixe(NomTable) ;
Q:=OpenSQL('SELECT DH_NOMCHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE="'+
           Pref + '" AND DH_CONTROLE LIKE "%L%"',True,-1,'',true) ;
While not Q.EOF do
   begin
   if Trim(Q.FindField('DH_LIBELLE').AsString)='' then FCombo.Items.Add(Q.FindField('DH_NOMCHAMP').AsString)
                                                  else FCombo.Items.Add(Q.FindField('DH_LIBELLE').AsString) ;
   FCombo.Values.Add(Q.FindField('DH_NOMCHAMP').AsString) ;
   Q.Next ;
   end ;
Ferme(Q) ;
end ;

procedure TFParPieceGrp.ChargeComboChamps(NomTable : String) ;
var i_ind : Integer ;
    CC : THValComboBox ;
begin
CC := THValComboBox(FindComponent('FCombo'+Copy(NomTable,1,3))) ;
if CC = nil then exit ;
FCChamp.Clear ;
FCChamp.Items.Clear ;
FCChamp.Values.Clear;
for i_ind := 0 to CC.Items.Count - 1 do
   begin
   FCChamp.Items.Add (CC.Items[i_ind]) ;
   FCChamp.Values.Add(CC.Values[i_ind]) ;
   end ;
end ;

procedure TFParPieceGrp.ChargeComboChps(NomTable : String) ;
var i_ind : Integer ;
    CC : THValComboBox ;
begin
CC := THValComboBox(FindComponent('FCombo'+Copy(NomTable,1,3))) ;
if CC = nil then exit ;
FCChp.Clear ;
FCChp.Items.Clear ;
FCChp.Values.Clear;
for i_ind := 0 to CC.Items.Count - 1 do
   begin
   FCChp.Items.Add (CC.Items[i_ind]) ;
   FCChp.Values.Add(CC.Values[i_ind]) ;
   end ;
end ;

function TFParPieceGrp.ValueToItem(CC : THValComboBox ; St : String) : String ;
var i_ind : Integer ;
begin
i_ind := CC.Values.IndexOf(St) ;
if i_ind >= 0 then result := CC.Items[i_ind] else result := '' ;
end ;

{Charge dans le grid les conditions stockées dans le TMemoField}
procedure TFParPieceGrp.GetConditions ;
var i_ind   : Integer ;
    st,s1,s2 : String ;
begin
EffaceGrid ;
for i_ind := 0 to GPR_CONDGRP.Lines.Count-1 do
   begin
   st := GPR_CONDGRP.Lines[i_ind] ;
   s1 := ReadTokenSt(st) ; // table
   s2 := ValueToItem(FCTable,s1) ;
   ChargeComboChamps(s1) ; // Charge ponctuellement les champs de la table s1
   G_COND.Cells[0, i_ind + 1] := s2 ;
   s1 := ReadTokenSt(st) ; // champ
   s2 := ValueToItem(FCChamp,s1) ;
   G_COND.Cells[1, i_ind + 1] := s2 ;
   s1 := ReadTokenSt(st) ; // opérateur
   s2 := ValueToItem(FCOpe,s1) ;
   G_COND.Cells[2, i_ind + 1] := s2 ;
   s1 := ReadTokenSt(st) ; // valeur
   G_COND.Cells[3, i_ind + 1] := s1 ;
   // Appel de la première ligne
   G_COND.Row := 1 ;
   ChangeLigneGrid(G_COND.Row) ;
   end ;
end ;

procedure TFParPieceGrp.PutConditions ;
var i_ind    : Integer ;
    St_Lig, st : String ;
    F : TBlobField ;
begin
if Chargement then exit;
GPR_CONDGRP.Lines.Clear ;
for i_ind := 1 to G_COND.RowCount - 1 do
    begin
    if Trim(G_COND.Cells[0, i_ind]) = '' then continue ;
    st := ItemToValue(FCTable,G_COND.Cells[0,i_ind]) ;
    ChargeComboChps(st) ;
    St_Lig := st + ';'
              + ItemToValue (FCChp, G_COND.Cells[1, i_ind]) + ';'
              + ItemToValue (FCOpe, G_COND.Cells[2, i_ind]) + ';'
              + G_COND.Cells [3, i_ind] ;
    GPR_CONDGRP.Lines.Add (St_Lig) ;
    end ;
F:=TBlobField (ta.FindField('GPR_CONDGRP'));
F.Value:=GPR_CONDGRP.Text;
end ;

{==============================================================================================}
{=============================== Evènements du Grid ===========================================}
{==============================================================================================}
procedure TFParPieceGrp.G_CONDRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
  inherited;
if (G_COND.Row-1 >= 1) and (Trim(G_COND.Cells[0,G_COND.Row-1])='') then Cancel:=True ;
if not Cancel then ChangeLigneGrid(G_COND.Row) ;
end;

procedure TFParPieceGrp.G_CONDKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  inherited;
Case Key of
  VK_DELETE : SupprimeLigne ;
  end ;
end;

{==============================================================================================}
{=============================== Actions liés au Grid =========================================}
{==============================================================================================}
procedure TFParPieceGrp.EffaceGrid ;
var lig, col : Integer ;
begin
for lig := 1 to G_COND.RowCount - 1 do
   begin
   for col := 0 to G_COND.ColCount - 1 do G_COND.Cells[col, lig] := '' ;
   end ;
G_COND.Row := 1 ;
ChangeLigneGrid(G_COND.Row) ;
end ;

procedure TFParPieceGrp.EffaceCombo(CC : THValComboBox) ;
begin
CC.Items.Clear ;
CC.Values.Clear ;
end ;

function TFParPieceGrp.ItemToValue(CC : THValComboBox ; St : String) : String ;
var i_ind : Integer ;
begin
i_ind := CC.Items.IndexOf(St) ;
if i_ind >= 0 then result := CC.Values[i_ind] else result := '' ;
end ;

procedure TFParPieceGrp.ChangeLigneGrid(Lig : Integer) ;
begin
FCTable.Value := ItemToValue(FCTable,G_COND.Cells[0,Lig]) ;
FCTableChange(nil) ;
FCChamp.Value := ItemToValue(FCChamp,G_COND.Cells[1,Lig]) ;
FCOpe.Value := ItemToValue(FCOpe,G_COND.Cells[2,Lig]) ;
end ;

procedure TFParPieceGrp.SupprimeLigne;
var lig, i_ind, j_ind : Integer ;
begin
{Effacement d'une ligne}
lig := G_COND.Row ;
for i_ind := lig to G_COND.RowCount - 1 do
    begin
    for j_ind := 0 to G_COND.ColCount - 1 do
        G_COND.Cells[j_ind, i_ind] := G_COND.Cells[j_ind, i_ind + 1] ;
    end ;
ChangeLigneGrid (Lig) ;
PutConditions;
end;

{==============================================================================================}
{=============================== Evènements du Pied ===========================================}
{==============================================================================================}
procedure TFParPieceGrp.FCTableChange(Sender: TObject);
begin
if G_COND.Cells[0,G_COND.Row] = FCTable.Items[FCTable.ItemIndex] then exit ;
if Sta.State = dsBrowse then Sta.Edit ;
G_COND.Cells[0,G_COND.Row] := FCTable.Items[FCTable.ItemIndex] ;
if FCTable.Value = '' then EffaceCombo(FCChamp) else ChargeComboChamps(FCTable.Value) ;
PutConditions;
end;

procedure TFParPieceGrp.FCChampChange(Sender: TObject);
begin
if G_COND.Cells[1,G_COND.Row] = FCChamp.Items[FCChamp.ItemIndex] then exit ;
if Sta.State = dsBrowse then Sta.Edit ;
G_COND.Cells[1,G_COND.Row] := FCChamp.Items[FCChamp.ItemIndex] ;
PutConditions ;
end;

procedure TFParPieceGrp.FCOpeChange(Sender: TObject);
Var St : string ;
begin
//if sta.State = dsBrowse then exit ;
if G_COND.Cells[2,G_COND.Row] = FCOpe.Items[FCOpe.ItemIndex] then exit ;
St:=FCOpe.Items[FCOpe.ItemIndex] ;
if Sta.State = dsBrowse then Sta.Edit ;
G_COND.Cells[2,G_COND.Row]:=St ;
PutConditions ;
end;

{==============================================================================================}
{=============================== Gestion enregistrement =======================================}
{==============================================================================================}
procedure TFParPieceGrp.ChargeEnreg;
BEGIN
  inherited;
Chargement:=True;
GetConditions ;
Chargement:=False;
END;

{==============================================================================================}
{================================= Ponit d'entrée =============================================}
{==============================================================================================}
Procedure ParPieceGrp_SaisieGrp ( Parms : array of variant ; nb : integer ) ;
var StA,NatPiece : string ;
    i_ind  : integer ;
    Action : TActionFiche ;
    F : TForm;
BEGIN
Action:=taConsult ;
F := TForm (Longint (Parms[0]));
NatPiece := String (Parms[1]) ;
if NatPiece='' then exit;
StA:=String(Parms[2]) ;
if (F.Name = 'GCPARPIECE') then
    BEGIN
    i_ind:=Pos('ACTION=',StA) ;
    if i_ind>0 then
       BEGIN
       Delete(StA,1,i_ind+6) ; StA:=uppercase(ReadTokenSt(StA)) ;
       if StA='CREATION' then BEGIN Action:=taCreat ; END ;
       if StA='MODIFICATION' then BEGIN Action:=taModif ; END ;
       if StA='CONSULTATION' then BEGIN Action:=taConsult ; END ;
       END ;
    EntreeParPieceGrp (NatPiece, Action) ;
    END;
END ;

procedure InitParPieceGrp ();
begin
RegisterAglProc('ParPieceGrp_SaisieGrp', True , 2, ParPieceGrp_SaisieGrp);
end;

Initialization
InitParPieceGrp();
end.
