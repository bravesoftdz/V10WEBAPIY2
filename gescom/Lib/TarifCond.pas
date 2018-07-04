unit TarifCond;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, HPanel, Grids, Hctrls, HTB97, StdCtrls,
  ComCtrls,HRichEdt, HRichOLE,Hqry,
{$IFDEF EAGLCLIENT}
{$ELSE}
  DBCtrls,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  UIUtil, Hent1, UTOB, Mask,EntGc;

type
  TFTarifCond = class(TForm)
    PENTETE: THPanel;
    Dock971: TDock97;
    Toolbar972: TToolWindow97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    PPIED: THPanel;
    PGRILLE: THPanel;
    G_COND: THGrid;
    FTable: THValComboBox;
    FChamp: THValComboBox;
    FOpe: THValComboBox;
    FVal: TEdit;
    FComboTIE: THValComboBox;
    FComboART: THValComboBox;
    FComboLIG: THValComboBox;
    FComboPIE: THValComboBox;
    GF_CONDAPPLIC: THRichEditOLE;
    GF_LIBELLE: THCritMaskEdit;
    GF_DATEDEBUT: THCritMaskEdit;
    GF_DATEFIN: THCritMaskEdit;
    TGF_DATEDEBUT: THLabel;
    TGF_DATEFIN: THLabel;
    GF_DEPOT: THValComboBox;
    TGF_DEPOT: THLabel;
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FTableChange(Sender: TObject);
    procedure FChampChange(Sender: TObject);
    procedure FOpeChange(Sender: TObject);
    procedure FValChange(Sender: TObject);
    procedure G_CONDRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure G_CONDKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    TOBTarifOrig : TOB ;
    procedure InitComboChamps ;
    procedure RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
    procedure ChargeComboChamps(NomTable : String) ;
    procedure GetConditions ;
    procedure PutConditions ;
    procedure EffaceGrid ;
    procedure EffaceCombo(CC : THValComboBox) ;
    function  ValueToItem(CC : THValComboBox ; St : String) : String ;
    function  ItemToValue(CC : THValComboBox ; St : String) : String ;
    procedure ChangeLigneGrid(Lig : Integer) ;
    procedure SupprimeLigne;
  public
    { Déclarations publiques }
    Action : TActionFiche ;
    sTiersArticle : string;
  end;

  procedure EntreeTarifCond(Action : TActionFiche; var TOBTarf : TOB; sTiersArticle: string='TiersArticle'); 

var
  FTarifCond: TFTarifCond;

implementation

{$R *.DFM}

procedure EntreeTarifCond (Action : TActionFiche; var TOBTarf : TOB; sTiersArticle: string='TiersArticle');
var FF : TFTarifCond;
    PPANEL  : THPanel ;
BEGIN
SourisSablier;
FF := TFTarifCond.Create(Application) ;
FF.Action:=Action ;
FF.sTiersArticle:=sTiersArticle;
FF.TOBTarifOrig := TOBTarf;
PPANEL := FindInsidePanel ; // permet de savoir si la forme dépend d'un PANEL
if PPANEL = Nil then        // Le PANEL est le premier ecran affiché
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
END;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}
procedure TFTarifCond.FormShow(Sender: TObject);
var FF : TForm ;
begin
PENTETE.Visible := (sTiersArticle='TiersArticle');
FTable.Items.Clear ; FTable.Values.Clear;
if (pos('Tiers'  ,sTiersArticle)>0) then begin FTable.Items.Add('Pièce'  ); FTable.Values.Add('PIECE'  ); end;
if (pos('Tiers'  ,sTiersArticle)>0) then begin FTable.Items.Add('Tiers'  ); FTable.Values.Add('TIERS'  ); end;
if (pos('Article',sTiersArticle)>0) then begin FTable.Items.Add('Ligne'  ); FTable.Values.Add('LIGNE'  ); end;
if (pos('Article',sTiersArticle)>0) then begin FTable.Items.Add('Article'); FTable.Values.Add('ARTICLE'); end;     

InitComboChamps ;
FF := TForm (PENTETE.Owner);
TOBTarifOrig.PutEcran (FF, PENTETE) ;
GetConditions ;
G_COND.SetFocus ;
if not VH_GC.GCMultiDepots then
   begin    //mcd 25/08/03
   TGF_DEPOT.enabled := false;
   GF_DEPOT.enabled := false;
   TGF_DEPOT.Visible := false;
   GF_DEPOT.Visible:= false;
   end;
end;

procedure TFTarifCond.FormCreate(Sender: TObject);
begin
//
end;

procedure TFTarifCond.FormClose(Sender: TObject; var Action: TCloseAction);
begin
//
end;

{==============================================================================================}
{=============================== Préparation du Grid ==========================================}
{==============================================================================================}
procedure TFTarifCond.InitComboChamps ;
begin
SourisSablier ;
FChamp.Items.Clear ; FChamp.Values.Clear ;
if (pos('Tiers'  ,sTiersArticle)>0) then RemplitComboChamps('PIECE'  ,FComboPIE);
if (pos('Tiers'  ,sTiersArticle)>0) then RemplitComboChamps('TIERS'  ,FComboTIE);
if (pos('Article',sTiersArticle)>0) then RemplitComboChamps('LIGNE'  ,FComboLIG);
if (pos('Article',sTiersArticle)>0) then RemplitComboChamps('ARTICLE',FComboART);  
SourisNormale ;
end ;

procedure TFTarifCond.RemplitComboChamps(NomTable : String ; FCombo : THValComboBox) ;
begin
ExtractFields(NomTable,'T',FCombo.Items,FCombo.Values,Nil,False);
end ;

procedure TFTarifCond.ChargeComboChamps(NomTable : String) ;
var i_ind : Integer ;
    CC : THValComboBox ;
begin
CC := THValComboBox(FindComponent('FCombo'+Copy(NomTable,1,3))) ;
if CC = nil then exit ;
FChamp.Clear ;
FChamp.Items.Clear ;
FChamp.Values.Clear;
for i_ind := 0 to CC.Items.Count - 1 do
   begin
   FChamp.Items.Add (CC.Items[i_ind]) ;
   FChamp.Values.Add(CC.Values[i_ind]) ;
   end ;
end ;

function TFTarifCond.ValueToItem(CC : THValComboBox ; St : String) : String ;
var i_ind : Integer ;
begin
i_ind := CC.Values.IndexOf(St) ;
if i_ind >= 0 then result := CC.Items[i_ind] else result := '' ;
end ;

{Charge dans le grid les conditions stockées dans le TMemoField}
procedure TFTarifCond.GetConditions ;
var i_ind   : Integer ;
    st,s1,s2 : String ;
begin
EffaceGrid ;
GF_CONDAPPLIC.Text := TOBTarifOrig.GetValue ('GF_CONDAPPLIC');
for i_ind := 0 to GF_CONDAPPLIC.Lines.Count-1 do
   begin
   st := GF_CONDAPPLIC.Lines[i_ind] ;
   s1 := ReadTokenSt(st) ; // table
   s2 := ValueToItem(FTable,s1) ;
   If s2='Client' then s2 :='Tiers'; //mcd 26/08/03 .. traduction à tort du nom de la table dans la fiche
   ChargeComboChamps(s1) ; // Charge ponctuellement les champs de la table s1
   G_COND.Cells[0, i_ind + 1] := s2 ;
   s1 := ReadTokenSt(st) ; // champ
   s2 := ValueToItem(FChamp,s1) ;
   G_COND.Cells[1, i_ind + 1] := s2 ;
   s1 := ReadTokenSt(st) ; // opérateur
   s2 := ValueToItem(FOpe,s1) ;
   G_COND.Cells[2, i_ind + 1] := s2 ;
   s1 := ReadTokenSt(st) ; // valeur
   G_COND.Cells[3, i_ind + 1] := s1 ;
   // Appel de la première ligne
   G_COND.Row := 1 ;
   ChangeLigneGrid(G_COND.Row) ;
   end ;
end ;

{==============================================================================================}
{=============================== Evènements du Grid ===========================================}
{==============================================================================================}
procedure TFTarifCond.G_CONDRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
begin
if (G_COND.Row-1 >= 1) and (Trim(G_COND.Cells[0,G_COND.Row-1])='') then Cancel:=True ;
if not Cancel then ChangeLigneGrid(G_COND.Row) ;
end;

procedure TFTarifCond.G_CONDKeyDown(Sender: TObject; var Key: Word;
                                    Shift: TShiftState);
begin
Case Key of
  VK_DELETE : SupprimeLigne ;
  end ;
end;

{==============================================================================================}
{=============================== Actions liés au Grid =========================================}
{==============================================================================================}
procedure TFTarifCond.EffaceGrid ;
var lig, col : Integer ;
begin
for lig := 1 to G_COND.RowCount - 1 do
   begin
   for col := 0 to G_COND.ColCount - 1 do G_COND.Cells[col, lig] := '' ;
   end ;
G_COND.Row := 1 ;
ChangeLigneGrid(G_COND.Row) ;
end ;

procedure TFTarifCond.EffaceCombo(CC : THValComboBox) ;
begin
CC.Items.Clear ;
CC.Values.Clear ;
end ;

function TFTarifCond.ItemToValue(CC : THValComboBox ; St : String) : String ;
var i_ind : Integer ;
begin
i_ind := CC.Items.IndexOf(St) ;
if i_ind >= 0 then result := CC.Values[i_ind] else result := '' ;
end ;

procedure TFTarifCond.ChangeLigneGrid(Lig : Integer) ;
begin
FTable.Value := ItemToValue(FTable,G_COND.Cells[0,Lig]) ;
FTableChange(nil) ;
FChamp.Value := ItemToValue(FChamp,G_COND.Cells[1,Lig]) ;
FOpe.Value := ItemToValue(FOpe,G_COND.Cells[2,Lig]) ;
FVal.Text := G_COND.Cells[3,Lig] ;
end ;

procedure TFTarifCond.SupprimeLigne;
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
end;

{==============================================================================================}
{=============================== Evènements du Pied ===========================================}
{==============================================================================================}
procedure TFTarifCond.FTableChange(Sender: TObject);
begin
G_COND.Cells[0,G_COND.Row] := FTable.Items[FTable.ItemIndex] ;
if FTable.Value = '' then EffaceCombo(FChamp) else ChargeComboChamps(FTable.Value) ;
end;

procedure TFTarifCond.FChampChange(Sender: TObject);
begin
G_COND.Cells[1,G_COND.Row] := FChamp.Items[FChamp.ItemIndex] ;
end;

procedure TFTarifCond.FOpeChange(Sender: TObject);
begin
G_COND.Cells[2,G_COND.Row] := FOpe.Items[FOpe.ItemIndex] ;
end;

procedure TFTarifCond.FValChange(Sender: TObject);
begin
G_COND.Cells[3,G_COND.Row] := FVal.Text ;
end;

{==============================================================================================}
{=================================== Validation ===============================================}
{==============================================================================================}
procedure TFTarifCond.BValiderClick(Sender: TObject);
begin
PutConditions ;
end;

{Sauve dans le TMemoField les conditions stockées dans le grid}
procedure TFTarifCond.PutConditions ;
var i_ind    : Integer ;
    St_Lig, st : String ;
begin
GF_CONDAPPLIC.Lines.Clear ;
for i_ind := 1 to G_COND.RowCount - 1 do
    begin
    if Trim(G_COND.Cells[0, i_ind]) = '' then continue ;
    st := ItemToValue(FTable,G_COND.Cells[0,i_ind]) ;
    ChargeComboChamps(st) ;
    St_Lig := st + ';'
              + ItemToValue (FChamp, G_COND.Cells[1, i_ind]) + ';'
              + ItemToValue (FOpe, G_COND.Cells[2, i_ind]) + ';'
              + G_COND.Cells [3, i_ind] ;
    GF_CONDAPPLIC.Lines.Add (St_Lig) ;
    end ;
TOBTarifOrig.PutValue ('GF_CONDAPPLIC', GF_CONDAPPLIC.Text) ;
end ;

procedure TFTarifCond.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
