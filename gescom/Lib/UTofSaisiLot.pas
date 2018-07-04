unit UTofSaisiLot;

interface

uses  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, Grids, Vierge, Hctrls, ComCtrls, StdCtrls, Hent1, UTOB,UTOF, AglInit,
{$IFDEF EAGLCLIENT}
  MaineAgl,
{$ELSE}
  Fe_main,
{$ENDIF}
  SaisUtil, UtilPGI, NouveauLot, HMsgBox;

function Entree_SaisiLot(TobDispoLot : TOB) : boolean;

type
  TOF_SaisiLot = class(TOF)

   public
    { Déclarations publiques }

   iTableLigne : integer ;
   Action : TActionFiche ;

   procedure OnArgument (stArgument : String ) ; override ;
   procedure OnUpdate ; override ;
   procedure OnClose ; override;

  private
    { Déclarations privées }

    LesColGQL : string ;
    GQL : THGrid;
    bsolde, bRepartition,bsupprimer,bnouveau : TToolBarButton97;
    Total : THNumEdit;

    procedure RecupControl;

    //Gestion du grid
    procedure GQLRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GQLRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GQLColumnWidthsChanged(Sender: TObject);

    procedure EtudieColsListe ;
    procedure DessineCell ( ACol,ARow : Longint; Canvas : TCanvas ;
                                                        AState: TGridDrawState);

    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

    //Gestion des boutons
    procedure bsoldeClick(Sender: TObject);
    procedure bRepartitionClick(Sender: TObject);
    procedure bsupprimerClick(Sender: TObject);
    procedure bnouveauClick(Sender: TObject);

    //Validation des lignes
    function  TTLigneOK(ARow : integer) : Boolean;
    function  ValideQte (ARow, ACol : integer) : Boolean;

    //Traitements sur les lignes
    procedure ChangeDateVide;
    function  GQLSommeLot(Col : integer) : double;
    function  LigneVide (ARow : integer) : Boolean;
    procedure VideLigne (ARow : integer);

    function  FStr(Reel : double): string;
   end;


var
  TOBDesLots : TOB;
  QteDispo : double;
  Abandon : boolean;

  GQL_MOV  : integer ;
  GQL_NUM  : integer ;
  GQL_DAT  : integer ;
  GQL_QTE  : integer ;

Const

    // libellés des messages
	TexteMessage: array[1..3] of string 	= (
          {1}  'La somme des quantités saisies est supérieure à la quantité disponible',
          {2}  'Vous devez renseigner une quantité positive',
          {3}  'Vous devez renseigner une quantité correcte'
              );

implementation
uses CbpMCD
   ,CbpEnumerator;

function Entree_SaisiLot(TobDispoLot : TOB) : boolean;
var SomPhys : double;
    TobSauve : TOB;
begin
result := false;
TOBDesLots := TobDispoLot;
TobSauve:=TOB.Create('GCDISPOLOT',nil,-1);
TobSauve.Dupliquer(TobDispoLot,true,true);

AGLLanceFiche('GC','GCSAISILOT','','','');
if Abandon then TobDesLots.Dupliquer(TobSauve,true,true);
SomPhys := TobDesLots.Somme('GQL_PHYSIQUE',[''],[''],true);
if SomPhys = TobDesLots.GetValue('GZE_NEWQTE') then Result:=true;
TobSauve.Free;
end;


////////////////////////////////////////////////////////////////////////////////
//************************** Gestion de la fiche *****************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_SaisiLot.OnArgument (stArgument : String ) ;
var Tot : double;
    Cancel : boolean;
begin
Inherited;
Abandon:=True;
RecupControl;
Total:=THNumEdit(GetControl('TOTAL'));

GQL.PostDrawCell:=DessineCell;
iTableLigne := PrefixeToNum('GQL') ;
GQL.Tag:=1;
EtudieColsListe ;
GQL.ColFormats[GQL_DAT]:='';
AffecteGrid (GQL, action) ;

if TOBDesLots.Detail.Count>0 then
TOBDesLots.PutGridDetail(GQL, False, False, LesColGQL,true)
else GQL.RowCount:=2;
ChangeDateVide;

GQL.Cells[GQL_MOV,0]:='';
GQL.ColWidths[GQL_MOV]:=8;
GQL.ColLengths[GQL_NUM]:=-1;
GQL.ColLengths[GQL_DAT]:=-1;
GQL.ColAligns[GQL_NUM]:=taCenter;

SetControlText('HDisponible',Fstr(TobDesLots.GetValue('GZE_NEWQTE')));
Tot:= GQLSommeLot(GQL_QTE);
SetControlText('HReste',FStr(TobDesLots.GetValue('GZE_NEWQTE') - Tot));
Total.Text:=FStr(Tot);

GQL.Col := GQL_QTE;
Total.Masks.PositiveMask:=GQL.ColFormats[GQL_QTE];
TFVierge(Ecran).Hmtrad.ResizeGridColumns(GQL) ;

SetControlText('HArticle',RechDom('GCARTICLE', TobDesLots.GetValue('GZE_ARTICLE'), false));
SetControlText('HDepot',RechDom('GCDEPOT', TobDesLots.GetValue('GZE_DEPOT'), false));
GQLRowEnter(nil, 1, Cancel, False);
end;

Procedure TOF_SaisiLot.OnUpdate;
begin
Inherited;
if not TTLigneOK(GQL.Row) then
   begin
   Ecran.ModalResult := 0;
   Exit;
   end;

if Valeur(GetControlText('HReste')) < 0 then
   begin
   PGIBox('La somme des quantités saisies est supérieure à la quantité disponible',Ecran.Caption);
   //LastError:=1 ; LastErrorMsg:=TexteMessage[LastError];
   Ecran.ModalResult := 0;
   exit;
   end;

if valeur(GetControlText('HReste')) > 0 then
   if PGIAsk('La quantité totale n''est pas répartie intégralement#13 Voulez-vous continuer?',Ecran.Caption)=mrno then
      begin
      Ecran.ModalResult := 0;
      Exit;
      end;

if ((GQL.RowCount >= 2) and (not LigneVide(1))) then
   TOBDesLots.GetGridDetail(GQL,GQL.RowCount-1,'DISPOLOT',LesColGQL);
Abandon:=False;
end;

procedure TOF_SaisiLot.OnClose;
var i_ind : integer;
begin
Inherited;
GQL.VidePile(False) ;
for i_ind:=TobDesLots.Detail.Count-1 downto 0 do
    if TobDesLots.Detail[i_ind].GetValue('GQL_PHYSIQUE')<=0 then
         TobDesLots.Detail[i_ind].Free;
end;

procedure TOF_SaisiLot.RecupControl;
begin
bsolde:=TToolBarButton97(GetControl('BSOLDE'));
bsolde.OnClick:=bsoldeClick;
bRepartition:=TToolBarButton97(GetControl('BREPARTITION'));
bRepartition.OnClick:=bRepartitionClick;
bsupprimer:=TToolBarButton97(GetControl('BDELETE'));
bsupprimer.OnClick:=bsupprimerClick;
bnouveau:=TToolBarButton97(GetControl('BINSERT'));
bnouveau.OnClick:=bnouveauClick;
GQL:=THGrid(GetControl('GQL'));
GQL.OnRowExit:=GQLRowExit;
GQL.OnRowEnter:=GQLRowEnter;
GQL.OnColumnWidthsChanged:=GQLColumnWidthsChanged;
end;


////////////////////////////////////////////////////////////////////////////////
//**************************** Gestion du grid *******************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_SaisiLot.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp : integer ;
		Field     : IFieldCOM ;
		Mcd : IMCDServiceCOM;
		stTable : string;
begin
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
//
LesCols:=GQL.Titres[0] ; LesColGQL:=LesCols ; icol:=0 ;
stTable := mcd.PrefixetoTable('GQL');
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    begin
    if mcd.fieldExists(NomCol) then
       begin
       if NomCol='GQL_NUMEROLOT' then GQL_NUM:=icol else
        if NomCol='GQL_DATELOT' then GQL_DAT:=icol else
         if NomCol='GQL_PHYSIQUE' then GQL_QTE:=icol ;
       end ;
    end ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;
GQL_MOV :=0;
end;

procedure TOF_SaisiLot.DessineCell ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < GQL.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=GQL.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := GQL.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = GQL.row) then
         begin
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=((ARect.Left+ARect.Right) div 2) ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         end ;
    end;
end;

procedure TOF_SaisiLot.GQLRowExit(Sender: TObject; Ou: Integer;
                                            var Cancel: Boolean; Chg: Boolean);
begin
TTLigneOK(Ou);
GQL.InvalidateRow(Ou);
end;

procedure TOF_SaisiLot.GQLRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
GQL.InvalidateRow(Ou);
end;

procedure TOF_SaisiLot.GQLColumnWidthsChanged(Sender: TObject);
var Coord : TRect;
begin
Coord:=GQL.CellRect(GQL_QTE,0);
Total.Left:=Coord.Left + 1;
Total.Width:=GQL.ColWidths[GQL_QTE] + 1;
end;



//=============================================================================
//==============================  Gestion des champs ==========================
//=============================================================================


function TOF_SaisiLot.FStr(Reel : double): string;
begin
if Reel = 0 then Result := '0,00'
else Result := StrS0(Reel);
end;

procedure TOF_SaisiLot.ChangeDateVide;
var i_ind : integer;
begin
for i_ind:=0 to GQL.RowCount-1 do
    begin
    if GQL.Cells[GQL_NUM,i_ind]='' then continue;
    if GQL.Cells[GQL_DAT,i_ind]='' then GQL.Cells[GQL_DAT,i_ind]:='01/01/1900';
    end;
end;

function TOF_SaisiLot.ValideQte (ARow, ACol : integer) : Boolean;
begin
Result := true;
if not IsNumeric(GQL.Cells[ACol, ARow]) and (GQL.Cells[ACol, ARow]<>'') then
    begin
    GQL.Col := ACol; GQL.Row := ARow;
    PGIBox('Vous devez renseigner une quantité correcte',Ecran.Caption);
    //LastError:=3 ; LastErrorMsg:=TexteMessage[LastError];
    Result := False;
    end;
if Valeur(GQL.Cells[ACol, ARow]) = 0 then exit;
if Valeur(GQL.Cells[ACol, ARow]) < 0 then
    begin
    GQL.Col := ACol; GQL.Row := ARow;
    //LastError:=2 ; LastErrorMsg:=TexteMessage[LastError];
    PGIBox('Vous devez renseigner une quantité positive',Ecran.Caption);
    Result := False;
    exit;
    end;
end;



////////////////////////////////////////////////////////////////////////////////
//************************** Gestion des lignes ******************************//
////////////////////////////////////////////////////////////////////////////////

function TOF_SaisiLot.TTLigneOK(ARow : integer) : Boolean;
var Reste, Tot : double;
begin
Result := True;
if LigneVide(Arow) and (Valeur(GQL.Cells[GQL_QTE,Arow])<>0) then
   begin
   PGIBox('Vous devez ajouter un lot',Ecran.Caption);
   GQL.Row := ARow;
   Result := False ;
   exit ;
   end;

if not ValideQte(ARow,GQL_QTE) then
   begin
   GQL.Row := ARow;
   Result := False ;
   exit ;
   end;
GQL.Cells[GQL_QTE , ARow]:=FStr(Valeur(GQL.Cells[GQL_QTE , ARow]));

Tot:=GQLSommeLot(GQL_QTE);
Reste := TobDesLots.GetValue('GZE_NEWQTE') - Tot;
SetControlText('HReste',FStr(Reste));
Total.Text:=FStr(Tot);
end;


Function TOF_SaisiLot.GQLSommeLot(Col : integer) : double;
var i_ind : integer;
    Total : double;
begin
Total := 0;
for i_ind := 1 to GQL.RowCount do
    Total := Total + Valeur(GQL.Cells[Col,i_ind]);

Result := ToTal;
end;


function TOF_SaisiLot.LigneVide (ARow : integer) : Boolean;
begin
Result := True;
if GQL.Cells[GQL_NUM, ARow] <> '' then Result := False;
end;

procedure TOF_SaisiLot.VideLigne (ARow : integer);
begin
GQL.Cells[GQL_NUM, ARow] := '';
GQL.Cells[GQL_QTE, ARow] := '';
GQL.Cells[GQL_DAT, Arow] := '';
end;


//=============================================================================
//===========================  Gestion des boutons ============================
//=============================================================================


procedure TOF_SaisiLot.bsoldeClick(Sender: TObject);
var QteSai : double;
begin
if Valeur(GetControlText('HReste'))<=0 then exit;
QteSai := Valeur(GQL.Cells[GQL_QTE,GQL.Row]);

GQL.Cells[GQL_QTE,GQL.Row] := FStr(QteSai + valeur(GetControlText('HReste')));
SetControlText('HReste','0,00');
Total.Text:=FStr(GQLSommeLot(GQL_QTE));
end;

procedure TOF_SaisiLot.bRepartitionClick(Sender: TObject);
var  i_ind, NbLot : integer;
    Reste, Restant, RepartTot, RepartEnt : double;
begin
if Valeur(GetControlText('HReste'))<=0 then exit;
Reste := Valeur(GetControlText('HReste'));
NbLot := (GQL.RowCount-1);
RepartTot := Reste / NbLot;
RepartEnt := int(RepartTot);
Restant := RepartEnt + (Frac(RepartTot) * NbLot);
GQL.Cells[GQL_QTE,1] := FStr(Valeur(GQL.Cells[GQL_QTE,1])+Restant);
for i_ind := 2 to NbLot do
    GQL.Cells[GQL_QTE,i_ind] := FStr(Valeur(GQL.Cells[GQL_QTE,i_ind])+RepartEnt);
SetControlText('HReste',FStr(0));
Total.Text:=FStr(GQLSommeLot(GQL_QTE));
end;


procedure TOF_SaisiLot.bsupprimerClick(Sender: TObject);
var ARow : integer;
begin
ARow := GQL.Row;
if GQL.RowCount <= 1 then Exit ;
if GQL.RowCount = 2 then  VideLigne(ARow)
  else
  begin
  GQL.CacheEdit ; GQL.SynEnabled := False;
  GQL.DeleteRow (ARow);
  GQL.MontreEdit; GQL.SynEnabled := True;
   end;
if TobDesLots.Detail.Count>0 then TobDesLots.Detail[Arow-1].free;
SetControlText('HReste',FStr(TobDesLots.GetValue('GZE_NEWQTE') - GQLSommeLot(GQL_QTE)));
Total.Text:=FStr(GQLSommeLot(GQL_QTE));
end;


procedure TOF_SaisiLot.bnouveauClick(Sender: TObject);
var TobTemp : TOB;
begin
if TobDesLots.Detail.Count > 0 then
   if not TTLigneOK(GQL.Row) then exit;
if CreatLot(TOBDesLots,TobDesLots.GetValue('GZE_ARTICLE'),TobDesLots.GetValue('GZE_DEPOT')) then
   begin
   TobTemp:=TOBDesLots.FindFirst(['GQL_PHYSIQUE'],['0'],true);
   if TobTemp<>nil then
      TobTemp.PutValue('GQL_PHYSIQUE',TobTemp.GetValue('QTESAISIE'));
   GQL.ColFormats[GQL_DAT]:='';
   GQL.RowCount:=TOBDesLots.Detail.Count + 1 ;
   GQL.Row := TOBDesLots.detail.Count;
   TOBDesLots.Detail[TOBDesLots.Detail.Count-1].PutLigneGrid(GQL,GQL.RowCount-1,false,false,LesColGQL);
   if GQL.Cells[GQL_DAT,GQL.Row]='' then GQL.Cells[GQL_DAT,GQL.Row]:='01/01/1900';
   end;
end;

procedure TOF_SaisiLot.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
Case Key of
   VK_RETURN : if Screen.ActiveControl=GQL then Key:=VK_TAB ;
   END ;
end;


Initialization
registerclasses([TOF_SaisiLot]) ;

end.

