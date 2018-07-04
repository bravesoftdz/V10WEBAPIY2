unit UTofAjustStock;

interface

uses Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, Grids, Vierge, Hctrls, ComCtrls, StdCtrls, Hent1, UTOB,UTOF, AglInit,
{$IFDEF EAGLCLIENT}
  MaineAgl,UtileAGL,
{$ELSE}
  Fe_Main, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$IFDEF V530}
      EdtEtat,
  {$ELSE}
      EdtREtat,
  {$ENDIF}
{$ENDIF}
  SaisUtil, UtilPGI, HMsgBox, HDimension, HPanel,UTofSaisiLot,EntGC;


function ValideStkAjust(TobRecap : TOB ; Parametrage : string) : boolean ;

type
  TOF_AjustStock = class(TOF)

     public
    { Déclarations publiques }
    //Gestion de la fiche
    procedure OnArgument (stArgument : String ) ; override ;
    procedure OnLoad ; override;
    procedure OnUpdate ; override ;
    procedure OnClose ; override ;

    //Gestion du grid
    procedure G_RecapRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;
                                                                 Chg: Boolean);
    procedure G_RecapRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;
                                                                Chg: Boolean);
    procedure InitGrid;

    //Gestion des boutons
    procedure bZoomClick(Sender: TObject);
    procedure bLotClick(Sender: TObject);
    procedure bfermerClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);

  private
    { Déclarations privées }

    G_Recap : THGrid;
    bZoom, bLot, bferme, bimprimer : TToolBarButton97;
    LesColRec : string;

    //Gestion du grid
    procedure EtudieColsListe ;
    procedure DessineCell ( ACol,ARow : Longint; Canvas : TCanvas ;
                                                        AState: TGridDrawState);

    procedure GereLesDoublons;
    procedure GereLesLots;

    //Gestion des tobs
    procedure TobRecapInit(TOBT,TOBR,TOBL : TOB ; cpteur : integer);
    procedure InsereFilleTobRec(TOBR : TOB);
    procedure ChargeInfoArticles;
    function  RechercheDimensions(TOBArt : TOB) : string;

  end;

var
  Abandon : boolean;
  stParam : string;
  TobRec : TOB;
  GRE_MOV :    integer;
  GRE_LOT :    integer;
  GRE_ART :    integer;
  GRE_DEP :    integer;
  GRE_LIBQTE : integer;
  GRE_OLDV :   integer;
  GRE_NEWV :   integer;

Const  stValSel = 'X';
       StUnValSel = '-';

Const
	// libellés des messages
	TexteMessage: array[1..3] of string 	= (
          {1}  'Vous devez réajuster les quantités en stock des articles gérés par lot',
          {2}  '',
          {3}  ''
              );

implementation
uses CbpMCD
   ,CbpEnumerator;

function ValideStkAjust(TobRecap : TOB ; Parametrage : string) : boolean ;
begin
Abandon := true;
TobRec:=TobRecap;
stParam:=Parametrage;
V_PGI.ZoomOLE := true;  //Affichage en mode modal
AGLLanceFiche('GC','GCAJUSTSTOCK','','','');
V_PGI.ZoomOLE := false;
if not(Abandon) then Result:=true
else Result:=false;
end;

////////////////////////////////////////////////////////////////////////////////
//************************** Gestion de la fiche *****************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_AjustStock.OnArgument (stArgument : String ) ;
begin
Inherited;
G_Recap:=THGrid(GetControl('G_RECAP'));
G_Recap.OnRowEnter:=G_RecapRowEnter;
G_Recap.OnRowExit:=G_RecapRowExit;

bzoom:=TToolBarButton97(GetControl('BZOOM'));
bzoom.OnClick:=bzoomClick;
blot:=TToolBarButton97(GetControl('BLOT'));
blot.OnClick:=blotClick;
bferme:=TToolBarButton97(GetControl('BFERME'));
bferme.OnClick:=bfermerClick;
bimprimer:=TToolBarButton97(GetControl('BIMPRIMER'));
bimprimer.OnClick:=BImprimerClick;
end;

procedure TOF_AjustStock.OnLoad;
var Cancel : boolean;
begin
Inherited;
G_Recap.Tag:=1;
EtudieColsListe ;
TobRec.PutGridDetail(G_Recap,false,false,LesColRec,true);
ChargeInfoArticles;
G_Recap.PostDrawCell:=DessineCell;
GereLesDoublons;
GereLesLots;
InitGrid;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(G_Recap);
G_RecapRowEnter(nil, 1, Cancel, False);
end;


procedure TOF_AjustStock.OnUpdate;
var OkLot : boolean;
    i_ind : integer;
begin
Inherited;
OkLot:=true;
for i_ind:=1 to G_Recap.RowCount-1 do    //Recherche des articles gérés par lot
    if G_Recap.Cells[GRE_LOT,i_ind]=StUnValSel then //dont la somme des qtés est <> de la qté phys
       begin
       OkLot:=false;
       break;
       end;
if not OkLot then
   begin
   LastError:=1 ; LastErrorMsg:=TexteMessage[LastError];
   G_Recap.InvalidateRow(G_Recap.Row);
   G_Recap.Row:=i_ind;
   blot.Enabled:=true;
   Ecran.ModalResult:=0;
   end
   else if PGIAsk('Confirmez vous la mise à jour des stocks?',Ecran.Caption)=mrno then
   Ecran.ModalResult:=0 else abandon:=false;
end;

procedure TOF_AjustStock.OnClose ;
begin
inherited;
ExecuteSQL('DELETE FROM GCTMPAJUSTSTOCK WHERE GZE_UTILISATEUR = "'+V_PGI.USer+'"');
end;

////////////////////////////////////////////////////////////////////////////////
//**************************** Gestion du grid *******************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_AjustStock.G_RecapRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;
                                                                Chg: Boolean);
begin
G_Recap.InvalidateRow(Ou);
end;


procedure TOF_AjustStock.G_RecapRowEnter(Sender: TObject; Ou: Integer;
  var Cancel: Boolean; Chg: Boolean);
var stDim : string;
begin
G_Recap.InvalidateRow(Ou);

blot.Enabled:= Not (TobRec.Detail[Ou-1].GetValue('LOT')='-') ;

SetControlText('TARTICLE','Article : '+TobRec.Detail[Ou-1].GetValue('LIBELLE'));

stDim:='';
stDim:=TobRec.Detail[Ou-1].GetValue('GZE_DIM');
if stDim='' then SetControlText('TCODEDIM','')
else  begin
      stDim:='Dimension(s) : ' + stDim;
      SetControlText('TCODEDIM',stDim);
      end;
end;


procedure TOF_AjustStock.InitGrid;
var i_ind : integer;
begin
for i_ind:=1 to G_Recap.RowCount-1 do
    begin
    if G_Recap.Cells[GRE_ART,i_ind]<>'' then
    G_Recap.Cells[GRE_ART,i_ind]:=copy(G_Recap.Cells[GRE_ART,i_ind],0,18);
    end;

G_Recap.Cells[GRE_LOT,0]:='Lot';
G_Recap.ColAligns[GRE_LOT]:=taCenter;
G_Recap.ColWidths[GRE_LOT]:=40;
G_Recap.Cells[GRE_MOV,0]:='';
G_Recap.ColWidths[GRE_MOV]:=20;
end;

procedure TOF_AjustStock.EtudieColsListe ;
Var NomCol,LesCols : String ;
    ichamp, iTableLigne, icol : integer ;
		Mcd : IMCDServiceCOM;
begin
MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
//
LesCols:=G_Recap.Titres[0] ; LesColRec:=LesCols+'LOT;' ; icol:=0 ;
iTableLigne := PrefixeToNum('GZE') ;
Repeat
 NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
 if NomCol<>'' then
    begin
    if mcd.fieldExists(NomCol) then
       begin
       if NomCol='GZE_ARTICLE' then GRE_ART:=icol else
       if NomCol='GZE_DEPOT'   then GRE_DEP:=icol else
       if NomCol='GZE_RUBRIQUE'   then GRE_LIBQTE:=icol else
       if NomCol='GZE_OLDQTE'  then GRE_OLDV:=icol else
       if NomCol='GZE_NEWQTE'  then GRE_NEWV:=icol;
       end ;
    end ;
 Inc(icol) ;
Until ((LesCols='') or (NomCol='')) ;
GRE_MOV:=0;
GRE_LOT:=icol;
end;

procedure TOF_AjustStock.GereLesDoublons;
var i_ind : integer;
    stArt, stDep : string;
begin
stArt:='';
stDep:='';
for i_ind:=1 to G_Recap.RowCount do
    begin
    if (G_Recap.Cells[GRE_ART,i_ind]=stArt) and (G_Recap.Cells[GRE_DEP,i_ind]=stDep) then
       begin
       G_Recap.Cells[GRE_ART,i_ind]:='';
       G_Recap.Cells[GRE_DEP,i_ind]:='';
       end else
       begin
       stArt:=G_Recap.Cells[GRE_ART,i_ind];
       stDep:=G_Recap.Cells[GRE_DEP,i_ind];
       end;
    end;
end;

procedure TOF_AjustStock.GereLesLots;
var i_ind : integer;
begin
for i_ind:=1 to G_Recap.RowCount-1 do
    begin
    if G_Recap.Cells[GRE_LOT,i_ind]='X' then G_Recap.Cells[GRE_LOT,i_ind]:=StUnValSel
    else G_Recap.Cells[GRE_LOT,i_ind] := '';
    end;
end;

procedure TOF_AjustStock.DessineCell ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
BEGIN
If Arow < G_Recap.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=G_Recap.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := G_Recap.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = G_Recap.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=((ARect.Left+ARect.Right) div 2) ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;


////////////////////////////////////////////////////////////////////////////////
//************************** Gestion des boutons *****************************//
////////////////////////////////////////////////////////////////////////////////


procedure TOF_AjustStock.bZoomClick(Sender: TObject);
var stArt, stDep : string;
begin
stArt:=TobRec.Detail[G_Recap.Row-1].GetValue('GZE_ARTICLE');
stDep:=TobRec.Detail[G_Recap.Row-1].GetValue('GZE_DEPOT');
//AGLLanceFiche('GC','GCDISPO','',stArt+';'+ stDep+';'
//                                  +'-'+';'+ '01/01/1900','ACTION=CONSULTATION');
AGLLanceFiche('BTP','BTDISPO','',stArt+';'+ stDep+';'
                                  +'-'+';'+ '01/01/1900','ACTION=CONSULTATION');

end;

procedure TOF_AjustStock.bfermerClick(Sender: TObject);
begin
if PGIAsk('Voulez-vous abandonner le traitement?',Ecran.Caption)=mrno then
   Ecran.ModalResult:=0;
end;

procedure TOF_AjustStock.bLotClick(Sender: TObject);
var TobR : TOB;
begin
TobR:=TobRec.Detail[G_Recap.Row-1];
if TobR.GetValue('LOT') <> 'X' then exit;
///if Entree_SaisiLot(TobR,TobR.GetValue('GZE_ARTICLE'),TobR.GetValue('GZE_DEPOT'),
   ///                          TobR.GetValue('GZE_NEWQTE')) then
    /// G_Recap.Cells[GRE_LOT,G_Recap.Row]:=stValSel
if Entree_SaisiLot(TobR) then G_Recap.Cells[GRE_LOT,G_Recap.Row]:=stValSel
else G_Recap.Cells[GRE_LOT,G_Recap.Row]:=stUnValSel;
end;

procedure TOF_AjustStock.BImprimerClick(Sender: TObject);
var i_ind1,i_ind2,i_cpte : integer;
    TOBR,TOBT,TobLot,TobTemp : TOB;
    stWhere : string;
begin
if TobRec.Detail.Count<=0 then exit;
ExecuteSQL('DELETE FROM GCTMPAJUSTSTOCK WHERE GZE_UTILISATEUR = "'+V_PGI.USer+'"');
stWhere:='GZE_UTILISATEUR="'+V_PGI.USer+'"';
i_cpte:=0;
for i_ind1:=0 to TobRec.Detail.Count-1 do
    begin
    if TobRec.Detail[i_ind1].GetValue('GZE_COMPTEUR') > i_cpte then
       i_cpte:=TobRec.Detail[i_ind1].GetValue('GZE_COMPTEUR');
    end;

for i_ind1:=0 to TobRec.Detail.Count-1 do    //Recherche lots
    begin
    TOBR:=TobRec.Detail[i_ind1];
    InsereFilleTobRec(TOBR);
    if TOBR.Detail.Count>0 then    //Si lots alors ajout à tobrecap pour insertion dans
       begin                       //la table
       TobLot:=TOBR.Detail[0];
       TobTemp:=TOB.Create('GCTMPAJUSTSTOCK',nil,-1);
       if G_Recap.Cells[GRE_LOT,i_ind1+1]=StUnValSel then
          begin
          inc(i_cpte);
          TobRecapInit(TobTemp,TOBR,TobLot,i_cpte);
          TobTemp.PutValue('GZE_NUMEROLOT','QTESLOTNONVALIDE');
          TobTemp.PutValue('GZE_NEWQTE',0);
          end else
          begin
          for i_ind2:=0 to TobLot.Detail.Count-1 do
            begin
            TOBT:=TOB.Create('GCTMPAJUSTSTOCK',TobTemp,-1);
            inc(i_cpte);
            TobRecapInit(TOBT,TOBR,TobLot.Detail[i_ind2],i_cpte);
            TOBT.PutValue('GZE_NEWQTE',TobLot.Detail[i_ind2].GetValue('GQL_PHYSIQUE'));
            end;
          end;
       TobTemp.InsertOrUpdateDB();
       TobTemp.Free;
       end;
    end;
LanceEtat('E','GZE','GZE',true,false,false,nil,stWhere,'',false,0,'XX_VARIABLE1='+stParam);
end;

////////////////////////////////////////////////////////////////////////////////
//**************************** Gestion des tobs*******************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_AjustStock.TobRecapInit(TOBT,TOBR,TOBL : TOB ; cpteur : integer);
begin
TOBT.PutValue('GZE_UTILISATEUR',V_PGI.USer);
TOBT.PutValue('GZE_COMPTEUR',cpteur);
TOBT.PutValue('GZE_ARTICLE',TOBR.GetValue('GZE_ARTICLE'));
TOBT.PutValue('GZE_DIM','');
TOBT.PutValue('GZE_DEPOT',TOBR.GetValue('GZE_DEPOT'));
TOBT.PutValue('GZE_RUBRIQUE','');
TOBT.PutValue('GZE_OLDQTE',0);
TOBT.PutValue('GZE_NUMEROLOT',TOBL.GetValue('GQL_NUMEROLOT'));
end;

procedure TOF_AjustStock.InsereFilleTobRec(TOBR : TOB);
var TobTemp : TOB;
begin
TobTemp:=TOB.Create('GCTMPAJUSTSTOCK',nil,-1);
TobTemp.Dupliquer(TOBR,false,true,true);
TobTemp.InsertOrUpdateDB();
TobTemp.Free;
end;

procedure TOF_AjustStock.ChargeInfoArticles;
var i_ind : integer;
    TobArticle : TOB;
begin
TobRec.Detail[0].AddChampSup('LIBELLE',true);
TobArticle:=TOB.Create('ARTICLE',nil,-1);
for i_ind:=0 to TobRec.Detail.Count-1 do
    begin
    TobArticle.SelectDB('"' + TobRec.Detail[i_ind].GetValue('GZE_ARTICLE') + '"', nil);
    TobRec.Detail[i_ind].PutValue('GZE_DIM',RechercheDimensions(TobArticle));
    TobRec.Detail[i_ind].PutValue('LIBELLE',TobArticle.GetValue('GA_LIBELLE'));
    end;
TobArticle.Free;
end;

function TOF_AjustStock.RechercheDimensions(TOBArt : TOB) : string;
var stDim : string;
    i_ind : integer;
    Sep : string;
    GrilleDim,CodeDim,LibDim : String ;
begin
Result:='';
if TOBArt.GetValue('GA_STATUTART') <> 'DIM' then exit;
Sep := '';
stDim := '';
for i_ind := 1 to MaxDimension do
    begin
    GrilleDim:=TOBArt.GetValue('GA_GRILLEDIM'+IntToStr(i_ind)) ;
    CodeDim:=TOBArt.GetValue('GA_CODEDIM'+IntToStr(i_ind)) ;
    if GrilleDim<>'' then
        begin
        LibDim:=GCGetCodeDim(GrilleDim,CodeDim,i_ind) ;
        if LibDim<>'' then stDim:=stDim+Sep+LibDim ;
        end ;
    Sep := ' - ';
    end;
Result:=stDim;
end;

Initialization
registerclasses([TOF_AjustStock]) ;

end.
