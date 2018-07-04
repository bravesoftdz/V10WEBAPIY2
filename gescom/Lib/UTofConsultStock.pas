unit UTofConsultStock;

interface

uses {$IFDEF VER150} variants,{$ENDIF}  StdCtrls,Controls,Classes,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_Main,Mul,HDB,
{$ENDIF}
      forms,sysutils,
      ComCtrls,Hpanel, Math,
      HCtrls,HEnt1,HMsgBox,UTOF,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics,
      grids,windows,
      M3FP,HTB97,Dialogs, AGLInitGC, ExtCtrls, Hqry, LicUtil,
      FactUtil, StockUtil, FactComm, Facture,
      InvUtil, UTofOptionEdit,HDimension,uEntCommun,UtilTOBPiece;

Type
    TOF_ConsultStock = Class (TOF)
    private
        SPLITTER  : TSplitter;
        GARTICLE  : THGRID ;
        GLIGNE    : THGRID ;
        PSAICUMUL : THPanel;
        CUMQTERES : THNumEdit;
        CUMQTEFOU : THNumEdit;
        QTEPRO    : THNumEdit;
        BImprimer : TToolbarButton97;
        CODEARTICLE : THEdit;
        CODEARTICLE_: THEDIT;

        procedure OnElipsisClick(Sender: TObject);
        procedure OnElipsisClick_(Sender: TObject);        
        
    public
        TobVue, TobLigne, TobDispo : TOB;
        LesColsLigne,LesColsDispo, stListSelect : string;
        bFormCreate : boolean;
        procedure OnArgument (Arguments : String ) ; override;
        procedure OnClose ; override;
        procedure OnUpdate ; override;

    // Initialisation
        procedure CreeTotal (ColLig : integer; Cumul : THNumEdit);
        procedure EtudieColsListe ;
        procedure InitialiseGrille;

    // Impression
        procedure BImprimerClick(Sender: TObject);
        function PrepareImpression : integer ;

    // Actions liées au Grid
        procedure GARTICLERowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure GLIGNEColumnWidthsChanged(Sender: TObject);
        procedure GARTICLEDblClick(Sender: TObject);
        procedure GLIGNEDblClick (Sender : TObject);

    // Gestion des données
        procedure AffectePied;
        procedure AjoutSelect (stChamp : string);
        procedure CalculStockPro;
        procedure ChargeTobDispo;
        procedure ChargeGrille;
        procedure ChargeTob;
        Function  GetTob (Grid : THGrid; ARow, ARowMax : integer) : TOB ;
        function  GetTobArt(stArt : string) : TOB;
        procedure LoadTobVue;
        function  RechDim(TobArt : TOB) : string;
        procedure RecupComposants;
        function  GetParentQte(TOBLigneEtNomen : TOB ; pref : string) : Double;
        function  RecupCriteres(pref : string) : string;
    end;

//Construction d'une clause where sur les natures de pièces dont le stock est
//est affecté par stQteAffect
function RecupWhereNaturePiece(QteAffect, Prefixe : string;  ExceptInv : boolean=false) : string;

Var PEtat : RPARETAT;
    ColCli : integer;
    ColFou : integer;
    ColPro : integer;
    ColArt : integer;
    ColDep : integer;

const
// libellés des messages
TexteMessage: array[1..2] of string 	= (
          {1}  'Vous devez saisir un code article'
          {2} ,''
              );

implementation


procedure TOF_ConsultStock.OnArgument (Arguments : String ) ;
begin
inherited ;
bFormCreate := true;
PEtat.Tip:='E';
PEtat.Nat:='GZH';
PEtat.Modele:='GZH';
PEtat.Titre:=TFmul(Ecran).Caption;
PEtat.Apercu:=True;
PEtat.DeuxPages:=False;
PEtat.First:=True;
PEtat.stSQL:='GZH_UTILISATEUR = "'+V_PGI.USer+'"';

GARTICLE := THGRID(GetControl('GARTICLE'));
GLIGNE := THGRID(GetControl('GLIGNE'));

GARTICLE.ListeParam := 'GCCONSULTSTOCK';
GLIGNE.ListeParam := 'GCCONSULTSTOCKLIG';

GARTICLE.OnRowEnter := GARTICLERowEnter;
GARTICLE.OnDblClick := GARTICLEDblClick;

GLIGNE.OnDblClick := GLIGNEDblClick;
GLIGNE.OnColumnWidthsChanged := GLIGNEColumnWidthsChanged;

CUMQTERES := THNumEdit(GetControl('CUMQTERES'));
CUMQTEFOU := THNumEdit(GetControl('CUMQTEFOU'));
QTEPRO := THNumEdit(GetControl('QTEPRO'));

BImprimer:=TToolbarButton97(GetControl('Bimprimer')) ;
BImprimer.OnClick:=BImprimerClick;

SPLITTER := TSplitter.Create(THPanel(GetControl('PGRID')));
SPLITTER.Name         := 'Splitter';
SPLITTER.Parent       := THPanel(GetControl('PGRID')) ;
SPLITTER.AutoSnap     := True;
SPLITTER.Beveled      := True;
SPLITTER.ResizeStyle  := rsUpdate;
SPLITTER.Left   := THGRID(GetControl('GARTICLE')).Left ;
SPLITTER.Align  := THGRID(GetControl('GARTICLE')).Align;
SPLITTER.Width  := THGRID(GetControl('GARTICLE')).Width;
SPLITTER.Height := 3;
SPLITTER.Cursor := crVSplit;
SPLITTER.Color  := clActiveCaption;
PSAICUMUL       := THPANEL(GetControl('PSAICUMUL'));

THValComboBox(GetControl('NATUREPIECEG')).ItemIndex := 0;

if not (ctxMode in V_PGI.PGIContexte) then
begin
  SetControlVisible('NATUREPIECEG',False);
  SetControlVisible('TGL_NATUREPIECEG',False);
end;

if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
  TToolbarButton97(GetControl('BParamListe')).Visible:=True
else
  TToolbarButton97(GetControl('BParamListe')).Visible:=False;

TobVue := Tob.Create ('', Nil, -1);
TobDispo := Tob.Create ('', Nil, -1);
TobLigne := TOB.Create ('', nil, -1);

stListSelect:= 'GL_DEPOT,GL_DATELIVRAISON,GL_NATUREPIECEG,GL_NUMERO,GL_NUMLIGNE,GL_SOUCHE,GL_INDICEG,GL_DATEPIECE,'+
                             'GL_QUALIFQTEACH,GL_QUALIFQTEVTE,GL_QUALIFQTESTO,GL_QTERESTE, GL_MTRESTE';
InitialiseGrille;

CODEARTICLE := THEdit(GETCOntrol('GA_CODEARTICLE'));
CODEARTICLE_:= THEdit(GETCOntrol('GA_CODEARTICLE_'));
CodeArticle.OnElipsisClick := OnElipsisClick;
CodeArticle_.OnElipsisClick:= OnElipsisClick_;

end;

procedure TOF_ConsultStock.OnClose;
begin
inherited;
SPLITTER.Free;
TobDispo.Free;
TobVue.Free;
TobLigne.Free;
end;

procedure TOF_ConsultStock.OnUpdate ;
begin
inherited;

//TFMul(Ecran).Q.UpdateCriteres;
if (GetControlText('GA_CODEARTICLE') = '') and (not bFormCreate) then
   begin
   //MessageAlerte('Aucun élément sélectionné');
   HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[1]+';W;O;O;O;','','') ;
   exit;
   end;
bFormCreate := false;
TobVue.ClearDetail;
TobLigne.ClearDetail;
TobDispo.ClearDetail;
GARTICLE.VidePile (False);
GLIGNE.VidePile (False);
ChargeTobDispo;
if TobDispo.Detail.Count > 0 then
   begin
   GARTICLE.Row:=1;
   LoadTobVue;
   ChargeTob;
   end;
AffectePied;
end;

{========================================================================================}
{========================= Impression ===================================================}
{========================================================================================}
procedure TOF_ConsultStock.BImprimerClick(Sender: TObject);
begin
EntreeOptionEdit(PEtat.Tip,PEtat.Nat,PEtat.Modele,PEtat.Apercu,PEtat.DeuxPages,PEtat.First,
                 TPageControl(GetControl('Pages')),PEtat.stSQL,PEtat.Titre, PrepareImpression);
ExecuteSQL('DELETE FROM GCTMPCONSULTSTK WHERE GZH_UTILISATEUR = "'+V_PGI.USer+'"');
end;

function TOF_ConsultStock.PrepareImpression : integer ;
var i_ind : integer;
    TobGZH, TobD, TobV : TOB;
    stChamp, stField, stLesCol, stCodeArticle : string;
begin
Result:=0;
stCodeArticle := '';
ExecuteSQL('DELETE FROM GCTMPCONSULTSTK WHERE GZH_UTILISATEUR = "'+V_PGI.USer+'"');
TobGZH:=TOB.Create('',Nil,-1);
for i_ind:=0 to TobVue.Detail.Count-1 do
    begin
    TobV:=TobVue.Detail[i_ind] ;
    TobD:=TOB.Create('GCTMPCONSULTSTK',TobGZH,-1);
    TobD.PutValue('GZH_UTILISATEUR',V_PGI.USer);
    TobD.PutValue('GZH_COMPTEUR',i_ind);
    stLesCol := stListSelect;
    //ReadTokenSt (stLesCol);   // On enlève la fixedrow
    stChamp := ReadTokenPipe(stLesCol,',');
    while stChamp <> '' do
        begin
        stField := Copy(stChamp, 4, Length(stChamp) - 3);
        if TobD.FieldExists('GZH_' + stField) then
           TobD.PutValue ('GZH_' + stField , tobV.GetValue (stChamp));
        stChamp := ReadTokenPipe(stLesCol,',');
        end;
    TobD.PutValue ('GZH_ARTICLE', TobV.GetValue ('ARTICLE'));
    TobD.PutValue ('GZH_CODEARTICLE', TobV.GetValue ('CODEARTICLE'));
    TobD.PutValue ('GZH_PHYSIQUE', TobV.GetValue ('GQ_PHYSIQUE'));
    if GetTobArt(TobV.GetValue ('ARTICLE')).GetValue('GA_STATUTART')='DIM' then
        TobD.PutValue ('GZH_LIBDIM', RechDim(GetTobArt(TobV.GetValue ('ARTICLE'))))
        else
        TobD.PutValue ('GZH_LIBDIM', '');
    TobD.PutValue ('GZH_LIBELLE', (GetTobArt(TobV.GetValue ('ARTICLE'))).GetValue('GA_LIBELLE'));
    TobD.PutValue ('GZH_RATIO', TobV.GetValue ('RATIO'));
    TobD.PutValue ('GZH_VA', TobV.GetValue ('VA'));
    TobD.PutValue('GZH_STOCKNET', Valeur(TobV.GetValue ('STOCKPRO')));

    if TobV.GetValue ('VA') = 'ACH' then
        begin
        //TobD.PutValue('GZH_QTEFOU', TobD.GetValue ('GZH_QTESTOCK'));
        TobD.PutValue('GZH_QTEFOU', TobV.GetValue ('QTEFOU'));
        TobD.PutValue('GZH_QTESTOCK', 0);
        end else
        begin
        TobD.PutValue('GZH_QTESTOCK', Tobv.GetValue ('GL_QTERESTE'));
        TobD.PutValue('GZH_QTEFOU', 0);
        end;
    end;
TobGZH.InsertOrUpdateDB(True);
TobGZH.Free;
end;

{========================================================================================}
{========================= Actions liées au Grid ========================================}
{========================================================================================}

procedure TOF_ConsultStock.GLIGNEColumnWidthsChanged(Sender: TObject);
var Coord : TRect;
begin
if PSAICUMUL.ControlCount<=0 then exit ;
Coord:=GLIGNE.CellRect(ColCLI,0);
CUMQTERES.Left:=Coord.Left + 1;
CUMQTERES.Width:=GLIGNE.ColWidths[ColCLI] + 1;

Coord:=GLIGNE.CellRect(ColFOU,0);
CUMQTEFOU.Left:=Coord.Left + 1;
CUMQTEFOU.Width:=GLIGNE.ColWidths[ColFOU] + 1;

Coord:=GLIGNE.CellRect(ColPRO,0);
QTEPRO.Left:=Coord.Left + 1;
QTEPRO.Width:=GLIGNE.ColWidths[ColPRO] + 1;
end;

procedure TOF_ConsultStock.GARTICLERowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
ChargeGrille;
TobLigne.PutGridDetail (GLIGNE, False, False, LesColsLigne, False);
AffectePied;
end;

procedure TOF_ConsultStock.GARTICLEDblClick(Sender: TObject);
var TobA : TOB;
begin
TobA := GetTob (GARTICLE, GARTICLE.Row, TobDispo.Detail.Count);
if (TobA <> nil) and (TobA.GetValue ('GA_ARTICLE') <> '') then
{$IFNDEF GPAO}  
{$IFDEF BTP}
    V_PGI.DispatchTT(7, taConsult, TobA.GetValue('GA_ARTICLE'), 'MONOFICHE', '');
{$ELSE}
    AglLanceFiche ('GC', 'GCARTICLE', '', TobA.GetValue ('GA_ARTICLE'),
                   'ACTION=CONSULTATION;MONOFICHE');
{$ENDIF}
{$ELSE}
    V_PGI.DispatchTT(7, taConsult, TobA.GetValue('GA_ARTICLE'), 'MONOFICHE', '');
{$ENDIF GPAO}
end;

procedure TOF_ConsultStock.GLIGNEDblClick (Sender : TObject);
var CleDoc : R_CleDoc;
    stNature, stDate, stSouche, stNumero, stIndiceg : string;
    TobL : TOB;
begin
TobL := GetTob (GLIGNE, GLIGNE.Row, TobLigne.Detail.Count);
if TobL = nil then exit;
stNature := TobL.GetValue ('GL_NATUREPIECEG');
stDate := TobL.GetValue ('GL_DATEPIECE');
stSouche := TobL.GetValue ('GL_SOUCHE');
stNumero := IntToStr(TobL.GetValue ('GL_NUMERO'));
stIndiceg := IntToStr(TobL.GetValue ('GL_INDICEG'));
if (stNature <> '') and (stDate <> '') and (stSouche <> '') and (stNumero <> '') and
   (stIndiceg <> '') then
    begin
    StringToCleDoc (
        stNature + ';' + stDate + ';' + stSouche + ';' + stNumero + ';' + stIndiceg + ';',
        CleDoc);
    SaisiePiece (CleDoc, taConsult) ;
    end;
end;

{========================================================================================}
{===================      Initialisation    =============================================}
{========================================================================================}

procedure TOF_ConsultStock.CreeTotal (ColLig : integer; Cumul : THNumEdit);
var Coord : TRect;
begin
Cumul.ParentColor := True;
Cumul.Font.Style := PSAICUMUL.Font.Style;
Cumul.Font.Size := PSAICUMUL.Font.Size;
Cumul.Masks.PositiveMask := GLIGNE.ColFormats[ColLig];
Cumul.Ctl3D := False;
Cumul.Top := 0;
Coord := GLIGNE.CellRect (ColLig, 0);
Cumul.Left := Coord.Left;
Cumul.Width := GLIGNE.ColWidths[ColLig] + 1;
Cumul.Height := PSAICUMUL.Height;
Cumul.Enabled := False;
end;

procedure TOF_ConsultStock.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp : integer ;
begin
LesColsDispo := GARTICLE.Titres[0];
LesCols := LesColsDispo;
icol:=1 ;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        begin
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then
           begin
           if NomCol='GA_ARTICLE' then ColArt:=icol else
           if NomCol='GQ_DEPOT' then ColDep:=icol;
           end ;
        end ;
    Inc(icol) ;
    Until ((LesCols='') or (NomCol='')) ;
end ;


{========================================================================================}
{================================= Gestion des Données ==================================}
{========================================================================================}
procedure TOF_ConsultStock.AffectePied;
var stDim,StArt : string;
begin
stDim := '';
stArt := GARTICLE.Cells[ColArt,GARTICLE.Row];
if stArt <> '' then
   if GetTobArt(stArt).GetValue('GA_STATUTART')='DIM' then
      stDim := RechDim(GetTobArt(stArt));

SetControlText('TDIMENSION',stDim);
if stDim <> '' then
     THPanel(GetControl('PPIED')).Visible := True
else THPanel(GetControl('PPIED')).Visible := False;

CUMQTERES.Value := TobLigne.Somme ('GL_QTERESTE', [''], [''], False);
CUMQTEFOU.Value := TobLigne.Somme ('QTEFOU', [''], [''], False);
if TobLigne.Detail.Count > 0 then
     QTEPRO.Value := TobLigne.Detail[TobLigne.Detail.Count - 1].GetValue ('STOCKPRO')
else QTEPRO.Value := 0;
end;

procedure TOF_ConsultStock.AjoutSelect (stChamp : string);
begin
if Pos (stChamp, stListSelect) = 0 then
    if stListSelect <> '' then stListSelect := stListSelect + ', '+ stChamp;
end;

procedure TOF_ConsultStock.ChargeTobDispo;
var stSQL, stWhere : string;
    QDispo : TQuery;
begin
if TFMul(Ecran).Q.Eof then exit;
stSQL := 'SELECT GA_ARTICLE,GA_CODEARTICLE,GA_LIBELLE,GQ_PHYSIQUE,GQ_DEPOT '+
',GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5,'+
'GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,GA_STATUTART '+
'FROM ARTICLE LEFT JOIN DISPO ON GA_ARTICLE=GQ_ARTICLE WHERE GQ_CLOTURE="-" ';
stWhere := RecupWhereCritere (TFMul(Ecran).Pages);
if stWhere <> '' then
    begin
    System.Delete (stWhere, pos ('WHERE ', stWhere), 6);
    stWhere := ' AND ' + stWhere;
    end;
stSQL := stSQL + stWhere + ' ORDER BY GQ_DEPOT';
QDispo := OpenSQL(stSQL,true);
TobDispo.LoadDetailDB('','','',QDispo,false);
Ferme(QDispo);
TobDispo.PutGridDetail(GARTICLE, False, False, LesColsDispo, False);
end;

procedure TOF_ConsultStock.ChargeGrille;
var TobFind, TobF  : TOB;
    stArticle,stDepot : string;
begin
TobLigne.ClearDetail;
GLIGNE.VidePile (False);
stArticle := GARTICLE.Cells[ColArt,GARTICLE.Row];
stDepot   := GARTICLE.Cells[ColDep,GARTICLE.Row];
TobFind:= TobVue.FindFirst(['ARTICLE','GL_DEPOT'],[stArticle,stDepot],True);
While TobFind <> nil do
    begin
    TobF:=TOB.Create ('', TobLigne,-1) ;
    TobF.Dupliquer(TobFind, False, True);
    TobFind:= TobVue.FindNext(['ARTICLE','GL_DEPOT'],[stArticle,stDepot],True);
    end;
end;

procedure TOF_ConsultStock.ChargeTob;
begin
if GARTICLE.Cells[ColArt,GARTICLE.Row] = '' then exit;
ChargeGrille;
TobLigne.PutGridDetail (GLIGNE, False, False,LesColsLigne, False);
end;

procedure TOF_ConsultStock.CalculStockPro;
var stDepot, stArticle : string;
    i_ind1, i_ind2, i_max : integer;
begin
stDepot := '';
stArticle := '';
i_ind1 := 0;
i_max := TobVue.Detail.Count - 1;
while i_ind1 <=  i_max do
    begin
    i_ind2 := 0;
    stDepot := TobVue.Detail[i_ind1].GetValue('GL_DEPOT');
    stArticle := TobVue.Detail[i_ind1].GetValue('ARTICLE');
     while (stDepot=TobVue.Detail[i_ind1].GetValue('GL_DEPOT')) and
          (stArticle=TobVue.Detail[i_ind1].GetValue('ARTICLE'))  do
        begin
        if i_ind2 > 0 then
            begin
            TobVue.Detail[i_ind1].PutValue ('STOCKPRO',
            Valeur(TobVue.Detail[i_ind1 - 1].GetValue ('STOCKPRO')) -
            TobVue.Detail[i_ind1].GetValue ('GL_QTERESTE') +
            Valeur(TobVue.Detail[i_ind1].GetValue ('QTEFOU')));
            end else
            begin
            TobVue.Detail[i_ind1].PutValue ('STOCKPRO',
            TobVue.Detail[i_ind1].GetValue ('GQ_PHYSIQUE') -
            TobVue.Detail[i_ind1].GetValue ('GL_QTERESTE') +
            Valeur(TobVue.Detail[i_ind1].GetValue ('QTEFOU')));
            end;
        i_ind2 := i_ind2 + 1;
        i_ind1 := i_ind1 + 1;
        if i_ind1 > i_max then exit;
        end;
    end;
end;

Function TOF_ConsultStock.GetTob (Grid : THGrid; ARow, ARowMax : integer) : TOB ;
begin
Result := Nil ;
if ((ARow <= 0) or (ARow > ARowMax)) then Exit ;
Result := TOB (Grid.OBjects [0,ARow]);
end ;

function TOF_ConsultStock.GetTobArt(stArt : string) : TOB;
begin
result := TobDispo.FindFirst(['GA_ARTICLE'],[stArt],false);
end;

procedure TOF_ConsultStock.InitialiseGrille;
var icol : integer;
    stChamp,stLesCols : string;
begin
EtudieColsListe;

stLesCols := GLIGNE.Titres[0];
icol := 0;
While stLesCols <> '' do
    begin
    stChamp := ReadTokenSt(stLesCols);
    if (pos ('GL_', stChamp) = 1) then
        begin
        icol := icol + 1;
        AjoutSelect(stChamp);
        if pos ('DATE', stChamp) > 0 then GLIGNE.ColFormats[icol] := '';
        end
    end;

LesColsLigne := GLIGNE.Titres[0] + 'GL_QTERESTE;QTEFOU;STOCKPRO';
ColCli := icol+1;
ColFou := icol+2;
ColPro := icol+3;

GARTICLE.ColLengths[ColArt] := -1;
GARTICLE.ColWidths[ColArt] := -1;
GLIGNE.ColCount := GLIGNE.ColCount + 3; // pour QteCli, Qtefour, Qtepro
GLIGNE.ColWidths[ColCli] := 130;
GLIGNE.ColFormats[ColCli] := '#,##0.00;; ;';
GLIGNE.Cells[ColCli, 0] := 'Livraison client';
GLIGNE.ColWidths[ColFou] := 130;
GLIGNE.ColFormats[ColFou] := '#,##0.00;; ;';
GLIGNE.Cells[ColFou, 0] := 'Qté fournisseur';
GLIGNE.ColWidths[ColPro] := 130;
GLIGNE.ColFormats[ColPro] := '#,##0.00;; ;';
GLIGNE.Cells[ColPro, 0] := 'Stock Progressif';

GLIGNE.ColAligns[ColCli] := taRightJustify;
GLIGNE.ColAligns[ColFOU] := taRightJustify;
GLIGNE.ColAligns[ColPRO] := taRightJustify;

TFMul(Ecran).Hmtrad.ResizeGridColumns(GLIGNE) ;
TFMul(Ecran).Hmtrad.ResizeGridColumns(GARTICLE) ;

CreeTotal (ColCli, CUMQTERES);
CreeTotal (ColFou, CUMQTEFOU);
CreeTotal (ColPro, QTEPRO);
end;


procedure TOF_ConsultStock.LoadTobVue;
var stWhere, stSelect, stOrder : string;
    TSql : TQuery;
    TobV,TobD : TOB;
    iIndex : integer;
    phys : double;
    Ratio : Double;
begin
stWhere := RecupCriteres ('GL');
stSelect := 'SELECT GL_ARTICLE AS ARTICLE,' + stListSelect + ',GL_CODEARTICLE AS CODEARTICLE,'+
            'GL_QUALIFQTEACH as GLN_QUALIFQTEACH,GL_QUALIFQTEVTE as GLN_QUALIFQTEVTE, ' +
            'GL_QUALIFQTESTO as GLN_QUALIFQTESTO FROM LIGNE ' +
            'WHERE GL_TENUESTOCK="X" AND GL_VIVANTE="X" AND GL_QTERESTE > 0 AND GL_QUALIFMVT<>"ANN" ' ; {DBR NEWPIECE QTERESTE > 0}
stOrder := ' ORDER BY GL_DEPOT, GL_ARTICLE,GL_DATELIVRAISON';

stSelect := stSelect + stWhere + stOrder;

TSql := OpenSql (stSelect, True);
if not TSql.Eof then TobVue.LoadDetailDB ('', '', '', TSql, False);
Ferme (TSql);
if TobVue.Detail.count > 0 then TobVue.Detail[0].AddChampSupValeur ('NOMEN', '-',True);
RecupComposants;
if TobVue.Detail.count = 0 then exit;
TobVue.Detail.Sort('ARTICLE;GL_DEPOT;GL_DATELIVRAISON;GL_NATUREPIECEG;GL_NUMERO;GL_NUMLIGNE');
TobVue.Detail[0].AddChampSup ('VA', True);
TobVue.Detail[0].AddChampSup ('RATIO', True);
TobVue.Detail[0].AddChampSup ('QTEFOU', True);
TobVue.Detail[0].AddChampSup ('STOCKPRO', True);
TobVue.Detail[0].AddChampSup ('GQ_PHYSIQUE', True);

for iIndex := 0 to TobVue.Detail.Count - 1 do
    begin
    Ratio := 0;
    TobV := TobVue.Detail[iIndex];
    TobV.PutValue('VA',GetInfoParPiece (TobV.GetValue('GL_NATUREPIECEG'),'GPP_VENTEACHAT'));
    //
    if TobV.GetValue('NOMEN') = 'X' then
      Ratio := GetRatio(TobV,TobV, trsStock)
    else
      Ratio := GetRatio(TobV, nil, trsStock);

    TobV.PutValue ('RATIO', Ratio);
    //
    if TobV.GetValue('VA') = 'ACH' then
       begin
       TobV.PutValue ('QTEFOU', TobV.GetValue ('GL_QTERESTE')/TobV.GetValue('RATIO'));
       TobV.PutValue ('GL_QTERESTE', StrToFloat ('0,0'));
       end else
       begin
       TobV.PutValue('GL_QTERESTE', TobV.GetValue ('GL_QTERESTE')/TobV.GetValue('RATIO'));
       TobV.PutValue ('QTEFOU', StrToFloat ('0,0'));
       end;
    TobD := TobDispo.FindFirst(['GA_ARTICLE','GQ_DEPOT'],[TobV.GetValue('ARTICLE'),TobV.GetValue('GL_DEPOT')],false);
    if TobD <> nil then phys := TobD.GetValue('GQ_PHYSIQUE') else phys := 0;
    TobV.PutValue('GQ_PHYSIQUE',phys);
    end;
CalculStockPro;
end;

function TOF_ConsultStock.RechDim(TobArt : TOB) : string;
var i_indDim : integer;
    GrilleDim,CodeDim,LibDim,StDim : string;
begin
StDim:='';
for i_indDim := 1 to MaxDimension do
    begin
    GrilleDim := TOBArt.GetValue ('GA_GRILLEDIM' + IntToStr (i_indDim));
    CodeDim := TOBArt.GetValue ('GA_CODEDIM' + IntToStr (i_indDim));
    LibDim := GCGetCodeDim (GrilleDim, CodeDim, i_indDim);
    if LibDim <> '' then
       if StDim='' then StDim:=StDim + LibDim
       else StDim := StDim + ' - ' + LibDim;
    end;
Result := stDim;
end;

// Récupère les composants des nomenclatures dans LigneNomen
procedure TOF_ConsultStock.RecupComposants;
var QNomen : TQuery;
    i_ind : integer;
    stSelectNomen, stWhere, stOrder : string;
    TobNomen, TOBLN : TOB;
begin
stSelectNomen := 'SELECT GLN_ARTICLE AS ARTICLE,' + stListSelect +
   ',GLN_CODEARTICLE AS CODEARTICLE,GLN_QUALIFQTEACH,GLN_QUALIFQTEVTE,GLN_QUALIFQTESTO,' +
   'GLN_COMPOSE,GLN_ORDRECOMPO,GLN_QTE'+
   ' FROM LIGNENOMEN LEFT JOIN LIGNE ON GLN_NATUREPIECEG=GL_NATUREPIECEG ' +
   ' AND GLN_SOUCHE=GL_SOUCHE AND GLN_NUMERO=GL_NUMERO' +
   ' AND GLN_INDICEG=GL_INDICEG AND GLN_NUMLIGNE=GL_NUMLIGNE' +
   ' WHERE GLN_TENUESTOCK="X" AND GL_VIVANTE="X" AND GL_QUALIFMVT<>"ANN" ';
stOrder := ' ORDER BY GL_DEPOT, GLN_ARTICLE,GL_DATELIVRAISON';

stWhere:=RecupCriteres ('GLN');
stSelectNomen := stSelectNomen + stWhere + stOrder;
QNomen := OpenSQL(stSelectNomen,true);
if not QNomen.Eof then
   begin
   TobNomen := TOB.Create('',nil,-1);
   TobNomen.LoadDetailDB('','','',QNomen,false);
   if TobNomen.Detail.count > 0 then TobNomen.Detail[0].AddChampSupValeur ('NOMEN', 'X',True);
   for i_ind := 0 to TobNomen.Detail.Count -1 do
       begin
       TOBLN := TobNomen.Detail[i_ind];
       TOBLN.PutValue('GL_QTERESTE',TOBLN.GetValue('GL_QTERESTE') * GetParentQte(TOBLN,'GL'));
       end;
   for i_ind := TobNomen.Detail.Count -1 downto 0 do
       TobNomen.Detail[i_ind].ChangeParent(TobVue,-1);
   TobNomen.Free;
   end;
Ferme(QNomen);
end;

function TOF_ConsultStock.GetParentQte(TOBLigneEtNomen : TOB ; pref : string) : Double;
var Q : TQuery;
    TOBParent : TOB;
begin
result := TOBLigneEtNomen.GetValue('GLN_QTE');
if (TOBLigneEtNomen.GetValue('GLN_COMPOSE') = null) or
   (TOBLigneEtNomen.GetValue('GLN_COMPOSE') = '') then exit;
Q := OpenSQL('SELECT GLN_NATUREPIECEG,GLN_SOUCHE,GLN_NUMERO,GLN_INDICEG,GLN_NUMLIGNE,GLN_COMPOSE,GLN_ORDRECOMPO,GLN_QTE '+
             'FROM LIGNENOMEN '+
             'WHERE GLN_ARTICLE="'+TOBLigneEtNomen.GetValue('GLN_COMPOSE')+'" '+
                   'AND GLN_NUMORDRE='+IntToStr(TOBLigneEtNomen.GetValue('GLN_ORDRECOMPO'))+' '+
                   'AND GLN_NATUREPIECEG="'+TOBLigneEtNomen.GetValue(pref+'_NATUREPIECEG')+'" '+
                   'AND GLN_SOUCHE="'+TOBLigneEtNomen.GetValue(pref+'_SOUCHE')+'" '+
                   'AND GLN_NUMERO='+IntToStr(TOBLigneEtNomen.GetValue(pref+'_NUMERO'))+' '+
                   'AND GLN_INDICEG='+IntToStr(TOBLigneEtNomen.GetValue(pref+'_INDICEG'))+' '+
                   'AND GLN_NUMLIGNE='+IntToStr(TOBLigneEtNomen.GetValue(pref+'_NUMLIGNE')), true);
TOBParent := TOB.Create('LIGNE', nil, -1);
if not TOBParent.SelectDB('', Q) then begin TOBParent.Free; TOBParent := nil; end;
Ferme(Q);
if TOBParent <> nil then begin result := result * GetParentQte(TOBParent,'GLN');
                               TOBParent.Free; end;
end;

function TOF_ConsultStock.RecupCriteres(pref : string) : string;
var stWhere, stWhereNat : string;
begin
stWhere := '';
if GetControlText ('GA_CODEARTICLE') <> '' then
    stWhere := stWhere + ' AND ' + pref + '_ARTICLE>="' + GetControlText ('GA_CODEARTICLE') + '"';
if GetControlText ('GA_CODEARTICLE_') <> '' then
    stWhere := stWhere + ' AND ' + pref + '_ARTICLE<="' + format('%-18.18s%16.16s',[GetControlText ('GA_CODEARTICLE_'),StringOfChar('Z',16)]) + '"';

if GetControlText('GQ_DEPOT') <> '' then
    stWhere := stWhere + 'AND ' + 'GL_DEPOT>="' + GetControlText('GQ_DEPOT') + '"';

if GetControlText('GQ_DEPOT_') <> '' then
    stWhere := stWhere + 'AND ' + 'GL_DEPOT<="' + GetControlText('GQ_DEPOT_') + '"';

stWhereNat := RecupWhereNaturePiece(GetControlText('NATUREPIECEG'),'GL');
if stWhereNat <> '' then stWhereNat := ' AND '+stWhereNat;
result := stWhere + stWhereNat;
end;


{***********A.G.L.***********************************************
Auteur  ...... : JS
Créé le ...... : 04/07/2001
Modifié le ... :   /  /
Description .. : Retourne une condition SQL sur les natures de pièce
Suite ........ : incrémentant ou décrémentant les quantités en stock
Suite ........ : "QteAffect". Qteaffect est une chaîne tokenisée de
Suite ........ : codes tablette GCCOLSTOCK.
Mots clefs ... : PARPIECE;STOCK
*****************************************************************}
function RecupWhereNaturePiece(QteAffect, Prefixe : string;  ExceptInv : boolean=false) : string;
var stQte,stN,stWhere : string;
    i_ind : integer;
    TobPP : TOB;
begin
stWhere := ''; Result := '';
if PrefixeToNum(Prefixe) <= 0  then exit;
stQte := QteAffect;
stN := ReadTokenSt(stQte);
while stN <> '' do
    begin
    for i_ind := 0 to VH_GC.TOBParPiece.Detail.Count -1 do
        begin
        TobPP := VH_GC.TOBParPiece.Detail[i_ind];
        if (TobPP.GetString('GPP_NATUREPIECEG')='INV') and (ExceptInv) then Continue;
        if (Pos(stN, TobPP.GetValue('GPP_QTEPLUS')) > 0)  or
           (Pos(stN, TobPP.GetValue('GPP_QTEMOINS')) > 0) then
            begin
            if stWhere <> '' then stWhere := stWhere + ' OR ';
            stWhere := stWhere + Prefixe + '_NATUREPIECEG="'+TobPP.GetValue('GPP_NATUREPIECEG')+'"';
            end;
        end;
    stN := ReadTokenSt(stQte);
    end;
if stWhere <> '' then stWhere := '(' + stWhere + ')';
result := stWhere;
end;

//FV1 : 13/02/2014 - FS#818 - Fiche catalogue : la recherche sur article ne doit faire apparaître que des articles
Procedure TOF_ConsultStock.OnElipsisClick(Sender: TObject);
Var CodeArt : String;
    StChamps: String;
begin

  CodeArt := CODEARTICLE.text;
  StChamps:= '';

  if CodeArt <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(CodeArt, 1, 18)) + ';';

  StChamps:= StChamps + 'XX_WHERE=GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="ARP" OR GA_TYPEARTICLE="OUV"';

  CodeArt := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps + ';GA_TENUESTOCK=X;ACTION=CONSULTATION;STATUTART=UNI,DIM');

  if codeArt <> '' then
  begin
    CODEARTICLE.text := CodeArt;
    CODEARTICLE_.text:= CodeArt;
  end;      

end;

//FV1 : 13/02/2014 - FS#818 - Fiche catalogue : la recherche sur article ne doit faire apparaître que des articles
Procedure TOF_ConsultStock.OnElipsisClick_(Sender: TObject);
Var CodeArt : String;
    StChamps: String;
begin

  CodeArt := CODEARTICLE_.text;
  StChamps:= '';

  if CodeArt <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(CodeArt, 1, 18)) + ';';

  StChamps:= StChamps + 'XX_WHERE=GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="ARP" OR GA_TYPEARTICLE="OUV"';

  CodeArt := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps + ';GA_TENUESTOCK=X;ACTION=CONSULTATION;STATUTART=UNI,DIM');

  if codeArt <> '' then
  begin
    CODEARTICLE.Text := CodeArt;
    CODEARTICLE_.Text := CodeArt;
  end;

end;



Initialization
registerclasses([TOF_ConsultStock]);

end.
