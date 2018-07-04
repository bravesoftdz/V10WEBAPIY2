{***********UNITE*************************************************
Auteur  ...... : JS
Créé le ...... : 13/05/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCARTNONMVTES_MUL ()
                 Consultation des articles non mouvementés
Mots clefs ... : TOF;GCARTNONMVTES_MUL
*****************************************************************}
Unit ArtNonMvtes_Tof ;

Interface

uses {$IFDEF VER150} variants,{$ENDIF} StdCtrls,
     Controls,
     Classes,Buttons,
{$IFNDEF EAGLCLIENT}
     db,mul,FE_Main,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}dbGrids,EdtREtat,
{$ELSE}
     emul,MaineAGL,UtileAgl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,Hqry,FactUtil,HDimension,Menus,
     HMsgBox,UTOF,UTOB,Ent1,HTB97,HPanel,HSysMenu,ParamSoc,
     UTofOptionEdit,EntGC,ParamDBG,LicUtil,UtilGC,HStatus;

function GCLanceFiche_ArtNonMvtes(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
  TOF_GCARTNONMVTES_MUL = Class (TOF)
  private
    GA : THGrid;
    HMTrad: THSystemMenu;
    BImprimer,BSelectAll : TToolbarButton97;
    BOuvrir,BValideDateSup : TToolbarButton97;
    BZoomArticle,BZoomPiece : TBitBtn;
    TobA : TOB;
    ListeChampGrid,ListeChampSQL : string;
    NaturePiece : string;

    /// Boutons
    procedure BZoomArticleClick(Sender: TObject);
    procedure BZoomPieceClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BSelectAllClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure BValideDateSupClick(Sender: TObject);
    /// Grid
    procedure GDblClick(Sender: TObject);
    procedure GARowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure RefreshGrille;
    function CreerSelectSQL : string;
    /// Tob
    procedure ChargeTobArticle;
    function RecupWhereCritereLigne : string;
    function GetCondition(Where,Plus,NomChamp : string) : string;
    /// Dimensions article
    procedure AfficheDimension;
    function RechDim(TobArt : TOB) : string;
    procedure CodeArticleRech(Sender: Tobject);

  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

Implementation

function GCLanceFiche_ArtNonMvtes(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
Result:='';
if Nat='' then exit;
if Cod='' then exit;
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_GCARTNONMVTES_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCARTNONMVTES_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCARTNONMVTES_MUL.OnUpdate ;
begin
  Inherited ;
RefreshGrille;
SetControlText('GA_ARTICLE','');
end ;

procedure TOF_GCARTNONMVTES_MUL.OnLoad ;
begin
 Inherited ;
end ;

procedure TOF_GCARTNONMVTES_MUL.OnArgument (S : String ) ;
var DateCloture,DateFinExo : TDateTime;
begin
  Inherited ;
//Initialisation form
if (V_PGI.PassWord = CryptageSt(DayPass(Date))) then
    begin
    TToolbarButton97(GetControl('BParamListe')).Visible:=True ;
    end else
    begin
    TToolbarButton97(GetControl('BParamListe')).Visible:=False;
    end;
MajChampsLibresArticle(Ecran);
SetControlText('GA_DATECREATION', DateToStr(DebutDeMois(PlusMois(V_PGI.DateEntree,-1))));
TFMul(Ecran).FListe.Tag:=1;
TFMul(Ecran).Q.Manuel := True;
//Affectation
GA:=THGrid(GetControl('GA'));
GA.OnDblClick := GDblClick;
GA.OnRowEnter := GARowEnter;
DateCloture := GetParamSoc('SO_GCDATECLOTURESTOCK');
DateFinExo := iDate1900;
if DateCloture < VH^.Encours.Fin then DateFinExo:=VH^.Encours.Fin
else if DateCloture < VH^.Suivant.Fin then DateFinExo:=VH^.Suivant.Fin;

if DateCloture = iDate1900 then SetControlText('DATEPIECE', DateToStr(DateFinExo))
else SetControlText('DATEPIECE', DateToStr(DateCloture));
//Boutons
BImprimer  := TToolbarButton97(GetControl('BImprimer'));
BOuvrir  := TToolbarButton97(GetControl('BOuvrir'));
BValideDateSup  := TToolbarButton97(GetControl('BVALIDEDATESUP'));
BSelectAll := TToolbarButton97(GetControl('BSelectAll'));
BZoomArticle := TBitBtn(GetControl('BZOOMARTICLE'));
BZoomPiece := TBitBtn(GetControl('BZOOMPIECE'));
BImprimer.OnClick := BImprimerClick;
BOuvrir.OnClick := BOuvrirClick;
BValideDateSup.OnClick := BValideDateSupClick;
BSelectAll.OnClick := BSelectAllClick;
BZoomArticle.OnClick := BZoomArticleClick;
BZoomPiece.OnClick := BZoomPieceClick;
//Grids
GA.ListeParam := 'GCMULARTNONMVTES';
//Tobs
TobA := TOB.Create('',nil,-1) ;
//Etat
//PEtat.stSQL := 'GZK_UTILISATEUR = "'+ V_PGI.USer +'"';
//Menu zoom
PopZoom97(TToolbarButton97(GetControl('BZOOM')),TPopupMenu(GetControl('POPMENU'))) ;
  THEdit(GetControl('CODEARTICLE')).onElipsisClick := CodeArticleRech;
  THEdit(GetControl('CODEARTICLE_')).onElipsisClick := CodeArticleRech;

end ;

procedure TOF_GCARTNONMVTES_MUL.OnClose ;
begin
  Inherited ;
TobA.Free ;
end ;

///////////////////////// Boutons
procedure TOF_GCARTNONMVTES_MUL.BZoomArticleClick(Sender: TObject);
var TobAD : TOB;
begin
TobAD := TOB(GA.Objects[0,GA.Row]);
if TobAD = nil then exit;
ZoomArticle(TobAD.GetValue('GA_ARTICLE'),TobAD.GetValue('GA_TYPEARTICLE'),taConsult);
end;

procedure TOF_GCARTNONMVTES_MUL.BZoomPieceClick(Sender: TObject);
var TobAD : TOB;
    stDate : string;
begin
TobAD := TOB(GA.Objects[0,GA.Row]);
if TobAD = nil then exit;
try    stDate:=DateToStr(TobAD.GetValue('DATEPIECE'));
except stDate:=DateToStr(iDate1900);
end;
AGLLanceFiche('GC','GCPIECESCOURS_ART','GL_ARTICLE='+TobAD.GetValue('GA_ARTICLE')
                +';GL_DATEPIECE='+stDate
                +';GL_DEPOT='+GetControlText('DEPOT'),'','GL_NATUREPIECEG='+NaturePiece+';GL_VIVANTE=');
end;

procedure TOF_GCARTNONMVTES_MUL.BImprimerClick(Sender: TObject);
begin
{if TobA.Detail.Count = 0 then exit;
EntreeOptionEdit(PEtat.Tip,PEtat.Nat,PEtat.Modele,PEtat.Apercu,PEtat.DeuxPages,PEtat.First,
                 TPageControl(GetControl('Pages')),PEtat.stSQL,PEtat.Titre,PrepareImpression);
ExecuteSQL('DELETE FROM GCTMPARTNONMVTES WHERE ' + PEtat.stSQL);}
LanceEtatTob('E','GZL','GZL',TobA,True,False,False,TPageControl(GetControl('Pages')),'',Ecran.Caption,False);
end;

procedure TOF_GCARTNONMVTES_MUL.BSelectAllClick(Sender: TObject);
begin
GA.AllSelected := not GA.AllSelected ;
end;

procedure TOF_GCARTNONMVTES_MUL.BOuvrirClick(Sender: TObject);
var iRow : integer;
    okShow : boolean;
begin
okShow := False;
for iRow := 1 to GA.RowCount -1 do
    if GA.isSelected(iRow) then
       begin okShow := True ; break ; end;
if okShow then TToolWindow97(GetControl('PDATESUP')).Visible:=True
else PGIBox('Aucun article sélectionné',Ecran.Caption);
end;

procedure TOF_GCARTNONMVTES_MUL.BValideDateSupClick(Sender: TObject);
var iRow : integer;
    TobModif,TobAD,TobF : TOB;
begin
if not isValidDate(GetControlText('DATESUP')) then
   begin
   PGIBox('La date est incorrecte',Ecran.Caption);
   exit;
   end;
TobModif := TOB.Create('',nil,-1) ;
for iRow := 1 to GA.RowCount -1 do
    begin
    if GA.isSelected(iRow) then
       begin
       TobAD:=TOB(GA.Objects[0,iRow]);
       if TobAD=nil then continue;
       TobF := TOB.Create('ARTICLE',TobModif,-1);
       TobF.PutValue('GA_ARTICLE',TobAD.GetValue('GA_ARTICLE'));
       TobF.LoadDB();
       TobF.PutValue('GA_DATESUPPRESSION',StrToDate(GetControlText('DATESUP')));
       end;
    end;
if TobModif.Detail.count > 0 then
   begin
   if TobModif.UpdateDB() then PGIInfo(IntToStr(TobModif.Detail.Count)+' article(s) modifié(s)',Ecran.Caption)
   else PGIBox('La mise à jour ne s''est pas correctement effectuée',Ecran.Caption);
   RefreshGrille;
   end;
TToolWindow97(GetControl('PDATESUP')).Visible:=False;
TobModif.Free;
end;


///////////////////////// Grid
procedure TOF_GCARTNONMVTES_MUL.GDblClick(Sender: TObject);
var TobAD : TOB;
begin
TobAD := TOB(GA.Objects[0,GA.Row]);
if TobAD = nil then exit;
ZoomArticle(TobAD.GetValue('GA_ARTICLE'),TobAD.GetValue('GA_TYPEARTICLE'),taConsult);
end;

procedure TOF_GCARTNONMVTES_MUL.GARowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
AfficheDimension;
end;

function TOF_GCARTNONMVTES_MUL.CreerSelectSQL : string;
var TokenListe,st : string;
begin
//Champs obligatoires
ListeChampSQL := 'GA_ARTICLE,GA_CODEARTICLE,GA_TYPEARTICLE,GA_STATUTART,'
               + 'GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5,'
               + 'GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5';
TokenListe := ListeChampGrid;
st := Trim(ReadTokenSt(TokenListe));
while (st <> '') do
    begin
    if PrefixeToNum(Copy(st,1,pos('_',st)-1))>0 then
       ListeChampSQL := ListeChampSQL + ',' + st;
    st := Trim(ReadTokenSt(TokenListe));
    end;
result := ListeChampSQL;
end;

procedure TOF_GCARTNONMVTES_MUL.RefreshGrille;
var F : TFMul;
    ColDat,ColPhy : integer;
    NomList,FRecordSource,FLien,FSortBy,FFieldList,FLargeur,FAlignement,FParams,FPerso : string ;
    Ftitre,TT,NC : Hstring;
    OkTri,OkNumCol : boolean ;
    //
    ColDPA  : Integer;
    ColPMAP : Integer;
    NbCol   : Integer;
begin
  Inherited ;

  GA.VidePile(True) ;
  GA.AllSelected := False;

  GA.ColCount := 1;

  TobA.ClearDetail ;

  F:=TFMul(Ecran);

  ListeChampGrid := '';

  NomList := GA.ListeParam;

  ChargeHListe(NomList,FRecordSource,FLien,FSortBy,FFieldList,FTitre,FLargeur,FAlignement,FParams,tt,NC,FPerso,OkTri,OkNumCol);

  ListeChampGrid := FFieldList;

  {$IFDEF EAGLCLIENT}
  //GA.ColCount := F.FListe.ColCount+3;
  NbCol := F.FListe.Columns.ColCount;
  {$ELSE}
  //GA.ColCount := F.FListe.Columns.Count+3;
  NbCol := F.FListe.Columns.Count;
  {$ENDIF}

  ListeChampGrid := ListeChampGrid + 'QTEPHY';

  Inc(NbCol);
  GA.ColCount := NbCol+1;
  //
  ColPhy := NbCol;
  if GetControlText('DEPOT')='' then
    GA.Cells[ColPhy, 0] := TraduireMemoire('Total stock')
  else
    GA.Cells[ColPhy, 0] := TraduireMemoire('Stock');
  GA.ColWidths[ColPhy]:=130;
  GA.ColAligns[ColPhy] := taRightJustify;

  if Pos('DPA',ListeChampGrid) > 0  then
  begin
    ListeChampGrid := ListeChampGrid + ';TOTALDPA';
    Inc(NbCol);
    GA.ColCount := NbCol+1;
    ColDPA := NbCol;
    GA.Cells[ColDPA, 0]  := TraduireMemoire('Total DPA');
    GA.ColTypes[ColDPA]  := 'D';
    GA.ColWidths[ColDPA] := 130;
    GA.ColAligns[ColDPA] := taRightJustify;
  end;

  if Pos('PMAP', ListeChampGrid) > 0  then
  Begin
    ListeChampGrid := ListeChampGrid + ';TOTALPMAP';
    Inc(NbCol);
    GA.ColCount := NbCol+1;
    ColPMAP := NbCol;
    GA.Cells[ColPMAP, 0]  := TraduireMemoire('Total PMAP');
    GA.ColTypes[ColPMAP]  := 'D';
    GA.ColWidths[ColPMAP] := 130;
    GA.ColAligns[ColPMAP] := taRightJustify;
  end;

  ListeChampGrid  := ListeChampGrid + ';DATEPIECE';

  GA.Titres[0]    := ListeChampGrid;

  Inc(NbCol);
  GA.ColCount := NbCol+1;

  ColDat := NbCol;
  GA.Cells[ColDat, 0]   := TraduireMemoire('Dernier mouvement');
  GA.ColTypes[ColDat]   := 'D';
  GA.ColWidths[ColDat]  :=130;
  GA.ColAligns[ColDat]  := taCenter;

  ListeChampSQL := CreerSelectSQL;

  HMTrad.ResizeGridColumns(GA) ;

  ChargeTobArticle;

  TobA.Detail.Sort('DATEPIECE;GA_CODEARTICLE');
  TobA.PutGridDetail(GA,false,false,ListeChampGrid);

  AfficheDimension;

  THPanel(GetControl('PCUMULART')).Caption := 'Totaux  ('+IntToStr(TobA.Detail.Count)+' lignes)';

end ;

///////////////////////// Tob
procedure TOF_GCARTNONMVTES_MUL.ChargeTobArticle;
var TobAD : TOB;
    stSql,stWhere, stWhereDate, stWhereLigne : string;
    xx_where,Depot,SelectDispo,WhereDispo : string;
    iInd : integer;
    QQ : TQuery;
    QteStock : Double;
    DPA      : Double;
    PMAP     : Double;
begin
Depot := GetControlText('DEPOT');
if Depot <> '' then
   begin
   SelectDispo := 'GQ_PHYSIQUE';
   WhereDispo := ' AND GQ_DEPOT="'+Depot+'"';
   end else
   begin
   SelectDispo := 'SUM(GQ_PHYSIQUE)';
   WhereDispo := '';
   end;
stWhereLigne := RecupWhereCritereLigne;
stWhereDate := ' AND GL_DATEPIECE>="' + USDateTime(strToDate(GetControlText('DATEPIECE'))) + '"';
//xx_where:=GetControlText('XX_WHERE') ; SetControlText('XX_WHERE','');
//SetControlText('XX_WHERE',xx_where);

stWhere := RecupWhereCritere(TFMul(Ecran).Pages);
if stWhere <> '' then  System.Delete(stWhere, pos ('WHERE ', stWhere), 6);

if GetControlText('CODEARTICLE') <> '' then
    stWhere := stWhere + ' AND GA_ARTICLE>="' + GetControlText('CODEARTICLE') + '"';
if GetControlText('CODEARTICLE_') <> '' then
    stWhere := stWhere + ' AND GA_ARTICLE<="' + format('%-18.18s%16.16s',[GetControlText ('CODEARTICLE_'),StringOfChar('Z',16)]) + '"';

stSql :=  'SELECT ' + ListeChampSQL + ' FROM ARTICLE WHERE ' + stWhere
       + 'AND NOT EXISTS '
       + '(SELECT GL_ARTICLE FROM LIGNE WHERE GL_ARTICLE=ARTICLE.GA_ARTICLE'+stWhereLigne+stWhereDate+') AND NOT EXISTS '
       + '(SELECT GLN_ARTICLE FROM LIGNENOMEN LEFT JOIN LIGNE ON LIEN(LIGNENOMEN,LIGNE) WHERE GLN_ARTICLE=ARTICLE.GA_ARTICLE'+stWhereLigne+stWhereDate+')';

  DPA   := 0;
  PMAP  := 0;

  TobA.LoadDetailFromSQL(stSql,false,true);
  InitMove(TobA.Detail.Count,'');
for iInd := 0 to TobA.Detail.Count -1 do
    begin
    TobAD := TobA.Detail[iInd];
    //
    if TOBAD.FieldExists('GA_DPA')  then DPA   := TOBAD.GetDouble('GA_DPA');
    if TOBAD.FieldExists('GA_PMAP') then PMAP  := TOBAD.GetDouble('GA_PMAP');
    //
    TobAD.AddChampSupValeur('DATEPIECE',iDate1900);
    QQ := OpenSQL('SELECT MAX(GL_DATEPIECE) AS DATEPIECE FROM LIGNE WHERE GL_ARTICLE="'+TobAD.GetValue('GA_ARTICLE')+'" ' + stWhereLigne,True,-1,'',true);
    if not QQ.Eof then TobAD.PutValue('DATEPIECE',QQ.FindField('DATEPIECE').AsDateTime);
    Ferme(QQ);
    //
    TobAD.AddChampSupValeur('DATEPIECECOMPOSE',iDate1900);
    QQ := OpenSQL('SELECT MAX(GL_DATEPIECE) AS DATEPIECECOMPOSE FROM LIGNENOMEN LEFT JOIN LIGNE ON LIEN(LIGNENOMEN,LIGNE) WHERE GLN_ARTICLE="'+TobAD.GetValue('GA_ARTICLE')+'" ' + stWhereLigne,True,-1,'',true);
    if not QQ.Eof then TobAD.PutValue('DATEPIECECOMPOSE',QQ.FindField('DATEPIECECOMPOSE').AsDateTime);
    Ferme(QQ);
    //
    QteStock := 0;
    TobAD.AddChampSupValeur('QTEPHY',0);
    TobAD.AddChampSupValeur('TOTALDPA',0);
    TobAD.AddChampSupValeur('TOTALPMAP',0);
    //
    QQ := OpenSQL('SELECT '+SelectDispo+' AS QTEPHY FROM DISPO WHERE GQ_ARTICLE="' + TobAD.GetValue('GA_ARTICLE') + '" AND GQ_CLOTURE="-"' +WhereDispo,True,-1,'',true);
    QteStock := QQ.FindField('QTEPHY').AsFloat;
    if not QQ.Eof then
    begin
      TobAD.PutValue('QTEPHY',    Arrondi(QteStock, V_PGI.OkDecQ));
      TobAD.PutValue('TOTALDPA',  Arrondi(QteStock*DPA, V_PGI.OkDecP));
      TobAD.PutValue('TOTALPMAP', Arrondi(QteStock*PMAP,V_PGI.OkDecP));
    end;
    Ferme(QQ);
    //
    MoveCur(False) ;
    {$IFDEF EAGLCLIENT}
    if not isValidDate(TobAD.GetValue('DATEPIECE')) then
    {$ELSE}
    if TobAD.GetValue('DATEPIECE') = null then
    {$ENDIF}
       TobAD.PutValue('DATEPIECE',idate1900);
    {$IFDEF EAGLCLIENT}
    if not isValidDate(TobAD.GetValue('DATEPIECECOMPOSE')) then
    {$ELSE}
    if TobAD.GetValue('DATEPIECECOMPOSE') = null then
    {$ENDIF}
       TobAD.PutValue('DATEPIECECOMPOSE',idate1900);

    if TobAD.GetValue('DATEPIECECOMPOSE') > TobAD.GetValue('DATEPIECE') then
       TobAD.PutValue('DATEPIECE',TobAD.GetValue('DATEPIECECOMPOSE'));
    if TobAD.GetValue('QTEPHY') = null then
       TobAD.PutValue('QTEPHY',Arrondi(0,V_PGI.OkDecQ));
    //
    end;

  FiniMove;

end;

function TOF_GCARTNONMVTES_MUL.RecupWhereCritereLigne : string;
var stFlux,stWhere : string;
    i_ind : integer;
    TobPP : TOB;
begin
stWhere := ''; NaturePiece := '';
stFlux := GetControlText('VENTEACHAT');
for i_ind := 0 to VH_GC.TOBParPiece.Detail.Count -1 do
    begin
    TobPP := VH_GC.TOBParPiece.Detail[i_ind];
    if Pos(TobPP.GetValue('GPP_VENTEACHAT'),stFlux) > 0 then
        begin
        if stWhere <> '' then stWhere := stWhere + ' OR ';
        stWhere := stWhere + 'GL_NATUREPIECEG="'+TobPP.GetValue('GPP_NATUREPIECEG')+'"';
        if NaturePiece = '' then NaturePiece := TobPP.GetValue('GPP_NATUREPIECEG')
        else NaturePiece := NaturePiece + ',' + TobPP.GetValue('GPP_NATUREPIECEG');
        end;
    end;
if stWhere <> '' then stWhere := ' AND (' + stWhere + ')';
stWhere := stWhere + GetCondition(stWhere,'AND','DEPOT');
result := stWhere;
end;

function TOF_GCARTNONMVTES_MUL.GetCondition(Where,Plus,NomChamp : string) : string;
begin
result:='';
if GetControlText(NomChamp) <> '' then
   begin
   if Where <> '' then result := ' ' + Plus + ' ';
   result := result + 'GL_'+NomChamp+'="' + GetControlText(NomChamp) + '"';
   end;
end;


procedure TOF_GCARTNONMVTES_MUL.AfficheDimension;
var TobAD : TOB;
begin
TobAD := TOB(GA.Objects[0,GA.Row]);
if TobAD = nil then
   begin
   SetControlVisible('PDIMENSION',False);
   SetControlText('TDIMENSION','');
   exit;
   end;
if TobAD.GetValue('GA_STATUTART')='DIM' then
   begin
   SetControlVisible('PDIMENSION',True);
   SetControlText('TDIMENSION',RechDim(TobAD));
   end else
   begin
   SetControlVisible('PDIMENSION',False);
   SetControlText('TDIMENSION','');
   end;
end;

function TOF_GCARTNONMVTES_MUL.RechDim(TobArt : TOB) : string;
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


procedure TOF_GCARTNONMVTES_MUL.CodeArticleRech(Sender: Tobject);
Var CodeArt : String;
    StChamps: String;
begin

  CodeArt := THEdit(sender).text;
  StChamps:= '';

  if CodeArt <> '' then StChamps := 'GA_CODEARTICLE=' + Trim(Copy(CodeArt, 1, 18)) + ';';

  StChamps:= StChamps + 'XX_WHERE=GA_TYPEARTICLE="MAR" OR GA_TYPEARTICLE="ARP" OR GA_TYPEARTICLE="OUV"';

  CodeArt := AGLLanceFiche('BTP', 'BTARTICLE_RECH', '', '', StChamps + ';GA_TENUESTOCK=X;ACTION=CONSULTATION;STATUTART=UNI,DIM');

  if codeArt <> '' then
  begin
    THEdit(Sender).text := CodeArt;
  end;

end;


Initialization
  registerclasses ( [ TOF_GCARTNONMVTES_MUL ] ) ;
end.

