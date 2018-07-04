unit UTofGcTarifMode_Mul;

interface
uses  StdCtrls,Controls,Classes,db,forms,sysutils,dbTables,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, mul, DBGrids, HDimension,Entgc,Windows,Utob,AGLInit,
      Fe_Main,M3FP,UtilDimArticle,TarifTiersMode,TarifArticleMode;

Type
        TOF_GcTarifMode_Mul = Class (TOF)

                Grid: THGrid ;
                procedure OnArgument(St:String) ; override;
                procedure OnLoad ; override ;
                procedure OnVisualiseTarif(CodeTarif,CodeArticle,TarifArticle,CodeTiers,TarifTiers: String) ;
                procedure InitEntete ;

     END ;

implementation

procedure TOF_GcTarifMode_Mul.OnArgument(St: String) ;
Begin
InitEntete ;
end ;

procedure TOF_GcTarifMode_Mul.OnLoad ;
begin
Grid:=THGrid(TFmul(Ecran).FListe) ;
SetControlText('XX_WHERE','GA_STATUTART="GEN" or GA_STATUTART="UNI" or GF_ARTICLE is null or GF_ARTICLE=""') ;
end ;

procedure TOF_GcTarifMode_Mul.InitEntete ;
var QTyp,QPer: TQuery ;
Begin
QTyp:=OpenSQL('Select GFT_DEVISE,GFT_CODETYPE from TarifTypMode where GFT_DERUTILISE="X"',True) ;
if Not QTyp.EOF then
   begin
   THValComboBox(GetControl('GFM_TYPETARIF')).Value:=QTyp.FindField('GFT_CODETYPE').AsString ;
   THValComboBox(GetControl('GF_DEVISE')).Value:=QTyp.FindField('GFT_DEVISE').AsString ;
   end;
QPer:=OpenSQL('Select * from TarifPer where GFP_DERUTILISE="X"',True) ;
if Not QTyp.EOF then
   begin
   THValComboBox(GetControl('GFM_PERTARIF')).Value:=QPer.FindField('GFP_CODEPERIODE').AsString ;
   THEdit(GetControl('GF_DATEDEBUT')).Text:=QPer.FindField('GFP_DATEDEBUT').AsString ;
   THEdit(GetControl('GF_DATEFIN')).Text:=QPer.FindField('GFP_DATEFIN').AsString ;
   end ; 
Ferme(QTyp) ;
Ferme(QPer) ;
End ;


Procedure TOF_GcTarifMode_Mul.OnVisualiseTarif(CodeTarif,CodeArticle,TarifArticle,CodeTiers,TarifTiers: String) ;
var tobTemp:TOB ;
begin
if (CodeArticle<>'') or (TarifArticle<>'') and (Codetiers='') and (TarifTiers='')  then
 begin
  TobTemp:=TOB.Create('_INIT',NIL,-1);
  TobTemp.AddChampSup('_CodeTarif',False) ;
  TobTemp.AddChampSup('_CodeArticle',False) ;
  TobTemp.AddChampSup('_TarifArticle',False) ;
  TobTemp.PutValue('_CodeTarif',CodeTarif) ;
  TobTemp.PutValue('_CodeArticle',CodeArticle) ;
  TobTemp.PutValue('_TarifArticle',TarifArticle) ;
  TheTob:=TobTemp ;
  EntreeTarifArticleMode(taModif,TRUE) ;
 end ;
if (CodeTiers<>'') or (TarifTiers<>'') then
  begin
  TobTemp:=TOB.Create('_INIT',NIL,-1);
  TobTemp.AddChampSup('_CodeTarif',False) ;
  TobTemp.AddChampSup('_CodeTiers',False) ;
  TobTemp.AddChampSup('_TarifTiers',False) ;
  TobTemp.PutValue('_CodeTarif',CodeTarif) ;
  TobTemp.PutValue('_CodeTiers',CodeTiers) ;
  TobTemp.PutValue('_TarifTiers',TarifTiers) ;
  TheTob:=TobTemp ;
  EntreeTarifTiersMode(taModif,TRUE) ;
  end ;
TobTemp.Free ; TobTemp:=nil ;
end ;
{var CoordonneEcran: TPoint ;
Top,Left,Height,Width,i_ind: Integer ;
TobTarif, TobTarfArtDim: Tob;
QArtDim,QTarifDim: TQuery ;
begin
TOBTarif := TOB.Create ('', Nil, -1) ;
TobTarfArtDim:= TOB.Create ('', TOBTarif, -1) ;
QArtDim:=OpenSql('Select GA_ARTICLE from Article where GA_CODEARTICLE="'+TRIM(copy(CodeArticle,1,18))+'" And GA_STATUTART="DIM"', True) ;
if QArtDim.EOF then
begin
ferme(QArtDim);
if TFmul(Ecran).Q.FindField('GF_Article').asstring<>'' then
   aglLanceFiche('GC','GCARTICLE','',TFmul(Ecran).Q.FindField('GF_Article').asstring,'ACTION=CONSULTATION') ;
exit ;
End ;
while not QartDim.EOF do
    begin
    QTarifDim := OpenSQL('SELECT * FROM TARIF WHERE GF_ARTICLE="'+QArtDim.FindField('GA_ARTICLE').AsString+'" AND GF_TARFMODE="'+TFmul(Ecran).Q.FindField('GF_TARFMODE').asstring+'" AND GF_DEPOT="'+TFmul(Ecran).Q.FindField('GF_DEPOT').asstring+'"',True) ;
    //if QTarifDim.EOF then exit ;
    TobTarfArtDim.LoadDetailDB('TARIF',QTarifDim.FindField('GF_TARIF').AsString,'',nil,True,False) ;
    QArtDim.next ;
    ferme(QTarifDim) ;
    end ;
ferme(QArtDim) ;
TheTob:=TobTarfArtDim ;
CoordonneEcran:=RetourneCoordonneeCellule(2,Grid)  ;
Top:=CoordonneEcran.y ;
Left:=CoordonneEcran.x ;
Height:=Grid.Height-Grid.RowHeights [0]-Grid.RowHeights [1];
Width:=Grid.Width-Grid.ColWidths [0]  ;
V_PGI.FormCenter:=False ;
AglLanceFiche ('GC','GCMSELECTDIMDOC','','', 'GA_CODEARTICLE='+Trim(Copy(CodeArticle,1,18))+';ACTION=CONSULT;TOP='+IntToStr(Top)+';LEFT='+IntToStr(Left)+';HEIGTH='+IntToStr(Height)+';WIDTH='+IntToStr(Width)+';TARIF=X;TYPEPARAM=TAF') ;
TobTarfArtDim:=TheTob ;
TheTob:=Nil ;
end ; }

procedure AGLOnVisualiseTarif( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFMul) then TOTOF:=TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_GcTarifMode_Mul) then TOF_GcTarifMode_Mul(TOTOF).OnVisualiseTarif(parms[1],parms[2],parms[3],parms[4],parms[5])
  else PGIBOx('Essai','essai') ;
end;

Initialization
registerclasses([TOF_GcTarifMode_Mul]) ;
RegisterAglProc('OnVisualiseTarif',TRUE,5,AGLOnVisualiseTarif) ;
end.
