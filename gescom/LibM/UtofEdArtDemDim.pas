unit UtofEdArtDemDim;

interface       
uses  StdCtrls,Controls,Classes,Windows,
{$IFDEF EAGLCLIENT}
{$ELSE}
      dbTables,DBGrids, db,
{$ENDIF}
      forms,sysutils,ComCtrls,EntGC,
      HCtrls,HEnt1,HMsgBox,UTOF, vierge, HDimension,UTOB,UtilDimArticle, UtilArticle,
      UdimArticle, AglInit,M3FP;
Type
     TOF_EDARTDEMDIM = Class (TOF)
        DimensionsArticle : TODimArticle ;
        Q : TQuery ;
     private
        dimAction : string ;  // SAISIE, MODIF
        dimChamp : string ;
        dimMasque : string ;
        CodeArticle, CodeDepot, NumLigne : string ;
        Top, Left,Height,Width: Integer ;
        AuDessus: Boolean ;
        procedure InitDimensionsArticle ;
        Procedure AdapterFiche ;
        procedure RechargeNbEtq ;
     public
        procedure OnArgument (Arguments : String ) ; override ;
        procedure OnUpdate; override ;
        procedure OnClose; override ;
        Destructor Destroy ; override ;
     END ;

procedure AGLOnClickParamGrilleEti( parms: array of variant; nb: integer ) ;

implementation

destructor TOF_EDARTDEMDIM.destroy ;
begin
inherited ;
end;

procedure TOF_EDARTDEMDIM.OnArgument(Arguments : String ) ;
var Lequel : string ;
    Critere : string ;
    ChampMul,ValMul : string;
    x : integer ;
    SQL : string ;
begin
inherited ;
dimAction:='SELECT';
dimChamp:='' ;

Repeat
    Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));

           if ChampMul='TOP' then Top:=StrToInt(ValMul) ;
           if ChampMul='LEFT' then Left:=StrToInt(ValMul) ;
           if ChampMul='HEIGTH' then Height:=StrToInt(ValMul) ;
           if ChampMul='WIDTH' then Width:=StrToInt(ValMul) ;
           if ChampMul='OU' then Audessus:=Boolean(ValMul='X') ;
           if ChampMul='GA_CODEARTICLE' then Lequel := ValMul;
           if ChampMul='ACTION' then dimAction := ValMul;
           if ChampMul='CHAMP' then dimChamp := ValMul;
           if ChampMul='DEPOT' then CodeDepot := ValMul;
           if ChampMul='NUMLIGNE' then NumLigne := ValMul;
           end;
        end;
until  Critere='';

SQL := 'SELECT GA_ARTICLE, GA_CODEARTICLE, GA_DIMMASQUE ' ;
if dimChamp<>'' then SQl:=SQL+','+dimChamp ;
SQL := SQL+' from Article where GA_CODEARTICLE="'+Lequel+'"' ;
Q:=OpenSQl(SQL,True) ;
if Q.EOF  then begin Ferme(Q) ; TFVierge(Ecran).Close ; Exit ; end;
if GetControl('GA_ARTICLE') <> nil then SetControlText('GA_ARTICLE',Lequel);
//SetControlText('DIM_CHAMP_AFFICHE',dimChamp);
CodeArticle:=Q.FindField('GA_CODEARTICLE').AsString ;
SetControlText('GA_CODEARTICLE',CodeArticle) ;
dimMasque:=Q.FindField('GA_DIMMASQUE').AsString ;
Ferme(Q) ;
InitDimensionsArticle ;
AdapterFiche;
end;

procedure TOF_EDARTDEMDIM.OnClose ;
begin
Inherited;
if DimensionsArticle<>nil then
   BEGIN
   DimensionsArticle.destroy;
   DimensionsArticle:=nil;
   END;
end;

Procedure TOF_EDARTDEMDIM.AdapterFiche ;
Var CoordDernCol,CoordDernLign: TRect ;
GridDim: THGrid ;
ValOnglet : Integer ;
begin
Ecran.Width:=1000 ; // pour pouvoir voir toutes les colonnes
Ecran.Height:=1000 ; // pour pouvoir voir toutes les lignes
THDimension(GetControl('FDIM')).Align:=alNone ;
GridDim:=DimensionsArticle.Dim.GridDim ;
if DimensionsArticle.Dim.DimOngl=nil then ValOnglet:=0 else ValOnglet:=GridDim.DefaultRowHeight ;
CoordDernCol:=GridDim.CellRect (GridDim.VisibleColCount+GridDim.FixedCols ,1) ;
CoordDernLign:=GridDim.CellRect (2,GridDim.VisibleRowCount+GridDim.FixedRows) ; // Dernière ligne
GridDim.ScrollBars:=ssNone ;

If (CoordDernLign.Bottom>Height) or (CoordDernLign.Bottom+76>Height) then
     begin
       Ecran.Height:=Height ;
       GridDim.ScrollBars:=ssVertical ;
     end else
     Ecran.Height:=CoordDernLign.Bottom  + 76 + ValOnglet ; // 76= titre et panel bouton de ecran + encadrment objet Dim
If (CoordDernCol.Right>Width) or (CoordDernCol.Right+20>Width) then
     begin
       Ecran.Width:=Width ;
       if GridDim.ScrollBars=ssVertical then GridDim.ScrollBars:=ssBoth else GridDim.ScrollBars:=ssHorizontal ;
     end
     else if CoordDernCol.Right< 230 then Ecran.Width:=230
     else Ecran.Width:=CoordDernCol.Right+20 ;

// Ajustement taille si ajout d'un ascenseur
if (GridDim.ScrollBars=ssVertical) or (GridDim.ScrollBars=ssBoth) then
   begin
     if (CoordDernCol.Right+16 >Width) or (CoordDernCol.Right+16+20 >Width) then Ecran.Width:=Width
     else Ecran.Width:=CoordDernCol.Right+16+20 ; //ScrollBars+encadrement Dim
   end ;
if (GridDim.ScrollBars=ssHorizontal) or (GridDim.ScrollBars=ssBoth) then
  begin
    if (CoordDernLign.Bottom+16>Height) or (CoordDernLign.Bottom+16+76>Height) then Ecran.Height:=Height
    else Ecran.Height:=CoordDernLign.Bottom+ValOnglet+16+76 ;//ScrollBars+titre,panel bouton de ecran et encadrment objet Dim
  end ;
THDimension(GetControl('FDIM')).Align:=alClient ;
if Audessus then Ecran.Top:= Top - Ecran.Height else Ecran.Top:=Top ;
  Ecran.Left:=Left ;
End ;

procedure TOF_EDARTDEMDIM.InitDimensionsArticle ;
begin
    if DimensionsArticle<>nil then
        begin
        DimensionsArticle.free;
        end;
    DimensionsArticle:=TODimArticle.Create(THDimension(GetControl('FDIM'))
                         ,CodeArticle
                         ,dimMasque
                         ,dimChamp,'GCDIMCHAMP','ETA','', '', '', False);
If DimensionsArticle.TOBArticleDim=nil then exit ;
    DimensionsArticle.Dim.PopUp.Items[2].Visible:=False ; // Menu Existant invisible
    DimensionsArticle.Dim.PopUp.Items[3].Visible:=False ; // Menu Inexistant invisible
//    DimensionsArticle.Dim.OnChange := OnChangeItemdim;

SetControlVisible('BPARAMDIM',True) ;
SetControlVisible('CBDETAIL',False) ;
if dimAction = 'MODIF' then RechargeNbEtq;
OnChangeAfficheChampDimMul(DimensionsArticle,'ETA') ;
end;

procedure TOF_EDARTDEMDIM.OnUpdate ;
var TobT,TobSelect : TOB;
     i, y : integer;
begin
DimensionsArticle.ChangeChampDimMul (false);
if DimensionsArticle.TOBArticleDim <> nil then
  begin
  if not DimensionsArticle.IsModified then exit;
  for i:=0 to DimensionsArticle.TOBArticleDim.detail.count-1 do
    begin
    TobT:=DimensionsArticle.TOBArticleDim.detail[i] ;
    TobSelect := LaTob.FindFirst(['GA_ARTICLE','DEPOT','NUMLIGNE'],[TobT.GetValue('GA_ARTICLE'),CodeDepot,NumLigne],false);
    if TobSelect <> nil then
      begin
      y := TobSelect.GetIndex;
      LaTob.Detail[y].free ;
      end;
    end;
  for i:=0 to DimensionsArticle.TOBArticleDim.detail.count-1 do
    begin
    TobT:=DimensionsArticle.TOBArticleDim.detail[i] ;
    if TobT.GetValue('NBETIQ') <> 0 then
      begin
      TobSelect:=tob.create('articles select',LaTob, -1 ) ;
      TobSelect.AddChampSup( 'GA_CODEARTICLE',false);
      TobSelect.AddChampSup( 'GA_ARTICLE',false);
      TobSelect.AddChampSup( 'DEPOT',false);
      TobSelect.AddChampSup( 'NBETIQDIM',false);
      TobSelect.AddChampSup( 'NUMLIGNE',false);
      TobSelect.PutValue('GA_CODEARTICLE',TobT.GetValue('GA_CODEARTICLE')) ;
      TobSelect.PutValue('GA_ARTICLE',TobT.GetValue('GA_ARTICLE')) ;
      TobSelect.PutValue('DEPOT',CodeDepot) ;
      TobSelect.PutValue('NBETIQDIM',TobT.GetValue('NBETIQ')) ;
      TobSelect.PutValue('NUMLIGNE',NumLigne) ;
      end;
    end;
  end;
TheTob := LaTob;
TFVierge(Ecran).close ;
end ;

procedure TOF_EDARTDEMDIM.RechargeNbEtq ;
var TobT,TobSelect : TOB;
     i : integer;
begin
if (DimensionsArticle.TOBArticleDim <> nil)
 and (LaTob <> nil) then
  begin
  for i:=0 to DimensionsArticle.TOBArticleDim.detail.count-1 do
    begin
    TobT:=DimensionsArticle.TOBArticleDim.detail[i] ;
    TobSelect := LaTob.FindFirst(['GA_ARTICLE','DEPOT','NUMLIGNE'],[TobT.GetValue('GA_ARTICLE'),CodeDepot,NumLigne],false);
    if TobSelect <> nil then
      begin
      TobT.PutValue('NBETIQ',TobSelect.GetValue('NBETIQDIM')) ;
//      TobT.PutValue('NUMLIGNE',TobSelect.GetValue('NUMLIGNE')) ;
      end;
    end;
  end;
end ;

procedure AGLOnClickParamGrilleEti( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF : TOF ;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit;
{parm[1] contient la nature des champs affichés : préférences article/stock/achat/vente/autre}
  if (TOTOF is TOF_EDARTDEMDIM) then
  begin
    ParamGrille (TOF_EDARTDEMDIM(TOTOF).DimensionsArticle, 'ETA', '');
    TOF_EDARTDEMDIM(TOTOF).AdapterFiche ;
  end else exit ;
end;

Initialization
registerclasses([TOF_EDARTDEMDIM]);
RegisterAglProc('OnClickParamGrilleEti',TRUE,0,AGLOnClickParamGrilleEti);
end.
