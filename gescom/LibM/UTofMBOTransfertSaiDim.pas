unit UTofMBOTransfertSaiDim;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF,Windows,ent1, SaisUtil,
{$IFDEF EAGLCLIENT}
{$ELSE}
      DBGrids, dbTables,db,
{$ENDIF}
       vierge, HDimension,UTOB, UtilArticle,UdimArticle, UtilPGI,
       AglInit,entgc,UtilDimArticle,M3FP,UtofMBOParamDim,ParamSoc,FactUtil,UtilTarif ,Dialogs;

Type
     TOF_MBOTransfertSaiDim = Class (TOF)
        DimensionsArticle : TODimArticle ;
        Q : TQuery ;
     private
        Dim_EStkIni,Dim_EStkFin,Dim_EStkMin,Dim_EStkMax,Dim_EProp,
        Dim_RStkIni,Dim_RStkFin,Dim_RStkMin,Dim_RStkMax,Dim_RProp : integer;
        dimAction : string ;  // SELECT, SAISIE, CONSULT, MULTI
        CodeArticle : string ;
        NatureDoc,NaturePiece,CodeDepot,MultiDepot: String ;
        Top, Left,Height,Width: Integer ;
        AuDessus : Boolean ;
        NotClose,bValider : Boolean ;
        procedure InitDimensionsArticle ;
        procedure InitialiserColonne;
        procedure AdapterFiche ;
        procedure OnChangeItem(Sender: TObject);
        procedure RegroupeTousLesDepots;
        function  VerifStock(ItemDim:THDimensionItem):Boolean ;
        Procedure AffecteDEP(TOBMere:TOB; CodeArt:String; EProp,EManquant,ESurplus,EFinal:Double);
     public
        procedure OnArgument (Arguments : String ) ; override ;
        procedure OnLoad ; override ;
        procedure OnUpdate; override ;
        procedure OnClose ; override ;
     END ;

const RepCodeEmett:string='...';

implementation

procedure TOF_MBOTransfertSaiDim.InitialiserColonne;
var iChamp : Integer ;
    NomChamp : string;
begin
Dim_EStkIni:=-1; Dim_EStkFin:=-1; Dim_EStkMin:=-1; Dim_EStkMax:=-1; Dim_EProp:=-1;
Dim_RStkIni:=-1; Dim_RStkFin:=-1; Dim_RStkMin:=-1; Dim_RStkMax:=-1; Dim_RProp:=-1;
for iChamp := 1 to MaxDimChamp do
  begin
  NomChamp:=DimensionsArticle.OldDimChamp [iChamp];
  if NomChamp='GTL_STOCKINITIAL' then Dim_RStkIni:=iChamp
  else if NomChamp='STOCKFINAL' then Dim_RStkFin:=iChamp
  else if NomChamp='GTL_STOCKMINI' then Dim_RStkMin:=iChamp
  else if NomChamp='GTL_STOCKMAXI' then Dim_RStkMax:=iChamp
  else if NomChamp='GTL_PROPOSITION' then Dim_RProp:=iChamp
  else if NomChamp='DEP_STOCKINITIAL' then Dim_EStkIni:=iChamp
  else if NomChamp='DEP_STOCKFINAL' then Dim_EStkFin:=iChamp
  else if NomChamp='DEP_STOCKMINI' then Dim_EStkMin:=iChamp
  else if NomChamp='DEP_STOCKMAXI' then Dim_EStkMax:=iChamp
  else if NomChamp='DEP_PROPOSITION' then Dim_EProp:=iChamp
  ;
  end ;
end;

procedure TOF_MBOTransfertSaiDim.OnArgument(Arguments : String) ;
var Critere : string ;
    ChampMul,ValMul : string;
    x : integer ;
begin
inherited ;
dimAction:='SAISIE';
Repeat
  Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
  if Critere<>'' then
      begin
      x:=pos('=',Critere);
      if x<>0 then
         begin
         ChampMul:=copy(Critere,1,x-1);
         ValMul:=copy(Critere,x+1,length(Critere));
         if ChampMul='TYPEPARAM' then NatureDoc:=ValMul
         else if ChampMul='GA_CODEARTICLE' then CodeArticle:=ValMul
         else if ChampMul='ACTION' then dimAction:=ValMul
         else if ChampMul='TOP' then Top:=StrToInt(ValMul)
         else if ChampMul='LEFT' then Left:=StrToInt(ValMul)
         else if ChampMul='HEIGTH' then Height:=StrToInt(ValMul)
         else if ChampMul='WIDTH' then Width:=StrToInt(ValMul)
         else if ChampMul='OU' then Audessus:=Boolean(ValMul='X')
         else if ChampMul='CODEDEPOT' then CodeDepot:=ValMul
         else if ChampMul='MULTIDEPOT' then MultiDepot:=ValMul //1=Multi-dépôts 2=Mono dépôt
         ;
         end;
      end;
until  Critere='';
if (NatureDoc=NAT_PROPSAI) and (MultiDepot='2') then NatureDoc:=NAT_PROPDEP ;
NaturePiece:=NatureDoc ;
bValider:=False ;
TheTob:=LaTob ;
InitDimensionsArticle ;
TheTob:=nil;
End ;

procedure TOF_MBOTransfertSaiDim.InitDimensionsArticle ;
var detail:string ;
begin
if DimensionsArticle<>nil then DimensionsArticle.free ;
DimensionsArticle:=TODimArticle.Create(THDimension(GetControl('FDIM'))
                     ,CodeArticle
                     ,''
                     ,'','GCDIMCHAMP',NatureDoc,NaturePiece,CodeDepot,'-',True) ;
If DimensionsArticle.TOBArticleDim=nil then exit ;
DimensionsArticle.Dim.PopUp.Items[2].Visible:=False ; // Menu Existant invisible
DimensionsArticle.Dim.PopUp.Items[3].Visible:=False ; // Menu Inexistant invisible
DimensionsArticle.Dim.OnChange:=OnChangeItem ;
//DimensionsArticle.Dim.OnDblClick:=OnDoubleClick ;
if DimAction='CONSULT' then DimensionsArticle.Action := taConsult ;
detail:='-' ;
{Affichage des champs par défaut de l'utilisateur}
AfficheUserPref(DimensionsArticle,NatureDoc,'') ;
SetControlVisible('BPARAMDIM',NaturePieceGeree(naturePiece)) ;
end;

procedure TOF_MBOTransfertSaiDim.OnLoad ;
var CaptionCompl : string;
begin
 inherited ;
If DimensionsArticle.TOBArticleDim=nil then
  begin
  SetControlVisible('LPasDimensions',True) ;
  SetControlEnabled('BPARAMDIM',False) ;
  Ecran.Caption:= 'Saisie articles: ' + LibelleArticleGenerique(CodeArticle) ;
  //if Audessus then Ecran.Top:=Top-Ecran.Height else Ecran.Top:=Top ;
  //Ecran.Left:=Left ;
  Exit ;
  end ;
Ecran.Width:=1000 ; // pour pouvoir diminuer la fiche dans Adapter fiche
Ecran.Height:=1000 ;
//MAJ:=False ;
if MultiDepot='1' then CaptionCompl:=' (Multi dépôts)' else CaptionCompl:=RechDom('GCDEPOT',CodeDepot,True);
Ecran.Caption:= 'Saisie articles: ' + LibelleArticleGenerique(CodeArticle) + CaptionCompl;
AdapterFiche ;
//if Audessus then Ecran.Top:=Top-Ecran.Height else Ecran.Top:=Top ;
//Ecran.Left:=Left ;
InitialiserColonne;
end ;

Procedure TOF_MBOTransfertSaiDim.AdapterFiche ;
var CoordDernCol,CoordDernLign : TRect ;
    GridDim : THGrid ;
    ValOnglet,H,W : Integer ;
begin
// pour voir toutes les colonnes et toutes les lignes
H:=Application.MainForm.Height; W:=Application.MainForm.Width;
Ecran.Width:=W; Ecran.Height:=H;
THDimension(GetControl('FDIM')).Align:=alNone ;
GridDim:=DimensionsArticle.Dim.GridDim ;
if DimensionsArticle.Dim.DimOngl=nil then ValOnglet:=0 else ValOnglet:=GridDim.DefaultRowHeight ;
CoordDernCol:=GridDim.CellRect (GridDim.VisibleColCount+GridDim.FixedCols ,1) ;
CoordDernLign:=GridDim.CellRect (2,GridDim.VisibleRowCount+GridDim.FixedRows) ; // Dernière ligne
GridDim.ScrollBars:=ssNone ;

If (CoordDernLign.Bottom+76+ValOnglet>(H-100)) then
     begin
     Ecran.Height:=H-100 ;
     GridDim.ScrollBars:=ssVertical ;
     end else
     Ecran.Height:=CoordDernLign.Bottom  + 76 + ValOnglet ; // 76= titre et panel bouton de ecran + encadrment objet Dim
If (CoordDernCol.Right+20>(W-100)) then
     begin
     Ecran.Width:=W-100 ;
     if GridDim.ScrollBars=ssVertical then GridDim.ScrollBars:=ssBoth else GridDim.ScrollBars:=ssHorizontal ;
     end
     else if CoordDernCol.Right<230 then Ecran.Width:=230
     else Ecran.Width:=CoordDernCol.Right+20 ;

// Ajustement pour XP
Ecran.Height := Ecran.Height + 3;
Ecran.Width := Ecran.Width + 3;

// Ajustement taille si ajout d'un ascenseur
if (GridDim.ScrollBars=ssVertical) or (GridDim.ScrollBars=ssBoth) then Ecran.Width:=Ecran.Width+16;
if (GridDim.ScrollBars=ssHorizontal) or (GridDim.ScrollBars=ssBoth) then Ecran.Height:=Ecran.Height+16;

THDimension(GetControl('FDIM')).Align:=alClient ;
End ;

Procedure TOF_MBOTransfertSaiDim.OnClose ;
begin
if NotClose then LastError:=-1
            else BEGIN
                 LastError:=0 ;
                 // Sortie avec valider -> TobArticleDim conservé et retourné par TheTob
                 DimensionsArticle.ConserveTobArtDim:=bValider ;
                 if (DimensionsArticle<>nil) then DimensionsArticle.Destroy ;
                 END ;
NotClose:=False ;
end ;

procedure TOF_MBOTransfertSaiDim.OnUpdate ;
var ItemDim: THDimensionItem ;
Begin
ItemDim:=DimensionsArticle.Dim.CurrentItem ;
if ItemDim<>nil then
  begin
  if Not VerifStock(ItemDim) then begin NotClose:=True; exit; end;
  if MultiDepot='1' then RegroupeTousLesDepots;
  TheTob:=DimensionsArticle.TOBArticleDim;
  //TheTob:=LaTob;
  bValider:=True ;
  end;
inherited;
end ;

procedure TOF_MBOTransfertSaiDim.RegroupeTousLesDepots;
var i : integer;
begin
with DimensionsArticle do
  begin
  for i:=TobDimensions.Detail.Count-1 DownTo 0 do
    TobDimensions.Detail[i].ChangeParent(TOBArticleDim,-1);
  end;
end;

procedure TOF_MBOTransfertSaiDim.OnChangeItem(Sender: TObject);
var ItemDim :THDimensionItem ;
begin
  ItemDim:=THDimensionItem(Sender) ;
  if ItemDim = nil then exit ;
  if not VerifStock(ItemDim) then exit;
end ;

Function TOF_MBOTransfertSaiDim.VerifStock(ItemDim:THDimensionItem):Boolean ;
Var TobSelect: TOB;
    RProp,RPropTOB,RDiffProp,EStkFin,EProp,RStockALF : double;
    EManquant,ESurplus,EFinal : double;
    iItem,index : integer;
    ItemDimEmet:THDimensionItem;
Begin
Result:=False;
if ItemDim = nil then exit;
if Dim_RProp=-1 then exit;
TobSelect:=TOB(ItemDim.data);
RProp:=Valeur(ItemDim.Valeur[Dim_RProp]);
RPropTOB:=TobSelect.GetValue('GTL_PROPOSITION');
RStockALF:=TobSelect.GetValue('GTL_STOCKALF');
if RProp=RPropTOB then begin result:=true; exit; end;
//cas dépôt émetteur
if TobSelect.GetValue('GTL_DEPOTDEST')=RepCodeEmett then
  begin
  ItemDim.Valeur[Dim_RProp]:=TobSelect.GetValue('DEP_PROPOSITION');;
  result:=true; exit;
  end;
//cas d'une proposition négative
if RProp<0 then
  begin
  PGIInfo(TraduireMemoire('Les propositions négatives ne sont pas acceptées'),Ecran.Caption);
  ItemDim.Valeur[Dim_RProp]:=RPropTOB;
  result:=true; exit;
  end;
RDiffProp:=RProp-RPropTOB;
EStkFin:=TobSelect.GetValue('DEP_STOCKFINAL');
//Proposition > Stock dispo
if RDiffProp>EStkFin then
  begin
  PGIInfo(TraduireMemoire('Le stock disponible est insuffisant'),Ecran.Caption);
  ItemDim.Valeur[Dim_RProp]:=RPropTOB;
  DimensionsArticle.Dim.FocusDim(ItemDim,Dim_EProp);
  end
else
  begin
  EProp:=TobSelect.GetValue('DEP_PROPOSITION')-RDiffProp;
  EManquant:=TobSelect.GetValue('DEP_MANQUANT')-RDiffProp;
  ESurplus:=TobSelect.GetValue('DEP_SURPLUS')-RDiffProp;
  EFinal:=TobSelect.GetValue('DEP_STOCKINITIAL')+EProp;

  TobSelect.PutValue('GTL_PROPOSITION',RProp);
  TobSelect.PutValue('GTL_MANQUANT',TobSelect.GetValue('GTL_MANQUANT')-RDiffProp);
  TobSelect.PutValue('STOCKFINAL',TobSelect.GetValue('GTL_STOCKINITIAL')+RProp+RStockALF);

  if Dim_EProp<>-1 then ItemDim.Valeur[Dim_EProp]:=EProp;
  if Dim_EStkFin<>-1 then ItemDim.Valeur[Dim_EStkFin]:=TobSelect.GetValue('DEP_STOCKFINAL');
  if Dim_RStkFin<>-1 then ItemDim.Valeur[Dim_RStkFin]:=TobSelect.GetValue('STOCKFINAL');

  with DimensionsArticle do
    begin
    AffecteDEP(TOBArticleDim,TobSelect.GetValue('GTL_ARTICLE'),EProp,EManquant,ESurplus,EFinal);
    if MultiDepot='1' then AffecteDEP(TobDimensions,TobSelect.GetValue('GTL_ARTICLE'),EProp,EManquant,ESurplus,EFinal);

    ChangeChampDimMul(True);
    //RefreshGrid(HideUnUsed);
    //Dans le cas du multi-dépôt MAJ du dépôt émetteur au niveau de l'affichage
    if bMasqueDepot then
      begin
      iItem:=0 ; index:=-1 ;
      while (index=-1) and (iItem<TableDim.count) do
        begin
        ItemDimEmet:=THDimensionItem(TableDim.Items[iItem]) ;
        if (TOB(ItemDimEmet.data)<>nil) and (TOB(ItemDimEmet.data).GetValue('GTL_ARTICLE')=TobSelect.GetValue('GTL_ARTICLE'))
          and (TOB(ItemDimEmet.data).GetValue('GTL_DEPOTDEST')=RepCodeEmett)
          then index:=iItem ;
        iItem:=iItem+1 ;
        end ;
      if index<0 then exit;
      with TableDim.Items[index] do
        begin
        if Dim_RProp<>-1 then Valeur[Dim_RProp]:=TOB(data).GetValue('DEP_PROPOSITION');
        if Dim_RStkFin<>-1 then Valeur[Dim_RStkFin]:=TOB(data).GetValue('DEP_STOCKFINAL');
        end;
      Refresh(TableDim.Items[index]);
      end;
    end;
  Result:=True;
  end;
end ;

Procedure TOF_MBOTransfertSaiDim.AffecteDEP(TOBMere:TOB; CodeArt:String; EProp,EManquant,ESurplus,EFinal:Double);
var TOBUneDim : TOB;
begin
if TOBMere=nil then Exit;
TOBUneDim:=TOBMere.FindFirst(['GTL_ARTICLE'],[CodeArt],False);
while TOBUneDim<>nil do
  begin
  TOBUneDim.PutValue('DEP_PROPOSITION',EProp);
  TOBUneDim.PutValue('DEP_MANQUANT',EManquant);
  TOBUneDim.PutValue('DEP_SURPLUS',ESurplus);
  TOBUneDim.PutValue('DEP_STOCKFINAL',EFinal);
  TOBUneDim:=TOBMere.FindNext(['GTL_ARTICLE'],[CodeArt],False);
  end;
end;


{==============================================================================================}
{===================================== Procedure AGL ==========================================}
{==============================================================================================}
procedure AGLParamGrilleArtDim( parms: array of variant; nb: integer ) ;
var F : TForm ;
    TOTOF : TOF ;
    CaptionCompl : string ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit ;
if (TOTOF is TOF_MBOTransfertSaiDim) then
    BEGIN
    with TOF_MBOTransfertSaiDim(TOTOF) do
        BEGIN
        ParamGrille(DimensionsArticle,DimensionsArticle.NatureDoc,DimensionsArticle.NatureDoc) ;
        InitialiserColonne ;
        AdapterFiche ;
        if MultiDepot='1' then CaptionCompl:=' (Multi dépôts)' else CaptionCompl:=RechDom('GCDEPOT',CodeDepot,True);
        Ecran.Caption:= 'Saisie articles: ' + LibelleArticleGenerique(CodeArticle) + CaptionCompl;
        END ;
    END else exit ;
end ;

Initialization
registerclasses([TOF_MBOTransfertSaiDim]);
RegisterAglProc('ParamGrilleArtDim',TRUE,0,AGLParamGrilleArtDim);
end.

