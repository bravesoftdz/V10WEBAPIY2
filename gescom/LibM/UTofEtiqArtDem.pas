unit UTofEtiqArtDem;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,Hpanel, Math
      ,HCtrls,HEnt1,HMsgBox,UTOF,vierge,UTOB,AglInit,LookUp,EntGC,SaisUtil,graphics
      ,grids,windows,FactUtil,Ent1, utildimarticle,CalcOleGescom,UtilGC,
{$IFDEF EAGLCLIENT}
     HPdfPrev,UtileAGL,Maineagl,
{$ELSE}
      dbTables,db,FE_Main,
{$IFDEF V530}
      EdtEtat,
{$ELSE}
      EdtREtat,
{$ENDIF}
{$ENDIF}
{$IFNDEF SANSCOMPTA}
      Ventil,
{$ENDIF}

      M3FP,HTB97,Dialogs,AGLInitGC,UtilArticle,UtilPGI;

Type
     TOF_EtiqArtDem = Class (TOF)
     private
        LesColArt : string ;
        imprim_etiq, valide_etiq : integer;
        G_ART : THGRID ;
        AArt,AFou,ADEP, AVAL, ANUM : integer ;
        PART  : THPanel ;
        BNewLine: TToolbarButton97;
        BDelLine: TToolbarButton97;
        BChercher: TToolbarButton97;
        BImprimer: TToolbarButton97;
        FindLigne: TFindDialog;
        TobPR, TobToDelete, TOBDim : TOB ;
        procedure BNewLineClick(Sender: TObject);
        procedure BDelLineClick(Sender: TObject);
        procedure BChercherClick(Sender: TObject);
        procedure BImprimerClick(Sender: TObject);
        procedure FindLigneFind(Sender: TObject);
        procedure G_ARTElipsisClick(Sender: TObject);
        procedure G_ARTOnDblClick(Sender: TObject);
        procedure G_ARTCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_ARTCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
        procedure G_ARTRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure G_ARTRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
        procedure DessineCellArt (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
        // Actions liées au grid
        procedure ChargeGrille;
        Function  GetTOBLigne ( ARow : integer) : TOB ;
        procedure InitRow (Row : integer; GS : THGrid) ;
        Procedure CreerTOBLigne (ARow : integer; GS : THGrid);
        Function  LigVide( Row : integer; GS : THGRID) : Boolean ;
        procedure InsertLigne (ARow : Longint; GS : THGrid) ;
        procedure SupprimeLigne (ARow : Longint; GS : THGrid) ;
        Procedure ChercheArticle(GS : THGRID);
        Function  SortDeLaLigne : boolean ;
        //Manipulation des Champs GRID
        procedure TraiterDepot (ACol, ARow : integer; GS : THGrid);
        procedure TraiterValeur (ACol, ARow : integer; GS : THGrid);
        procedure TraiterArticle (ACol, ARow : integer; GS : THGrid);
        // Validation
        procedure ValidePrix;
        procedure VerifTOB;
        Procedure DepileTOBLigne ;
        Procedure DepileTOBDim ;
        Function  RecalcQte(CodeArticle : string;CodeDepot : string;NumLigne : String): integer;
        procedure ChargeTOBDimEtiq ;
        Procedure AfficheDIM(CodeArt : string; CodeDep: string; NumLigne : String; Action : String; GS : THGRID);
     public
        Action   : TActionFiche ;
        FindDebut : Boolean;
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure OnArgument (Arguments : String ) ; override ;
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
        procedure Onclose  ; override ;

     END ;

     Function SelectMultiDimedt ( sCodeGene : String; sCodeDep : String; Action : String; GS: THGrid) : Boolean ;  //AC

const colRang=1 ;
      NbRowsInit = 50 ;
      NbRowsPlus = 20 ;

implementation

Procedure TOF_EtiqArtDem.OnArgument (Arguments : String ) ;
var i,NbCol : integer ;
    St,Nam : string ;
begin
inherited ;
St:=Arguments ;
Action:=taModif ;
i:=Pos('ACTION=',St) ;
if i>0 then
   BEGIN
   System.Delete(St,1,i+6) ;
   St:=uppercase(ReadTokenSt(St)) ;
   if St='MODIFICATION' then Action:=taModif
   else if St='CONSULTATION' then Action:=taConsult
   else if St='CREATION' then Action:=taCreat ;
   END ;

NbCol:=6;
LesColArt:='FIXED;GZD_COMPTEUR;GZD_DEPOT;GZD_CODEARTICLE;GZD_NBETIQ;GZD_NUMLIGNE' ;

BNewLine:=TToolbarButton97(GetControl('BNEWLINE'));
BNewLine.OnClick:=BNewLineClick;
BDelLine:=TToolbarButton97(GetControl('BDELLINE'));
BDelLine.OnClick:=BDelLineClick;
BChercher:=TToolbarButton97(GetControl('BCHERCHER'));
BChercher.OnClick:=BChercherClick;
BImprimer:=TToolbarButton97(GetControl('BIMPRIMER'));
BImprimer.OnClick:=BImprimerClick;
FindLigne:=TFindDialog.Create(Ecran);
FindLigne.OnFind:=FindLigneFind ;

PART:=THPanel(GetControl('PARTICLE'));
PART.Align:=alClient;

G_ART:=THGRID(GetControl('G_ART'));
G_ART.OnElipsisClick:=G_ARTElipsisClick  ;
G_ART.OnDblClick:=G_ARTOnDblClick ;
G_ART.OnCellEnter:=G_ARTCellEnter ;
G_ART.OnCellExit:=G_ARTCellExit ;
G_ART.OnRowEnter:=G_ARTRowEnter ;
G_ART.OnRowExit:=G_ARTRowExit ;
G_ART.PostDrawCell:= DessineCellArt;
G_ART.ColCount:=NbCol;
G_ART.ColAligns[colRang]:=taCenter;
G_ART.ColTypes[colRang]:='I' ;
G_ART.ColWidths[1]:=0;   // rang
G_ART.ColWidths[5]:=0;   // Numero de ligne
G_ART.ColWidths[0]:=15;


St:=LesColArt ;
for i:=0 to G_ART.ColCount-1 do
   BEGIN
   if (i>1) and (i<>5) then  G_ART.ColWidths[i]:=100;
   Nam:=ReadTokenSt(St) ;
   if Nam='GZD_DEPOT' then BEGIN G_ART.ColFormats[i]:='CB=GCDEPOT'; ADEP:=i; END
   else if Nam='GZD_CODEARTICLE' then AArt:=i
   else if Nam='GZD_NBETIQ' then
        BEGIN
        G_ART.ColFormats[i] := '0' ;
        G_ART.ColAligns[i]  := taRightJustify ;
        //G_ART.ColLengths[i] := -1 ;
        AVAL := i ;
        END
   else if Nam='GZD_NUMLIGNE' then
        BEGIN G_ART.ColLengths[i] := -1 ; ANUM := i; END;
   END ;
imprim_etiq := 0 ;
valide_etiq := 0 ;
AffecteGrid(G_ART,Action) ;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(G_ART);
TFVierge(Ecran).OnKeyDown:=FormKeyDown ;
end;

procedure TOF_EtiqArtDem.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var FocusGrid : Boolean;
    ARow : Longint;
    GS : THGrid ;
BEGIN
//FocusGrid := False;
FocusGrid := True;
GS := G_ART;
ARow := G_Art.Row;
Case Key of
    VK_F5 : if FocusGrid then GS.OnElipsisClick(Sender);
    VK_RETURN : Key:=VK_TAB ;
    VK_INSERT : BEGIN
                if FocusGrid then
                    BEGIN
                    Key := 0;
                    InsertLigne (ARow, GS);
                    END;
                END;
    VK_DELETE : BEGIN
                if ((FocusGrid) and (Shift=[ssCtrl])) then
                    BEGIN
                    Key := 0 ;
                    SupprimeLigne (ARow, GS) ;
                    END ;
                END;
    END;
END;

Procedure TOF_EtiqArtDem.OnLoad  ;
BEGIN
inherited ;
TOBDim:=TOB.Create('',Nil,-1) ;
PART.Visible:=True;
Transactions (ChargeGrille, 1);
TFVierge(Ecran).Hmtrad.ResizeGridColumns(G_ART);       // Permet de formater les colonnes sur tout l'écran
END;

Procedure TOF_EtiqArtDem.OnUpdate  ;
BEGIN
inherited ;
Transactions (ValidePrix, 2);
TobPR.SetAllModifie(False);
END;

Procedure TOF_EtiqArtDem.Onclose  ;
var Fermer : boolean ;
    St : string ;
begin
inherited ;
Fermer:=True ;
if valide_etiq = 0 then
    BEGIN
    St:='1;?caption?;Confirmez-vous l''abandon de la saisie ?;Q;YN;Y;N;';
    if HShowMessage(St, Ecran.Caption, '') <> mrYes then Fermer:=False ;
    END;
if (Fermer) then
    BEGIN
    TobPR.free ; TobPR:=nil;
    TOBDim.free ; TOBDim:=nil;
    TobToDelete.free ;
    FindLigne.Destroy;
    END else
    BEGIN
    AfficheError:=False;
    LastError:=1; LastErrorMsg:='Non Fermeture' ;
    END;
VH_GC.TOBEdt.ClearDetail ;
initvariable;
end ;

{==============================================================================================}
{================================= Evènements du Grid =========================================}
{==============================================================================================}
procedure TOF_EtiqArtDem.G_ARTElipsisClick(Sender: TObject);
begin
if G_ART.Col = AART then
    BEGIN
    ChercheArticle (G_ART);
    END ;
end;

procedure TOF_EtiqArtDem.G_ARTOnDblClick(Sender: TObject);
begin
if G_ART.Col = AART then
    BEGIN
    if (G_ART.Cells [AART,G_ART.Row] = '') then  ChercheArticle (G_ART)
    else AfficheDIM(G_ART.Cells [AART,G_ART.Row],G_ART.CellValues [ADEP,G_ART.Row],IntToStr(G_ART.Row),'MODIF',G_ART);
    END ;
end;

procedure TOF_EtiqArtDem.G_ARTCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
//    if (G_ART.Col <> AART) AND (G_ART.Cells [AART,G_ART.Row] = '') then BEGIN G_ART.Col := AART;Cancel:=true; END;
    if (G_ART.Col <> AART) AND (G_ART.Cells [AART,G_ART.Row] = '') then BEGIN G_ART.Col := ADEP; END;
    G_ART.ElipsisButton:=((G_ART.Col=AART) or (G_ART.Col=AFOU)) ;
    end ;
    if (G_ART.Row = 1) and (G_ART.Cells [ADEP,G_ART.Row] = '') then
      BEGIN
      // modif le 13/08/02 ici c'est la notion d'établissement qui est la bonne
      G_ART.CellValues [ADEP,G_ART.Row] := VH^.EtablisDefaut; //VH_GC.GCDepotDefaut;
      END;
    if (G_ART.Row > 1) and (G_ART.Cells [ADEP,G_ART.Row] = '') then
      BEGIN
      G_ART.Cells [ADEP,G_ART.Row] := G_ART.Cells [ADEP,G_ART.Row-1];
      END;
    if (G_ART.Cells [AVAL,G_ART.Row] = '') or (G_ART.Cells [AVAL,G_ART.Row] = '0') then
      BEGIN
      G_ART.Cells [AVAL,G_ART.Row] := '1';
      END;
    if (G_ART.Cells [ANUM,G_ART.Row] = '') or (G_ART.Cells [ANUM,G_ART.Row] = '0') then
      BEGIN
      G_ART.Cells [ANUM,G_ART.Row] := IntToStr(G_ART.Row);
      END;
    if G_ART.Col = AART then TraiterDepot (ADep, ARow, G_ART);
end;

procedure TOF_EtiqArtDem.G_ARTCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if ACol = AART then TraiterArticle (ACol, ARow, G_ART);
end;


procedure TOF_EtiqArtDem.G_ARTRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var ARow : Integer;
begin
G_ART.InvalidateRow(ou) ;
if Ou >= G_ART.RowCount - 1 then G_ART.RowCount := G_ART.RowCount + NbRowsPlus ;
ARow := Min (Ou, TobPR.detail.count + 1);
if (ARow = TobPR.detail.count + 1) AND (not LigVide (ARow - 1, G_ART)) then
    BEGIN
    CreerTOBligne (ARow, G_ART);
    END;
if Ou > TobPR.detail.count then
    BEGIN
    G_ART.Row := TobPR.detail.count;
    END;
end;

procedure TOF_EtiqArtDem.G_ARTRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var diff : integer;
begin
diff:=G_ART.Row-Ou;
if (G_ART.Row > 1) and (G_Art.Cells[AArt,G_ART.Row -diff] <> '') then
  BEGIN

  G_Art.Cells[AVal,G_ART.Row -diff]:=InttoStr(ValeurI(G_Art.Cells[AVal,G_ART.Row -diff]));
  TraiterValeur (AVal, G_ART.Row -diff, G_ART) ;
  G_Art.Cells[ANUM,G_ART.Row -diff]:=InttoStr(ValeurI(G_Art.Cells[ANUM,G_ART.Row -diff]));
  TraiterValeur (ANUM, G_ART.Row -diff, G_ART) ;
  END;
G_ART.InvalidateRow(ou) ;
if LigVide (Ou, G_ART) Then G_ART.Row := Min (G_ART.Row,Ou);
end;


procedure TOF_EtiqArtDem.DessineCellArt(ACol, ARow: Longint; Canvas : TCanvas; AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < G_ART.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=G_ART.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := G_ART.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = G_ART.row) then
         BEGIN
         Canvas.Brush.Color := clBlack ;
         Canvas.Pen.Color := clBlack ;
         Triangle[1].X:=ARect.Right-2 ; Triangle[1].Y:=((ARect.Top+ARect.Bottom) div 2) ;
         Triangle[0].X:=Triangle[1].X-5 ; Triangle[0].Y:=Triangle[1].Y-5 ;
         Triangle[2].X:=Triangle[1].X-5 ; Triangle[2].Y:=Triangle[1].Y+5 ;
         if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle) ;
         END ;
    end;
end;


{==============================================================================================}
{================================ Actions liées au Grid =======================================}
{==============================================================================================}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 15/10/2001
Modifié le ... : 15/10/2001
Description .. : Chargement des étiquettes éventuellement en attente
Suite ........ : depuis la table temporaire GCTMPETQ
Mots clefs ... : BO;
*****************************************************************}
procedure TOF_EtiqArtDem.ChargeGrille ;
var QQ : TQuery ;
    St : string ;
BEGIN
  St := 'SELECT * FROM GCTMPETQ WHERE GZD_UTILISATEUR="'+V_PGI.USer+'" AND GZD_CODEARTICLE<>""'+
        ' AND GZD_COMPTEUR="0" ORDER BY  GZD_DEPOT, GZD_CODEARTICLE' ;
  QQ:=OpenSql(St, True) ;
  TobPR:=tob.create('',Nil,-1) ;
  if not QQ.Eof then TobPR.LoadDetailDB('GCTMPETQ','','',QQ,false,true) ;
  Ferme(QQ) ;
  If TobPR.detail.count <> 0 then
     ChargeTOBDimEtiq else
     Tob.create ('GCTMPETQ',TobPR,-1) ;
  TobPR.PutGridDetail(G_ART,True,True,LesColArt,True);
  G_ART.RowCount:=Max (NbRowsInit, G_ART.RowCount+1) ;
  TobToDelete:=tob.create('',Nil,-1) ;
END;

procedure TOF_EtiqArtDem.ChargeTOBDimEtiq ;
var i_ind : integer;
    TobSelect : TOB;
    QQ : TQuery ;
    St : string ;
    TobTmp : TOB;
BEGIN
St := 'SELECT * FROM GCTMPETQ WHERE GZD_UTILISATEUR="'+V_PGI.USer+'" AND GZD_CODEARTICLE<>""'+
    ' ORDER BY  GZD_DEPOT, GZD_CODEARTICLE' ;
QQ:=OpenSql(St, True) ;
Tobtmp:=tob.create('',Nil,-1) ;
if not QQ.Eof then Tobtmp.LoadDetailDB('GCTMPETQ','','',QQ,false,true) ;
Ferme(QQ) ;
for i_ind := 0 to  Tobtmp.detail.count-1 do
  begin
  if (Tobtmp.detail[i_ind].getvalue ('GZD_STATUTART') = 'DIM') then
    begin
    if TobDim.FindFirst(['GA_ARTICLE','DEPOT'],[Tobtmp.Detail[i_ind].GetValue('GZD_ARTICLE'),Tobtmp.Detail[i_ind].GetValue('GZD_DEPOT')],false) = nil then
      begin
      TobSelect:=tob.create('articles select',TobDim, -1 ) ;
      TobSelect.AddChampSup( 'GA_CODEARTICLE',false);
      TobSelect.AddChampSup( 'GA_ARTICLE',false);
      TobSelect.AddChampSup( 'DEPOT',false);
      TobSelect.AddChampSup( 'NBETIQDIM',false);
      TobSelect.AddChampSup( 'NUMLIGNE',false);
      TobSelect.PutValue('GA_CODEARTICLE',Tobtmp.Detail[i_ind].GetValue('GZD_CODEARTICLE')) ;
      TobSelect.PutValue('GA_ARTICLE',Tobtmp.Detail[i_ind].GetValue('GZD_ARTICLE')) ;
      TobSelect.PutValue('DEPOT',Tobtmp.Detail[i_ind].GetValue('GZD_DEPOT')) ;
      TobSelect.PutValue('NBETIQDIM',Tobtmp.Detail[i_ind].GetValue('GZD_NBETIQDIM')) ;
      TobSelect.PutValue('NUMLIGNE',Tobtmp.Detail[i_ind].GetValue('GZD_NUMLIGNE')) ;
      end;
    end;
  end;
Tobtmp.free ; //Tobtmp:=nil;
END;

Function TOF_EtiqArtDem.GetTOBLigne ( ARow : integer) : TOB ;
BEGIN
Result:=Nil ;
if ((ARow<=0) or (ARow>TOBPR.Detail.Count)) then Exit ;
Result:=TOBPR.Detail[ARow-1] ;
END ;

Procedure TOF_EtiqArtDem.InitRow (Row : integer; GS : THGrid) ;
var Col : integer ;
    TOBL : TOB;
BEGIN
TOBL:=GetTOBLigne(Row) ; if TOBL<>Nil then TOBL.InitValeurs ;
for Col:=0 to GS.ColCount do GS.cells[Col,Row]:='';
END ;

Procedure TOF_EtiqArtDem.CreerTOBLigne (ARow : integer; GS : THGrid);
BEGIN
if ARow <> TOBPR.Detail.Count + 1 then exit;
TOB.Create ('GCTMPETQ', TOBPR, ARow-1) ;
InitRow (ARow, GS) ;
END;

Function TOF_EtiqArtDem.LigVide(Row : integer; GS : THGRID) : Boolean ;
var Col : integer ;
BEGIN
Result:=True ;
Col:=ADEP ;
if (GS.Cells[Col,Row]<>'') then result:= False ;
END ;

procedure TOF_EtiqArtDem.InsertLigne (ARow : Longint; GS : THGrid) ;
var i_ind1,i_ind2: integer;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if LigVide (ARow, GS) then exit;
// Si on insert une ligne à la fin, on sort
if (ARow > TOBPR.Detail.Count) then Exit
else
// sinon on insert une ligne en plein milieu
begin
     for i_ind1 := TOBPR.Detail.Count downto ARow do
     // Pour toutes les lignes supérieur à la ligne inseré
     begin
              // Je met à jour le nouveau numéro de ligne dans TOBPR
              TOBPR.Detail[i_ind1-1].PutValue ('GZD_NUMLIGNE',i_ind1+1 );
              // Pour toutes les lignes de la tobdim
              for i_ind2 :=0  to TobDim.Detail.Count-1 do
              begin
                   // Si le depot n'est pas renseigné je le renseigne
                   if TOBPR.Detail[i_ind1-1].GetValue('GZD_DEPOT')='' then
                          TOBPR.Detail[i_ind1-1].PutValue ('GZD_DEPOT',TobDim.Detail[i_ind2].GetValue('DEPOT') );
                   // Si l'article de TobPR est égal à l'article de TobDim et que le numéro de ligne correspond alors
                   // j'incremente se numero de ligne dans TobDim afin d'être cohérant avec TobPR
                   // Cette mis à jour des numéro de ligne dans les 2 tob sont nécessaire si on veut inserer n'importe quel
                   // article notamment le même article que celui de la ligne suivante
                   if (TOBPR.Detail[i_ind1-1].GetValue('GZD_CODEARTICLE') = TobDim.Detail[i_ind2].GetValue('GA_CODEARTICLE'))
                       and (i_ind1 = TobDim.Detail[i_ind2].GetValue('NUMLIGNE'))then
                   begin
                      TobDim.Detail[i_ind2].PutValue('NUMLIGNE',i_ind1+1) ;
                   end;
              end;
              // Ceci permet de mettre à jour le champ numligne de l'ecran (même si on ne le voit pas)
             G_Art.Cells[ANUM,i_ind1]:=IntToStr(i_ind1+1);
     end;
end;
GS.CacheEdit; GS.SynEnabled := False;
TOB.Create ('GCTMPETQ', TOBPR, ARow-1) ;
GS.InsertRow (ARow); GS.Row := ARow;
G_Art.Cells[ANUM,ARow]:=IntToStr(ARow);
InitRow (ARow, GS) ;
GS.MontreEdit; GS.SynEnabled := True;
END;

procedure TOF_EtiqArtDem.SupprimeLigne (ARow : Longint; GS : THGrid) ;
Var i_ind, i_ind1,i_ind2: integer;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if (ARow > TOBPR.Detail.Count) then Exit;
GS.CacheEdit; GS.SynEnabled := False;
GS.DeleteRow (ARow);
// On supprime dans la TobDim toutes les lignes concernant l'article que l'on veut supprimer
// Sinon il reste dans la Tob et si on réinsère se même article à la même ligne
// il va les rajouter à celle existante
for i_ind2:=TobDim.Detail.Count-1 downto 0 do
begin
     if (TobDim.Detail[i_ind2].GetValue('NUMLIGNE')=ARow)
     then TobDim.Detail[i_ind2].Free;
end;
// Si on supprime la derniere ligne saisieon recré une tob ligne
if (ARow = TOBPR.Detail.Count) then CreerTOBLigne (ARow + 1, GS)
else
begin
// Sinon
     for i_ind1 := TOBPR.Detail.Count Downto ARow + 1 do
     // Pour toutes les lignes supérieur à la ligne supprimée
     begin
              // Je met à jour le nouveau numéro de ligne dans TOBPR
              TOBPR.Detail[i_ind1-1].PutValue ('GZD_NUMLIGNE',i_ind1-1 );
              for i_ind2 :=0 to TobDim.Detail.Count-1 do
              // Pour toutes les lignes de la tobdim
              begin
                   // Si le depot n'est pas renseigné, je le renseigne
                   if TOBPR.Detail[i_ind1-1].GetValue('GZD_DEPOT')=''
                   then TOBPR.Detail[i_ind1-1].PutValue ('GZD_DEPOT',TobDim.Detail[i_ind2].GetValue('DEPOT') );
                   // Si l'article de TobPR est égal à l'article de TobDim et que le numéro de ligne correspond alors
                   // je décremente se numero de ligne dans TobDim afin d'être cohérant avec TobPR
                   // Cette mise à jour des numéros de ligne dans les 2 tobs sont nécessaire si on veut inserer n'importe quel
                   // article notamment le même article que celui de la derniere ligne
                   if (TOBPR.Detail[i_ind1-1].GetValue('GZD_CODEARTICLE') = TobDim.Detail[i_ind2].GetValue('GA_CODEARTICLE'))
                       and (i_ind1 = TobDim.Detail[i_ind2].GetValue('NUMLIGNE'))
                   then TobDim.Detail[i_ind2].PutValue('NUMLIGNE',i_ind1-1) ;
              end;
          // Ceci permet de mettre à jour le champ numligne de l'ecran (même si on ne le voit pas)
          G_Art.Cells[ANUM,i_ind1-1]:=IntToStr(i_ind1-1);
     end;
end;
if TOBPR.Detail[ARow-1].GetValue ('GZD_COMPTEUR') <> 0 then
    BEGIN
    i_ind := TobToDelete.Detail.Count;
    TOB.Create ('GCTMPETQ', TobToDelete, i_ind) ;
    TobToDelete.Detail[i_ind].Dupliquer (TOBPR.Detail[ARow-1], False, True);
    END;
TOBPR.Detail[ARow-1].Free;
if GS.RowCount < NbRowsInit then GS.RowCount := NbRowsInit;
GS.MontreEdit; GS.SynEnabled := True;
END;

Procedure TOF_EtiqArtDem.ChercheArticle(GS : THGRID);
Var ARTICLE : THCritMaskEdit;
    Coord : TRect;
BEGIN
Coord := GS.CellRect (GS.Col, GS.Row);
ARTICLE := THCritMaskEdit.Create (ECRAN);
ARTICLE.Parent := GS;
ARTICLE.Top := Coord.Top;
ARTICLE.Left := Coord.Left;
ARTICLE.Width := 3; ARTICLE.Visible := False;
ARTICLE.Text:= GS.Cells[GS.Col,GS.Row] ;
ARTICLE.DataType:='GCMULARTICLE_MODE';
ARTICLE.Text:=DispatchArtMode(1,'','','SELECTION;GA_CODEARTICLE='+ARTICLE.Text+';XX_WHERE_=GA_TYPEARTICLE = "MAR"');
if ARTICLE.Text <> '' then
  begin
  GS.Cells[GS.Col,GS.Row]:= Trim (Copy (ARTICLE.Text, 1, 18));
  // DCA - FQ MODE 10141
  TraiterArticle (GS.Col, GS.Row, G_ART);
  //AfficheDIM(UpperCase(GS.Cells [GS.Col, GS.Row]),UpperCase(GS.CellValues [GS.Col-1,GS.Row]),IntToStr(GS.Row),'SAISIE',GS);
  end ;
ARTICLE.Destroy;
END ;

Procedure TOF_EtiqArtDem.AfficheDIM(CodeArt : string; CodeDep: string; NumLigne : String;  Action : String; GS : THGRID);
Var QQ : TQuery;
    NbQte : integer;
BEGIN
//NbQte := 0;
QQ:=OpenSQL('Select GA_STATUTART from ARTICLE Where GA_ARTICLE like "'+CodeArt+' %"',True) ;
if Not QQ.EOF then
  begin
  if QQ.Findfield('GA_STATUTART').Asstring <> 'UNI' then
    begin
    TheTob := TobDim;
    if SelectMultiDimedt(CodeArt,CodeDep,Action,GS) then
      begin
      TOBDim := TheTOB ;
      TheTOB:=Nil ;
      end;
    NbQte := RecalcQte (CodeArt,CodeDep,NumLigne);
    // DCA - FQ MODE 10141
    //GS.Cells [GS.Col+1, GS.Row] := inttostr(NbQTE);
    GS.Cells [AVAL, GS.Row] := inttostr(NbQTE);
    end;
  end;
Ferme (QQ);
END;

Function TOF_EtiqArtDem.RecalcQte(CodeArticle : string;CodeDepot : string;NumLigne : String): integer;
Var i_dim, NbETQ : integer;
BEGIN
NbETQ := 0;
for i_dim :=0 to TobDim.detail.count -1 do
begin
if (CodeArticle = TobDim.Detail[i_dim].GetValue('GA_CODEARTICLE'))
And (CodeDepot = TobDim.Detail[i_dim].GetValue('DEPOT'))
and (NumLigne = TobDim.Detail[i_dim].GetValue('NUMLIGNE')) then
  begin
  NbETQ := NbETQ + TobDim.Detail[i_dim].GetValue('NBETIQDIM');
  end;
end;
Result := NbETQ;
END;

Function TOF_EtiqArtDem.SortDeLaLigne : boolean ;
Var ACol,ARow : integer ;
    Cancel : boolean ;
BEGIN
Result:=False ;
ACol:=G_Art.Col ; ARow:=G_Art.Row ; Cancel:=False ;
G_ArtCellExit(Nil,ACol,ARow,Cancel) ; if Cancel then Exit ;
G_ArtRowExit(Nil,ACol,Cancel,False) ; if Cancel then Exit ;
Result:=True ;
END ;

{==============================================================================================}
{============================ Manipulation des Champs GRID ====================================}
{==============================================================================================}

procedure TOF_EtiqArtDem.TraiterDepot (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St := GS.CellValues [ACol, ARow];
if St <> '' then
    BEGIN
    TOBL.PutValue ('GZD_DEPOT', St);
    END;
END;


procedure TOF_EtiqArtDem.TraiterValeur (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
if ACol=4 then
TOBL.PutValue ('GZD_NBETIQ', Valeur (GS.Cells [ACol, ARow]));
if ACol=5 then
TOBL.PutValue ('GZD_NUMLIGNE',Valeur (GS.Cells [ACol, ARow]));
END;

procedure TOF_EtiqArtDem.TraiterArticle (ACol, ARow : integer; GS : THGrid);
var TOBL : TOB;
    St : string;
    QQ : TQuery ;
    ArticleDim : Boolean ;
BEGIN
TOBL := GetTOBLigne (ARow); if TOBL=nil then exit;
St := UpperCase(GS.Cells [ACol, ARow]);
GS.Cells [ACol, ARow] := St;
if St=''then exit ;

if TOBL.GetValue ('GZD_CODEARTICLE') <> St then
    BEGIN
    ArticleDim := False ;
    QQ := OpenSQL('Select GA_ARTICLE, GA_CODEARTICLE from ARTICLE Where GA_CODEARTICLE="' +
                   St + '" AND GA_STATUTART in ("GEN","UNI")',True) ;
    if Not QQ.EOF then
        BEGIN
        ArticleDim := True ;
        TOBL.PutValue ('GZD_ARTICLE', QQ.Findfield('GA_ARTICLE').Asstring);
        TOBL.PutValue ('GZD_CODEARTICLE', QQ.Findfield('GA_CODEARTICLE').Asstring);
        END else
        BEGIN
        Ferme (QQ);
        QQ:=OpenSQL('Select GA_ARTICLE, GA_CODEARTICLE from ARTICLE Where GA_CODEBARRE="'+St+'"',True) ;
        if Not QQ.EOF then
          BEGIN
          ArticleDim := True ;
          TOBL.PutValue ('GZD_ARTICLE', QQ.Findfield('GA_ARTICLE').Asstring);
          TOBL.PutValue ('GZD_CODEARTICLE', QQ.Findfield('GA_CODEARTICLE').Asstring);
          END else
          BEGIN
          GS.Col:=ACol; GS.Row:=ARow ;
          ChercheArticle (GS);
          END;
        END;
    Ferme (QQ);
    // DCA - FQ MODE 10141 - Si article saisi existe, saisie des qtes dans objet DIM
    if ArticleDim then
      begin
      AfficheDIM(UpperCase(GS.Cells [ACol, ARow]),UpperCase(GS.CellValues [ACol-1, ARow]),IntToStr(ARow),'SAISIE',GS);
      end ;
    END;

END;

{==============================================================================================}
{============================ Evenement lié aux Boutons =======================================}
{==============================================================================================}
procedure TOF_EtiqArtDem.BChercherClick(Sender: TObject);
begin
if G_Art.RowCount < 3 then Exit;
FindDebut:=True ; FindLigne.Execute ;
end;

procedure TOF_EtiqArtDem.BImprimerClick(Sender: TObject);
var  st,CodeEtat, LibEtat : string;
     BApercu : boolean;
begin
VH_GC.TOBEdt.ClearDetail ;
initvariable;
Transactions (ValidePrix, 2);
CodeEtat := GetControlText ('CODETAT');
LibEtat := RechDom('TTMODELETIQART',CodeEtat,FALSE);
// Récupération de la coche "apercu avant impression"
BApercu := StringToCheck(GetControlText('FApercu'));
EditMonarchSiEtat (LibEtat);
LanceEtat('E','GED',CodeEtat,BApercu,False,False,nil,'GZD_UTILISATEUR="'+V_PGI.USer+'"','',False) ;
EditMonarch ('');
//TobPR.SetAllModifie(False);
imprim_etiq := 1;
St:='1;?caption?;Voulez vous relancer le traitement ultérieurement?;Q;YN;Y;N;';
if HShowMessage(St, Ecran.Caption, '') <> mrYes then
  begin
  G_Art.VidePile(True) ;
  G_Art.RowCount:= NbRowsInit ;
  DepileTOBLigne;
  DepileTOBDim;
  ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "'+V_PGI.USer+'"');
  CreerTOBligne (1, G_ART);
  // modif le 13/08/02 on parle d'établissemnt et non de dépôt
  G_ART.CellValues [ADEP,1] := VH^.EtablisDefaut;//VH_GC.GCDepotDefaut;
  G_ART.Cells [AVAL,1] := '1';
  G_ART.Cells [ANUM,1] := IntToStr(G_ART.Row);
  TraiterDepot (ADep, 1, G_ART);
  end;
end;

Procedure TOF_EtiqArtDem.DepileTOBLigne ;
var i_ind : integer;
BEGIN
for i_ind := TOBPR.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBPR.Detail[i_ind].Free ;
    END;
END;

Procedure TOF_EtiqArtDem.DepileTOBDim ;
var i_ind : integer;
BEGIN
for i_ind := TOBDim.Detail.Count - 1 Downto 0 do
    BEGIN
    TOBDim.Detail[i_ind].Free ;
    END;
END;

procedure TOF_EtiqArtDem.FindLigneFind(Sender: TObject);
begin
Rechercher (G_Art, FindLigne, FindDebut) ;
end;

procedure TOF_EtiqArtDem.BNewLineClick(Sender: TObject);
BEGIN
//if(Screen.ActiveControl=G_Art) then  InsertLigne (G_Art.Row, G_ART)
InsertLigne (G_Art.Row, G_ART) ;
end;

procedure TOF_EtiqArtDem.BDelLineClick(Sender: TObject);
begin
//if(Screen.ActiveControl = G_Art) then  SupprimeLigne (G_Art.Row, G_ART)
SupprimeLigne (G_Art.Row, G_ART) ;
end;

{==============================================================================================}
{================================= Validation =================================================}
{==============================================================================================}
procedure TOF_EtiqArtDem.ValidePrix;
begin
if Not SortDeLaLigne then Exit ;
TobToDelete.DeleteDB (False);
VerifTOB;
valide_etiq := 1;
end;

procedure TOF_EtiqArtDem.VerifTOB;
var i_ind,i_ind2,i_ind3,i_dim, nbex, nbexdim : integer;
    GS : THGrid ;
    TobEtiq, TOBE : TOB;
    RegimePrix,starticle,stdepot,stnumligne,CodeArticleGen,LibEtat : string;
BEGIN
LibEtat := RechDom('TTMODELETIQART',GetControlText ('CODETAT'),FALSE);
// Cette procédure permet d'insérer les enregistrements dans la table temporaire
GS:=G_Art ;
RegimePrix := 'TTC';
TobEtiq := TOB.Create('GCTMPETQ',nil,-1);
if EtatMonarchFactorise(LibEtat) then
   begin
   for i_ind := TOBPR.Detail.count - 1 Downto 0 do
      BEGIN
      if LigVide (i_ind + 1,GS) then
         BEGIN
         TOBPR.Detail[i_ind].Free ;
         END
      else
         BEGIN
         if TOBPR.Detail[i_ind].GetValue ('GZD_COMPTEUR') = 0 then
            BEGIN
            if (TOBPR.Detail[i_ind].GetValue('GZD_CODEARTICLE') <> '') then
               begin
               starticle := TOBPR.Detail[i_ind].GetValue('GZD_CODEARTICLE');
               stdepot := TOBPR.Detail[i_ind].GetValue('GZD_DEPOT');
               stnumligne := TOBPR.Detail[i_ind].GetValue('GZD_NUMLIGNE');
               //nbex := TOBPR.Detail[i_ind].GetValue('GZD_NBETIQ');
               i_ind2:=0;
               if TobDim.FindFirst(['GA_CODEARTICLE','DEPOT','NUMLIGNE'],[starticle,stdepot,stnumligne],false) = nil then
                  begin
                  TOBE:=TOB.Create('GCTMPETQ',TobEtiq,-1);
                  TOBE.PutValue('GZD_UTILISATEUR',V_PGI.USer);
                  TOBE.PutValue('GZD_COMPTEUR',i_ind2);
                  TOBE.PutValue('GZD_ARTICLE',TOBPR.Detail[i_ind].GetValue('GZD_ARTICLE'));
                  TOBE.PutValue('GZD_CODEARTICLE',TOBPR.Detail[i_ind].GetValue('GZD_CODEARTICLE'));
                  TOBE.PutValue('GZD_DEPOT',TOBPR.Detail[i_ind].GetValue('GZD_DEPOT'));
                  TOBE.PutValue('GZD_NBETIQ',TOBPR.Detail[i_ind].GetValue('GZD_NBETIQ'));
                  TOBE.PutValue('GZD_NBETIQDIM',TOBPR.Detail[i_ind].GetValue('GZD_NBETIQ'));      // j'insere aussi la valeur de NBETIQDIM car pour étiquette factorisée, on va se basé sur cette valeur
                  TOBE.PutValue('GZD_REGIMEPRIX',RegimePrix);
                  TOBE.PutValue('GZD_STATUTART','UNI');
                  TOBE.PutValue('GZD_NUMERO','0');
                  TOBE.PutValue('GZD_SOUCHE','0');
                  TOBE.PutValue('GZD_INDICEG','0');
                  TOBE.PutValue('GZD_NUMLIGNE',TobPR.Detail[i_ind].GetValue('GZD_NUMLIGNE'));
                  // Appel de la fonction qui transforme le codearticle en article générique
                  CodeArticleGen:=GCDimToGen(TOBPR.Detail[i_ind].GetValue('GZD_CODEARTICLE'));
                  TOBE.PutValue('GZD_CODEARTICLEGEN',CodeArticleGen);
                  end
               else
                  begin
                  for i_dim :=0 to TobDim.detail.count -1 do
                     begin
                     if (TOBPR.Detail[i_ind].GetValue('GZD_CODEARTICLE') = TobDim.Detail[i_dim].GetValue('GA_CODEARTICLE'))
                     and (TOBPR.Detail[i_ind].GetValue('GZD_DEPOT') = TobDim.Detail[i_dim].GetValue('DEPOT'))
                     and (TOBPR.Detail[i_ind].GetValue('GZD_NUMLIGNE') = TobDim.Detail[i_dim].GetValue('NUMLIGNE'))then
                        begin
                        //nbexdim := TobDim.Detail[i_dim].GetValue('NBETIQDIM');
                        //i_ind3:=0;
                        TOBE:=TOB.Create('GCTMPETQ',TobEtiq,-1);
                        TOBE.PutValue('GZD_UTILISATEUR',V_PGI.USer);
                        TOBE.PutValue('GZD_COMPTEUR',i_ind2);
                        TOBE.PutValue('GZD_ARTICLE',TobDim.Detail[i_dim].GetValue('GA_ARTICLE'));
                        TOBE.PutValue('GZD_CODEARTICLE',TobDim.Detail[i_dim].GetValue('GA_CODEARTICLE'));
                        TOBE.PutValue('GZD_DEPOT',TOBPR.Detail[i_ind].GetValue('GZD_DEPOT'));
                        TOBE.PutValue('GZD_NBETIQ',TOBPR.Detail[i_ind].GetValue('GZD_NBETIQ'));  //TobDim.Detail[i_dim].GetValue('NBETIQDIM'));
                        TOBE.PutValue('GZD_NBETIQDIM',TobDim.Detail[i_dim].GetValue('NBETIQDIM'));
                        TOBE.PutValue('GZD_REGIMEPRIX',RegimePrix);
                        TOBE.PutValue('GZD_STATUTART','DIM');
                        TOBE.PutValue('GZD_NUMERO','0');
                        TOBE.PutValue('GZD_SOUCHE','0');
                        TOBE.PutValue('GZD_INDICEG','0');
                        TOBE.PutValue('GZD_NUMLIGNE',TobPR.Detail[i_ind].GetValue('GZD_NUMLIGNE'));
                        // Appel de la fonction qui transforme le codearticle en article générique
                        CodeArticleGen:=GCDimToGen(TobDim.Detail[i_dim].GetValue('GA_CODEARTICLE'));
                        TOBE.PutValue('GZD_CODEARTICLEGEN',CodeArticleGen);
                        inc (i_ind2);
                        //inc (i_ind3);
                        end;  // fin if codearticle, depot , numligne
                     end; // fin du for i_dim
                  end;  //  fin else Tobdim.findfirst
               //inc (i_ind2);
               end; // fin du if codearticle<>''
            END;   // fin du if compteur=0
         END; // fin du else LigVide
      END;  // fin du for i_ind
   end  // fin si factorisé
else
   begin
   for i_ind := TOBPR.Detail.count - 1 Downto 0 do
      BEGIN
      if LigVide (i_ind + 1,GS) then
         BEGIN
         TOBPR.Detail[i_ind].Free ;
         END
      else
         BEGIN
         if TOBPR.Detail[i_ind].GetValue ('GZD_COMPTEUR') = 0 then
            BEGIN
            if (TOBPR.Detail[i_ind].GetValue('GZD_CODEARTICLE') <> '') then
               begin
               starticle := TOBPR.Detail[i_ind].GetValue('GZD_CODEARTICLE');
               stdepot := TOBPR.Detail[i_ind].GetValue('GZD_DEPOT');
               stnumligne := TOBPR.Detail[i_ind].GetValue('GZD_NUMLIGNE');
               nbex := TOBPR.Detail[i_ind].GetValue('GZD_NBETIQ');
               i_ind2:=0;
               while  i_ind2 < nbex do
                  begin
                  if TobDim.FindFirst(['GA_CODEARTICLE','DEPOT','NUMLIGNE'],[starticle,stdepot,stnumligne],false) = nil then
                     begin
                     TOBE:=TOB.Create('GCTMPETQ',TobEtiq,-1);
                     TOBE.PutValue('GZD_UTILISATEUR',V_PGI.USer);
                     TOBE.PutValue('GZD_COMPTEUR',i_ind2);
                     TOBE.PutValue('GZD_ARTICLE',TOBPR.Detail[i_ind].GetValue('GZD_ARTICLE'));
                     TOBE.PutValue('GZD_CODEARTICLE',TOBPR.Detail[i_ind].GetValue('GZD_CODEARTICLE'));
                     TOBE.PutValue('GZD_DEPOT',TOBPR.Detail[i_ind].GetValue('GZD_DEPOT'));
                     TOBE.PutValue('GZD_NBETIQ',TOBPR.Detail[i_ind].GetValue('GZD_NBETIQ'));
                     TOBE.PutValue('GZD_NBETIQDIM',TOBPR.Detail[i_ind].GetValue('GZD_NBETIQ'));
                     TOBE.PutValue('GZD_REGIMEPRIX',RegimePrix);
                     TOBE.PutValue('GZD_STATUTART','UNI');
                     TOBE.PutValue('GZD_NUMERO','0');
                     TOBE.PutValue('GZD_SOUCHE','0');
                     TOBE.PutValue('GZD_INDICEG','0');
                     TOBE.PutValue('GZD_NUMLIGNE',TobPR.Detail[i_ind].GetValue('GZD_NUMLIGNE'));
                     // Appel de la fonction qui transforme le codearticle en article générique
                     CodeArticleGen:=GCDimToGen(TOBPR.Detail[i_ind].GetValue('GZD_CODEARTICLE'));
                     TOBE.PutValue('GZD_CODEARTICLEGEN',CodeArticleGen);
                     end
                  else
                     begin
                     for i_dim :=0 to TobDim.detail.count -1 do
                        begin
                        if (TOBPR.Detail[i_ind].GetValue('GZD_CODEARTICLE') = TobDim.Detail[i_dim].GetValue('GA_CODEARTICLE'))
                        and (TOBPR.Detail[i_ind].GetValue('GZD_DEPOT') = TobDim.Detail[i_dim].GetValue('DEPOT'))
                        and (TOBPR.Detail[i_ind].GetValue('GZD_NUMLIGNE') = TobDim.Detail[i_dim].GetValue('NUMLIGNE'))then
                           begin
                           nbexdim := TobDim.Detail[i_dim].GetValue('NBETIQDIM');
                           i_ind3:=0;
                           while  i_ind3 < nbexdim do
                              begin
                              TOBE:=TOB.Create('GCTMPETQ',TobEtiq,-1);
                              TOBE.PutValue('GZD_UTILISATEUR',V_PGI.USer);
                              TOBE.PutValue('GZD_COMPTEUR',i_ind2);
                              TOBE.PutValue('GZD_ARTICLE',TobDim.Detail[i_dim].GetValue('GA_ARTICLE'));
                              TOBE.PutValue('GZD_CODEARTICLE',TobDim.Detail[i_dim].GetValue('GA_CODEARTICLE'));
                              TOBE.PutValue('GZD_DEPOT',TOBPR.Detail[i_ind].GetValue('GZD_DEPOT'));
                              TOBE.PutValue('GZD_NBETIQ',TOBPR.Detail[i_ind].GetValue('GZD_NBETIQ'));
                              TOBE.PutValue('GZD_NBETIQDIM',TobDim.Detail[i_dim].GetValue('NBETIQDIM'));
                              TOBE.PutValue('GZD_REGIMEPRIX',RegimePrix);
                              TOBE.PutValue('GZD_STATUTART','DIM');
                              TOBE.PutValue('GZD_NUMERO','0');
                              TOBE.PutValue('GZD_SOUCHE','0');
                              TOBE.PutValue('GZD_INDICEG','0');
                              TOBE.PutValue('GZD_NUMLIGNE',TobPR.Detail[i_ind].GetValue('GZD_NUMLIGNE'));
                              // Appel de la fonction qui transforme le codearticle en article générique
                              CodeArticleGen:=GCDimToGen(TobDim.Detail[i_dim].GetValue('GA_CODEARTICLE'));
                              TOBE.PutValue('GZD_CODEARTICLEGEN',CodeArticleGen);
                              inc (i_ind2);
                              inc (i_ind3);
                              end; // fin du while i_ind3
                           end;  // fin if codearticle, depot , numligne
                        end; // fin du for i_dim
                     end;  //  fin else Tobdim.findfirst
                  inc (i_ind2);
                  end;    //  fin du while i_ind2
               end; // fin du if codearticle<>''
            END;   // fin du if compteur=0
         END; // fin du else LigVide
      END;  // fin du for i_ind
   end;  // fin else factorisé
if TobEtiq.Detail.Count > 0 then
  begin
  ExecuteSQL('DELETE FROM GCTMPETQ WHERE GZD_UTILISATEUR = "'+V_PGI.USer+'"');
  TobEtiq.InsertDB(nil);  //TobEtiq.InsertOrUpdateDB();
  end;
TobEtiq.Free;
END;

Function SelectMultiDimedt ( sCodeGene : String;sCodeDep : String;Action : String; GS: THGrid ) : Boolean ;
var Top, Left,Height,Width: Integer ;
CelluleEcran: Tpoint ;
AuDessus: String ;
BEGIN
Result:=False ;

CelluleEcran:=RetourneCoordonneeCellule(1,GS)  ;
If CelluleEcran.y >= 372 then AuDessus:='X' else AuDessus:='-' ;
Top:=CelluleEcran.y ;
Left:=CelluleEcran.x ;
Height:=GS.Height-GS.RowHeights [0]-GS.RowHeights [1];
Width:=GS.Width -GS.ColWidths [0]  ;
V_PGI.FormCenter:=False ;
AglLanceFiche ('GC','GCEDARTDEMDIM','','', 'GA_CODEARTICLE='+sCodeGene+';ACTION='+Action+';CHAMP= ;DEPOT='+sCodeDep+';NUMLIGNE='+IntToStr(GS.Row)+';TOP='+IntToStr(Top)+';LEFT='+IntToStr(Left)+';OU='+AuDessus+';TYPEETAT=ETA;HEIGTH='+IntToStr(Height)+';WIDTH='+IntToStr(Width)+'') ;
if TheTOB<>Nil then Result:=(TheTOB.Detail.Count>0) ;
END ;

Initialization
registerclasses([TOF_EtiqArtDem]);

end.

