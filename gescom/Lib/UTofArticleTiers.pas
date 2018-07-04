{***********UNITE*************************************************
Auteur  ...... : JS
Créé le ...... : 12/02/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : ARTICLETIERS ()
Mots clefs ... : TOF;ARTICLETIERS
*****************************************************************}
Unit UTofArticleTiers ;

Interface

Uses Windows, StdCtrls, Controls, Classes,{ db,} Vierge, Graphics, forms, Grids, sysutils,
     ComCtrls, HCtrls, HEnt1, HMsgBox, HTB97, HPanel, UTOF, UTOB, UtilArticle, AGLInitGC,
     tarifUtil,HDimension,EntGC,
{$IFDEF EAGLCLIENT}
     maineagl,
{$ELSE}
     Fe_Main, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      AGLInit;
Type
  TOF_ARTICLETIERS = Class (TOF)

    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnClose                  ; override ;

    private
    G_Ref : THGrid;
    Ctrl : THCritMaskEdit;
    bValider,bFerme,bDelete,bInsert,bMultiSel : TToolBarButton97;
    bNew,Data_Modified,bValide : boolean;
    TobDelete: TOB;
    stFicSource,LesColsRef,Cell_Text : string;
    Article,LibArticle : String ;
    Gfixed,GInfo,GArt,GCodArt,GCod,GRefA,GRefT,GLibA : integer;

    ///Grid événements
    procedure G_RefRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
    procedure G_RefRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
    procedure G_RefCellEnter(Sender: TObject; var ACol, ARow: Integer;var Cancel: Boolean);
    procedure G_RefCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure G_RefElipsisClick(Sender: TObject);
    procedure G_RefDblClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
    ///Grid initialisation
    procedure InitGrid;
    procedure EtudieColsListe ;
    procedure AfficheLibelle;
    procedure DessineCell ( ACol,ARow : Longint; Canvas : TCanvas ;
                                                        AState: TGridDrawState);
    procedure AffichePiedInfo(ARow : integer);
    ///Boutons
    procedure bInsertClick(Sender: TObject);
    procedure bDeleteClick(Sender: TObject);
    procedure bValiderClick(Sender: TObject);
    procedure bMultiSelClick(Sender: TObject);
    ///Tob
    procedure ChargeTob(stArticle : string);
    procedure SupprimeTobFille;
    procedure CreerTobFille(ARow : integer);
    //Recherche
    function  Recherche_Art(ARow : LongInt)  : boolean;
    function  Recherche_Tiers(ARow : integer): boolean;
    function  RechDim(TobArt : TOB) : string;
    ///Traitements sur les lignes
    procedure GereCode(ACol,ARow : integer);
    procedure EntreeCell(Acol,ARow : integer);
    function  QuitLaLigne(ARow : integer) : boolean;
    function  ValideLaLigne(ARow : integer) : boolean;
    procedure AjouteLigne(TobAjout : TOB = nil);
    procedure VideLigne(ARow : integer);
    function  LigneVide(ARow : integer) : boolean;
    procedure SupprimeLigne(ARow : integer);
    function  ValideLesDonnees(ARow : integer) : boolean;
    //Actions de bvaliderclick
    function  ValideLaForm : boolean;
    procedure MAJTable;    //pour transaction

    end ;

//****************** Multi références ****************************************//

procedure ValideLesRef(TobMultiRef : TOB);

Type
  TOF_SAISIEREF = Class (TOF)
    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnClose                  ; override ;

    private
    G_Art : THGrid;
    TobComp : TOB;
    bValider,bFerme,bDelete : TToolBarButton97;
    i_art, ColCodArt,ColArt : integer;

    procedure G_ArtRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;Chg: Boolean);
    procedure G_ArtDessineCell( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState);
    Procedure G_ArtGetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;

    procedure bValiderClick(Sender: TObject);
    procedure bFermeClick(Sender: TObject);
    procedure bDeleteClick(Sender: TObject);

    procedure InitialiseGrid;
    procedure AfficheEcran;
    function  ValideLesDonnees(ARow : integer) : boolean;
    procedure SupprimeLigne(ARow : integer);
    end;

var TobRef,TobArticleTiers : TOB;
    CodeTiers : string;

const stLesColsArt   : string = 'GAT_ARTICLE;GA_CODEARTICLE;GAT_REFARTICLE;GAT_LIBELLE';
const stLesColsTiers : string = 'GAT_REFTIERS;GAT_REFARTICLE;GAT_LIBELLE';

const
// libellés des messages
TexteMessage: array[1..11] of string 	= (
          {1}  'Vous devez renseigner un '
          {2} ,'Vous devez renseigner la référence de cet article'
          {3} ,'La désignation doit être renseignée'
          {4} ,'Cette référence existe déjà chez ce tiers'
          {5} ,'L''article est déjà référencé pour ce client'
          {6} ,'Le code client n''existe pas.'
          {7} ,'Le code article n''existe pas.'
          {8} ,'Cet article est déjà référencé pour ce client,#13voulez vous le supprimer?'
          {9} ,'Voulez-vous enregistrer les modifications ?'
          {10} ,'ATTENTION : Mise à jour des données non effectuée !'
          {11} ,'ATTENTION : Ce référencement est en cours de modification par un autre utilisateur, '#13'mise à jour non effectuée !'
              );

Implementation


procedure TOF_ARTICLETIERS.OnArgument (StArgument : String ) ;
var i_ind : integer ;
    CodeArticle, Critere, ChampCritere, ValeurCritere : string ;
    TobA : TOB;
    Cancel : boolean;
begin
  inherited;
stFicSource:= '';
CodeArticle := '';
CodeTiers := '';
Article := '';
Data_Modified := false;
bValide := false;
//**** récupération des arguments ****//
Repeat
    Critere:=Trim(ReadTokenSt(StArgument)) ;
    if Critere<>'' then
        begin
        i_ind:=pos('=',Critere);
        if i_ind<>0 then
           begin
           ChampCritere:=copy(Critere,1,i_ind-1);
           ValeurCritere:=copy(Critere,i_ind+1,length(Critere));
           if ChampCritere='CODE_TIERS' then CodeTiers := ValeurCritere ;
           if ChampCritere='ARTICLE' then Article := ValeurCritere ;
           end;
        end;
until Critere = '';

//SetControlText('TINFO','');
SetControlProperty('PNDIMTOP', 'Caption', '');

if (CodeTiers<>'') then
   begin
   stFicSource:= 'TIERS';
   LesColsRef := stLesColsArt;
   end;

if (Article='') then
   Ecran.caption := Ecran.caption + ' pour le client ' + CodeTiers + ' ' +
                                   RechDom('GCTIERSCLI',CodeTiers,false)
else begin
     CodeArticle := Trim(Copy(Article,0,18));
     LibArticle :=  RechDom('GCARTICLEGENERIQUE',CodeArticle,False);
     //Affichage des dimensions
     Ecran.caption := Ecran.caption + ' de l''article  ' + CodeArticle + ' ' + LibArticle;
     TobA := TOB.Create('ARTICLE',nil,-1);
     TobA.PutValue('GA_ARTICLE',Article);
     TobA.LoadDB();
     SetControlProperty('PNDIMTOP', 'Caption', RechDim(TobA));
     SetControlVisible('GRIDDIM', False);
     SetControlProperty('PNBOTTOM', 'Height', THPanel(GetControl('PNDIMTOP')).Height);
     SetControlVisible('PNBOTTOM', True);
     stFicSource:='ARTICLE';
     LesColsRef := stLesColsTiers;
     TobA.Free;
     end;

//Initialisation de la fiche
G_Ref:=THGrid(GetControl('G_REF'));
G_Ref.OnRowEnter := G_RefRowEnter;
G_Ref.OnRowExit  := G_RefRowExit;
G_Ref.OnCellExit := G_RefCellExit;
G_Ref.OnCellEnter := G_RefCellEnter;
G_Ref.OnElipsisClick := G_RefElipsisClick;
G_Ref.OnDblClick := G_RefDblClick;
G_Ref.PostDrawCell := DessineCell;
G_Ref.OnKeyDown:=FormKeyDown ;
//Boutons
bInsert  := TToolBarButton97(GetControl('bInsert'));
bDelete  := TToolBarButton97(GetControl('bDelete'));
bValider := TToolBarButton97(GetControl('btValid'));
bFerme   := TToolBarButton97(GetControl('bFerme'));
bMultiSel := TToolBarButton97(GetControl('btMultiRef'));
if Assigned(bInsert) then
  bInsert.OnClick  := bInsertClick;
if Assigned(bDelete) then
  bDelete.OnClick  := bDeleteClick;
if Assigned(bValider) then
  bValider.OnClick := bValiderClick;
if Assigned(bMultiSel) then
  bMultiSel.OnClick := bMultiSelClick;

Ctrl := THCritMaskEdit.Create (nil);
if stFicSource = 'ARTICLE' then
   begin
   Ctrl.DataType := 'GCTIERSCLI';
   bMultiSel.Visible := false;
   end else Ctrl.DataType := 'GCARTICLE';
Ctrl.Text := '';
bNew := false;
ChargeTob(Article);
InitGrid;
TFVierge(Ecran).Hmtrad.ResizeGridColumns(G_Ref);
G_RefRowEnter(nil, 1, Cancel, False);
end ;

procedure TOF_ARTICLETIERS.OnClose ;
begin
  Inherited ;
if not bValide then
   if Data_Modified then
      Case PGIAskCancel(TexteMessage[9],Ecran.Caption) of
        mrYes :    if not ValideLaForm then
                      begin
                      Ecran.ModalResult := 0;
                      exit;
                      end;
        mrCancel : begin
                   Ecran.ModalResult := 0;
                   exit;
                   end;
        end;
TobArticleTiers.Free;
TobDelete.Free;
Ctrl.Free;
end ;

////////////////////////////////////////////////////////////////////////////////
//**************************** Gestion des événements ************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_ARTICLETIERS.G_RefRowEnter(Sender: TObject; Ou: Integer;
                                                var Cancel: Boolean;Chg: Boolean);
begin
  G_Ref.InvalidateRow(Ou);
  AffichePiedInfo(G_Ref.Row);
  SupprimeTobFille;
if bnew then
   begin
   bNew := false;
   bInsert.Enabled := true;
   end;
end;

procedure TOF_ARTICLETIERS.G_RefRowExit(Sender: TObject; Ou: Integer;
                                               var Cancel: Boolean;Chg: Boolean);
begin
G_Ref.InvalidateRow(Ou);
if ValideLaLigne(Ou) and ValideLesDonnees(Ou) then CreerTobFille(Ou);
end;

procedure TOF_ARTICLETIERS.G_RefCellEnter(Sender: TObject; var ACol, ARow: Integer;
                                                              var Cancel: Boolean);
begin
G_Ref.ElipsisButton := (G_Ref.Col = GCod);
Ctrl.Text := G_Ref.Cells[G_Ref.Col,G_Ref.Row];
Cell_Text := G_Ref.Cells[G_Ref.Col,G_Ref.Row];
end;

procedure TOF_ARTICLETIERS.G_RefCellExit(Sender: TObject; var ACol, ARow: Integer;
                                                          var Cancel: Boolean);
begin
if (Cell_Text <> '') and (G_Ref.Cells[ACol,ARow] <> Cell_Text) then Data_Modified := true;
if (ACol = GCod )or (ACol = GRefA) then
   G_Ref.Cells[ACol,ARow] := UpperCase(G_Ref.Cells[ACol,ARow]);
if ACol <> GCod then exit;
if G_Ref.Cells[GCod,ARow] = Ctrl.Text then exit;
if G_Ref.Cells[GCod,ARow] = '' then
   begin
   VideLigne(ARow);
   exit;
   end;
GereCode(GCod,ARow);
end;

procedure TOF_ARTICLETIERS.G_RefElipsisClick(Sender: TObject);
begin
GereCode(G_Ref.Col,G_Ref.Row);
end;

procedure TOF_ARTICLETIERS.G_RefDblClick(Sender: TObject);
begin
GereCode(G_Ref.Col,G_Ref.Row);
end;

procedure TOF_ARTICLETIERS.FormKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
Var OkG,Vide : Boolean ;
begin
OkG:=(Screen.ActiveControl=G_Ref) ; Vide:=(Shift=[]) ;
Case Key of
   VK_RETURN : Key:=VK_TAB ;
   VK_F5     : if ((OkG) and (Vide)) then begin Key:=0;GereCode(G_Ref.Col,G_Ref.Row);end;
   //VK_INSERT : if ((OkG) and (Vide)) then begin Key:=0;bInsertClick(Sender) ;end;
   VK_DELETE : if ((OkG) and (Shift=[ssCtrl])) then begin Key:=0;bDeleteClick(Sender) ;end;
   end ;
end;


////////////////////////////////////////////////////////////////////////////////
//*************************** Initialisation du Grid  ************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_ARTICLETIERS.InitGrid;
begin
EtudieColsListe ;
AfficheLibelle;
if stFicSource <> 'ARTICLE' then
   begin
   G_Ref.FixedCols := 2;
   G_Ref.ColWidths[GArt] := -1;
   end;
G_Ref.ColWidths[GInfo] := -1;
G_Ref.ColLengths[GInfo] := -1;
if stFicSource = 'ARTICLE' then
   if TobArticleTiers.Detail.Count = 0 then G_Ref.Cells[GLibA,1] := LibArticle;
end;


procedure TOF_ARTICLETIERS.EtudieColsListe;
Var NomCol, LesCols : String;
    icol : Integer;
begin
LesCols := 'FIXED;' + LesColsRef + ';INFO';
icol := 0;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        begin
        if NomCol='FIXED'          then GFixed:=icol   else
        if NomCol='INFO'           then GInfo:=icol    else
        if NomCol='GAT_REFARTICLE' then GRefA:=icol    else
        if NomCol='GAT_ARTICLE'    then GArt:=icol     else
        if NomCol='GA_CODEARTICLE' then GCodArt:=icol  else
        if NomCol='GAT_REFTIERS'   then GRefT:=icol    else
        if NomCol='GAT_LIBELLE'    then GLibA:=icol;
        end;
    Inc(icol) ;
    Until ((LesCols='') or (NomCol='')) ;
G_Ref.ColCount := icol;
if stFicSource <> 'ARTICLE' then GCod := GCodArt
else GCod := GRefT;

end;

procedure TOF_ARTICLETIERS.AfficheLibelle;
var LesCols,NomCol : string;
    icol : integer;
begin
LesCols := 'FIXED;'+LesColsRef+';INFO';
icol := 0;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        begin
        if NomCol='FIXED' then
           begin
           G_Ref.Cells[icol,0]:='';
           G_Ref.ColWidths[icol] := 8;
           end else
        if NomCol='GAT_REFARTICLE' then
           begin
           G_Ref.Cells[icol,0]:='Référence';
           G_Ref.ColWidths[icol] := 70;
           end else
        if NomCol='GA_CODEARTICLE' then
           begin
           G_Ref.Cells[icol,0]:='Article';
           G_Ref.ColWidths[icol] := 70;
           end else
        if NomCol='GAT_REFTIERS'   then
           begin
           G_Ref.Cells[icol,0]:='Code client';
           G_Ref.ColWidths[icol] := 70;
           end else
        if NomCol='GAT_LIBELLE'    then
           begin
           G_Ref.Cells[icol,0]:='Désignation';
           G_Ref.ColWidths[icol] := 120;
           end;
        end;
    Inc(icol) ;
    Until ((LesCols='') or (NomCol='')) ;
end;

procedure TOF_ARTICLETIERS.DessineCell( ACol,ARow : Longint; Canvas : TCanvas ;
                                                             AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < G_Ref.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=G_Ref.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := G_Ref.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = G_Ref.row) then
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

////////////////////////////////////////////////////////////////////////////////
//**************************** Gestion des boutons  **************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_ARTICLETIERS.bInsertClick(Sender: TObject);
begin
if LigneVide(G_Ref.Row) then exit;
if not QuitLaLigne(G_Ref.Row) then exit;
binsert.Enabled := false;
AjouteLigne();
end;

procedure TOF_ARTICLETIERS.bDeleteClick(Sender: TObject);
begin
SupprimeLigne(G_Ref.Row);
if bnew then
   begin
   bNew := false;
   bInsert.Enabled := true;
   end;
end;

procedure TOF_ARTICLETIERS.bValiderClick(Sender: TObject);
begin
if not ValideLaForm then Ecran.ModalResult := 0
else bValide := true;
end;

procedure TOF_ARTICLETIERS.bMultiSelClick(Sender: TObject);
var i_art,i_ind1, i_ind2 : integer;
    stNbArt,stChamp,ValChamp : string;
    Cancel : boolean;
    TobMulti,TobM, TobArt, TobVerif : TOB;
    QArt : TQuery;
begin
if not QuitLaLigne(G_Ref.Row) then exit;
stNbArt := AGLLanceFiche('GC','GCMULTISELART','','',CodeTiers+';'+RechDom('GCTIERSCLI',CodeTiers,false));
if stNbArt = '' then
   begin
   G_RefRowEnter(nil, G_Ref.Row, Cancel, False);
   exit;
   end;
i_art := StrToInt(stNbArt);
TobMulti := TOB.Create('ARTICLETIERS',nil,-1);
TobArt := TOB.Create('ARTICLE',nil,-1);
for i_ind1 := 1 to i_art do
    begin
    TobM := TOB.Create('ARTICLETIERS',TobMulti,-1);
    TobM.AddChampSup('GA_CODEARTICLE',false);
    TobM.AddChampSup('LIBELLE',false);
    TobM.AddChampSup('INFO',false);
    stChamp := 'ArticleNo'+ IntToStr(i_ind1);
    ValChamp := TheTOB.GetValue(stChamp);
    TobM.PutValue('GAT_ARTICLE',ValChamp);
    TobM.PutValue('GA_CODEARTICLE',Copy(ValChamp,0,18));
    TobM.PutValue('GAT_REFTIERS',CodeTiers);
    QArt := OpenSQL('SELECT GA_ARTICLE,GA_LIBELLE,GA_CODEARTICLE,GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3'+
                 ',GA_GRILLEDIM4,GA_GRILLEDIM5,GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3'+
                 ',GA_CODEDIM4,GA_CODEDIM5 FROM ARTICLE WHERE GA_ARTICLE="'+ValChamp+'"',false);
    if TOBArt.SelectDB('',QArt) then
       begin
       TobM.PutValue('INFO',RechDim(TobArt));
       TobM.PutValue('LIBELLE',TobArt.GetValue('GA_LIBELLE'));
       end;
    Ferme(QArt);
    end;
if TobMulti.Detail.Count = 1 then AjouteLigne(TobMulti.Detail[0])
else  begin
      ValideLesRef(TobMulti);
      if TobMulti.Detail.Count > 0 then
        begin
        for i_ind2 := TobMulti.Detail.Count -1 downto 0 do
          begin
          TobVerif := TobDelete.FindFirst(['GAT_ARTICLE','GAT_REFTIERS','GAT_REFARTICLE'],
                   [TobMulti.Detail[i_ind2].GetValue('GAT_ARTICLE'),
                    TobMulti.Detail[i_ind2].GetValue('GAT_REFTIERS'),
                    TobMulti.Detail[i_ind2].GetValue('GAT_REFARTICLE')],false);
          if TobVerif <> nil then TobVerif.Free;
          if (G_Ref.RowCount-1 >= 1) and not (LigneVide(G_Ref.RowCount-1)) then
             G_Ref.RowCount:=G_Ref.RowCount+1;
          TobMulti.Detail[i_ind2].PutLigneGrid(G_Ref,G_Ref.RowCount-1,false,false,'FIXED;'+LesColsRef+';INFO');
          TobMulti.Detail[i_ind2].ChangeParent(TobArticleTiers,-1);
          end;
        G_Ref.SortGrid(GCod,false);
        Data_Modified := true;
        end;
      G_RefRowEnter(nil, 1, Cancel, False);
      end;
TobMulti.Free;
TheTOB := Nil;
TobArt.Free;
end;

////////////////////////////////////////////////////////////////////////////////
//***************************** Gestion des Tobs  ****************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_ARTICLETIERS.ChargeTob;
var QRef : TQuery;
    TobR : TOB;
    stSelect,stChampDim,stJoin,stWhere : string;
    i_ind : integer;
begin
TobArticleTiers := TOB.Create('',nil,-1);
if stFicSource = 'ARTICLE' then
   begin
   stWhere := ' GAT_ARTICLE="' + Article + '"';
   end else
   begin
   stWhere := ' GAT_REFTIERS="'   + CodeTiers   + '"';
   stChampDim := ',GA_STATUTART,GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3'+
                 ',GA_GRILLEDIM4,GA_GRILLEDIM5,GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3'+
                 ',GA_CODEDIM4,GA_CODEDIM5';
   end;
stJoin :=  'LEFT JOIN ARTICLE ON GAT_ARTICLE=GA_ARTICLE';
stWhere := stWhere + ' AND GA_STATUTART<>"GEN"';
stSelect := StringReplace(LesColsRef,';',',',[rfReplaceAll,rfIgnoreCase]);
if stFicSource = 'ARTICLE' then stSelect := stSelect + ',GAT_ARTICLE'
else stSelect := stSelect + ',GAT_REFTIERS';

QRef := OpenSQL('SELECT ' + stSelect + stChampDim + ' FROM ARTICLETIERS ' + stJoin +
                ' WHERE ' + stWhere + ' ORDER BY ' + stSelect,true);
if not QRef.Eof then TobArticleTiers.LoadDetailDB('ARTICLETIERS','','',QRef,false);
Ferme(QRef);
if stFicSource <> 'ARTICLE' then
  for i_ind := 0 to TobArticleTiers.Detail.Count -1 do
      begin
      TobR := TobArticleTiers.Detail[i_ind];
      if TobR.GetValue('GA_STATUTART') <> 'DIM' then continue;
      TobR.AddChampSup('INFO',false);
      TobR.PutValue('INFO',RechDim(TobR));
      end;
TobArticleTiers.PutGridDetail(G_Ref,false,false,LesColsRef+';INFO',true);
//TobArticleTiers.PutGridDetail(G_Ref,false,false,'FIXED;'+LesColsRef+';INFO',true);

//Création de la tob de suppression des enreg dans la table
TobDelete := TOB.Create('',nil,-1);
end;

procedure TOF_ARTICLETIERS.SupprimeTobFille;
var TobA : TOB;
    RefTiers,RefArt : string;
begin
if stFicSource = 'ARTICLE' then RefTiers := G_Ref.Cells[GRefT,G_Ref.Row]
else RefTiers := CodeTiers;
RefArt := G_Ref.Cells[GRefA,G_Ref.Row];
TobA := TobArticleTiers.FindFirst(['GAT_REFTIERS','GAT_REFARTICLE'],[RefTiers,RefArt],false);
if TobA <> nil then TobA.ChangeParent(TobDelete,-1);
end;

procedure TOF_ARTICLETIERS.CreerTobFille(ARow : integer);
var TobD : TOB;
    Art, RefTiers, RefArt : string;
begin
if stFicSource = 'ARTICLE' then
   begin
   Art := Article;
   RefTiers := G_Ref.Cells[GRefT,ARow];
   end else
   begin
   RefTiers := CodeTiers;
   Art := G_Ref.Cells[GArt,ARow];
   end;
RefArt := G_Ref.Cells[GRefA,ARow];
TobD := TobDelete.FindFirst(['GAT_REFTIERS','GAT_REFARTICLE','GAT_ARTICLE'],[RefTiers,RefArt,Art],false);
if TobD <> nil then
     TobD.ChangeParent(TobArticleTiers,-1)
else TobD := TOB.Create('ARTICLETIERS',TobArticleTiers,-1);
TobD.PutValue('GAT_ARTICLE',Art);
TobD.PutValue('GAT_REFARTICLE',RefArt);
TobD.PutValue('GAT_REFTIERS',RefTiers);
TobD.PutValue('GAT_LIBELLE',G_Ref.Cells[GLibA,ARow]);
end;

////////////////////////////////////////////////////////////////////////////////
//******************************** Recherches  *******************************//
////////////////////////////////////////////////////////////////////////////////

function TOF_ARTICLETIERS.Recherche_Art(ARow : LongInt) : boolean;
var RechArt : T_RechArt ;
    OkArt   : Boolean ;
    TOBArt : TOB;
    stArt,stOldArt : string;
begin
OkArt:=False ;
TOBArt := TOB.Create('ARTICLE',nil,-1);
RechArt := TrouverArticle (Ctrl.Text, TOBArt);
case RechArt of
     traOk : OkArt:=True ;

     traAucun : begin
                stOldArt := Ctrl.Text;
{$IFDEF GPAO}
	              stArt := DispatchRecherche(Ctrl, 1, '', 'GA_CODEARTICLE=' + Ctrl.Text + ';RECHERCHEARTICLE;', NomMulRechArticle);
{$ELSE}
                stArt := DispatchRecherche(Ctrl, 1, '', 'GA_CODEARTICLE=' + Ctrl.Text, '');
{$ENDIF}
                if (stArt <> '') and (stArt <> stOldArt) then
                   begin
                   TOBArt.PutValue('GA_ARTICLE',stArt);
                   TOBArt.LoadDB();
                   Ctrl.Text:=TobArt.GetValue('GA_CODEARTICLE');
                   if TobArt.GetValue('GA_STATUTART')='GEN'  then
                      ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE'), TOBArt);
                   OkArt := True;
                   end;
                end;

     traGrille : begin
                 if ChoisirDimension (TOBArt.GetValue ('GA_ARTICLE'), TOBArt) then
                    OkArt := True;
                 end;
         end ;

if (OkArt) then
   begin
   G_Ref.Cells[GCodArt,ARow]:=TobArt.GetValue('GA_CODEARTICLE');
   G_Ref.Cells[GArt,ARow]:=TobArt.GetValue('GA_ARTICLE');
   G_Ref.Cells[GLibA,ARow]:=TobArt.GetValue('GA_LIBELLE');
   if TobArt.GetValue('GA_STATUTART')='DIM' then
      G_Ref.Cells[GInfo,ARow] := RechDim(TobArt);
   end else  VideLigne(Arow);
Result := okArt;
end;

function TOF_ARTICLETIERS.Recherche_Tiers(ARow : integer) : boolean;
var stTiers : string;
begin
result := true;
stTiers := G_Ref.Cells[GCod,ARow];
if ExisteSQL('SELECT T_TIERS FROM TIERS WHERE T_TIERS="' + stTiers + '"') then exit;
Ctrl.Text:=DispatchRecherche(Ctrl, 2, 'T_NATUREAUXI="CLI"','T_TIERS="' + Ctrl.Text + '"', '');
if stTiers = Ctrl.Text then
   begin
   VideLigne(Arow);
   result := false;
   exit;
   end;
G_Ref.Cells[GCod,ARow]:=Ctrl.Text;
end;

function TOF_ARTICLETIERS.RechDim(TobArt : TOB) : string;
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

////////////////////////////////////////////////////////////////////////////////
//**************************** Gestion des lignes  ***************************//
////////////////////////////////////////////////////////////////////////////////

procedure TOF_ARTICLETIERS.GereCode(ACol,ARow : integer);
var Coord : TRect;
    okCode : boolean;
    QTiers :TQuery;
    CodeAuxi : string;
begin
OkCode := false;
CodeAuxi:='';
if (Ctrl.Text = G_Ref.Cells[GCod,ARow]) and (G_Ref.Cells[GCod,ARow] <> '') then
   begin
   if stFicSource <> 'ARTICLE' then
  {$IFNDEF GPAO}
    AGLLanceFiche('GC','GCARTICLE','',G_Ref.Cells[GArt, ARow],'ACTION=CONSULTATION;MONOFICHE')
  {$ELSE}
	  V_PGI.DispatchTT(7, taConsult, G_Ref.Cells[GArt, ARow], 'MONOFICHE', '')
  {$ENDIF}
   else begin
         QTiers := OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+G_Ref.Cells[GCod,ARow]+'"',false);
         if not QTiers.Eof then CodeAuxi := QTiers.FindField('T_AUXILIAIRE').AsString;
         Ferme(QTiers);
         AGLLanceFiche('GC','GCTIERS','',CodeAuxi,'ACTION=CONSULTATION;MONOFICHE');
         end;
   exit;
   end;

if ACol = GCod then
    begin
    Coord := G_Ref.CellRect (ACol, ARow);
    Ctrl.Parent := G_Ref;
    Ctrl.Top := Coord.Top;
    Ctrl.Left := Coord.Left;
    Ctrl.Width := 3; Ctrl.Visible := False;
    Ctrl.Text:=G_Ref.Cells[ACol,ARow];
    if stFicSource <> 'ARTICLE' then okCode := Recherche_Art(ARow)
    else okCode := Recherche_Tiers(ARow);
    end ;
AffichePiedInfo(ARow);
if bNew and okCode then binsert.Enabled := true;
end;

procedure TOF_ARTICLETIERS.EntreeCell(Acol,ARow : integer);
begin
G_Ref.Col := ACol;
G_Ref.Row := ARow;
G_Ref.ElipsisButton := (G_Ref.Col = GCod);
Ctrl.Text := G_Ref.Cells[G_Ref.Col,G_Ref.Row];
end;

function TOF_ARTICLETIERS.QuitLaLigne(ARow : integer) : boolean;
var ACol : integer;
begin
ACol := G_Ref.Col;
result := true;
if LigneVide(ARow) and (G_Ref.RowCount-1 <= 1) then exit;
if (ACol = GCod )or (ACol = GRefA) then
   G_Ref.Cells[ACol,ARow] := UpperCase(G_Ref.Cells[ACol,ARow]);
if ACol = GCod then
   if G_Ref.Cells[GCod,ARow] <> Ctrl.Text then
      if G_Ref.Cells[GCod,ARow] = '' then VideLigne(ARow)
      else GereCode(GCod,ARow);
if ValideLaLigne(ARow) and ValideLesDonnees(ARow) then CreerTobFille(ARow)
else result := false;
end;

procedure TOF_ARTICLETIERS.AffichePiedInfo(ARow : integer);
begin
if stFicSource = 'ARTICLE' then exit;
if G_Ref.Cells[GInfo,ARow] <> '' then
  SetControlProperty('PNDIMTOP', 'Caption', G_Ref.Cells[GInfo,ARow])
//  SetControlText('TINFO',G_Ref.Cells[GInfo,ARow])
else
//  SetControlText('TINFO','');
    SetControlProperty('PNDIMTOP', 'Caption', '');
end;

function TOF_ARTICLETIERS.ValideLaLigne(ARow : integer) : boolean;
var stRech : string;
begin
result := true;
if bNew and LigneVide(ARow) then
   begin
   SupprimeLigne(ARow);
   Result:=false;
   exit;
   end;

if stFicSource <> 'ARTICLE' then stRech := 'article'
else stRech := 'tiers';
//pas d'article ou tiers
if G_Ref.Cells[GCod,ARow]='' then
   begin
   HShowMessage('0;'+TFVierge(Ecran).Caption+';'+TexteMessage[1]+stRech+';W;O;O;O;','','') ;
   EntreeCell(GCod,ARow);
   G_Ref.InvalidateRow(G_Ref.Row);
   result := false;
   exit;
   end;

//pas de référence tiers de l'article
if G_Ref.Cells[GRefA,ARow]='' then
   begin
   HShowMessage('0;'+TFVierge(Ecran).Caption+';'+TexteMessage[2]+';W;O;O;O;','','') ;
   EntreeCell(GRefA,ARow);
   G_Ref.InvalidateRow(G_Ref.Row);
   result := false;
   exit;
   end;

//pas de désignation de l'article
if G_Ref.Cells[GLibA,ARow]='' then
   begin
   HShowMessage('0;'+TFVierge(Ecran).Caption+';'+TexteMessage[3]+';W;O;O;O;','','') ;
   EntreeCell(GLibA,ARow);
   G_Ref.InvalidateRow(G_Ref.Row);
   result := false;
   exit;
   end;
end;

procedure TOF_ARTICLETIERS.AjouteLigne(TobAjout : TOB = nil);
var ARow : integer;
begin
Data_Modified := true;
bNew := true;
if (TobAjout <> nil) and LigneVide(G_Ref.Row) then ARow := G_Ref.RowCount -1
else begin
     G_Ref.RowCount := G_Ref.RowCount + 1;
     G_Ref.InvalidateRow(G_Ref.Row);
     ARow := G_Ref.RowCount -1;
     end;

if TobAjout <> nil then //Cas de multi-références mais un article retourné
   begin
   G_Ref.Cells[GCodArt,ARow]:= TobAjout.GetValue('GA_CODEARTICLE');
   G_Ref.Cells[GArt,ARow]   := TobAjout.GetValue('GAT_ARTICLE');
   G_Ref.Cells[GLibA,ARow]  := TobAjout.GetValue('LIBELLE');
   G_Ref.Cells[GInfo,ARow]  := TobAjout.GetValue('INFO');
   EntreeCell(GRefA,ARow);
   end else
   begin
   if stFicSource = 'ARTICLE' then G_Ref.Cells[GLibA,ARow] := LibArticle;
   EntreeCell(GCod,ARow);
   end;
AffichePiedInfo(ARow);
end;

procedure TOF_ARTICLETIERS.VideLigne(ARow : integer);
var i_ind : integer;
begin
for i_ind := 1 to G_Ref.ColCount -1 do
    begin
    if i_ind = GRefA then continue;
    G_Ref.Cells[i_ind,ARow] := '';
    end;
end;

function TOF_ARTICLETIERS.LigneVide(ARow : integer) : boolean;
var i_ind : integer;
begin
result := true;
for i_ind := 1 to G_Ref.ColCount -1 do
    if (G_Ref.Cells[i_ind,ARow] <> '') and (i_ind <> GLibA) then
       begin
       result := false;
       exit;
       end;
end;

procedure TOF_ARTICLETIERS.SupprimeLigne(ARow : integer);
var Cancel : boolean;
begin
if G_Ref.RowCount <= 1 then Exit ;
if G_Ref.RowCount = 2 then
   begin
   G_Ref.VidePile(False);
   if stFicSource = 'ARTICLE' then G_Ref.Cells[GLibA,1] := LibArticle;
   AffichePiedInfo(1);
   end else
   begin
   G_Ref.CacheEdit ; G_Ref.SynEnabled := False;
   G_Ref.DeleteRow (ARow);
   G_Ref.MontreEdit; G_Ref.SynEnabled := True;
   G_RefRowEnter(nil, G_Ref.Row, Cancel, False);
   Ctrl.Text := G_Ref.Cells[GCod,G_Ref.Row];
   end;
end;

function TOF_ARTICLETIERS.ValideLesDonnees(ARow : integer) : boolean;
var Art,RefTiers,RefArt : string;
    TobVerif : TOB;
begin
result := true;
if stFicSource = 'ARTICLE' then Art := Article
else RefTiers := CodeTiers;
if stFicSource = 'ARTICLE' then RefTiers := G_Ref.Cells[GRefT,ARow]
else Art := G_Ref.Cells[GArt,ARow];
RefArt := G_Ref.Cells[GRefA,ARow];

TobVerif := TobArticleTiers.FindFirst(['GAT_REFTIERS','GAT_REFARTICLE'],[RefTiers,RefArt],false);
if TobVerif <> nil then
    begin
    HShowMessage('0;'+TFVierge(Ecran).Caption+';'+TexteMessage[4]+';W;O;O;O;','','') ;
    EntreeCell(GRefA,ARow);
    G_Ref.InvalidateRow(G_Ref.Row);
    result := false;
    exit;
    end;

TobVerif := TobArticleTiers.FindFirst(['GAT_ARTICLE','GAT_REFTIERS'],[Art,RefTiers],false);
if TobVerif <> nil then
    begin
    HShowMessage('0;'+TFVierge(Ecran).Caption+';'+TexteMessage[5]+';W;O;O;O;','','') ;
    EntreeCell(GCod,ARow);
    G_Ref.InvalidateRow(G_Ref.Row);
    result := false;
    exit;
    end;
end;

function  TOF_ARTICLETIERS.ValideLaForm : boolean;
var ioerr : TIOErr ;
begin
result := true;
if not QuitLaLigne(G_Ref.Row) then
   begin
   result := false;
   exit;
   end;
ioerr := Transactions (MAJTable, 3);
Case ioerr of
    oeUnknown : begin
                HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[10]+';W;O;O;O;','','') ;
                result := false;
                exit;
                end;

    oeSaisie  : begin
                HShowMessage('0;'+Ecran.Caption+';'+TexteMessage[11]+';W;O;O;O;','','') ;
                result := false;
                exit;
                end ;
    end;
end;

procedure TOF_ARTICLETIERS.MAJTable;
begin
TobDelete.DeleteDB();
TobArticleTiers.InsertOrUpdateDB();
end;

////////////////////////////////////////////////////////////////////////////////
///*************************** Multi références *****************************///
////////////////////////////////////////////////////////////////////////////////

//Entrée
procedure ValideLesRef(TobMultiRef : TOB);
begin
TobRef := TobMultiRef;
if TobRef.Detail.Count < 1 then exit;
AGLLanceFiche('GC','GCSAISIEREF','','','');
end;

procedure TOF_SAISIEREF.OnArgument(StArgument : String ) ;
var Cancel : boolean;
begin
///Fiche
Ecran.caption := CodeTiers + ' ' + RechDom('GCTIERSCLI',CodeTiers,false);
///Grid
G_Art:=THGrid(GetControl('G_ART'));
G_Art.OnRowEnter := G_ArtRowEnter;
G_Art.PostDrawCell := G_ArtDessineCell;
G_Art.GetCellCanvas := G_ArtGetCellCanvas;
///Boutons
bValider := TToolBarButton97(GetControl('bValider'));
bFerme   := TToolBarButton97(GetControl('bFerme'));
bDelete  := TToolBarButton97(GetControl('bDelete'));
bValider.OnClick  := bValiderClick;
bFerme.OnClick    := bFermeClick;
bDelete.OnClick   := bDeleteClick;
///Initialisation
i_art := 0;
TobComp := TOB.Create('',nil,-1);
InitialiseGrid;
G_Art.Row := i_art;
G_ArtRowEnter(nil,i_art,Cancel,false);
end;

procedure TOF_SAISIEREF.OnClose;
begin
Inherited;
TobComp.Free;
end;

///**** Gestion du Grid ****///
procedure TOF_SAISIEREF.G_ArtRowEnter(Sender: TObject; Ou: Integer;
                                              var Cancel: Boolean;Chg: Boolean);
begin
G_Art.InvalidateRow(Ou);
AfficheEcran;
SetFocusControl('GAT_REFARTICLE');
end;

procedure TOF_SAISIEREF.G_ArtDessineCell( ACol,ARow : Longint; Canvas : TCanvas ;
                                                             AState: TGridDrawState);
var Triangle : array[0..2] of TPoint ;
    Arect: Trect ;
begin
If Arow < G_Art.Fixedrows then exit ;
if (gdFixed in AState) and (ACol = 0) then
    begin
    Arect:=G_Art.CellRect(Acol,Arow) ;
    Canvas.Brush.Color := G_Art.FixedColor;
    Canvas.FillRect(ARect);
      if (ARow = G_Art.row) then
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

Procedure TOF_SAISIEREF.G_ArtGetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
begin
if (ARow = G_Art.Row) and (ACol = ColCodArt) then
    Canvas.Font.Style := [fsBold];
end;

///**** Gestion des boutons ****///
procedure TOF_SAISIEREF.bValiderClick(Sender: TObject);
var ARow : integer;
    Cancel : boolean;
    TobR : TOB;
begin
ARow := G_Art.Row;
TobR := TobRef.Detail[ARow];
TobR.PutValue('GAT_REFARTICLE',GetControlText('GAT_REFARTICLE'));
TobR.PutValue('GAT_LIBELLE',GetControlText('GAT_LIBELLE'));
if ValideLesDonnees(ARow) then
   begin
   if ARow >= TobRef.Detail.Count -1 then exit;
   G_Art.InvalidateRow(ARow);
   ARow := ARow + 1;
   G_Art.Row := ARow;
   end else if G_Art.RowCount > TobRef.Detail.Count then
               begin
               TFVierge(Ecran).Close; //Cas de la suppression dans ValideLesDonnees
               exit;                  //si référence déjà existante
               end;
G_ArtRowEnter(nil,ARow,Cancel,false);
Ecran.ModalResult :=0;
end;

procedure TOF_SAISIEREF.bFermeClick(Sender: TObject);
begin
TobRef.ClearDetail;
end;

procedure TOF_SAISIEREF.bDeleteClick(Sender: TObject);
var Cancel,bStop : boolean;
begin
bStop := (G_Art.Row = TobRef.Detail.Count-1);
SupprimeLigne(G_Art.Row);
if bStop then
   begin
   TFVierge(Ecran).Close;
   exit;
   end;
G_ArtRowEnter(nil, G_Art.Row, Cancel, False);
end;

///**** Gestion de l'affichage et contrôle des données ****///
procedure TOF_SAISIEREF.AfficheEcran;
var TobR : TOB;
begin
TobR := TobRef.Detail[G_Art.Row];
if TobR = nil then exit;
SetControlCaption('TGAT_ARTICLE',Trim(TobR.GetValue('GA_CODEARTICLE')) + '  ' +
                                 TobR.GetValue('LIBELLE'));
SetControlCaption('TDIMENSIONS',TobR.GetValue('INFO'));
SetControlText('GAT_REFARTICLE','');
SetControlText('GAT_LIBELLE',TobR.GetValue('LIBELLE'));
end;

procedure TOF_SAISIEREF.InitialiseGrid;
begin
G_Art.ColWidths[0] := 6;
ColArt := 1; ColCodArt := 2;
G_Art.ColWidths[ColArt] := -1;

TobRef.PutGridDetail(G_Art,false,false,'FIXED;GAT_ARTICLE;GA_CODEARTICLE');
TFVierge(Ecran).Hmtrad.ResizeGridColumns(G_Art);
end;

function  TOF_SAISIEREF.ValideLesDonnees(ARow : integer) : boolean;
var RefTiers, RefArt, Art : string;
    TobVerif, TobC : TOB;
begin
result := true;
RefTiers := CodeTiers;
RefArt := GetControlText('GAT_REFARTICLE');
Art := G_Art.Cells[ColArt,ARow];
if RefArt = '' then
   begin
   HShowMessage('0;'+TFVierge(Ecran).Caption+';'+TexteMessage[2]+';W;O;O;O;','','') ;
   SetFocusControl('GAT_REFARTICLE');
   result := false;
   exit;
   end;

if GetControlText('GAT_LIBELLE') = '' then
   begin
   HShowMessage('0;'+TFVierge(Ecran).Caption+';'+TexteMessage[3]+';W;O;O;O;','','') ;
   SetFocusControl('GAT_LIBELLE');
   result := false;
   exit;
   end;

TobVerif := TobArticleTiers.FindFirst(['GAT_REFTIERS','GAT_REFARTICLE'],[RefTiers,RefArt],false);
if TobVerif = nil then
   TobVerif := TobComp.FindFirst(['GAT_REFTIERS','GAT_REFARTICLE'],[RefTiers,RefArt],false);
if TobVerif <> nil then
    begin
    HShowMessage('0;'+TFVierge(Ecran).Caption+';'+TexteMessage[4]+';W;O;O;O;','','') ;
    SetFocusControl('GAT_REFARTICLE');
    result := false;
    exit;
    end;

TobVerif := TobArticleTiers.FindFirst(['GAT_ARTICLE','GAT_REFTIERS'],[Art,RefTiers],false);
if TobVerif = nil then
   TobVerif := TobComp.FindFirst(['GAT_ARTICLE','GAT_REFTIERS'],[Art,RefTiers],false);
if TobVerif <> nil then
    begin
    result := false;
    if PGIAsk(TexteMessage[8],Ecran.Caption)=mrYes then
       begin
       SupprimeLigne(ARow);
       exit;
       end;
    SetFocusControl('GAT_REFARTICLE');
    exit;
    end;
TobC := TOB.Create('',TobComp,-1);
TobC.Dupliquer(TobRef.Detail[Arow],true,true,false);
end;

procedure TOF_SAISIEREF.SupprimeLigne(ARow : integer);
begin
if ARow = TobRef.Detail.Count -1 then
   begin
   G_Art.VidePile(False);
   TobRef.Detail[ARow].Free;
   end else
   begin
   G_Art.CacheEdit ; G_Art.SynEnabled := False;
   G_Art.DeleteRow (ARow);
   G_Art.MontreEdit; G_Art.SynEnabled := True;
   TobRef.Detail[ARow].Free;
   end;
end;

Initialization
registerclasses ( [ TOF_ARTICLETIERS ] ) ;
registerclasses ( [ TOF_SAISIEREF ] ) ;

end.
