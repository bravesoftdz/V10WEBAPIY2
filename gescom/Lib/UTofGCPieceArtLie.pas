{***********UNITE*************************************************
Auteur  ...... : Olivier Tarcy
Créé le ...... : 30/05/2000
Modifié le ... : 16/06/2000
Description .. : Sélection d'articles liés lors de la saisie de pièces :
Suite ........ :   - les articles obligatoires seront sélectionnés d'office
Suite ........ :   - les autres seront à sélectionner par l'utilisateur
Suite ........ :   - dans le cas d'un article générique, la fiche de sélection 
Suite ........ :     de dimension s'affichera
Suite ........ : 
Suite ........ : La sélection est renvoyée dans une String comportant le 
Suite ........ : code unique de chaque article sélectionné
Mots clefs ... : PIECE;ARTICLELIE
*****************************************************************}
unit UTofGCPieceArtLie;

interface

uses  SysUtils, Controls, Classes, HCtrls, HMsgBox, UTOF, UTOB, Vierge,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
{$ENDIF}
      UtilArticle, FactUtil ;

Type

     TOF_GCPieceArtLie = Class (TOF)
       private
         CodeArticle : string ;
         GridArtLie : THGrid ;
         procedure OnBeforeFlip( Sender : TObject ; ARow : Longint ; Var Cancel : Boolean ) ;
       public
         procedure OnArgument (stArgument : string) ; override;
         procedure OnLoad ; override ;
         procedure OnUpdate ; override ;
         function SelectionIntoListe : String ;
END ;

const colCodeArticle = 0 ;
      colDesignation = 1 ;
      colObligatoire = 2 ;
      colQteRef      = 3 ;
      colCodeUnique  = 4 ;

function SelectArticleslies (Article : string) : String ;

implementation

{Récupération des arguments}
procedure TOF_GCPieceArtLie.OnArgument(stArgument : String);
var x : integer ; Critere, ChampCritere, ValeurCritere : string ;
begin
  inherited;
  Repeat
    Critere:=Trim(ReadTokenSt(StArgument)) ;
    if Critere<>'' then
      begin
        x:=pos('=',Critere);
        if x<>0 then
          begin
            ChampCritere:=copy(Critere,1,x-1);
            ValeurCritere:=copy(Critere,x+1,length(Critere));
            if ChampCritere='CODE_ARTICLE' then CodeArticle := ValeurCritere ;
          end;
      end;
  until  Critere='';
Ecran.Caption := 'Articles liés de : ' + CodeArticle + ' ' + RechDom('GCARTICLEGENERIQUE',CodeArticle,False) ;
GridArtLie := THGRID(GetControl('ARTICLES_LIES')) ;
GridArtLie.ColWidths[2] := 0 ;
GridArtLie.ColWidths[3] := 0 ;
// JLD 12/09/2001 GridArtLie.ColWidths[4] := 0 ;

{$IFDEF GIGI}
  // PL le 16/10/03 : en GI/GA on ne veut pas le même libellé en entête
  if (GetControl('LABEL_NOTE') <> nil) then
    SetControlProperty ('LABEL_NOTE', 'caption', 'Les articles ci-dessous sont liés à l''article choisi. Ceux ayant été définis comme "obligatoires" sont pré-sélectionnés (en italique). Vous pouvez sélectionner / dé-sélectionner des articles à saisir par la barre d''espace.');
{$ENDIF}

SetControlVisible ('HelpBtn', False); // DBR Fiche 10489
end ;

procedure TOF_GCPieceArtLie.OnBeforeFlip( Sender : TObject ; ARow : Longint ; Var Cancel : Boolean ) ;
var CodeUnique : string ;
begin
//if GridArtLie.cells[colObligatoire,ARow] = 'X' then Cancel:=True ;
if (GridArtLie.cells[colCodeUnique,ARow] <> '') and (not GridArtLie.IsSelected(ARow)) then
  begin
    CodeUnique:=CodeArticleUnique2(GridArtLie.CellValues[colCodeArticle,Arow],'') ;
    CodeUnique:=SelectUneDimension (CodeUnique) ;
    if CodeUnique=''then Cancel:=True else GridArtLie.cells[colCodeUnique,ARow] := CodeUnique ;
  end ;
end;

{Affichage des articles liés dans le THGrid}
procedure TOF_GCPieceArtLie.OnLoad ;
var QARTLIE : TQuery ;
    CodeArtLie,LibArtLie,ObligArtLie,QteRef,CodeUnique : string ;
    ListeArtLies,TobFille : TOB ;
    I : Integer ;
begin
  inherited;
  ListeArtLies := TOB.Create('ARTICLELIE',Nil,-1);
  QARTLIE:=OpenSQL('SELECT GAL_ARTICLELIE,GAL_LIBELLE,GAL_OBLIGATOIRE,GAL_QTEREF,GA_STATUTART FROM ARTICLELIE LEFT JOIN ARTICLE ON GA_CODEARTICLE=GAL_ARTICLELIE AND GA_STATUTART<>"DIM" WHERE GAL_TYPELIENART="LIE" AND GAL_ARTICLE="'+CodeArticle+'"',TRUE) ;
  While Not QARTLIE.EOF do
     begin
       CodeUnique := '' ;
       CodeArtLie := QARTLIE.Findfield('GAL_ARTICLELIE').AsString ;
       LibArtLie :=  QARTLIE.Findfield('GAL_LIBELLE').AsString ;
       ObligArtLie :=  QARTLIE.Findfield('GAL_OBLIGATOIRE').AsString ;
       QteRef :=  QARTLIE.Findfield('GAL_QTEREF').AsString ;
       if QARTLIE.findField('GA_STATUTART').asString = 'GEN' then CodeUnique:=CodeArticleUnique2(CodeArtLie,'') ;
       TobFille := TOB.Create('ARTICLELIE',ListeArtLies,-1) ;
       TobFille.addChampSup('ArtGen',TRUE);
       TobFille.PutValue('GAL_ARTICLELIE',CodeArtLie);
       TobFille.PutValue('GAL_LIBELLE',LibArtLie);
       TobFille.PutValue('GAL_OBLIGATOIRE',ObligArtLie);
       TobFille.PutValue('GAL_QTEREF',QteRef);
       TobFille.PutValue('ArtGen',CodeUnique);
       QARTLIE.Next ;
     end ;
  Ferme(QARTLIE) ;
  ListeArtLies.PutGridDetail(GridArtLie,False,False,'GAL_ARTICLELIE;GAL_LIBELLE;GAL_OBLIGATOIRE;GAL_QTEREF;ArtGen',False);
  ListeArtLies.Free ;

  For i:=GridArtLie.fixedrows to GridArtLie.RowCount-1 do
    if GridArtLie.cells[colObligatoire,i]='X' then
      begin
        if GridArtLie.cells[colCodeUnique,i] <> '' then
          begin
            CodeUnique:=SelectUneDimension (GridArtLie.cells[colCodeUnique,i]) ;
            GridArtLie.cells[colCodeUnique,i] := CodeUnique ;
          end ;
          GridArtLie.FlipSelection(i)
      end ;
  GridArtLie.OnBeforeFlip:=OnBeforeFlip ;
  GridArtLie.ColWidths[colCodeUnique] := 0 ;
end ;

procedure TOF_GCPieceArtLie.OnUpdate ;
begin
  inherited ;
  TFVierge(Ecran).Fretour:=SelectionIntoListe ;
end;

{Récupération de la sélection d'articles liés dans une String}
function TOF_GCPieceArtLie.SelectionIntoListe : String ;
var  i_row(*, reponse *) : integer ;
     pv, CodeUnique : string ;
begin
  result := '' ;                                    
  pv:='' ;
  if GridArtLie = Nil then exit ;
  for i_row := colDesignation to GridArtLie.RowCount-1 do
    if  GridArtLie.IsSelected(i_row) then
      begin
        if GridArtLie.CellValues[colCodeUnique,i_row] <> '' then // cas d'un article Dimensionné
            CodeUnique := GridArtLie.CellValues[colCodeUnique,i_row] 
        else CodeUnique:=CodeArticleUnique2(GridArtLie.CellValues[colCodeArticle,i_row],'') ;
        Result:=Result+pv+CodeUnique+';'+GridArtLie.CellValues[colQteRef,i_row];
        pv:=';' ;
        
      end ;
{  if (GridArtLie.nbSelected = 0) and (GridArtLie.CellValues[colCodeArticle,1] <> '') then
    begin
      reponse := PGIAsk('Aucun article n''a été sélectionné, confirmez-vous la sélection automatique des articles obligatoires ?','ATTENTION');
      if reponse = mrNo then Result:='';
    end;   }
end;

function SelectArticleslies (Article : string) : String ;
begin
  if not ExisteSQL('SELECT GAL_ARTICLELIE, GAL_LIBELLE, GAL_OBLIGATOIRE FROM ARTICLELIE WHERE GAL_TYPELIENART="LIE" AND GAL_ARTICLE="'+Article+'"') then
    begin result := '' ; {PGIInfo('Aucun article lié','INFORMATION'); } end
  else result:=AGLLanceFiche('GC','GCPIECEARTLIE','','','CODE_ARTICLE='+Article) ;
end;

Initialization
registerclasses([TOF_GCPieceArtLie]);

end.
