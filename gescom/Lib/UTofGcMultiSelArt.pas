unit UTofGcMultiSelArt;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,HStatus,
      HCtrls,HEnt1,HMsgBox,UTOF, UTOB, HDimension,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
      mul, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DBGrids, db,Fe_Main,
{$ENDIF}
      M3FP,M3VM,AGLInit,Dialogs,HDB;
TYPE
     TOF_GCMULTISELART = Class (TOF)
        private
               FArgument:String;
               CCTiers:THLabel;
        public
              procedure OnArgument (StArgument:String) ; override ;
              procedure OnUpdate ; override ;
              function SaisieMultiple : integer ;
     END ;

implementation

procedure TOF_GCMULTISELART.OnArgument (StArgument:String);
var CCLibelle:THLabel;
    F:TFMul;
begin
inherited;
if not (Ecran is TFMul) then exit;
F:=TFMul(Ecran) ;
FArgument:=StArgument;
CCTiers:= THLabel(TFMul(F).FindComponent('TIERS'));
CCLibelle:= THLabel(TFMul(F).FindComponent('LIBELLE'));
CCTiers.Caption:=ReadTokenSt(FArgument);
CCLibelle.Caption:=ReadTokenSt(FArgument);
Ecran.Caption:=Ecran.Caption+CCTiers.Caption+CCLibelle.Caption;
end;

procedure TOF_GCMULTISELART.OnUpdate;
var  F : TFMul ;
    i:integer;
    CC:THLabel ;
begin
inherited ;
if not (Ecran is TFMul) then exit;
F:=TFMul(Ecran) ;
{$IFDEF EAGLCLIENT}
{AFAIREEAGL}
{$ELSE}
for i:=0 to F.FListe.Columns.Count-1 do
    begin
    if copy(F.FListe.Columns[i].Title.caption,1,14)='Famille niveau' then
        begin
        CC:=THLabel(F.FindComponent('TGA_FAMILLENIV'+copy(F.FListe.Columns[i].Title.caption,16,1))) ;
        F.Fliste.columns[i].visible:=CC.visible ;
        F.Fliste.columns[i].Field.DisplayLabel:=CC.Caption ;
        end;
    end;
{$ENDIF}
end;

function TOF_GCMULTISELART.SaisieMultiple : integer;
var F:TFMul ;
    SelectionArticles : TOB ;
    i,nbEnrgts : integer;
    NouveauChamp,ValeurChamp : string;
begin
{$IFDEF EAGLCLIENT}
{AFAIREEAGL}
{$ELSE}
  Result:=0;
  F:=TFMul(Ecran);
  if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
     begin
       PGIInfo('Aucun article n''a été sélectionné',Ecran.Caption);
       exit;
     end;
  SelectionArticles := TOB.Create('_REFERENCEMENT',Nil,-1);
  //**** tout a été sélectionné avec le bouton tout sélectionner ****
  if (F.FListe.AllSelected) then
    begin
      F.Q.DisableControls;
      nbEnrgts:=F.Q.RecordCount; F.Q.First ;
      for i:=0 to nbEnrgts-1 do
        begin
          //ValeurChamp:=TFmul(Ecran).Q.FindField('GA_CODEARTICLE').AsString;
          ValeurChamp:=TFmul(Ecran).Q.FindField('GA_ARTICLE').AsString;
          NouveauChamp:='ArticleNo'+IntToStr(i+1);
          SelectionArticles.AddChampSup(NouveauChamp,TRUE);
          SelectionArticles.PutValue(NouveauChamp,ValeurChamp);
          F.Q.Next;
        end;
      F.Q.EnableControls;
    end
    else
    begin
      nbEnrgts:=F.FListe.nbSelected ;
      for i:=0 to F.FListe.NbSelected-1 do
        begin
          F.FListe.GotoLeBOOKMARK(i);
          MoveCur(False);
          ValeurChamp:=TFmul(Ecran).Q.FindField('GA_ARTICLE').AsString;
          NouveauChamp:='ArticleNo'+IntToStr(i+1);
          SelectionArticles.AddChampSup(NouveauChamp,TRUE);
          SelectionArticles.PutValue(NouveauChamp,ValeurChamp);
        end;
    end;
  TheTob:=SelectionArticles;
  result:=nbEnrgts;
{$ENDIF}
End;

function AGLSaisieMultiple(parms:array of variant; nb: integer ) : variant ;
var  F : TForm ;
     TOTOF  : TOF;
begin
  F:=TForm(Longint(Parms[0])) ;
  if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
  if (TOTOF is TOF_GCMULTISELART) then result:=TOF_GCMULTISELART(TOTOF).SaisieMultiple() else exit;
end;

Initialization
registerclasses([TOF_GCMULTISELART]);
RegisterAglFunc('SaisieMultiple',TRUE,0,AGLSaisieMultiple);
end.
 
