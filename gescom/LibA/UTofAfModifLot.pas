unit UTofAfModifLot;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF, HDimension,UTOM,AGLInit,
      Utob,Messages,HStatus,GCMzsUtil, 
{$IFDEF EAGLCLIENT}
      Maineagl,eMul,
{$ELSE}
      Fiche,mul, DBGrids, Fe_Main,HDB,db,dbTables,
{$ENDIF}
      M3VM , M3FP ,UtofAfTraducChampLibre
      ,AfUtilArticle
       ;

Type

     TOF_AfModifLot = Class (TOF_AFTRADUCCHAMPLIBRE)

        procedure OnArgument ( S: String ); override ;
        procedure OnLoad ; override ;
        procedure OnClose ; override ;
        procedure AfModificationParLot ;
     END ;



    
procedure AGLAfModificationParLot(parms:array of variant; nb: integer ) ;
Procedure AFLanceFiche_Modiflot_Article;


implementation


procedure TOF_AfModifLot.OnLoad;
begin
Ecran.Caption:='Saisie en lot d''article';
end;

procedure TOF_AfModifLot.OnArgument( S: String );
var    ComboTypeArticle : THValComboBox;
begin
inherited;
  //mcd 05/03/02
ComboTypeArticle:=THValComboBox(GetControl('GA_TYPEARTICLE'));
ComboTypeArticle.plus:=PlusTypeArticle;
end;
procedure TOF_AfModifLot.OnClose;
begin
V_PGI.ExtendedFieldSelection:='' ;
end;


/////////////// ModificationParLotDesTiers //////////////
procedure TOF_AfModifLot.AfModificationParLot;
Var F : TFMul ;
    Parametrages : String;
    TheModifLot : TO_ModifParLot;
begin
F:=TFMul(Ecran);
if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
  begin
  MessageAlerte('Aucun élément sélectionné');
  V_PGI.ExtendedFieldSelection:='' ;   //mcd 06/10/2003 10912
  exit;
  end;
//if HShowMessage('0;Confirmation;Confirmez vous la modification des clients?;Q;YN;N;N;','','')<>mrYes then exit;
TheModifLot := TO_ModifParLot.Create;
TheModifLot.F := F.FListe; TheModifLot.Q := F.Q;
TheModifLot.Titre := 'Modification par lot des Articles';
TheModifLot.FCode := 'GA_ARTICLE';
TheModifLot.Nature := 'AFF';
TheModifLot.FicheAOuvrir := 'AFARTICLE';

ModifieEnSerie(TheModifLot, Parametrages) ;
V_PGI.ExtendedFieldSelection:='' ;   //mcd 06/10/2003 10912

end;



/////////////// Procedure appellé par le bouton Validation //////////////
procedure AGLAfModificationParLot(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFmul) then TOTOF:=TFMul(F).LaTOF else exit;
if (TOTOF is TOF_AfModifLot) then TOF_AfModifLot(TOTOF).AfModificationParLot else exit;
end;


Procedure AFLanceFiche_Modiflot_Article;
begin
AGLLanceFiche('AFF','AFARTICLE_MODIFLO','','','');
end;




Initialization
registerclasses([TOF_AfModifLot]);
RegisterAglProc('AfModificationParLot',TRUE,2,AGLAfModificationParLot);
end.




