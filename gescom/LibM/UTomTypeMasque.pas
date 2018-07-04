unit UTomTypeMasque;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOM, HDimension, Dialogs ,
{$IFDEF EAGLCLIENT}
      eFichList,
{$ELSE}
      HDB,dbTables,db,FichList,
{$ENDIF}
      M3FP,EntGC,uTotGc,UTob,UtilPGI;
                                  
Type
     TOM_TypeMasque = Class (TOM)
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnLoadRecord ; override ;
       procedure OnUpdateRecord ; override ;
       procedure OnDeleteRecord ; override ;
       procedure OnClose ; override ;
     private
       bReloadTablette : boolean ;
     END ;

{     procedure RecupTabletteChoixCodeGCTYPEMASQUE ; }

const TexteMessage: array[1..2] of string 	= (
        {1}    'Impossible de supprimer cet élément, sa présence est obligatoire !'
        {2}   ,'Des masques utilisent ce type de masques. La suppression de ce type de masques'+#13+#10+
                     'supprimera également les masques de ce type.'+#13+#10+
                     'Confirmez-vous la suppression de ce type de masques ?'
);

implementation

{ TOM_TypeMasque }

procedure TOM_TypeMasque.OnArgument(stArgument:String) ;
begin
Inherited OnArgument (StArgument) ;
bReloadTablette:=False ;
if not ExisteSQL ('select GMQ_TYPEMASQUE from TYPEMASQUE where GMQ_TYPEMASQUE="'+VH_GC.BOTypeMasque_Defaut+'"')
    then ExecuteSQL('insert into TYPEMASQUE(GMQ_TYPEMASQUE,GMQ_LIBELLE,GMQ_ABREGE,GMQ_MULTIETAB) values ("...","Défaut","Défaut","-")') ;

// Appel pour test fonction de reprise
//RecupTabletteChoixCodeGCTYPEMASQUE ;

end ;

procedure TOM_TypeMasque.OnLoadRecord ;
begin
SetControlEnabled('GMQ_MULTIETAB',(DS.State=DsInsert)) ;
end ;

procedure TOM_TypeMasque.OnUpdateRecord;
var stTypeMasque : string ;
    bMultiEtab : boolean ;
begin
if DS.State=DsInsert then SetControlEnabled('GMQ_MULTIETAB',False) ;
stTypeMasque:=GetField('GMQ_TYPEMASQUE') ;
bMultiEtab:=(GetField('GMQ_MULTIETAB')='X') ;
CreerMasqueAuto(stTypeMasque,bMultiEtab,'',nil) ;
bReloadTablette:=True ;
end ;

procedure TOM_TypeMasque.OnDeleteRecord  ;
begin
if GetField('GMQ_TYPEMASQUE')=VH_GC.BOTypeMasque_Defaut then
    BEGIN LastError:=1; LastErrorMsg:=TexteMessage[LastError]; exit ; END ;
if ExisteSQL ('select GDM_MASQUE from DIMMASQUE where GDM_TYPEMASQUE="'+GetField('GMQ_TYPEMASQUE')+'"') then
    BEGIN
    if PGIAsk(TexteMessage[2], '') = mrYes
    then ExecuteSQL('delete from DIMMASQUE where GDM_TYPEMASQUE="'+GetField('GMQ_TYPEMASQUE')+'"')
    else BEGIN LastError:=2 ; exit ; END ;
    END ;
bReloadTablette:=True ;
end ;

procedure TOM_TypeMasque.OnClose ;
begin
if bReloadTablette then
    BEGIN
    AvertirTable('GCTYPEMASQUE') ;
    END ;
end ;

{ Pour test avant mise en place dans MajHalley ...
procedure RecupTabletteChoixCodeGCTYPEMASQUE ;
var iTob : integer ;
    QQ : TQuery ;
    TobChoixCod,TobTypeMasque,TF : Tob ;
begin
// Chargement en Tob des types de masques
QQ:=OpenSQL('select CC_CODE,CC_LIBELLE,CC_ABREGE from CHOIXCOD where CC_TYPE="GTQ"',True) ;
TobChoixCod:=TOB.Create('_CHOIXCOD',nil,-1) ;
TobChoixCod.LoadDetailDB('CHOIXCOD','','',QQ,False) ;
Ferme(QQ) ;
TobTypeMasque:=TOB.Create('_TYPEMASQUE',nil,-1) ;
for iTob:=0 to TobChoixCod.Detail.Count-1 do
    BEGIN
    TF:=TOB.Create('TYPEMASQUE',TobTypeMasque,-1) ;
    TF.PutValue('GMQ_TYPEMASQUE',TobChoixCod.Detail[iTob].GetValue('CC_CODE')) ;
    TF.PutValue('GMQ_LIBELLE',TobChoixCod.Detail[iTob].GetValue('CC_LIBELLE')) ;
    TF.PutValue('GMQ_ABREGE',TobChoixCod.Detail[iTob].GetValue('CC_ABREGE')) ;
    TF.PutValue('GMQ_MULTIETAB',CheckToString(ExisteSql('select GDM_MASQUE from DIMMASQUE where GDM_TYPEMASQUE="'+TobChoixCod.Detail[iTob].GetValue('CC_CODE')+'" and GDM_POSITION6<>""'))) ;
    END ;
if not TobTypeMasque.InsertDB(nil) then PgiBox('Error InsertDB','RecupTabletteChoixCodeGCTYPEMASQUE') ;
TobChoixCod.Free ;
TobTypeMasque.Free ;
end ;
}

Initialization
registerclasses([TOM_TypeMasque]) ;
end.

