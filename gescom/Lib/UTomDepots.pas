{***********UNITE*************************************************
Auteur  ...... : Jean-Christian FOURCADE
Créé le ...... : 22/05/2000
Modifié le ... : 23/05/2000
Description .. : TOM pour DEPOTS :
Suite ........ :
Mots clefs ... : DEPOT
*****************************************************************}
unit UTomDepots;

interface
uses
     HCtrls,HEnt1,HMsgBox,UTOM,
{$IFDEF EAGLCLIENT}
     eFiche,emul,MaineAGL,
{$ELSE}
     Fiche,FichList,Fe_Main,db,dbTables,Mul,DBCtrls,MAJTable,HDB,DBGrids,
{$ENDIF}
     HDimension,UTob,AglInit,
     StdCtrls,Controls,Classes,forms,sysutils,
     ComCtrls,UtilArticle,UdimArticle,menus,
     AglInitGC,Entgc, graphics, extctrls, UtilPGI,Hqry,FACTUTIL,LookUp,
     Grids,windows,M3FP,choix,TarifUtil,UtilGC;

Type
     TOM_Depots = Class (TOM)
       private
           procedure AffichePhoto ;
       public
           procedure OnArgument (Arguments : String )  ; override ;
           procedure OnNewRecord  ; override ;
           procedure OnLoadRecord  ; override ;
           procedure OnChangeField(F: TField)  ;  override ;
           procedure OnUpdateRecord  ; override ;
           procedure OnDeleteRecord  ; override ;
     END ;

const TexteMessage: array[1..3] of string 	= (
        {1}    'Ce code existe déjà ! Vous devez le mettre à jour',
        {2}    'Ce dépôt est utilisé comme dépôt par défaut',
        {3}    'Ce dépôt est déjà utilisé dans les pièces ou possède du stock'
      );

implementation

uses ParamSoc;
procedure TOM_Depots.OnArgument (Arguments : String );
begin
inherited ;
GCMAJChampLibre (TForm (Ecran), False, 'COMBO', 'GDE_LIBREDEP', 10, '');
GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GDE_VALLIBRE', 3, '');
GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GDE_DATELIBRE', 3, '');
GCMAJChampLibre (TForm (Ecran), False, 'BOOL', 'GDE_BOOLLIBRE', 3, '');
GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'GDE_CHARLIBRE', 3, '');
end;

// Affiche la photo du dépôt
procedure TOM_Depots.AffichePhoto ;
var QQ :TQuery ;
    SQL : string ;
    CC : TImage ;
begin
SetControlVisible('P_PHOTO',false);
CC:=TImage(GetControl('LAPHOTO')) ;
if CC=nil then exit ;
SQL:='SELECT * from LIENSOLE where LO_TABLEBLOB="GDE" AND LO_IDENTIFIANT="'+getField('GDE_DEPOT')+'"' ;
SQL:=SQL+' AND LO_QUALIFIANTBLOB="PHO" AND LO_EMPLOIBLOB="'+VH_GC.GCPHOTOFICHE+'"' ;
QQ:=OpenSQL(SQL,true) ;
IF not QQ.EOF then
    begin
    LoadBitMapFromChamp(QQ, 'LO_OBJET',CC) ;
    if CC.Picture<>Nil then SetControlVisible('P_PHOTO',true);
    end ;
ferme(QQ) ;
end;

///////////////////////////////////////////////////////////////////////
procedure TOM_Depots.OnLoadRecord  ;
begin
// affiche Photo
if VH_GC.GCPHOTOFICHE<> '' then AffichePhoto else SetControlVisible('P_PHOTO',false);
if (ctxMode in V_PGI.PGIContexte) then
  begin
  SetControlProperty('GDE_SURSITE', 'Visible', True) ;
  end;
SetControlProperty('PDEPOTSASSOCIES', 'TabVisible', GetParamsoc('SO_GCTRV'));  
end;

///////////////////////////////////////////////////////////////////////
procedure TOM_Depots.OnNewRecord  ;
begin
inherited;
SetField('GDE_SURSITE', 'X');
end;

//////////////////////////////////////////////////////////////////////
procedure TOM_Depots.OnChangeField(F: TField)  ;
var QQ :TQuery ;
    SQL : string ;
begin
inherited ;
// En création, on vérifie que ce code n'existe pas
if (F.FieldName='GDE_DEPOT') and (DS.State in [dsInsert]) then
begin
  SQL:='SELECT GDE_DEPOT from DEPOTS where GDE_DEPOT="'+getField('GDE_DEPOT')+'"' ;
  QQ:=OpenSQL(SQL,true) ;
  if not QQ.EOF then
  begin
    LastError:=1 ; LastErrorMsg:=TexteMessage[LastError] ;
    SetFocusControl ('GDE_DEPOT');
  end ;
  ferme(QQ) ;
end;

// Renseigner le libellé abrégé si vide
if (F.FieldName='GDE_LIBELLE') and (GetField('GDE_LIBELLE')<>'')
                               and (GetField('GDE_ABREGE')='') then SetField('GDE_LIBELLE', GetField('GDE_ABREGE'));
end;

///////////////////////////////////////////////////
procedure TOM_Depots.OnUpdateRecord ;
begin
  inherited ;
  SetField('GDE_DEPOT',Trim(GetField('GDE_DEPOT'))) ;
  if not (ctxMode in V_PGI.PGIContexte) then
  begin
    if      (Trim(GetField('GDE_DEPOTTRANSIT')) = '') and (Trim(GetField('GDE_DEPOTLITIGE')) = '') then
    begin
      SetField('GDE_SURSITE','-');
      SetField('GDE_SURSITEDISTANT','-');
    end
    else if (Trim(GetField('GDE_DEPOTTRANSIT')) <> '') and (Trim(GetField('GDE_DEPOTLITIGE')) <> '') then
    begin
      SetField('GDE_SURSITE','X');
      SetField('GDE_SURSITEDISTANT','X');
    end;
  end;
end;

procedure TOM_Depots.OnDeleteRecord ;
Var Depot,ListeDepotLie : String ;
    Exi : Boolean ;
    QQ : TQuery;
    TOBDepotLie : TOB;
    i_ind,PositionDepot : integer;
begin
Depot:=GetField('GDE_DEPOT') ; if Depot='' then Exit ;
// Vérification que ce n'est pas le dépôt par défaut
if Depot=VH_GC.GCDepotDefaut then BEGIN LastError:=2 ; LastErrorMsg:=TexteMessage[LastError] ; exit; END
else
   BEGIN
   // Vérification que ce dépôt n'a pas de stock
   Exi:=ExisteSQL('SELECT GQ_DEPOT FROM DISPO WHERE GQ_DEPOT="'+Depot+'"') ;
   if Exi then BEGIN LastError:=3 ; LastErrorMsg:=TexteMessage[LastError] ; exit; END
   else
      // Vérification que ce dépôt n'a aucun mouvement
      BEGIN
      Exi:=ExisteSQL('SELECT GP_DEPOT FROM PIECE WHERE GP_DEPOT="'+Depot+'"') ;
      if Exi then BEGIN LastError:=3 ; LastErrorMsg:=TexteMessage[LastError] ; exit; END
      else
         BEGIN
         Exi:=ExisteSQL('SELECT GL_DEPOT FROM LIGNE WHERE GL_DEPOT="'+Depot+'"') ;
         if Exi then BEGIN LastError:=3 ; LastErrorMsg:=TexteMessage[LastError] ; exit; END ;
         END ;
      END;
   END ;
inherited ;
// Si on peut supprimer le dépôt alors
// Suppression des emplacements liés à ce dépôt
Exi:=ExisteSQL('SELECT GEM_EMPLACEMENT FROM EMPLACEMENT WHERE GEM_DEPOT="'+Depot+'"');
if Exi then
begin
if PGIAsk(TraduireMemoire('Il y a des emplacements pour ce dépôt, Etes-vous sûr de vouloir supprimer le dépôt et ses emplacements ?'),'')=mrNo
then
   begin
   AfficheError:=False;
   LastError:=4;
   LastErrorMsg:='';
   exit;
   end;
end;


ExecuteSQL('DELETE FROM EMPLACEMENT WHERE GEM_DEPOT="'+Depot+'"') ;

TobDepotLie:=TOB.CREATE('Dépôts liés',NIL,-1) ;
// Récupération de toutes les lignes d'établissement qui sont lié au dépôt que l'on souhaite supprimer
QQ:=OpenSQL('SELECT * FROM ETABLISS WHERE ET_DEPOTLIE<>""',True);
if QQ<>nil then TobDepotLie.LoadDetailDB('ETABLISS','','',QQ,False)
else begin Ferme(QQ); exit; end;
Ferme(QQ);
for i_ind:=0 to TobDepotLie.detail.count-1 do
    begin
    ListeDepotLie:=TobDepotLie.detail[i_ind].GetValue('ET_DEPOTLIE');
    PositionDepot:=pos (Depot, listeDepotLie);
    // Si le dépôt existe dans le champ ET_DEPOLIE, on l'enlève de cette liste
    if PositionDepot > 0 then
       begin
       if PositionDepot=1 then ListeDepotLie:=copy(ListeDepotLie,5,length(ListeDepotLie))
       else ListeDepotLie:=concat(copy(ListeDepotLie,1,PositionDepot-2),copy(ListeDepotLie,PositionDepot+3,length(ListeDepotLie))) ;
       TobDepotLie.detail[i_ind].PutValue('ET_DEPOTLIE',ListeDepotLie);
       end;
    end;
TobDepotLie.UpdateDB();
TobDepotLie.free;

end;

procedure AGLControleAffPhoto_GDE(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     OM  : TOM;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_DEPOTS) then TOM_DEPOTS(OM).AffichePhoto else exit;
end;

Initialization
registerclasses([TOM_DEPOTS]);
RegisterAglProc('ControleAffPhoto_GDE',TRUE,0,AGLControleAffPhoto_GDE);
end.

