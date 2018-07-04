unit UTomStock;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOM, HDimension, AglInit,UTOB,
      Dialogs,Menus, M3FP, EntGC,
{$IFDEF EAGLCLIENT}
      eFiche,maineagl,eFichList,uHttp,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Mul,Fiche, FE_main,FichList,
{$ENDIF}
      ParamSoc,
      HDB,ExtCtrls, graphics, LookUP, HTB97,DISPO_SERIE_TOF,
      AglInitGC, UtilArticle, AssistStockAjust ;


Type
     TOM_Dispo = Class (TOM)
      private
       TOBArticle : TOB;
{$IFDEF EAGLCLIENT}
       MaxTobFilles : integer ;
{$ENDIF}
       BVERIFSTOCK : TToolbarButton97;
{$IFDEF BTP}
       MVerifStock : TMenuItem;
{$ENDIF}
       RGUNITE : TRadioGroup;
       Procedure CompareDispolot ;
       function  RechercheEmplacement : string;
       Procedure InitFiche(TOBArt : TOB);
       Function  RechercheArticle (Champ_CodeArt : string) : Boolean;
       Function  VerifArticle (CodeArticle : string) : Boolean;
       procedure VerifMenu ;
       Procedure AppelDISPOLOT (Action : string) ;
       procedure BVERIFSTOCKClick (Sender: TObject);
       procedure RGUNITEClick(Sender: TObject);
      public
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnNewRecord  ; override ;
       procedure OnUpdateRecord  ; override ;
       procedure OnLoadRecord  ; override ;
       procedure OnClose  ; override ;
       procedure OnCancelRecord ; override ;
       procedure OnChangeField (F : TField)  ; override ;
       procedure OnDeleteRecord ; override ;

END ;
       var CoordDim  : Array of integer ;//Tableau contenant les positions des
                //Controls liés aux dimensions de l'article

       Function  TraiterArticle (CodeArticle : string; var TOBArt : TOB) : Boolean;
       Function  ChoisirDimension (St_CodeArt : string; var TOBArt : tob) : Boolean;
       Procedure AffiGrillCodedimension(F:TForm; TOBArt : TOB);
       procedure RedimensionePanel(F : TForm ; StatutArt : string) ;

       Var HautForm, HautPanel, nbenreg : integer ; //HautPanel et HautForm contiennent
            // respectivement les dimensions du Panel lié aux dimensions de l'article
            //et de la fiche
            //nbenreg contient la dimension du tableau TabDispo

           Nouveau : Boolean ; //Passe à True en création pour éventuellement détruire
                              //des lots sur le prochain LoadRecord

           NewArticle, NewDepot, MemoArticle : string ;//NewArticle et NewDepot gardent
                //en mémoire les valeurs correspondantes si on est en création

        var TabDispo : array of record//Ce tableau enregistre l'index des lots
           Art : string ;  //que l'utilisateur peut créer pendant son traitement
           Dep : string ;  //et qu'il faut détruire si l'article disponible en
        end ;              //cours n'est pas validé

Type
     TOM_DispoLot = Class (TOM)
      private
       TOBArticle : TOB;
       mnSerie : TMenuItem;
       BSERIE : TToolbarButton97;
       ActionLot  : string;
       procedure VerifMenu ;
       procedure mnSerieClick (Sender: TObject);
       procedure BSERIEClick (Sender: TObject);
       function MsgConfSupp : integer;

      public
       procedure OnArgument (stArgument : String ) ; override ;
       procedure OnLoadRecord  ; override ;
       procedure OnClose ; override ;
       procedure OnDeleteRecord ; override ;
       end;
const
  NbDecVSup : integer = 2;
	// libellés des messages
	TexteMessage: array[1..10] of string 	= (
          {1}         'Article vide ou inexistant'
          {3}        ,'Dépôt vide ou inexistant'
          {2}        ,'Le stock mini doit être inférieur au stock maxi'
          {4}        ,'La suppression est impossible : les stocks ne sont pas vides'
          {5}        ,'Emplacement vide ou inexistant'
          {6}        ,'La somme des lots n''est pas égale à la quantité physique'#13' de l''article. Voulez vous continuer?'
          {7}        ,'La suppression est impossible : le lot est présent sur une pièce encours.'
          {8}        ,'La suppression est impossible : le lot contient des numéros de série en "Réservé" ou en "Préparation".'
          {9}        ,'Des modifications ont été apportées à la fiche. Celles-ci pourront être perdues.'#13' Confirmez-vous la vérification?'
         {10}        ,'Le nombre de lignes retournées par le serveur est limité à %d.'#13' Le résultat de la vérification peut-être incomplet.'#13' Confirmez-vous la vérification?'
                 );

implementation

{$IFDEF GPAO}
uses
  wStock;
{$ENDIF GPAO}


{==============================================================================================}
{================================== Procédure de TOM_DispoLot =================================}
{==============================================================================================}

procedure TOM_DispoLot.OnArgument (stArgument : String) ;
Var F : TFfiche;
    Param : string;
Begin
Inherited ;
F:=TFfiche(Ecran) ;
if (F.Name <> 'GCDISPOLOT') then exit;
TOBArticle:=Nil;
Param   := stArgument;
ActionLot  := ReadTokenSt(Param);
mnSerie := TMenuItem(GetControl('mnserie'));
mnSerie.Onclick := mnSerieClick;
BSERIE:=TToolbarButton97(GetControl('BSERIE'));
BSERIE.Onclick := BSERIEClick;
end;

procedure TOM_DispoLot.OnLoadRecord  ;
Var QArt : TQuery ;
Begin
Inherited ;
if  TOBArticle=Nil then
    begin
    TOBArticle := TOB.Create ('ARTICLE', nil, -1) ;
    QArt := OpenSQL('SELECT GA_CODEARTICLE,GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,'+
                   'GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5,GA_STATUTART,'+
                   'GA_LOT,GA_NUMEROSERIE FROM ARTICLE '+
                   'WHERE GA_ARTICLE="'+GetField('GQL_ARTICLE')+'"',true);
     TOBArticle.SelectDB ('',QArt);
     Ferme(QArt);
     end ;
end;

procedure TOM_DispoLot.OnDeleteRecord;
var Qry : TQuery;
    TOBSerie,TOBRechRSV,TOBRechPREPA : TOB;
    Reponse : integer;
begin
Inherited ;
If ExisteSQL('select GP_NUMERO from PIECE, LIGNELOT where '+ //Test si le lot est sur une pièce vivante
            'GP_VIVANTE="X" and GP_NUMERO=GLL_NUMERO '+
            'and GLL_ARTICLE="'+GetField('GQL_ARTICLE')+'" '+
            'and GLL_NUMEROLOT="'+GetField('GQL_NUMEROLOT')+'"') then
     begin
     LastError    := 7;
     LastErrorMsg := TexteMessage[LastError];
     exit;
     end;

TOBSerie := TOB.Create('',nil,-1); //Test s'il y a des numéros de séries
Qry := OpenSQL('select * from DISPOSERIE where '+
               'GQS_ARTICLE="'+GetField('GQL_ARTICLE')+'"'+
               'and GQS_NUMEROLOT="'+GetField('GQL_NUMEROLOT')+'"',True);
TOBSerie.LoadDetailDB('DISPOSERIE','','',Qry,False,False);
Ferme(Qry);
if TOBSerie.Detail.count>0 then
   begin
   TOBRechRSV   := TOBSerie.FindFirst(['GQS_ENRESERVECLI'],['X'],True);
   TOBRechPREPA := TOBSerie.FindFirst(['GQS_ENPREPACLI'],['X'],True);
   if (TOBRechRSV<>nil) or (TOBRechPREPA<>nil) then
      begin
      TOBSerie.Free;
      LastError    := 8;
      LastErrorMsg := TexteMessage[LastError];
      exit;
      end
   else
      begin
      Reponse := MsgConfSupp;
      Case Reponse of
           mrYes : begin TOBSerie.DeleteDB(False); TOBSerie.Free; end;
           mrNo  : begin TOBSerie.Free; LastError := 1; end;
           end;
      end;
   end;
end;

function TOM_DispoLot.MsgConfSupp : Integer;
var Msg : string ;
begin
Msg:=TraduireMemoire('ATTENTION;Les numéros de série liés à ce lot seront aussi détruits. Voulez-vous continuer la destruction ?');
Msg:='0;'+Msg+';W;YN;N;N;';
Result := HShowMessage(Msg, Ecran.Caption, '');
end;

procedure TOM_DispoLot.OnClose  ;
Begin
Inherited ;
if TFFicheListe(Ecran).FTypeAction<>taModif then exit;
if not ExisteSQL('SELECT GQ_ARTICLE FROM DISPO WHERE GQ_ARTICLE="'+GetField('GQL_ARTICLE')+'"'+
       ' AND GQ_DEPOT="'+GetField('GQL_DEPOT')+'" AND GQ_CLOTURE="-" AND GQ_PHYSIQUE=(SELECT SUM(GQL_PHYSIQUE) FROM DISPOLOT'+
       ' WHERE GQL_ARTICLE="'+GetField('GQL_ARTICLE')+'" AND GQL_DEPOT="'+GetField('GQL_DEPOT')+'")') then
   if HShowMessage('6;'+Ecran.Caption+';'+TexteMessage[6]+';W;YN;Y;N;;;','','') = mrNo then
      Ecran.ModalResult := 0;
TOBArticle.Free ; TOBArticle:=Nil;
end;

{==============================================================================================}
{================================ Gestion du menu =============================================}
{==============================================================================================}
procedure TOM_DispoLot.VerifMenu ;
begin
NextPrevControl(Ecran) ;
mnSerie.enabled := (TOBArticle.GetValue('GA_LOT') = 'X') or
                   (TOBArticle.GetValue('GA_NUMEROSERIE') = 'X') ;
End ;

procedure TOM_DispoLot.BSERIEClick (Sender: TObject);
begin
VerifMenu;
end;

procedure TOM_DispoLot.mnSerieClick (Sender: TObject);
var Article,Depot,Lot : String;
    Qte               : integer;
begin
  Article := GetField('GQL_ARTICLE');
  Depot   := GetField('GQL_DEPOT');
  Lot     := GetField('GQL_NUMEROLOT');
  Qte     := GetField('GQL_PHYSIQUE');
  if Lot=''  then
     PGIBox('Le numéro de lot est obligatoire.','Erreur')
  else
     GCLanceFiche_DispoSerie ('GC','GCDISPOSERIE','','',ActionLot+';PROV=LOT'+';ARTICLE='+Article+';DEPOT='+Depot+';LOT='+Lot+';QTE='+IntToStr(Qte));
end;

{==============================================================================================}
{================================== Procédure de TOM_Dispo ====================================}
{==============================================================================================}
Procedure TOM_Dispo.OnArgument (stArgument : String) ;
Var F : TFfiche;
    FF : string ;
    ii : integer ;
{$IFDEF EAGLCLIENT}
    ServerCwas : THttpServer;
    TobServer : TOB ;
{$ENDIF}
Begin
Inherited ;
SetControlProperty('GQ_EMPLACEMENT','Plus',GetControlText('GQ_DEPOT')) ;
F:=TFfiche(Ecran) ; F.MonoFiche:=true;
Nouveau:=False ;
HautForm:=F.Height;
HautPanel:=TPanel(F.FindComponent('PDIM')).Height ;
NewArticle:='';
NewDepot:='';
nbenreg:=1;
Setlength(TabDispo,nbenreg);
TOBArticle := TOB.Create ('ARTICLE', nil, -1) ;
FF:='#,##0.';
// BBI : alignement sur le nb de decimale de la compta....
for ii:=1 to GetParamSoc('SO_DECPRIX') do FF:=FF+'0';
//for ii:=1 to V_PGI.OkDecV+NbDecVSup do FF:=FF+'0';
SetControlProperty('GQ_DPA','DisplayFormat',FF);
SetControlProperty('GQ_DPR','DisplayFormat',FF);
SetControlProperty('GQ_PMAP','DisplayFormat',FF);
SetControlProperty('GQ_PMRP','DisplayFormat',FF);
FF:='#,##0.';
for ii:=1 to GetParamSoc('SO_DECQTE') do FF:=FF+'0';
//for ii:=1 to V_PGI.OkDecV+NbDecVSup do FF:=FF+'0';
SetControlProperty('GQ_STOCKMIN','DisplayFormat',FF);
SetControlProperty('GQ_STOCKMAX','DisplayFormat',FF);
SetControlProperty('GQ_PHYSIQUE','DisplayFormat',FF);
SetControlProperty('GQ_RESERVECLI','DisplayFormat',FF);
SetControlProperty('GQ_RESERVEFOU','DisplayFormat',FF);
SetControlProperty('GQ_PREPACLI','DisplayFormat',FF);
SetControlProperty('GQ_LIVRECLIENT','DisplayFormat',FF);
SetControlProperty('GQ_LIVREFOU','DisplayFormat',FF);
SetControlProperty('GQ_TRANSFERT','DisplayFormat',FF);
SetControlProperty('GQ_QTE1','DisplayFormat',FF);
SetControlProperty('GQ_QTE2','DisplayFormat',FF);
SetControlProperty('GQ_QTE3','DisplayFormat',FF);
//SetControlProperty('GQ_STOCKINITIAL','DisplayFormat',FF);
SetControlProperty('GQ_STOCKINV','DisplayFormat',FF);
SetControlProperty('GQ_CUMULSORTIES','DisplayFormat',FF);
SetControlProperty('GQ_CUMULENTREES','DisplayFormat',FF);
SetControlProperty('GQ_ENTREESORTIES','DisplayFormat',FF);
SetControlProperty('GQ_ECARTINV','DisplayFormat',FF);
SetControlProperty('GQ_QTEDISPO','DisplayFormat',FF);
SetControlProperty('GQ_QTENET','DisplayFormat',FF);
{$IFDEF GPAO}
  wSetCaptionStock(F); { Mise à jour des libellés attendus et réservés }
{$ENDIF GPAO}
{$IFDEF EAGLCLIENT}
MaxTobFilles:=0 ;
ServerCwas:=AppServer;
TobServer:=ServerCwas.GetServerParams;
If TobServer <> Nil then
Begin
     if TobServer.FieldExists('EAGLMAXTOBFILLES') then
        MaxTobFilles:=StrToInt(TobServer.GetValue('EAGLMAXTOBFILLES'));
     TobServer.Free;
End else MaxTobFilles:=1;
if (MaxTobFilles>0) and (MaxTobFilles<75000) then
        SetControlVisible('BVERIFSTOCK',False) ;
{$ENDIF}
BVERIFSTOCK:=TToolbarButton97(GetControl('BVERIFSTOCK'));
if BVERIFSTOCK<>nil then BVERIFSTOCK.Onclick := BVERIFSTOCKClick;
{$IFDEF BTP}
if BVERIFSTOCK=nil then
begin
	MVERIFSTOCK:=TMenuItem (GetControl('mnVerifStock'));
	MVERIFSTOCK.Onclick := BVERIFSTOCKClick;
end;
{$ENDIF}
RGUNITE := TRadioGroup(GetControl('RGUNITE'));
if RGUNITE <> nil then RGUNITE.Onclick := RGUNITEClick;
End ;

Procedure Tom_Dispo.OnLoadRecord ;
var stCaption : string ;
    F : TFfiche;
    QArt : TQuery;
    CodeArticle : string ;
Begin
Inherited ;
F:=TFfiche(ecran) ;
if (F.Name <> 'BTDISPO') and (F.Name <> 'FODISPO') then exit;
if F.Name = 'FODISPO' then SetActiveTabSheet ('PSTOCK') else SetActiveTabSheet ('PGENERAL') ;
if  TOBArticle=Nil then exit;

QArt := OpenSQL('SELECT GA_CODEARTICLE,GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,'+
     'GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5,GA_STATUTART,'+
     'GA_LOT,GA_NUMEROSERIE,GA_QUALIFUNITESTO,GA_QUALIFUNITEVTE FROM ARTICLE '+
     'WHERE GA_ARTICLE="'+GetField('GQ_ARTICLE')+'"',true);
TOBArticle.SelectDB ('',QArt);
Ferme(QArt);

RedimensionePanel(F,TOBArticle.GetValue('GA_STATUTART')) ;

SetControlText('LIB_EMPLACEMENT',RechercheEmplacement);
SetControlProperty('GQ_STOCKINITIAL','ReadOnly',True) ;
SetControlProperty('GQ_STOCKINITIAL','Color',clmenu) ;
//
SetControlProperty('GQ_STOCKINV','ReadOnly',True) ;
SetControlProperty('GQ_STOCKINV','Color',clmenu) ;
if getField('GQ_STOCKINV')=0 then
begin
  if not (DS.State in [dsInsert, dsEdit]) then DS.edit; // pour passer DS.state en mode dsEdit
	SetField ('GQ_STOCKINV',getField('GQ_STOCKINITIAL'));
end;

if Nouveau and (nbenreg>1) then CompareDispolot;

if getfield('GQ_ARTICLE')<>'' then
    begin
    Nouveau:=False ;
{$IFDEF EAGLCLIENT}
    if (MaxTobFilles=0) or (MaxTobFilles>=75000) then SetControlVisible('BVERIFSTOCK',True) ;
{$ELSE}
    SetControlVisible('BVERIFSTOCK',True) ;
{$ENDIF}

    CodeArticle := TOBArticle.GetValue('GA_CODEARTICLE') ;
    SetControlText('GQ_CODEARTICLE',CodeArticle) ;

    SetControlProperty('GQ_STOCKINITIAL','ReadOnly',True) ;
    SetControlProperty('GQ_STOCKINV','ReadOnly',True) ;
    //
    SetControlEnabled('GQ_CODEARTICLE',False) ;
    SetControlEnabled('GQ_DEPOT',False) ;

    AffiGrillCodedimension(F,TOBArticle);

    end ;

  stCaption := TraduireMemoire('Article disponible') + ' : ';
  if DS.State <> dsInsert then
    stCaption := stCaption + GetControlText('GQ_CODEARTICLE') + ' ' + GetControlText('NOMARTICLE') + '  ' + RechDom('GCDEPOT', GetField('GQ_DEPOT'), false);
Ecran.caption:=stCaption;

setcontroltext('GQ_QTEDISPO',getfield('GQ_PHYSIQUE')-getfield('GQ_RESERVECLI')-getfield('GQ_PREPACLI')) ;
SetControlText('GQ_QTENET',getfield('GQ_PHYSIQUE')-getfield('GQ_RESERVECLI')-getfield('GQ_PREPACLI')+getfield('GQ_RESERVEFOU')) ;

{$IFDEF GPAO}
SetActiveTabSheet('PSTOCK'); { Active l'onglet de visu des quantités }
{$ENDIF}

end;

Procedure Tom_Dispo.OnNewRecord ;
var i : Integer ;
Begin
Inherited ;
Nouveau:=True ;
SetControlVisible('BVERIFSTOCK',False);
SetControlEnabled('GQ_CODEARTICLE',True) ;
SetControlEnabled('GQ_DEPOT',True) ;
SetControlText('GQ_CODEARTICLE','') ;
NewArticle:='';
NewDepot:='';
MemoArticle:='';
for i:=1 to MaxDimension do
   begin
   SetControlText('GQ_CODEDIM'+inttostr(i),'') ;
   SetControlText('GRILLIBEL'+inttostr(i),'') ;
   end ;
End;

Procedure Tom_Dispo.OnUpdateRecord ;
var SQL : string ;
Begin
Inherited ;
SQL := 'SELECT GA_ARTICLE FROM ARTICLE WHERE GA_ARTICLE ="'+GetField('GQ_ARTICLE')+'"' ;

If (GetField('GQ_ARTICLE')='') Or Not ExisteSql(SQL) then
   begin
   SetFocusControl ('GQ_CODEARTICLE');
   LastError:=1 ; LastErrorMsg:=TexteMessage[LastError] ;
   exit;
   end;

If (GetField('GQ_DEPOT')='') Or (RechDom('GCDEPOT',GetField('GQ_DEPOT'),FALSE)='') then
   begin
   SetFocusControl ('GQ_DEPOT');
   LastError:=2 ; LastErrorMsg:=TexteMessage[LastError] ;
   exit;
   end;

if (GetField('GQ_EMPLACEMENT')<>'') and (RechercheEmplacement='') then
   begin
   SetFocusControl ('GQ_EMPLACEMENT');
   LastError:=5 ; LastErrorMsg:=TexteMessage[LastError] ;
   exit;
   end;

If (GetField('GQ_STOCKMIN')>GetField('GQ_STOCKMAX')) then
   begin
   SetActiveTabsheet('PSTOCK');
   SetFocusControl ('GQ_STOCKMIN');
   LastError:=3 ; LastErrorMsg:=TexteMessage[LastError] ;
   exit;
   end;
ExecuteSQl('UPDATE EMPLACEMENT Set GEM_EMPLACEOCCUPE="X" where GEM_EMPLACEMENT="'+
            GetField('GQ_EMPLACEMENT')+'" AND GEM_MONOARTICLE="X"');
End;

procedure TOM_Dispo.OnCancelRecord  ;
var StSQL : string ;
Begin
Inherited ;
if Not Nouveau then exit;
if (TOBArticle.GetValue('GA_LOT')='X') then
   begin
   StSQL:='DELETE DISPOLOT Where GQL_ARTICLE="'+TOBArticle.GetValue('GA_ARTICLE')+'" AND GQL_DEPOT="'+GetField('GQL_DEPOT')+'"';
   ExecuteSQl(StSQL);
   end;
if (TOBArticle.GetValue('GA_NUMEROSERIE')='X') then
   begin
   StSQL:='DELETE DISPOSERIE Where GQS_ARTICLE="'+TOBArticle.GetValue('GA_ARTICLE')+'" AND GQS_DEPOT="'+GetField('GQL_DEPOT')+'"';
   ExecuteSQl(StSQL);
   end;
end;

Procedure Tom_Dispo.OnClose ;
Begin
Inherited ;
TOBArticle.Free ;
End;

Procedure TOM_DISPO.OnChangeField (F : TField) ;
Var St,RefArt,TypeEmplace : string ;
    QQ : TQuery ;
Begin
Inherited ;
If (F.Fieldname='GQ_DEPOT') or (F.Fieldname='GQ_ARTICLE') then
    begin
    RefArt:=GetControlText('GQ_ARTICLE');
    St:=GetControlText('GQ_DEPOT');
    St:=St+'" AND (GEM_MONOARTICLE<>"X" OR (GEM_MONOARTICLE="X" AND GEM_EMPLACEOCCUPE="-"))';
    if RefArt <> '' then
        begin
        QQ:=OpenSQL('SELECT GA_TYPEEMPLACE FROM ARTICLE WHERE GA_ARTICLE="'+RefArt+'"', True) ;
        if not QQ.Eof then
            begin
            TypeEmplace:=QQ.Fields[0].AsString;
            if TypeEmplace<>'' then St:=St+' AND GEM_TYPEEMPLACE="'+TypeEmplace+'"'
            end;
        Ferme(QQ);
        end;
    St:=St+' AND GEM_EMPLACEMENT <>"';
    //SetControlProperty('GQ_EMPLACEMENT','Plus',GetControlText('GQ_DEPOT')) ;
    SetControlProperty('GQ_EMPLACEMENT','Plus',St) ;
    if Nouveau then NewDepot:=GetControlText('GQ_DEPOT') ;
    end;

  If F.Fieldname='GQ_EMPLACEMENT' then
 SetControlText('LIB_EMPLACEMENT',RechercheEmplacement);
//
  If (F.FieldName='GQ_STOCKINITIAL') and (DS.State in [dsInsert]) then
   SetField('GQ_PHYSIQUE',GetField('GQ_STOCKINITIAL')) ;
//
  If (F.FieldName='GQ_STOCKINV') and (DS.State in [dsInsert]) then
    SetField('GQ_PHYSIQUE',GetField('GQ_STOCKINV')) ;


end ;


//Ne detruit que les articles que si les stocks sont vides
procedure TOM_DISPO.OnDeleteRecord ;
Var  Statut : string ;
     i : integer ;
Begin
Inherited ;
If    (GetField('GQ_PHYSIQUE')<>0)    or (GetField('GQ_TRANSFERT')<>0)
   or (GetField('GQ_RESERVECLI')<>0)  or (GetField('GQ_RESERVEFOU')<>0)
   or (GetField('GQ_LIVRECLIENT')<>0) or (GetField('GQ_LIVREFOU')<>0) then
    begin
    SetActiveTabsheet('PSTOCK');
    SetFocusControl ('GQ_PHYSIQUE');
    LastError:=4 ; LastErrorMsg:=TexteMessage[LastError] ;
    exit;
    end;
SetControlText('GQ_CODEARTICLE','');
For i:=0 to maxdimension do
    begin
    SetControlText('GA_CODEDIM'+InttoStr(i),'');
    SetControlText('GRILLIBEL'+InttoStr(i),'');
    end;
Statut:='';
RedimensionePanel(TFFiche(Ecran),'') ;
Ecran.Caption :='Dispo de ';
End ;

procedure TOM_Dispo.BVERIFSTOCKClick (Sender: TObject);
Var Verif : Boolean;
begin
Verif:=True;
{$IFDEF EAGLCLIENT}
nextPrevcontrol(TFFiche(Ecran)) ;
if TFFiche(Ecran).QFiche.Modified then
{$ELSE}
if (DS.State<>dsBrowse) then
{$ENDIF}
    if PGIAsk(TraduireMemoire(TexteMessage[9]),Ecran.Caption)<>mrYes  then Verif:=False;

{$IFDEF EAGLCLIENT}
if Verif then
  if MaxTobFilles>0 then
    if PGIAsk(format(TraduireMemoire(TexteMessage[10]),[MaxTobFilles]),Ecran.Caption)<>mrYes  then Verif:=False;
{$ENDIF}
if Verif then
   if EntreeStockAjustUnArticle(GetControlText('GQ_ARTICLE'),GetControlText('GQ_DEPOT')) then
      begin
      TFFiche(Ecran).bDefaire.OnClick(Sender);
      RefreshDb;
      end;
end;

// swap l'affichage des qte en unites de stock, vente, ou achat de la fiche article
procedure TOM_Dispo.RGUNITEClick(Sender: TObject);
begin
;
end;

//Détruit les lots d'article si l'utilisateur en a créé et a par la suite annulé
// son traitement
Procedure TOM_Dispo.CompareDispolot ;
Var Trouve : boolean;
    i : integer ;
Begin
For i:=0 to High(TabDispo) do
    begin
    Trouve:=ExisteSQL('SELECT GQ_ARTICLE FROM DISPO WHERE GQ_ARTICLE="'
             +TabDispo[i].Art+'" AND GQ_DEPOT="'+TabDispo[i].Dep+'"') ;
    if not Trouve then
        begin
        Trouve:=ExisteSQL('SELECT GQL_ARTICLE FROM DISPOLOT WHERE GQL_ARTICLE="'
           +TabDispo[i].Art+'" AND GQL_DEPOT="'+TabDispo[i].Dep+'"') ;
        If Trouve then
           ExecuteSQL('DELETE FROM DISPOLOT WHERE GQL_ARTICLE="'
             +TabDispo[i].Art+'" AND GQL_DEPOT="'+TabDispo[i].Dep+'"') ;
        end;
    end ;

For i:=0 to High(TabDispo) do
    begin
    TabDispo[i].Art:='' ;  TabDispo[i].Dep:='' ;
    end ;
NewArticle:='';
NewDepot:='';
nbenreg:=1;
Setlength(TabDispo,nbenreg);
End;

function TOM_Dispo.RechercheEmplacement : string;
var QEmp : TQuery;
begin
QEmp := OpenSQL('SELECT GEM_LIBELLE FROM EMPLACEMENT WHERE GEM_DEPOT="'+GetField('GQ_DEPOT')
                        +'" AND GEM_EMPLACEMENT="'+GetField('GQ_EMPLACEMENT')+'"',false);
if not QEmp.eof then result := QEmp.FindField('GEM_LIBELLE').AsString
else result := '';
Ferme(QEmp);
end;

//Cache les champs liés aux dimensions de l'article
Procedure RedimensionePanel(F : TForm ; StatutArt : string) ;
Begin
If StatutArt<>'DIM' then
    begin
    TPanel(F.FindComponent('PDIM')).visible:=False;
    F.Height:=HautForm-HautPanel;
    end else
        begin
        TPanel(F.FindComponent('PDIM')).visible:=True;
        F.Height:=HautForm
        end;
End;

//Effectue des traitements, conséquences du changement du Code Article
Procedure TOM_DISPO.InitFiche(TOBArt : TOB);
begin
  SetField('GQ_CODEARTICLE', TOBArt.GetValue('GA_CODEARTICLE'));
  SetField('GQ_ARTICLE', TOBArt.GetValue('GA_ARTICLE')) ;

  if Nouveau then NewArticle:= TOBArt.GetValue('GA_ARTICLE') ;

  RedimensionePanel(TFFiche(Ecran),TOBArt.GetValue('GA_STATUTART')) ;

  AffiGrillCodedimension(TFFiche(Ecran),TOBArt);

  if string(TOBArt.GetValue('GA_LOT'))='X' then
    begin
    SetControlProperty('GQ_STOCKINITIAL','readonly',True) ;
    SetControlProperty('GQ_STOCKINITIAL','color',clmenu) ;
    //
    SetControlProperty('GQ_STOCKINV','readonly',True) ;
    SetControlProperty('GQ_STOCKINV','color',clmenu) ;
  end
  else
    begin
    SetControlProperty('GQ_STOCKINITIAL','readonly',False) ;
    SetControlProperty('GQ_STOCKINITIAL','color',clwindow) ;
    //
    SetControlProperty('GQ_STOCKINV','readonly',False) ;
    SetControlProperty('GQ_STOCKINV','color',clwindow) ;
    end ;

end ;

//Recherche a travers le MUL Article
Function TOM_DISPO.RechercheArticle (Champ_CodeArt : string) : Boolean;
var G_CodeArticle, G_Article : THCritMaskEdit;
BEGIN
Result := False;
G_CodeArticle := THCritMaskEdit (GetControl(Champ_CodeArt));
G_Article:= THCritMaskEdit.Create (G_CodeArticle.Parent);
G_Article.Parent := G_CodeArticle.Parent;
G_Article.Text := G_CodeArticle.Text;
G_Article.visible:=False;
G_Article.Top := G_CodeArticle.Top;
G_Article.Left := G_CodeArticle.Left;
DispatchRecherche (G_Article, 1, '',
                   'GA_CODEARTICLE=' + Trim (Copy (G_CodeArticle.Text, 1, 18)), '');
if G_Article.Text <> '' then
    BEGIN
    TOBArticle.SelectDB ('"' + G_Article.Text + '"', Nil) ;
    G_CodeArticle.Text := TOBArticle.GetValue ('GA_CODEARTICLE') ;
    if TraiterArticle (G_CodeArticle.Text, TOBArticle) then
        begin
        result:=true ;
        InitFiche(TOBArticle);
        end ;
    END;
G_Article.Destroy ;
MemoArticle:=G_CodeArticle.Text ;
END;

//Valide sur l'événement OnExit du Code Article
Function TOM_DISPO.VerifArticle (CodeArticle : string) : Boolean;
BEGIN
Result := False;
if CodeArticle = ''  then
    BEGIN
    MemoArticle:=CodeArticle;
    SetControlProperty('GQ_STOCKINITIAL','readonly',True) ;
    SetControlProperty('GQ_STOCKINITIAL','color',clmenu) ;

    SetControlProperty('GQ_STOCKINV','readonly',True) ;
    SetControlProperty('GQ_STOCKINV','color',clmenu) ;

    SetField ('GQ_ARTICLE', '');
    RedimensionePanel(TFFiche(Ecran),'') ;
    END ;

If (CodeArticle=MemoArticle) or (CodeArticle='') then
    begin
    Result:=true ;
    exit ;
    end ;

if TraiterArticle (CodeArticle, TOBArticle) then
    begin
    InitFiche(TOBArticle) ;
    result:=true ;
    MemoArticle:=CodeArticle
    end else if not ExisteSQL('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_ARTICLE="'
      +GetField('GQ_ARTICLE')+'" AND GA_CODEARTICLE="'
      +CodeArticle+'"') then SetControlText('GQ_CODEARTICLE',MemoArticle) ;
END;

procedure TOM_Dispo.VerifMenu ;
var CodArt : string ;
    mnpiece, mnarticle, mndispolot : TMenuItem ;
Begin
NextPrevControl(Ecran) ;
CodArt:=GetControlText('GQ_CODEARTICLE') ;
mnpiece:=TMenuItem(GetControl('mnpiecesencours')) ;
mnarticle:=TMenuItem(GetControl('mnarticlecorrespondant')) ;
mndispolot:=TMenuItem(GetControl('mndispolot')) ;

if (TOBArticle.GetValue('GA_CODEARTICLE')='') or (TOBArticle.GetValue('GA_CODEARTICLE')<>CodArt) then
    begin
    mndispolot.enabled:=False ;
    mnpiece.enabled:=False ;
    mnarticle.enabled:=False ;
    end else
    begin
    mndispolot.enabled := (TOBArticle.GetValue('GA_LOT')='X') or
                          (TOBArticle.GetValue('GA_NUMEROSERIE') = 'X') ;
    mnpiece.enabled:=True ;
    mnarticle.enabled:=True ;
    end ;
End ;

//Appel de la fiche GCDISPOLOT et enregistre dans le tableau TabDispo les index utilisé
//en paramètre de ce passage afin de les détruire si non validation de l'article dispo
Procedure TOM_Dispo.AppelDISPOLOT (Action : string) ;
var QSommeQte : TQuery ;
    i : integer ;
Begin
If Nouveau then
    begin
    NewDepot:=GetField('GQ_DEPOT') ;
    if NewDepot=''  then
       begin
       PGIBox('Vous devez renseigner un dépôt','Article disponible : ') ;
       SetFocusControl('GQ_DEPOT') ;
       exit;
       end;
    if ExisteSQL('SELECT GQ_ARTICLE FROM DISPO WHERE GQ_ARTICLE="'+NewArticle
        +'" AND GQ_DEPOT="'+NewDepot+'"') then
        begin
        PGIBox('Cet article existe déjà','Article disponible : ') ;
        SetFocusControl('GQ_CODEARTICLE') ;
        exit;
        end ;
    if TOBArticle.GetValue('GA_LOT')='X' then
       begin
       AglLanceFiche ('GC','GCDISPOLOT',''+NewArticle+';'+NewDepot+'','', Action) ;
       QSommeQte:=OpenSQL('SELECT SUM(GQL_PHYSIQUE) AS CUMUL FROM DISPOLOT WHERE GQL_ARTICLE="'
                           +NewArticle+'" AND GQL_DEPOT="'+NewDepot+'"', True) ;
       If not QSommeQte.EOF then
          begin
          SetField('GQ_PHYSIQUE', QSommeQte.FindField('CUMUL').AsFloat) ;
//???          SetField('GQ_STOCKINITIAL', QSommeQte.FindField('CUMUL').AsFloat) ;
          SetField('GQ_STOCKINV', QSommeQte.FindField('CUMUL').AsFloat) ;
          end;
       Ferme(QSommeQte) ;
       For i:=0 to High(TabDispo) do
           begin
           if (TabDispo[i].Art='') and (TabDispo[i].Dep='') then
              begin
              TabDispo[i].Art:=NewArticle ;
              TabDispo[i].Dep:=NewDepot ;
              nbenreg:=nbenreg+1 ;
              Setlength(TabDispo,nbenreg);
              break ;
              end ;
           if  (TabDispo[i].Art=NewArticle) and (TabDispo[i].Dep=NewDepot) then break ;
           end ;
       end else
       begin
       GCLanceFiche_DispoSerie ('GC','GCDISPOSERIE','','',Action+';'+'PROV=DISPO'+';ARTICLE='+GetField('GQ_ARTICLE')+
                                ';DEPOT='+GetField('GQ_DEPOT')+';QTE='+IntToStr(GetField('GQ_PHYSIQUE')));
       end;
    end else
    begin
    if TOBArticle.GetValue('GA_LOT')='X' then
       AglLanceFiche ('GC','GCDISPOLOT',''+GetField('GQ_ARTICLE')+';'+GetField('GQ_DEPOT')+'','', Action)
       else
       GCLanceFiche_DispoSerie ('GC','GCDISPOSERIE','','',Action+';'+'PROV=DISPO'+';ARTICLE='+GetField('GQ_ARTICLE')+
                                ';DEPOT='+GetField('GQ_DEPOT')+';QTE='+IntToStr(GetField('GQ_PHYSIQUE')));
    end;
End ;

(*------------------------------RECHERCHE ARTICLES------------------------------
--------------------------------------------------------------------------------*)

Function TOMDISPO_RechercheArticle (Parms : Array of variant; nb: integer) : variant;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F.Name <> 'BTDISPO') and (F.Name <> 'FODISPO') then exit;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_Dispo) then TOM_Dispo(OM).RechercheArticle(string (Parms [1])) else exit;
end;

Function TOMDISPO_VerifArticle (Parms : Array of variant; nb: integer) : variant;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F.Name <> 'BTDISPO') and (F.Name <> 'FODISPO') then exit;
//if csDestroying in F.ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_Dispo) then TOM_Dispo(OM).VerifArticle(string (Parms [1])) else exit;
end;

Function TraiterArticle (CodeArticle : string; var TOBArt : TOB) : Boolean;
var RefUnique : string ;
    RechArt : T_RechArt ;
BEGIN
Result := False;
if CodeArticle = '' then
   Begin
   Result := True;
   exit;
   END;
RechArt := TrouverArticle (CodeArticle, TOBArt) ;
Case RechArt of
     traOk    : BEGIN
                Result := True;
                END;
     traGrille: BEGIN
                RefUnique := TOBArt.GetValue ('GA_ARTICLE');
                If ChoisirDimension (RefUnique, TOBArt) then Result := true;
                END;
     END;
END;

function ChoisirDimension (St_CodeArt : string; var TOBArt : tob) : Boolean;
var QQ   :TQuery;
BEGIN
TheTOB:=TOB.Create('', nil, -1);
AglLanceFiche ('GC','GCSELECTDIM','','', 'GA_ARTICLE='+ St_CodeArt +
  	           ';ACTION=SELECT;CHAMP= ') ;
if TheTOB = Nil then
    BEGIN
    Result := False;
    END else
    BEGIN
    Result := True;
    QQ := OpenSQL('SELECT GA_ARTICLE,GA_CODEARTICLE,GA_LIBELLE,GA_LOT,GA_CODEDIM1,GA_CODEDIM2,GA_CODEDIM3,GA_CODEDIM4,GA_CODEDIM5,'+
     'GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5,GA_STATUTART FROM ARTICLE '+
     'WHERE GA_ARTICLE="'+TheTOB.Detail[0].GetValue ('GA_ARTICLE')+'"',true);
    if Not QQ.EOF then
        BEGIN
        TOBArt.SelectDB ('', QQ);
        END;
    Ferme(QQ) ;
    TheTOB := Nil;
    END;
END;


(*---------------------------------------------------------------------------------*)

//Affiche les dimensions relatives à l'article en cours
Procedure AffiGrillCodedimension(F: Tform; TOBArt : TOB);
var TLIB, TLIB2 : THLabel;
    CHPS, CHPS2 : THEdit;
    i_ind, i_Dim : integer;
    GrilleDim,CodeDim,LibDim : String ;
BEGIN
i_Dim := 1;
for i_ind := 1 to MaxDimension do
    BEGIN
    TLIB := THLabel (F.FindComponent ('GRILLIBEL' + IntToStr (i_Dim)));
    CHPS := THEdit (F.FindComponent ('GQ_CODEDIM' + IntToStr (i_Dim)));
    TLIB2 := THLabel (F.FindComponent ('GRILLIBEL' + IntToStr (i_ind)));
    CHPS2 := THEdit (F.FindComponent ('GQ_CODEDIM' + IntToStr (i_ind)));
    TLIB2.Caption := ''; CHPS2.Text := ''; CHPS2.Visible := False;
    GrilleDim:=TOBArt.GetValue('GA_GRILLEDIM'+IntToStr(i_ind)) ;
    CodeDim:=TOBArt.GetValue('GA_CODEDIM'+IntToStr(i_ind)) ;
    if GrilleDim<> '' then
       BEGIN
       TLIB.Caption:=RechDom('GCGRILLEDIM'+IntToStr(i_ind),GrilleDim,FALSE) ;
       LibDim:=GCGetCodeDim(GrilleDim,CodeDim,i_ind) ;
       if LibDim<>'' then
          BEGIN
          CHPS.Text:=LibDim ; CHPS.Visible:=True ; Inc(i_Dim) ;
          END else
          BEGIN
          TLIB.Caption:='' ; CHPS.Text:='' ; CHPS.Visible:=False ;
          END;
        END ;
    END;
END;

//Enregistre dans le tableau Coordim les positions des Controls liés aux dimensions
//de l'article
Procedure TOMDispo_InitForm (Parms : Array of variant; nb: integer) ;
var F     : TForm ;
    i_ind : integer ;
    TLIB : THLabel;
    CHPS : THEdit;
BEGIN
F := TForm (Longint (Parms[0]));
if (F.Name <> 'BTDISPO') and (F.Name <> 'FODISPO') then exit;

SetLength (CoordDim, MaxDimension) ;
For i_ind := Low (CoordDim) to High (CoordDim) do
    BEGIN
    CoordDim[i_ind]:= THEdit(F.FindComponent ('GQ_CODEDIM' + IntToStr (i_ind + 1))).Left ;
    TLIB := THLabel(F.FindComponent('GRILLIBEL'+ IntTostr(i_ind + 1)));
    CHPS := THEdit(F.FindComponent('GQ_CODEDIM'+IntTostr(i_ind + 1)));
    TLIB.Caption := '';
    CHPS.Text := '';
    CHPS.Visible := Not (CHPS.Text = '');
    END;
MemoArticle:='';
{$IFDEF CCS3}
THLabel (F.FindComponent ('TGQ_EMPLACEMENT')).Visible := False;
THEdit (F.FindComponent ('GQ_EMPLACEMENT')).Visible := False;
{$ENDIF}
END;

Procedure TOMDispo_AppelDISPOLOT (Parms : Array of variant; nb: integer) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F.Name <> 'BTDISPO') and (F.Name <> 'FODISPO') then exit;
//if csDestroying in F.ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_Dispo) then TOM_Dispo(OM).AppelDISPOLOT (String(Parms[1])) else exit;
end;

//Grise les menus de BMenu si les tests ne sont pas vérifiés
Procedure TOMDispo_VerifMenu (Parms : Array of variant; nb: integer) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F.Name <> 'BTDISPO') and (F.Name <> 'FODISPO') then exit;
//if csDestroying in F.ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_Dispo) then TOM_Dispo(OM).VerifMenu else exit;
end;


Initialization
{$IFNDEF STK}
registerclasses([TOM_Dispo]) ;
{$ENDIF}
registerclasses([TOM_DispoLot]) ;
RegisterAglFunc( 'DispoRechercheArticle', True , 1, TOMDISPO_RechercheArticle);
RegisterAglFunc( 'DispoVerifArticle', True , 1, TOMDISPO_VerifArticle);
RegisterAglProc( 'DispoInitForm', True , 0, TOMDispo_InitForm);
RegisterAglProc('DispoAppelDISPOLOT',True,1,  TOMDispo_AppelDISPOLOT);
RegisterAglProc('DispoVerifMenu',True,0,  TOMDispo_VerifMenu);
end.
