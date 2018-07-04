{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 29/10/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : TRAITEMENT_SUR_ANNUAIRE ()
               : avec lien avec gescom (traduction des tables libres ...
Mots clefs ... : TOF;TRAITEMENT_SUR_ANNUAIRE
*****************************************************************}


Unit UTOFTRAITEMENT_SUR_ANNUAIRE2 ;

Interface

Uses
{$IFDEF EAGLCLIENT}
     MaineAGL, eMul,
{$ELSE}
     FE_Main, Mul,dbTables, db,HDB,
{$ENDIF}
     Utob, DicoAf,Traducaffaire, HStatus, forms, UTOF,M3fp, Classes,
     UtilMulTrt,Hctrls, Hent1,HMsgBox,HTB97,HQry,ComCtrls,
     UtofAfTraducChampLibre,SysUtils,DPUtil,Vierge, StdCtrls
     ,Utoftraitement_Sur_Annuaire1,Controls
     ;

Type
  // fiche aftiers_ann liste des clients n'ayant pas de lien annauire
  TOF_AFTiers_Ann = Class (TOF_AFTraducChampLibre)
     procedure Tiers_Ann;
     Procedure OnArgument(stArgument : String ); override ;
  end ;
 // fiche afSup_tiersann permet la suppression d'un lien tiers/ann fait à tort
  TOF_AFSup_Tiers_Ann = Class (TOF_AFTraducChampLibre)
     procedure Sup_Tiers_Ann;
     Procedure OnArgument(stArgument : String ); override ;
  end ;

Procedure AFLanceFiche_CltSanslienAnn;
Procedure AFLanceFiche_SupLienTiersAnn;

Implementation
uses AnnOutils, AnnDedoub;

const
  TexteMsg: array [1..2] of string = (
    {1}   'Aucun élément sélectionné'
    {2} , 'Confirmez-vous la suppression du lien '
    ) ;


//******************************************************************************
//***************** Gestion des clients sans lien annuaire *********************
//******************************************************************************
procedure TOF_AFTiers_ann.OnArgument(stArgument : String );
begin
  inherited;
end;


procedure TOF_AFTiers_ann.Tiers_Ann;
var wi : integer;
    stguidper,client,nature,retour,tmp : string;
    TobClt,TobDet:TOB ;
    F : TFMul ;
begin
  F:=TFMul(Ecran);
  TobClt := Tob.Create('les clients',NIL,-1);
  TraiteEnregMulListe (F, 'T_TIERS','TIERS', Tobclt, True);
  InitMove(Tobclt.Detail.count,' ');
  wi := 0;
  while wi < TobClt.Detail.count  do
    Begin
    MoveCur(False);
    TobDET := TobClt.Detail[wi];
    // traitement d'un client
    retour:=AFLanceFiche_TiersSansAnn (TobDEt.Getvalue('T_TIERS'));
    tmp:=ReadTokenSt(retour);
    if tmp='STOP' then
       begin
       Wi :=TobClt.Detail.count;     // on force l'arrêt
       end
    else if (tmp='DPTIERS') or (tmp='TIERSDP')  then
       begin
       StGuidper:=Readtokenst(retour);
       // on traite l'enrgt
       Client:=TobDEt.Getvalue('T_TIERS');
       Nature :=TobDEt.Getvalue('T_NATUREAUXI');
       if tmp = 'TIERSDP' then
         DpSynchro(FALSE, StGuidper, Client, Nature, Nil)
       else
         begin
         DpSynchro(TRUE, StGuidper, Client, Nature, Nil);
         ExecuteSql('Update annuaire set ann_tiers="'+TobDEt.Getvalue('T_TIERS')+
                    '" where ann_Guidper="'+StGuidper+'"');
         end;
       //mcd 12/2005 on change clé contact si existe sur annuaire pour la passer sur le client
       //en passant la clé numerocontact +100, pour éviter pb si contact existe dans tiers
       If ExisteSql ('Select c_typecontact,c_auxiliaire from contact where c_typecontact="ANN" and c_auxiliaire="'+ StGuidper+'"' )then
        begin
        ExecuteSql ('Update Contact set c_typecontact="T",c_auxiliaire="'
            + TobDEt.GetString('T_AUXILIAIRE')+'",c_tiers="'

            + TobDEt.Getvalue('T_TIERS')+'",c_natureauxi="'
            + TobDEt.Getvalue('T_NATUREAUXI') + '"  ,c_numerocontact=c_numerocontact+100 from contact where c_typecontact="ANN" and c_auxiliaire="'+ StGuidper+'"' );
        end;
       //fin mcd 12/2005
       //mcd 12760 si conatct existe sur tiers, il faut renseigne GUID
       If ExisteSql ('Select c_auxiliaire from contact where c_typecontact="T" and c_auxiliaire="'+ TobDEt.GetString('T_AUXILIAIRE')+'"' )then
        begin
        ExecuteSql ('Update Contact set c_guidper="'
            + StGuidper+'" '
            +'where c_typecontact="T" and c_auxiliaire="'+ TobDEt.GetString('T_AUXILIAIRE')+'"' );
        end;
       end;
    // sinon l'enrgt en cours n'est pas traité
    inc(wi);
    End;  // fin while, boucle sur la sélection

  Tobclt.free;
  FiniMove;
end;

procedure TOF_AFSup_Tiers_ann.OnArgument(stArgument : String );
begin
  inherited;
end;


procedure TOF_AFSup_Tiers_ann.Sup_Tiers_Ann;
var F : TFMul ;
    wi : integer;
    TobClt, TobDet : TOB;
begin
  F:=TFMul(Ecran);
  if(F.FListe.NbSelected=0)and(not F.FListe.AllSelected) then
    begin
    MessageAlerte(TexteMsg[1]);
    V_PGI.ExtendedFieldSelection:='' ;
    exit;
    end;
  if PGIAsk(TexteMsg[2],Ecran.Caption) <> mrYes  then  Exit;
  TobClt := Tob.Create('les clients',NIL,-1);
  TraiteEnregMulListe (F, 'T_TIERS','AFTIERSANNUAIRE', Tobclt, True);
  InitMove(Tobclt.Detail.count,' ');
  wi := 0;
  while wi < TobClt.Detail.count  do
    Begin
    MoveCur(False);
    TobDET := TobClt.Detail[wi];
    // traitement d'un client
    //mcd 17/05/05 champ supprimé ExecuteSql ('UPDATE ANNUAIRE SET ANN_TIERS="", ANN_AUXILIAIRE="" WHERE ANN_TIERS="'
    ExecuteSql ('UPDATE ANNUAIRE SET ANN_TIERS="" WHERE ANN_TIERS="'
     + TobDet.Getvalue('T_TIERS')+ '"');
    //mcd 12/2005 si des contact existe, on les laisse associé au tiers
    //mais on efface la zone GUidper de la table contact
     If ExisteSql ('Select c_typecontact,c_auxiliaire from contact where c_typecontact="T" and c_tiers="'+ TobDet.Getvalue('T_TIERS')+'"' )then
          begin
          ExecuteSql ('Update Contact set c_Guidper="" from contact where c_typecontact="T" and c_tiers="'+ TobDet.Getvalue('T_TIERS')+'"' );
          end;
    //fin mcd 12/2005
    inc(wi);
    End;  // fin while, boucle sur la sélection
  Tobclt.free;
  FiniMove;
end;


procedure AGLTiers_Ann (parms: array of variant; nb: integer );
var F : TForm;
    LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0]));
  if (F is TFMul) then Latof:=TFMul(F).Latof
   else exit;
  if (Latof is TOF_AFTiers_ann) then TOF_AFTiers_ann(LaTof).Tiers_Ann else exit;
end;

procedure AGLSup_Tiers_Ann (parms: array of variant; nb: integer );
var F : TForm;
    LaTof : TOF;
begin
  F:=TForm(Longint(Parms[0]));
  if (F is TFMul) then Latof:=TFMul(F).Latof
    else Exit;;
  if (Latof is TOF_AFSup_Tiers_ann) then TOF_AFSup_Tiers_ann(LaTof).Sup_Tiers_Ann else exit;
end;

Procedure AFLanceFiche_CltSansLienAnn;
begin
AGLLanceFiche('AFF','AFTIERS_ANN','','','');
end;

Procedure AFLanceFiche_SupLienTiersAnn;
begin
AGLLanceFiche('AFF','AFSUP_TIERSANN','','','');
end;


Initialization
RegisterClasses ( [Tof_AfTiers_ann,Tof_AfSUP_Tiers_ann] ) ;
RegisterAglProc( 'Tiers_Ann',True,0,AGLTiers_Ann);
RegisterAglProc( 'Sup_Tiers_Ann',True,0,AGLSup_Tiers_Ann);
end.
