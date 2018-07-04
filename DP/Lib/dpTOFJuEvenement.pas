unit dpTOFJUEvenement;

interface
uses
{$IFDEF EAGLCLIENT}
   MaineAGL, eMul, UTob,
{$ELSE}
   FE_Main, {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} Mul,
{$ENDIF}
   Classes, HCtrls, sysutils, HEnt1,
   UTOF, HTB97,
   UTOFEvenement, DpTofAnnuSel,
   AGLInit, HDB, Forms, Menus ;

Type
   TOF_JUEvenement = Class (TOF_Evenement)
      procedure OnArgument(sArguments_p : String ) ; override ;

      procedure DOS_LIBELLE_OnElipsisClick(Sender: TObject);
      procedure DOS_LIBELLE_OnExit(Sender: TObject);
      procedure ANN_NOMPER_OnElipsisClick(Sender: TObject);
      procedure BINSERT_OnClick(Sender: TObject);
   private
      NoDossier : String;    // si NoDossier passé en param
      GuidPerDos : String;   // guid personne du Dossier si NoDossier en param
  END;


//////////// IMPLEMENTATION //////////////
implementation

uses DpJurOutils, DpJurOutilsEve;


{*****************************************************************
Auteur ....... : BM
Date ......... : //
Procédure .... : OnArgument
Description .. :
Paramètres ... :
*****************************************************************}

procedure TOF_JUEvenement.OnArgument(sArguments_p : String ) ;
var
   QRYConf_l, Q : TQuery;
   XX_WhereConf_l : string;
   bChargeTaches_l, bChargeEtats_l, bChargeOutlook_l: Boolean;
   i : Integer;
begin
  inherited;

  //BM 5/12/02 : Gestion de la confidentialité pour les évènement juridiques
  QRYConf_l := OpenSQL('select JCN_GROUPE from JUCONFIDENTIEL where JCN_USER="'+V_PGI.User+'"',true, -1, '', True);
{  XX_WhereConf_l := '( JUR_GRPCONF is null or JUR_GRPCONF = "" ) OR ' +
         '( JEV_CODEDOS is null or JEV_CODEDOS = "" ) OR ' +
         '( ANN_GRPCONF is null or ANN_GRPCONF = "" )';    }
  XX_WhereConf_l := '';
  if (QRYConf_l<>nil) and (Not QRYConf_l.EOF) then
    begin
{    XX_WhereConf_l := '(( JUR_GRPCONF = "" OR JUR_GRPCONF is null ) AND ' +
           ' ( ANN_GRPCONF = "" OR ANN_GRPCONF is null )) ';   }
    XX_WhereConf_l := '( ANN_GRPCONF = "" OR ANN_GRPCONF is null )';
    while not QRYConf_l.EOF do
      begin
      if XX_WhereConf_l<>'' then
      XX_WhereConf_l := XX_WhereConf_l + ' OR ';
      XX_WhereConf_l := XX_WhereConf_l +
{                ' ( JUR_GRPCONF = "' + QRYConf_l.Fields[0].AsString + '" ' +
                ' OR ANN_GRPCONF = "'+ QRYConf_l.Fields[0].AsString + '") ';   }
                ' ( ANN_GRPCONF = "'+ QRYConf_l.Fields[0].AsString + '") ';
      QRYConf_l.Next;
      end;
    end;
  Ferme(QRYConf_l);

   SetControlText('XX_WHERE', XX_WhereConf_l);
   // Filtrage des combos
   FiltreCombo('JEV_FAMEVT',  'CC_CODE IN ("DOC", "REU", "TAC")');
   FiltreCombo('JEV_CODEEVT', 'JTE_FAMEVT IN ("DOC", "REU", "TAC")');
   // Filtrage des enregistrements
   FiltreEnreg('JEV_FAMEVT IN ("DOC", "REU", "TAC") AND (JEV_DOMAINEACT = "")');
   // Filtrage du bouton-combo NEW
   if Ecran.Name = 'EVENEMENT_MUL' then
      FiltreMenuBouton('DOC;REU;TAC');

  NoDossier := '';
  GuidPerDos := '';
  // Pour tout évènement :
  // si Dossier en param (=appel depuis le module "Dossier client"),
  // on ne prend en compte que les évènements du DOSSIER en cours
  // et on peut voir tous les types d'évènements
  if Copy(sArguments_p, 1, 14)='DOS_NODOSSIER=' then
    begin
    NoDossier := ReadTokenSt(sArguments_p);
    i := pos('=', NoDossier);
    if i>0 then NoDossier := Copy(NoDossier, i+1, Length(NoDossier)-i);
    SetControlText('JEV_NODOSSIER', NoDossier);
    SetControlVisible('DOS_LIBELLE', False);
    SetControlVisible('TDOS_LIBELLE', False);
    Q := OpenSQL('SELECT DOS_GUIDPER FROM DOSSIER WHERE DOS_NODOSSIER="'+NoDossier+'"', True);
    if Not Q.EOF then GuidPerDos := Q.FindField('DOS_GUIDPER').AsString;
    Ferme(Q);
    THValComboBox(GetControl('JEV_CODEEVT')).Vide := True;
    end
  else
    begin
    // sinon on affiche les évènements de l'UTILISATEUR en cours
    SetControlText('JEV_USER1', V_PGI.User);
  {  if XX_WhereConf_l<>'' then XX_WhereConf_l := ' AND ('+XX_WhereConf_l+')';
    XX_WhereConf_l := 'JEV_USER1="'+V_PGI.User+'"' + XX_WhereConf_l;     }
    SetControlVisible('JEV_USER1', False);
    SetControlVisible('TJEV_USER1', False);
    end;

  // #### Zones dont on ne sait quoi faire pour l'instant
  SetControlVisible('JEV_USER2', False);
  SetControlVisible('TJEV_USER2', False);

  // DP :
  SetControlVisible('TNOMDOSSIER', False);
  SetControlVisible('JUR_NOMDOS', False);
  SetControlVisible('TNOMOPERATION', False);
  SetControlVisible('JEV_CODEOP', False);

// #### pas de DOS_LIBELLE dans la fiche JUR JUEVENEMENT_MUL
//  THEdit(GetControl('DOS_LIBELLE')).OnElipsisClick := DOS_LIBELLE_OnElipsisClick;
//  THEdit(GetControl('DOS_LIBELLE')).OnExit := DOS_LIBELLE_OnExit;
  TToolBarButton97(GetControl('BINSERT')).OnClick := BINSERT_OnClick;

  // traitement particulier si dossier passé en param
  THEdit(GetControl('ANN_NOMPER')).OnElipsisClick := ANN_NOMPER_OnElipsisClick;

  // Menu déroulant sur liste
  bChargeTaches_l := ( Ecran.Name = 'EVENEMENT_MUL' ) ;

  bChargeEtats_l := ( Ecran.Name = 'YYEVEN_TAC_MUL' ) or
                    ( Ecran.Name = 'YYEVEN_REU_MUL' ) or
                    ( Ecran.Name = 'EVENEMENT_MUL' ) ;

  bChargeOutlook_l := False; // ( Ecran.Name = 'YYEVEN_MSG_MUL' );

  InitClickDroit( bChargeOutlook_l, bChargeTaches_l, bChargeEtats_l);

end;


procedure TOF_JUEvenement.DOS_LIBELLE_OnElipsisClick(Sender: TObject);
// idem principe zone ANN_NOMPER et champ caché JEV_GUIDPER
var nodoss, libdoss, retour, tmp : String;
begin
   // retourne NoDossier;GuidPer,Libellé
   retour := AGLLanceFiche('YY','YYDOSSIER_SEL', '','','');
   if retour<>'' then
     begin
     nodoss := READTOKENST(retour);
     tmp := READTOKENST(retour);
     libdoss := READTOKENST(retour);
     SetControlText('JEV_NODOSSIER', nodoss);
     SetControlText('DOS_LIBELLE', libdoss);
     TFMul(Ecran).BChercheClick(Sender);
     end;
end;


procedure TOF_JUEvenement.DOS_LIBELLE_OnExit(Sender: TObject);
// Recherche sur nom dossier
begin
   if GetControlText('DOS_LIBELLE') = '' then
      SetControlText('JEV_NODOSSIER', '');
   TFMul(Ecran).BChercheClick(Sender);
end;


procedure TOF_JUEvenement.ANN_NOMPER_OnElipsisClick(Sender: TObject);
var nouveau : String;
begin
   if NoDossier='' then
     //nouveau := AGLLanceFiche('YY','ANNUAIRE_SEL', '','','')
     nouveau := LancerAnnuSel ('','','')
   else
     //nouveau := AGLLanceFiche('YY','ANNUAIRE_SEL', '', '', 'LIENSPERSONNE;'+GuidPerDos);
     nouveau := LancerAnnuSel ('', '', 'LIENSPERSONNE;'+GuidPerDos);

   if nouveau<>'' then
     begin
     SetControlText('JEV_GUIDPER', nouveau);
     SetControlText('ANN_NOMPER', RechDom('ANNUAIRE', nouveau, False));
     TFMul(Ecran).BChercheClick( Sender );
     end;
end;


procedure TOF_JUEvenement.BINSERT_OnClick(Sender: TObject);
var sCle_l : String;
begin
  sCle_l := 'ACTION=CREATION;;' +GetControlText('JEV_GUIDPER') + ';MSG;;;MSGINTERNE';
  AGLLanceFiche( 'YY', 'YYEVEN_MSG', '', '', sCle_l );
  AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


Initialization
   registerclasses([TOF_JUEvenement]) ;
end.
