{***********UNITE*************************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 13/03/2003
Modifié le ... :   /  /
Description .. : TOM commune des liens de l'annuaire
Mots clefs ... : TOM;ANNULIEN
*****************************************************************}
unit UTOMYYAnnulien;
/////////////////////////////////////////////////////////////////
interface
/////////////////////////////////////////////////////////////////
uses
   Hctrls, HEnt1, SysUtils, hmsgbox,
   StdCtrls, Controls, HDB, Classes, hRichOLE, ComCtrls, Db, UTOM,
{$IFDEF EAGLCLIENT}
   MaineAGL, eFiche, eFichList,
{$ELSE}
   FE_Main, Fiche, FichList, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
   CLASSAnnuLien, UTOB, dbctrls, HTB97, UJurOutilsMvtRem;
/////////////////////////////////////////////////////////////////
type
   TOM_YYANNULIEN = class(TOM)
         procedure OnArgument (stArgument : String ) ; override;
         procedure OnClose ; override;
         procedure OnNewRecord; override;
         procedure OnLoadRecord ; override;
         procedure OnUpdateRecord ; override;
         procedure OnAfterUpdateRecord ; override;
         procedure OnChangeField ( Field_p : TField)  ; override ;

         procedure OnNewRecordDP;
         procedure OnNewRecordSTE;
         procedure OnLoadRecordDP;
         procedure OnLoadRecordSTE;
         procedure OnChangeFieldDP( Field_p : TField);
         procedure OnChangeFieldSTE( Field_p : TField);
         procedure OnUpdateRecordDP;
         procedure OnUpdateRecordSTE;

      public
         FLien_c : TAnnuLien;
      private

         bRemAutorise_c, bHistoRem_c : boolean;
         FCheckGroup_c   : TMyCheckGroup;
         OBRemOldN_c, OBRemOldN1_c : TOB;
         sGrpFiscal_c : string;
         bLoadFiche : boolean ;//LM20070712
         bConvOk_c : boolean;
         OBConventions_c : TOB;

         procedure OnChange_Convention(Sender : TObject);
         procedure OnElipsisClick_LIEN(Sender : TObject);
         procedure OnClick_BTREM(Sender : TObject);
         procedure OnClick_BTHISTOREM(Sender : TObject);
         procedure OnClick_BTHISTOTIT(Sender : TObject);
         procedure OnCheck_TRAVCONT(Sender : TObject);
         procedure OnCheck_CBGrpFiscal(Sender : Tobject);
         procedure OnChange_TITNATURE(Sender : Tobject);         
//         procedure OnExit_NoAhesion( Sender : TObject );
//         procedure OnExit_NoDhCompl( Sender : TObject );

//         function  CalculeCle( sNoAdmin_p, sNoAdhesion_p : string ) : string;
         function  ControleDetermine(sZoneLibelle_p : string) : string;
         procedure EllipsisSelonOngletChamp( sChamp_p : string);
         procedure BoutonSelonOngletChamp(sChamp_p : string; OnClickBT_p : TNotifyEvent);
//         procedure CheckSelonOngletChamp(sChamp_p : string; OnClickCB_p : TNotifyEvent);
         procedure FiltreListe;
         procedure GereGrpFiscal;

         procedure OngletAfficheDP( nNumOnglet_p : integer; sFonction_p : string );
         procedure OngletAfficheSTE( nNumOnglet_p : integer; sFonction_p, sForme_p : string);
         procedure PanelAfficheDP( sFonction_p : string );
//         procedure PanelAfficheSTE( nNumOnglet_p : integer; sFonction_p, sForme_p : string);

         procedure PersonneChoix(sChamp_p : string);
         procedure PersonneAffiche(sGuidPer_p : string);
         procedure PersonneLieeAffiche( sChamp_p : string);

         procedure RacineDetermine( nNumOnglet_p : integer; var sPrefixe_p, sRacine_p : string);

         function RemCharge(OBRem_p : TOB; sCondition_p : string) : boolean;
         function RemChargeN(OBRem_p : TOB) : boolean;
         function RemChargeN1(OBRem_p : TOB) : boolean;

         procedure RemChargeEtAffiche(iNumOnglet_p : integer); overload;
         procedure RemChargeEtAffiche(iNumOnglet_p : integer; OBRemN_p, OBRemN1_p : TOB); overload;

         procedure RemRechargeEtAffiche(iNumOnglet_p : integer);

         procedure RemAffiche(OBRem_p : TOB; sTypeRem_p, sPref_p : string; bVisible_p : boolean);
         procedure RemAfficheDetail(OBRem_p : TOB; sChamp_p, sPref_p : string);
         procedure RemAfficheTablette(OBRem_p : TOB; sTablette_p, sChamp_p, sPref_p : string);
         procedure RemCacheDetail(sChamp_p, sPref_p : string; bVisible_p : boolean);
         procedure RemRazDetail(sChamp_p, sPref_p : string);
         procedure RemRazTablette(sChamp_p, sPref_p : string);
         function  RemVerif(OBRemOld_p, OBRemNew_p : TOB; sTyperem_p : string; var fMontant_p : double) : boolean;
         // procedure HistoLien;
         procedure handlerChangeField (sender:TObject) ;//LM20070704

         function VerifTauxDetentionDirect : boolean;
         function VerifTauxDetentionIndirect : boolean;
         function VerifTauxDetentionSomme(sFromChamp_p : string) : boolean;
         function VerifTauxDetentionTotal : boolean;

         {+GHA 11/2007 - FQ 11860}
         procedure AjouteLienAnnuaire(Guidperdos, Guidper, typedos, fct: String;NoOrdre : integer);
         procedure AjouteDPFiscal(Guidperdos, Guidper, fct, IntegFisc : String);
         {-GHA 11/2007 - FQ 11860}

         // Convetions
         procedure ConventionsCharge;
  end;

function LanceFicheLien(Nature, Fiche, Lequel, Range, Argument : String; TomFiscale : TOM): String;

/////////////////////////////////////////////////////////////////
implementation
/////////////////////////////////////////////////////////////////
uses
   uSatUtil,
   galOutil, DpJurOutils, galSystem,
   dpOutils, AnnOutils,
   DpJurOutilsBlob, MaskUtils,UtilLiensAnnuaire, StrUtils;

const
   //Zones qui ont .name <> .DataField LM20070712
   ZoneDoublee = 'ANL_NOMDATE_P3,ANL_EXPDATE_P3,'
               + 'ANL_EFDEB_P6,ANL_TTNBTOT_P6,ANL_CONVTXT_P5,ANL_CONVTXT_P6,ANL_CONVSUITE_P7';

var
   TomDpFiscal : TOM;


// FQ 11995 : on passe la tom fiscale pour gestion des FRP
function LanceFicheLien(Nature, Fiche, Lequel, Range, Argument : String; TomFiscale : TOM): String;
begin
  TomDpFiscal := TomFiscale;
  Result := AGLLanceFiche(Nature, Fiche, Lequel, Range, Argument);
  TomDpFiscal := Nil;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 11/03/2003
Modifié le ... :   /  /
Description .. : OnArgument
                 Récupération des paramètres
                 Affiche / cache les zones, boutons, onglets...
                 Initialisation de la classe TAnnuLien
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnArgument(stArgument: String);
var
  sAction_l, sFiltre_l, s : string;
  ctrl : TControl ;
  i, iNbCol_l : integer ;
begin
   inherited;
   bLoadFiche :=true ;
//   if not V_PGI.SAV then
//      SetControlVisible('PAGEENCOURS', FALSE);

   sAction_l := ReadTokenSt(stArgument); //action =

   FLien_c := TAnnuLien.Create;
   FLien_c.sGuidPerDos_c := ReadTokenSt(stArgument);
   FLien_c.sTypeDos_c := ReadTokenSt(stArgument);
   FLien_c.nNoOrdre_c := StrToInt(ReadTokenSt(stArgument));
   FLien_c.sFonction_c := ReadTokenSt(stArgument);
   FLien_c.sGuidPer_c := ReadTokenSt(stArgument);
   FLien_c.sForme_c := ReadTokenSt(stArgument);
   sGrpFiscal_c := ReadTokenSt(stArgument);

   // Droits de création,
   GereDroitsConceptsAnnuaire(Self, Ecran);

   TFFiche(Ecran).Retour :=  '0';
   FLien_c.nOnglet_c := FLien_c.OngletChoix(FLien_c.sFonction_c);
   iNbCol_l := 3;

   if FLien_c.nOnglet_c = 0 then   // DP (fiches autres que FICHELIEN, notamment FICHINTERVENTION)
   begin
      // ADT non utilisée car Administration du travail n'est pas une famille de personne
      // => on utilise le type DDT dans la famille SOCial
      if (FLien_c.sFonction_c = 'FIS') or (FLien_c.sFonction_c = 'ADT') or
         (FLien_c.sFonction_c = 'SOC') or (FLien_c.sFonction_c = 'OGA') then
      begin
         sFiltre_l := 'JTP_FAMPER = "' + FLien_c.sFonction_c + '"';
      end
      else if (FLien_c.sFonction_c = 'DIV') then
      begin
         sFiltre_l := '(JTP_FAMPER <> "SOC" AND JTP_FAMPER <> "FIS" AND ' +
                      ' JTP_FAMPER <> "OGA" AND JTP_TYPEPER <> "CAC")';
      end;
      THDBValCombobox(GetControl('ANL_TYPEPER')).Plus := sFiltre_l;

      if (GetControl('ANL_GRPFISCAL') <> nil) and GetControlVisible('ANL_GRPFISCAL') then
         TDBCheckBox(GetControl('ANL_GRPFISCAL')).OnClick := OnCheck_CBGrpFiscal;

   end
   else // FICHELIEN
   begin
      if (FLien_c.nOnglet_c = 1) and (Ecran.Name = 'INTERVENANT') then    // Fonction intervenant
      begin
         iNbCol_l := 2;
         FiltreListe;
      end;

      if GetControl('TANL_GUIDPER') <> nil then
         SetControlEnabled('TANL_GUIDPER', (FLien_c.sGuidPer_c = ''));

      SetControlVisible('BINSERT', FLien_c.nOnglet_c <> 1);

      if GetControl('ANL_TITNATURE') <> nil then
      begin
         THDBValComboBox(GetControl('ANL_TITNATURE')).OnChange := OnChange_TITNATURE;
      end;

      bConvOk_c := (GetControl('OLE_CONVENTION') <> nil) and (GetControlVisible('OLE_CONVENTION'));
      if bConvOk_c then
      begin
         OBConventions_c := TOB.Create('LIENSOLE', nil, -1);
         THRichEditOle(GetControl('OLE_CONVENTION')).onChange := OnChange_Convention;
         THRichEditOle(GetControl('OLE_CONVLIBRE')).onChange := OnChange_Convention;
      end;
   end;

   // Nouvel onglet présent uniquement sur YY FICHELIEN et JUR INTERVENANT
   bRemAutorise_c := GetControl('PAGEREMUNERATION') <> Nil;
   if bRemAutorise_c then
   begin
     // Rémunérations
     TDBCheckBox(GetControl('ANL_TRAVCONT')).OnClick := OnCheck_TRAVCONT;
     if GetControl('BTREM') <> nil then
        TToolbarButton97(GetControl('BTREM')).OnClick := OnClick_BTREM;
     if GetControl('BTHISTOREM') <> nil then
        TToolbarButton97(GetControl('BTHISTOREM')).OnClick := OnClick_BTHISTOREM;

     FCheckGroup_c := TMyCheckGroup.Create(TGroupBox(GetControl('JMR_AVANTAGE')),
                                          'PGAVANTAGENATURE', 'JMR_AVANTAGE', iNbCol_l, false);

      // Pour historisation
      OBRemOldN_c := TOB.Create('JUMVTREM', nil, -1);
      OBRemOldN1_c := TOB.Create('JUMVTREM', nil, -1);
   end;

   // FQ 11508
   SetControlVisible('BINSERT', False);
   if GetControl('BTLIENANNU') <> nil then
      SetControlVisible('BTLIENANNU', False);

 {$ifdef EAGLCLIENT}
  //Gestion des zones qui ont .name <> .DataField
  s := '*' ;
  i:=1 ;//+LM20070712
  while s<>'' do
  begin
    s:=trim(gtfs(ZoneDoublee, ',', i));
    if s<>'' then
    begin
      ctrl := getControl(s) ;
      if ctrl<> nil then TEdit(ctrl).OnExit := handlerChangeField ;
    end ;
    inc(i) ;
  end ;            //-LM20070712


 {$endif EAGLCLIENT}
   //-LM20070704

end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 23/04/2003
Modifié le ... :   /  /    
Description .. : OnClose
                 Libération de la classe
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.OnClose;
begin
   inherited;
   //bLoadFiche

   // maj DOR_ATTACHEMENT et DOR_NBRATTACH si on a renseigné des filiales
   if (FLien_c.sGuidPerDos_c<>'') and (FLien_c.sFonction_c='FIL') then
     MajNbFilialesOrga(FLien_c.sGuidPerDos_c);

   FLien_c.free;
   if bRemAutorise_c then
   begin
      FCheckGroup_c.Destroy(true);
      OBRemOldN_c.Free;
      OBRemOldN1_c.Free;
   end;

   if OBConventions_c <> nil then
      OBConventions_c.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 14/03/2003
Modifié le ... :   /  /    
Description .. : OnNewRecord
                 Nouvel enregistrement, initialisation avec les valeurs par défaut,
                 détermination de la nature du lien (onglet), et appel de la procédure
                 adéquate
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.OnNewRecord;
begin
   inherited;
   if FLien_c.sGuidPer_c <> '' then
      SetField('ANL_GUIDPER', FLien_c.sGuidPer_c);

   SetField('ANL_GUIDPERDOS', FLien_c.sGuidPerDos_c);
   SetField('ANL_TYPEDOS', FLien_c.sTypeDos_c);
   SetField('ANL_NOORDRE', FLien_c.nNoOrdre_c);
   SetField('ANL_FONCTION', FLien_c.sFonction_c);
   SetField('ANL_CODEDOS', GetDosjuri (FLien_c.sGuidPerDos_c));
   if bConvOk_c then
      SetField('ANL_CONVTXT', AglGetGuid);

   if FLien_c.nOnglet_c = 0 then   // DP
      OnNewRecordDP
   else
      OnNewRecordSTE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 26/03/2003
Modifié le ... :   /  /    
Description .. : OnNewRecordDP
                 Nouvel enregistrement de type DP : données de l'organisme
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnNewRecordDP;
//var
//   TOBTypePer :TOB;
begin
   SetField('ANL_GRPFISCAL', sGrpFiscal_c);
   SetField('ANL_AFFICHE', RechDom('JUTYPEFONCTAFF', FLien_c.sFonction_c, FALSE));
   FLien_c.GetTypePer;
//   bFromCalc_c := true;
{   FLien_c.Organisme;
   SetControlText('ANN_NOADMIN', FLien_c.sNoAdmin_c );
   SetControlText('ANN_NOIDENTIF', FLien_c.sNoIdentif_c );
   SetControlText('ANN_COMPLTNOADMIN', FLien_c.sCompltNoAdmin_c );}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 26/03/2003
Modifié le ... :   /  /
Description .. : OnNewRecordSTE
                 Nouvel enregistrement de type STE : initialisation
                 des données par défaut
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnNewRecordSTE;
begin
   if FLien_c.sGuidPer_c = '' then
      PersonneChoix('TANL_GUIDPER');

   // Valeurs par défaut
   if FLien_c.LienValeurDefaut then
   begin
      if (FLien_c.sDefaut_c = 'X') and (FLien_c.sFonction_c <>' STE') then
         SetField('ANL_FORME', FLien_c.sForme_c);
      SetField('ANL_RACINE', FLien_c.sRacine_c);
      SetField('ANL_TYPEDOS', FLien_c.sTypeDos_c);
      SetField('ANL_TRI', FLien_c.nTri_c);
      SetField('ANL_TIERS', FLien_c.sTiers_c);
      SetField('ANL_AFFICHE', FLien_c.sAffiche_c );
      SetField('TJTF_FONCTLIBELLE', FLien_c.sFonctLibbelle_c );
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 13/03/2003
Modifié le ... :   /  /
Description .. : OnLoadRecord
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnLoadRecord;
var s:string ;
    i:integer ;
begin
   bLoadFiche :=true ;
   inherited;
   Ecran.Caption := 'Lien d''une personne : ' + GetNomCompPer(GetField('ANL_GUIDPER'));

   // On recharge l'onglet si jamais on a changé de fonction (cas de la fiche liste intervenant)
   FLien_c.nOnglet_c := FLien_c.OngletChoix( GetField('ANL_FONCTION') );
   if (FLien_c.nOnglet_c = 0) then    // DP
      OnLoadRecordDP
   else
      OnLoadRecordSTE;

   if bConvOk_c then
   begin
      ConventionsCharge;
   end;

  {$ifdef EAGLCLIENT}
  //Chargement des zones qui ont .name <> .DataField
  s := '*' ; i:=1 ;//+LM20070524
  while s<>'' do
  begin
    s:=trim(gtfs(ZoneDoublee, ',', i));
    if s<>'' then setcontrolText(s, getField(copy(s, 1, length(s)-3)) ) ;
    inc(i) ;
  end ;            //-LM20070524
  {$endif}

  bLoadFiche :=false ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnLoadRecordSTE;
var
   sDevise_l : string;
   dCapital_l : double;
   iCapNbTitre_l : integer;
   sForme_l : string;
   bForme_l : boolean;
begin
//   bMajNbVoix_c := (GetField('ANL_VOIXAGO')='0') and (GetField('ANL_VOIXAGE')='0');
   if GetCapital(GetField('ANL_GUIDPERDOS'), sDevise_l, dCapital_l, iCapNbTitre_l) then
   begin
      if GetControl('ANN_CAPITAL') <> nil then
         SetControlText('ANN_CAPITAL', FloatToStr(dCapital_l));
      if GetControl('ANN_CAPNBTITRE') <> nil then
         SetControlText('ANN_CAPNBTITRE', IntToStr(iCapNbTitre_l));
      if GetControl('ANN_CAPDEV') <> nil then
         SetControlText('ANN_CAPDEV', sDevise_l);
   end;

   FLien_c.sGuidPerDos_c := GetField('ANL_GUIDPERDOS');
   FLien_c.sCodeDos_c := GetField('ANL_CODEDOS');
   FLien_c.sTypeDos_c := GetField('ANL_TYPEDOS');
   FLien_c.nNoOrdre_c := GetField('ANL_NOORDRE');
   FLien_c.sFonction_c := GetField('ANL_FONCTION');
   FLien_c.sGuidPer_c := GetField('ANL_GUIDPER');

   sForme_l := GetField('ANL_FORME');
   if sForme_l = '' then
      sForme_l := FLien_c.sForme_c;
   // Affichages
   OngletAfficheSTE(FLien_c.nOnglet_c, FLien_c.sFonction_c, sForme_l);
//   PanelAfficheSTE(FLien_c.nOnglet_c, FLien_c.sFonction_c, sForme_l);

   EllipsisSelonOngletChamp('ANL_GUIDPER');
   EllipsisSelonOngletChamp('ANL_PERASS1GUID');
   if (FLien_c.nOnglet_c = 4) then
      EllipsisSelonOngletChamp('ANL_COOPTGUID');

   BoutonSelonOngletChamp('BTHISTOTIT', OnClick_BTHISTOTIT);

   // Nouvel onglet présent uniquement sur YY FICHELIEN et JUR INTERVENANT
   if bRemAutorise_c then
     RemChargeEtAffiche(FLien_c.nOnglet_c);
{
   RemChargeEtAffiche(FLien_c.nOnglet_c);
   BoutonSelonOngletChamp('BTHISTOREM', OnClick_BTHISTOREM);
   BoutonSelonOngletChamp('BTREM', OnClick_BTREM);
   RemChargeEtAffiche(FLien_c.nOnglet_c);
   CheckSelonOngletChamp('ANL_TRAVCONT', OnCheck_TRAVCONT);}

   FLien_c.TitresEtDroits( FLien_c.nNbTitresClot_c, FLien_c.nNbDroitsVote_c);
   if GetControl('TJUR_DOSLIBELLE') <> nil then
      SetControlText('TJUR_DOSLIBELLE', FLien_c.sDosLibelle_c );
   if GetControl('TJUR_FORME') <> nil then
      SetControlText('TJUR_FORME', FLien_c.sJurForme_c );
   if GetControl('TJTF_FONCTLIBELLE') <> nil then
      SetControlText('TJTF_FONCTLIBELLE', FLien_c.sFonctLibelle_c );

   bForme_l := (FLien_c.sForme_c = 'SA') or (FLien_c.sForme_c = 'SAS') or
               (FLien_c.sForme_c = 'SASU') or (FLien_c.sForme_c = 'SADIR');
   SetControlVisible('TANL_TITNATURE', bForme_l and (FLien_c.sFonction_c = 'ACT'));
   SetControlVisible('ANL_TITNATURE', bForme_l and (FLien_c.sFonction_c = 'ACT'));
   SetControlVisible('TANL_TITTYPE', bForme_l and (FLien_c.sFonction_c = 'ACT'));
   SetControlVisible('ANL_TITTYPE', bForme_l and (FLien_c.sFonction_c = 'ACT'));

   OnChange_TITNATURE(nil);      
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 25/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnLoadRecordDP;
var
   sFonction_l : string;
begin
   // fct sélectionnée dans la combo pour l'enreg en cours =
   // NPC avec Fonction demandée dans onargument lors de l'appel de la fiche
   sFonction_l := GetField('ANL_FONCTION');
   OngletAfficheDP( FLien_c.nOnglet_c, sFonction_l );
   PanelAfficheDP(sFonction_l);
   EllipsisSelonOngletChamp('ANL_INTERLOCUTEUR');
   // rappel du libellé venant de l'annuaire
//   sGuidPer_l := GetField ('ANL_GUIDPER');
//   SetControlText( 'ANL_NOMPER', RechDom('ANNUAIRE', sGuidPer_l, FALSE) );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 13/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnChangeField( Field_p: TField);
begin
   inherited;
   if (Field_p.FieldName = 'ANL_GUIDPER') then
   begin
      PersonneAffiche(GetField('ANL_GUIDPER'));
   end
   else if (Field_p.FieldName = 'ANL_CONV') and bConvOk_c then
   begin
      SetControlEnabled('OLE_CONVENTION', GetField('ANL_CONV') = 'X');
   end
   else if (Field_p.FieldName = 'ANL_CONVLIB') and bConvOk_c then
   begin
      SetControlEnabled('OLE_CONVLIBRE', GetField('ANL_CONVLIB') = 'X');
   end;

   if (FLien_c.nOnglet_c = 0) then    // DP
      OnChangeFieldDP( Field_p )
   else
      OnChangeFieldSTE( Field_p );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 26/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.OnChangeFieldDP( Field_p : TField);
begin
   if (Field_p.FieldName = 'ANL_TYPEPER')  then
   begin
      if GetField('ANL_TYPEPER') = 'BNQ' then
         SetControlVisible('TSBANQUE', TRUE)
      else
         SetControlVisible('TSBANQUE', false);
   end
   //
   else if (Field_p.FieldName = 'ANL_GRPFISCAL') then
   begin
      GereGrpFiscal;
   end
   //
   else if (Field_p.FieldName = 'ANL_TXDETDIRECT') then
   begin
      if (DS.State <> dsBrowse) and (GetField('ANL_GRPFISCAL') = 'X') then
         SetField('ANL_TXDETTOTAL', GetField('ANL_TXDETDIRECT') + GetField('ANL_TXDETINDIRECT'));

{      if not bLoadFiche and (DS.State <> dsBrowse) then
      begin
         if not VerifTauxDetentionDirect then
            exit; // SORTIE !
         if not bFromCalc_c and not VerifTauxDetentionSomme('ANL_TXDETDIRECT') then
            exit; // SORTIE !
         if not bFromCalc_c and (GetField('ANL_GRPFISCAL') = 'X') then
         begin
            bFromCalc_c := true;
            SetField('ANL_TXDETTOTAL', GetField('ANL_TXDETDIRECT') + GetField('ANL_TXDETINDIRECT'));
         end;
      end;}
   end
   //
   else if (Field_p.FieldName = 'ANL_TXDETINDIRECT') then
   begin
      if (DS.State <> dsBrowse) and (GetField('ANL_GRPFISCAL') = 'X') then
         SetField('ANL_TXDETTOTAL', GetField('ANL_TXDETDIRECT') + GetField('ANL_TXDETINDIRECT'));

{      if not bLoadFiche and (DS.State <> dsBrowse) then
      begin
         if not VerifTauxDetentionIndirect then
            exit; // SORTIE !
         if not bFromCalc_c and not VerifTauxDetentionSomme('ANL_TXDETINDIRECT') then
            exit; // SORTIE !
         if not bFromCalc_c and (GetField('ANL_GRPFISCAL') = 'X') then
         begin
            bFromCalc_c := true;
            SetField('ANL_TXDETTOTAL', GetField('ANL_TXDETDIRECT') + GetField('ANL_TXDETINDIRECT'));
         end;
      end;}
   end
   //
   else if (Field_p.FieldName = 'ANL_TXDETTOTAL') then
   begin
{      if not bLoadFiche and not bFromCalc_c and (DS.State <> dsBrowse) then
      begin
         if not VerifTauxDetentionIndirect then
            exit; // SORTIE !
         if not VerifTauxDetentionSomme then
            exit; // SORTIE !
         if not VerifTauxDetentionTotal then
           exit; // SORTIE !
      end;
      bFromCalc_c := false;}
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 26/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnChangeFieldSTE( Field_p : TField);
var
   nVoixTotal_l : integer;
begin
   if (Field_p.FieldName = 'ANL_TYPEPER') then
   begin
      if GetValChamp('JUTYPEFONCT','JTF_TIERS', 'JTF_FONCTION = "' + GetField('ANL_FONCTION') + '"') = 'X' then
         SetField('ANL_TYPEPERDOS', GetField('ANL_TYPEPER'));
   end
   //
   else if (Field_p.FieldName = 'ANL_PERASS1GUID') or (Field_p.FieldName = 'ANL_COOPTGUID') then
   begin
      PersonneLieeAffiche( Field_p.FieldName );
   end
   //
   else if (Field_p.FieldName = 'ANL_TTNBPP') or (Field_p.FieldName = 'ANL_TTNBUS') or
      (Field_p.FieldName = 'ANL_TTNBNP') then
   begin
      FLien_c.VoixCalcule( GetField('ANL_TTNBPP'), GetField('ANL_TTNBUS'), GetField('ANL_TTNBNP'),
                           FLien_c.nVoixAgo_c, FLien_c.nVoixAge_c, FLien_c.nTTNBTOTUS_c,
                           FLien_c.nTTNBTOTPP_c );

//      if bMajNbVoix_c then
//      begin
      SetField('ANL_VOIXAGO', FLien_c.nVoixAgo_c);
      SetField('ANL_VOIXAGE', FLien_c.nVoixAge_c);
//      end;

      SetField('ANL_TTNBTOTUS', FLien_c.nTTNBTOTUS_c);
      SetField('ANL_TTNBTOT', FLien_c.nTTNBTOTPP_c);

      FLien_c.TotauxPCTCalcule( FLien_c.nNbTitresClot_c, FLien_c.nTTNBTOTUS_c,
                        FLien_c.nTTNBTOTPP_c, FLien_c.dPctBenef_c, FLien_c.dPctCap_c );
      if (GetField('ANL_TTPCTBENEF') <> FLien_c.dPctBenef_c) or
         (GetField('ANL_TTPCTCAP') <> FLien_c.dPctCap_c) then
         ModeEdition(DS);
      SetField('ANL_TTPCTBENEF', FLien_c.dPctBenef_c);
      SetField('ANL_TTPCTCAP', FLien_c.dPctCap_c);
   end
   //
   else if (Field_p.FieldName = 'ANL_VOIXAGO') or (Field_p.FieldName = 'ANL_VOIXAGE') then
   begin
      nVoixTotal_l := {GetField('ANL_VOIXAGO') +} GetField('ANL_VOIXAGE');
      FLien_c.dPctVoix_c := FLien_c.TotauxVoixCalcule( FLien_c.nNbDroitsVote_c, nVoixTotal_l);
      if (GetField('ANL_TTPCTVOIX') <> FLien_c.dPctVoix_c) then
         ModeEdition(DS);
      SetField('ANL_TTPCTVOIX', FLien_c.dPctVoix_c );
   end
   //
   else if (Field_p.FieldName = 'ANL_TTPCTBENEF') then
      THNumEdit(GetControl('ANL_QTTPCTBENEF')).value := GetField('ANL_TTPCTBENEF')

   else if (Field_p.FieldName = 'ANL_TTPCTCAP') then
      THNumEdit(GetControl('ANL_QTTPCTCAP')).value := GetField('ANL_TTPCTCAP')

   else if (Field_p.FieldName = 'ANL_TTPCTVOIX') then
      THNumEdit(GetControl('ANL_QTTPCTVOIX')).value := GetField('ANL_TTPCTVOIX');
   //
{   if (DS.State <> dsBrowse) and
      ((Field_p.FieldName = 'ANL_TTNBPP')     or (Field_p.FieldName = 'ANL_TTNBUS') or
       (Field_p.FieldName = 'ANL_TTNBNP')     or (Field_p.FieldName = 'ANL_VOIXAGO') or
       (Field_p.FieldName = 'ANL_VOIXAGE')    or (Field_p.FieldName = 'ANL_TTNBTOTUS') or
       (Field_p.FieldName = 'ANL_TTNBTOT')    or (Field_p.FieldName = 'ANL_TTPCTBENEF') or
       (Field_p.FieldName = 'ANL_TTPCTCAP')   or (Field_p.FieldName = 'ANL_TTPCTVOIX') or
       (Field_p.FieldName = 'ANL_TTPCTBENEF') or (Field_p.FieldName = 'ANL_TTPCTCAP') or
       (Field_p.FieldName = 'ANL_TTPCTVOIX')) then
   begin
//      if OBHistoLien_c.GetValue(Field_p.FieldName) <> GetField(Field_p.FieldName) then
         bHistoRem_c := true;
   end;}

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnUpdateRecord ;
begin
   inherited;
    if (FLien_c.nOnglet_c = 0) then    // DP
      OnUpdateRecordDP
   else
      OnUpdateRecordSTE;

   // Conventions
   if bConvOk_c then
      BlobsEnreg(OBConventions_c);
      
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 26/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnUpdateRecordDP;
var
   strTypePer  :string; // $$$ JP 27/04/06
   sGuidPerDos_l, sGuidPer_l : string;
   sCDI_l, sRecette_l, sInspection_l, sCle_l, sDossier_l : string;
   TobDpFiscal : TOB;
   bIIndExo : boolean; // $$$ JP 27/04/06

begin
   sGuidPerDos_l := GetField('ANL_GUIDPERDOS');
   sGuidPer_l := GetField('ANL_GUIDPER');

   if sGuidPerDos_l=sGuidPer_l then
   begin
      LastError := 1;
      LastErrorMsg := 'Il n''est pas possible de déclarer un lien d''une fiche avec elle-même.';
      exit;
   end;

   // si RDI, anl_organisme et anl_noadhesion obligatoires
   // $$$ JP 27/04/06 - prendre en compte également le SIE
   strTypePer := GetField ('ANL_TYPEPER');

   // FQ 11768 15/10/07 - insortable autrement
   if Not FicheEstChargee('FICHFISCALE') then
   BEGIN
     if (strTypePer = 'RDI') or (strTypePer = 'CDI') or (strTypePer = 'SIE') then
     begin
          // $$$ JP 27/04/06 - chargement du dp fiscal plus tôt, pour contrôle DFI_IMPOTINDIR...
          try
             TobDpFiscal := TOB.Create('DPFISCAL', nil, -1);
             // MD 20/09/07 : mise en commentaire
             {if not TobDpFiscal.SelectDB ('"' + sGuidPerDos_l + '"', nil, TRUE) then
             begin
                  TobDpFiscal.InitValeurs;
                  // vu ce qu'il y a dans la tom, pas besoin de faire un TomDpFiscal.InitTOB(TobDpFiscal) !
                  TobDpFiscal.PutValue ('DFI_GUIDPER', sGuidPerDos_l);
                  TobDpFiscal.PutValue ('DFI_TAXEPROF', 'X');
             end;}
             // SelectDb renvoit False si l'enregistrement est en cours de modif !
             // il ne faut plus essayer de créer ni de modifier dans DPFISCAL dans ce cas,
             // car l'enregistrement peut être en cours de création/modif dans la FICHFISCALE
             // d'où les messages du type "L'enregistrement est inaccessible" FQ 11698
             // ou "Cette fiche existe déjà, vous devez la modifier" FQ 11185
             if TobDpFiscal.SelectDB ('"' + sGuidPerDos_l + '"', nil, TRUE) then
             begin

                 bIIndExo := TobDpFiscal.GetValue ('DFI_IMPOINDIR') = 'E';

                 // maj n° Recette et CDI (cf UpdateTableCa3)
                 sCDI_l        := GetControlText ('ANN_NOIDENTIF');
                 sInspection_l := GetControlText ('ANN_COMPLTNOADMIN');
                 sRecette_l    := GetControlText ('ANN_NOADMIN');
                 sDossier_l    := Trim (GetField('ANL_NOADHESION'));
                 sCle_l        := Trim (GetField ('ANL_NODHCOMPL'));

                 // $$$ JP 27/04/06 - FQ 10964 - contrôle et màj CA3 uniquement si exonéré impots indirects
                 if bIIndExo = FALSE then
                 begin
                      if sDossier_l = '' then
                      begin
                           ErreurChamp('ANL_NOADHESION', 'TSPAIE', Self, 'Le n° de dossier du code FRP est obligatoire');
                           exit;
                      end;
                      if sCle_l = '' then
                      begin
                           ErreurChamp('ANL_NOADHESION', 'TSPAIE', Self, 'La clé du code FRP est obligatoire');
                           exit;
                      end;
                 end;

                 // $$$ JP 27/04/06 - FQ 10964 - contrôle de la clé uniquement si quelque chose de saisi
                 if (ControleCleFRP (sRecette_l, sDossier_l, sCle_l) = FALSE) or ((sCle_l = '') and (sDossier_l <> '')) then
                    begin
                         ErreurChamp( 'ANL_NODHCOMPL', 'TSPAIE', Self, 'La clé du code FRP est incorrecte' );
                         exit;
                    end;

                 // RDI trouvé => met à jour coordonnées + n° Recette et CDI
                 // $$$ JP 27/04/06 - plus la peine
                 //UpdateCoordEtRecetteCDI (sGuidPerDos_l, sGuidPer_l, sRecette_l, sCDI_l, sInspection_l, sDossier_l, sCle_l, TobDpFiscal.GetValue ('DFI_IMPOINDIR') <> 'E');

                 // $$$ JP 16/10/2003: mise à jour n° FRP dans fiche dossier, à partir code FRP du CDI
                 // S.MASSON 18/11/2005: mise à jour n° inspection dans fiche fiscale, à partir code inspection du CDI
                 // $$$ JP 24/04/06 - pour un SIE aussi des mise à jour nécessaire

                 try
                     // $$$ JP 27/04/06 - si RDI ou SIE (le SIE=CDI+RDI à priori)
                     if (strTypePer = 'RDI') or (strTypePer = 'SIE') then
                     begin
                          // Mise à jour du N° FRP
                          // $$$ jP 27/04/06 - ne pas màj si impoindir = 'E'
                          if bIIndExo = FALSE then
                          begin
                               TobDpFiscal.PutValue ('DFI_NOFRP', sRecette_l + sDossier_l + sCle_l );
                               TobDpFiscal.InsertOrUpdateDB(False);
                          end;
                     end;

                     // $$$ JP 27/04/06 - si CDI ou SIE
                     if (strTypePer = 'CDI') or (strTypePer = 'SIE') then
                     begin
                          // Mise à jour du N° inspection // ajout SMA
                          TobDpFiscal.PutValue ('DFI_NOINSPECTION', sInspection_l ); // ajout SMA
                          TobDpFiscal.InsertOrUpdateDB(False);
                     end;
                 except
                       if strTypePer = 'SIE' then
                            PgiInfo ('Mise à jour impossible du numéro FRP et du numéro d''inspection du dossier')
                       else if strTypePer = 'RDI' then
                            PgiInfo ('Mise à jour impossible du numéro FRP du dossier')
                       else
                            PgiInfo ('Mise à jour impossible du numéro d''inspection du dossier'); // ajout SMA
                 end;
             end;
          finally
                 TobDpFiscal.Free;
          end;
     end;
   END
   ELSE
   // MD 08/02/08 - FQ 11995
   // on passe la tom de la fiche fiscale, et on est bien sur la fiche fiscale du dossier
   if (TomDpFiscal<>Nil) and (TomDpFiscal.GetField('DFI_GUIDPER')=sGuidPerDos_l) THEN
   BEGIN
     if (strTypePer = 'RDI') or (strTypePer = 'CDI') or (strTypePer = 'SIE') then
     begin
          // try
             bIIndExo := TomDpFiscal.GetField ('DFI_IMPOINDIR') = 'E';

             // maj n° Recette et CDI (cf UpdateTableCa3)
             sCDI_l        := GetControlText ('ANN_NOIDENTIF');
             sInspection_l := GetControlText ('ANN_COMPLTNOADMIN');
             sRecette_l    := GetControlText ('ANN_NOADMIN');
             sDossier_l    := Trim (GetField('ANL_NOADHESION'));
             sCle_l        := Trim (GetField ('ANL_NODHCOMPL'));

             // $$$ JP 27/04/06 - FQ 10964 - contrôle et màj CA3 uniquement si exonéré impots indirects
             if bIIndExo = FALSE then
             begin
                  if sDossier_l = '' then
                  begin
                       ErreurChamp('ANL_NOADHESION', 'TSPAIE', Self, 'Le n° de dossier du code FRP est obligatoire');
                       exit;
                  end;
                  if sCle_l = '' then
                  begin
                       ErreurChamp('ANL_NOADHESION', 'TSPAIE', Self, 'La clé du code FRP est obligatoire');
                       exit;
                  end;
             end;

             // $$$ JP 27/04/06 - FQ 10964 - contrôle de la clé uniquement si quelque chose de saisi
             if (ControleCleFRP (sRecette_l, sDossier_l, sCle_l) = FALSE) or ((sCle_l = '') and (sDossier_l <> '')) then
                begin
                     ErreurChamp( 'ANL_NODHCOMPL', 'TSPAIE', Self, 'La clé du code FRP est incorrecte' );
                     exit;
                end;

             // RDI trouvé => met à jour coordonnées + n° Recette et CDI
             try
                 // $$$ JP 27/04/06 - si RDI ou SIE (le SIE=CDI+RDI à priori)
                 if (strTypePer = 'RDI') or (strTypePer = 'SIE') then
                 begin
                      // Mise à jour du N° FRP
                      // $$$ jP 27/04/06 - ne pas màj si impoindir = 'E'
                      if bIIndExo = FALSE then
                      begin
                           TomDpFiscal.SetField ('DFI_NOFRP', sRecette_l + sDossier_l + sCle_l );
                      end;
                 end;

                 // $$$ JP 27/04/06 - si CDI ou SIE
                 if (strTypePer = 'CDI') or (strTypePer = 'SIE') then
                 begin
                      // Mise à jour du N° inspection // ajout SMA
                      TomDpFiscal.SetField ('DFI_NOINSPECTION', sInspection_l ); // ajout SMA
                 end;
             except
                   if strTypePer = 'SIE' then
                        PgiInfo ('Mise à jour impossible du numéro FRP et du numéro d''inspection du dossier')
                   else if strTypePer = 'RDI' then
                        PgiInfo ('Mise à jour impossible du numéro FRP du dossier')
                   else
                        PgiInfo ('Mise à jour impossible du numéro d''inspection du dossier'); // ajout SMA
             end;
          // finally

          // end;
     end;
   END;

   // taux de détention
   if ((Ecran.name = 'FICHINTERVENTION') or (Ecran.name = 'FICHINTERVENTION_1')) and (GetControl('GRPBOX1') <> nil) and GetControlVisible('GRPBOX1') then
   begin
      if not VerifTauxDetentionDirect then
         exit;
      if not VerifTauxDetentionIndirect then
         exit;
      if not VerifTauxDetentionSomme('ANL_TXDETINDIRECT') then
         exit;
      if not VerifTauxDetentionTotal then
         exit;
   end;

   FLien_c.LienCalculDefaut;
   SetField('ANL_RACINE', FLien_c.sRacine_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Meriaux
Créé le ...... : 30/10/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_YYANNULIEN.VerifTauxDetentionDirect : boolean;
begin
   result := true;
   if (GetField('ANL_GRPFISCAL') = 'X') then
   begin
      if (GetField ('ANL_TXDETDIRECT') > 100) or (GetField ('ANL_TXDETDIRECT') < 0)  then
      begin
         ErreurChamp('ANL_TXDETDIRECT', 'PGeneral', Self, 'Le taux de détention direct doit être #13#10 compris entre 0% et 100%.') ;
         result := false; // SORTIE !
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Meriaux
Créé le ...... : 30/10/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_YYANNULIEN.VerifTauxDetentionIndirect : boolean;
begin
   result := true;
   if (GetField('ANL_GRPFISCAL') = 'X') then
   begin
      if (GetField('ANL_TXDETINDIRECT') > 100) or (GetField('ANL_TXDETINDIRECT') < 0) then
      begin
         ErreurChamp('ANL_TXDETINDIRECT', 'PGeneral', Self, 'Le taux de détention indirect doit être #13#10 compris entre 0% et 100%');
         result := false; // SORTIE !
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Meriaux
Créé le ...... : 30/10/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_YYANNULIEN.VerifTauxDetentionSomme(sFromChamp_p : string) : boolean;
var
   eSomme_l : Extended;
begin
   if (GetField('ANL_GRPFISCAL') = 'X') then
   begin
      result := false;
      eSomme_l := GetField('ANL_TXDETDIRECT') + GetField ('ANL_TXDETINDIRECT');
      if (eSomme_l > 100) or (eSomme_l < 0) then
      begin
         ErreurChamp(sFromChamp_p, 'PGeneral', Self, 'La somme des taux de détention doit être #13#10 compris entre 0% et 100%.');
         Exit;
      end;
      if eSomme_l < 95 then
      begin
         ErreurChamp(sFromChamp_p, 'PGeneral', Self, 'Somme des taux de détention inférieur au seuil #13#10 autorisé par l''intégration fiscale (95%).') ;
         Exit; // SORTIE !
      end;
   end;
   result := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Meriaux
Créé le ...... : 30/10/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_YYANNULIEN.VerifTauxDetentionTotal : boolean;
begin
   result := false;
   if (GetField('ANL_TXDETTOTAL') > 100) or (GetField('ANL_TXDETTOTAL') < 0) then
   begin
      ErreurChamp('ANL_TXDETTOTAL', 'PGeneral', Self, 'Le taux de détention total doit être #13#10 compris entre 0% et 100%.');
      Exit;
   end;
   if ((GetField('ANL_GRPFISCAL') = 'X')) and (GetField ('ANL_TXDETTOTAL') < 95) then
   begin
      ErreurChamp('ANL_TXDETTOTAL', 'PGeneral', Self, 'Taux de détention total inférieur au seuil #13#10 autorisé par l''intégration fiscale (95%).') ;
      Exit; // SORTIE !
   end;
   result := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 26/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnUpdateRecordSTE;
var
   sGuidPerDos_l, sGuidPer_l : string;
begin
   sGuidPerDos_l := GetField('ANL_GUIDPERDOS');
   sGuidPer_l := GetField('ANL_GUIDPER');

   if sGuidPerDos_l = sGuidPer_l then
   begin
      LastError := 1;
      LastErrorMsg := 'Il n''est pas possible de déclarer un lien d''une fiche avec elle-même.';
      exit;
   end;

   if (FLien_c.nOnglet_c = 2) and (FLien_c.sForme_c = 'SCI') and (GetField('ANL_REGIMEFISC') = '') then
   begin
      LastError := 1;
      LastErrorMsg := 'Le régime fiscal doit être renseigné pour les associés des SCI.';
       SetFocusControl('ANL_REGIMEFISC');
      exit;
   end;

   FLien_c.nTTNBPP_c := GetField('ANL_TTNBPP');
   FLien_c.nTTNBNP_c := GetField('ANL_TTNBNP');
   FLien_c.nTTNBUS_c := GetField('ANL_TTNBUS');
   FLien_c.sPerass1Guid_c := GetField('ANL_PERASS1GUID');
   FLien_c.dtExpDate_c := GetField('ANL_EXPDATE');
   FLien_c.sTypePer_c := GetField('ANL_TYPEPERDOS');
   FLien_c.sInfo_c := GetField('ANL_INFO');
   FLien_c.sRacine_c := GetField('ANL_RACINE');
   FLien_c.sAffiche_c := GetField('ANL_AFFICHE');

   FLien_c.LienCalculDefaut;

   SetField('ANL_INFO', FLien_c.sInfo_c);
   SetField('ANL_RACINE', FLien_c.sRacine_c);
   SetField('ANL_AFFICHE', FLien_c.sAffiche_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 21/03/2003
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnAfterUpdateRecord;
var
   nNoOrdre_l : integer;
   sTypeDos_l, sFonction_l, sCodeDos_l,
   sGuidPerDos_l, sGuidPer_l, sIntFiscale : string;

begin
   sGuidPerDos_l := GetField('ANL_GUIDPERDOS');
   sCodeDos_l := GetField('ANL_CODEDOS');
   sTypeDos_l := GetField('ANL_TYPEDOS');
   nNoOrdre_l := GetField('ANL_NOORDRE');
   sFonction_l := GetField('ANL_FONCTION');
   sGuidPer_l := GetField('ANL_GUIDPER');

   sIntFiscale := GetField('ANL_GRPFISCAL');    {+GHA 11/2007 - FQ11860}

   if FLien_c.nOnglet_c <> 0 then
   begin

      if ( sFonction_l <> 'INT')  and ( sFonction_l <> 'DIV' ) and
         ( sFonction_l <> 'FIL' ) and ( sFonction_l <> 'FIS' ) and 
         ( sFonction_l <> 'SOC' ) and ( sFonction_l <> 'TRS' ) then
      begin
         // #### surtout pas passer GetField ('ANL_FORME') qui est souvent vide !
         CreationUnLien(nNoOrdre_l, sGuidPerDos_l, sGuidPer_l, sTypeDos_l,
                        'INT', FLien_c.sForme_c, sCodeDos_l );
      end;
   end;
   // Nouvel onglet présent uniquement sur YY FICHELIEN et JUR INTERVENANT
   if bRemAutorise_c then
     RemChargeEtAffiche(FLien_c.nOnglet_c);
   // HistoLien;
   TFFiche(Ecran).Retour := sGuidPer_l;

   {+GHA 11/2007 FQ 11860.}
   // Mise à jour table ANNULIEN
   AjouteLienAnnuaire(sGuidPerDos_l, sGuidPer_l, FLien_c.sTypeDos_c, FLien_c.sFonction_c, FLien_c.nNoOrdre_c);
   // Mise à jour table DPFISCAL
   AjouteDPFiscal(sGuidPerDos_l, sGuidPer_l, FLien_c.sFonction_c, sIntFiscale);
   {-GHA 11/2007 FQ 11860.}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 16/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
{ procedure TOM_YYANNULIEN.HistoLien;
var
   sCle_l : string;
   // iInd_l : integer;
   OBHistoLien_l : TOB;
begin
   OBHistoLien_l := TOB.Create('ANNULIEN', nil, -1);
   OBHistoLien_l.LoadDetailDB('ANNULIEN',
                          IntToStr(GetField('ANL_GUIDPERDOS')) + ';"' +
                          GetField('ANL_TYPEDOS') + '";' +
                          IntToStr(GetField('ANL_NOORDRE')) + ';"' +
                          GetField('ANL_FONCTION') + '";' +
                          IntToStr(GetField('ANL_GUIDPER')),
                          '', nil, false);

   sCle_l := ' HNL_GUIDPERDOS   = '  + IntToStr(GetField('ANL_GUIDPERDOS')) +
             ' AND HNL_TYPEDOS  = "' + GetField('ANL_TYPEDOS') + '"' +
             ' AND HNL_GUIDPER  = '  + IntToStr(GetField('ANL_GUIDPER')) +
             ' AND HNL_FONCTION = "' + GetField('ANL_FONCTION') + '"';


//   OBHistoLien_l.Detail[0].PutValue('HNL_NOMPROPTITRE',   OBRemOldN_c.Detail[iInd_l].GetValue('JMR_AVANTAGE') + ';' +
//   OBHistoLien_l.Detail[0].PutValue('HNL_CPPROPTITRE',    OBRemOldN_c.Detail[iInd_l].GetValue('JMR_TYPEREM') + ';' +
//   OBHistoLien_l.Detail[0].PutValue('HNL_VILLEPROPTITRE', OBRemOldN_c.Detail[iInd_l].GetValue('JMR_PERIODEREM') + ';' +
//   OBHistoLien_l.Detail[0].PutValue('HNL_NOMACQUERTITRE', + DateToStr(OBRemOldN_c.Detail[iInd_l].GetValue('JMR_DATEDEB')) +
//                     ' au ' + DateToStr(OBRemOldN_c.Detail[iInd_l].GetValue('JMR_DATEFIN'))+ ';' +
//   OBHistoLien_l.Detail[0].PutValue('HNL_VALEURTITRE',    + OBRemOldN_c.Detail[iInd_l].GetValue('JMR_MONTANT') + ';' +
//   OBHistoLien_l.Detail[0].PutValue('HNL_NATURETITRE',    + OBRemOldN_c.Detail[iInd_l].GetValue('JMR_NATURE');

//   Historisation(OBHistoLien_l.Detail[0], 'ANNULIEN', 'HISTOANNULIEN',
//                 sCle_l, ' HNL_NOORDRE DESC', 'HNL_NOORDRE', 'Modification');

   OBHistoLien_l.Free;
end;  }


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 17/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnElipsisClick_LIEN(Sender : TObject);
begin
   PersonneChoix( TControl(Sender).Name );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 02/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.OnCheck_TRAVCONT(Sender : TObject);
begin
   if bRemAutorise_c then
      SetControlEnabled('ANL_TRAVTXT', (GetControlText('ANL_TRAVCONT') = 'X'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 03/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.OnClick_BTREM(Sender : TObject);
var
   sCle1_l, sCle2_l : string;
begin
   sCle1_l := 'JMR_GUIDPERDOS=' + GetField('ANL_GUIDPERDOS') + ';' +
              'JMR_GUIDPER=' + GetField('ANL_GUIDPER') + ';' +
//              'JMR_TYPEREM=001;' +
              'JMR_FONCTION=' + GetField('ANL_FONCTION');

   sCle2_l := GetField('ANL_GUIDPERDOS') + ';' +
             GetField('ANL_GUIDPER') + ';' +
             GetField('ANL_FONCTION');

   AGLLanceFiche('JUR', 'JUMVTREM_MUL', sCle1_l, '',
                        'ACTION=MODIFICATION;' + sCle2_l + ';' +
                        GetControlText('TANL_GUIDPER'));
   RemRechargeEtAffiche(FLien_c.nOnglet_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 14/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.OnClick_BTHISTOTIT(Sender : TObject);
var
   sCle_l : STRING;
begin
   sCle_l := 'HNL_GUIDPERDOS=' + GetField('ANL_GUIDPERDOS') + ';' +
             'HNL_GUIDPER=' + GetField('ANL_GUIDPER') + ';' +
             'HNL_TYPEDOS=' + GetField('ANL_TYPEDOS') + ';' +
             'HNL_FONCTION=' + GetField('ANL_FONCTION');
   AGLLanceFiche('JUR', 'JUHISTOTITRES_MUL', sCle_l, '', '');
end;

procedure TOM_YYANNULIEN.OnClick_BTHISTOREM(Sender : TObject);
var
   sCle_l : STRING;
begin
   sCle_l := 'JHM_GUIDPERDOS=' + GetField('ANL_GUIDPERDOS') + ';' +
             'JHM_GUIDPER=' + GetField('ANL_GUIDPER') + ';' +
             'JHM_FONCTION=' + GetField('ANL_FONCTION');
   AGLLanceFiche('JUR', 'JUHISTOREM_MUL', sCle_l, '', '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 14/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.PersonneChoix(sChamp_p : string);
var
   sChampCode_l, sValeurCode_l, sValeurNom_l, sContact_l : string;
begin
   sChampCode_l := ControleDetermine(sChamp_p);
   sValeurCode_l := GetField(sChampCode_l);
   sContact_l:='';

   if FLien_c.nOnglet_c = 0 then
   begin
        // $$$ JP 11/10/05 - nature YY et non plus DP
        //mcd 12/2005 AglLanceFiche ('YY' {'DP'},'ANNUINTERLOC_PER', sValeurCode_l, '', 'ACTION=MODIFICATION');
        //AccesContact ( sValeurCode_l);
        //sValeurNom_l := TraiteChoixInterlocuteur( sValeurCode_l );
        sContact_l:=AccesContact ( sValeurCode_l);
        sValeurNom_l := TraiteSelectInterlocuteur( sContact_l );
        if sValeurNom_l <> GetField(sChamp_p) then
        begin
             ModeEdition(DS);
             SetField( sChamp_p, sValeurNom_l );
        end;
   end
   else
   begin
       sValeurNom_l := GetControlText(sChamp_p);
       SelectPersonne( sValeurCode_l, sValeurNom_l );
       if sValeurCode_l <> GetControlText(sChamp_p) then
       begin
           ModeEdition(DS);
           SetField( sChampCode_l, sValeurCode_l );
           SetControlText( sChamp_p, sValeurNom_l);
       end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 14/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.PersonneAffiche(sGuidPer_p : string);
var
   sLibelle_l, sNom_l, sTypePer_l : string;
begin
   TraiteChoixPersonne(sGuidPer_p, sLibelle_l, sNom_l, sTypePer_l);

   if GetField('ANL_NOMPER') <> sNom_l then
   begin
      ModeEdition(DS);
      SetField('ANL_NOMPER', sNom_l);
   end;

   // Fonctionnement juridique (FICHELIEN)
   if FLien_c.nOnglet_c <> 0 then
   begin
      if GetControl('TANL_GUIDPER') <> nil then
         SetControlText('TANL_GUIDPER', sLibelle_l);
      SetField('ANL_TYPEPER', sTypePer_l);
   end
   // Fonctionnement autres (FICHINTERVENTION...)
   else
   begin
       // $$$ JP 16/10/2003: si fonction fiscale, il faut avoir le type lien
       if (FLien_c.sFonction_c = 'FIS') and {FQ 11544} ( GetField('ANL_TYPEPER')='' )  then
       begin
           ModeEdition(DS);
           SetField('ANL_TYPEPER', sTypePer_l);
       end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 28/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.PersonneLieeAffiche( sChamp_p : string);
var
   sLibelle_l, sNom_l, sRac_l, sPref_l, sTypePer_l, sChampCode_l, sValeurCode_l : string;
begin
   RacineDetermine(FLien_c.nOnglet_c, sPref_l, sRac_l);
   if FLien_c.nOnglet_c = 0 then   //DP
   begin
      sChampCode_l := ControleDetermine(sChamp_p);
      sValeurCode_l := GetField(sChampCode_l);

      sNom_l := TraiteChoixInterlocuteur(sValeurCode_l);
      SetField( sPref_l + sChamp_p + sRac_l, sNom_l);
   end
   else
   begin
      if GetControl( sPref_l + sChamp_p + sRac_l) <> nil then
      begin
         TraiteChoixPersonne(GetField(sChamp_p), sLibelle_l, sNom_l, sTypePer_l);
         SetControlText( sPref_l + sChamp_p + sRac_l, sLibelle_l);
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 17/03/2003
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.PanelAfficheDP( sFonction_p : string );
begin
   if sFonction_p = 'FIS' then
   begin
      SetControlProperty('ANL_NOADHESION', 'MaxLength', 6);
      SetControlProperty('ANL_NODHCOMPL', 'MaxLength', 2);
//      if GetField('ANL_TYPEPER') = 'RDI' then
//      begin
//         SetControlCaption ('TANL_ORGANISME', 'Centre Des Impôts');
//         SetControlCaption ('TANL_NUMAFFILIATION', 'Recette');
//         SetControlProperty ('ANL_ORGANISME', 'MaxLength', 3);
//         SetControlProperty ('ANL_NOADHESION', 'MaxLength', 10);
//      end
//      else
//      begin
//         SetControlCaption ('TANL_NUMAFFILIATION', 'Numéro Administratif');
//      end;
//      SetControlText ('TANL_NUMINTERNE', 'Complément N° Administratif');
   end
   // FQ 11461
   else if (sFonction_p = 'PFC') or (sFonction_p = 'FIL') or (sFonction_p = 'TIF') then
   begin
      SetControlVisible('TANL_TYPEPER', False);
      SetControlVisible('ANL_TYPEPER', False);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 27/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{procedure TOM_YYANNULIEN.PanelAfficheSTE(nNumOnglet_p : integer; sFonction_p, sForme_p : string);
var
  OBFonction_l : TOB;
  sRequete_l : string;
begin
   if sForme_p = '' then
      exit;

   sRequete_l := 'select JFT_TITRE, JFT_CONV, JFT_TRAV, JFT_REP, JFT_COOP ' +
                 'from JUFONCTION ' +
                 'where JFT_FONCTION = "' + sFonction_p + '"' +
                 'and JFT_FORME = "' + sForme_p + '"';

   OBFonction_l := TOB.Create('JUFONCTION', nil, -1);
   OBFonction_l.LoadDetailFromSQL(sRequete_l);

   if OBFonction_l.Detail.Count > 0 then
   begin
      if nNumOnglet_p = 2 then
      begin
         SetControlVisible('PANELTITRE', OBFonction_l.Detail[0].GetString('JFT_TITRE') = 'X');
         SetControlVisible('PANELCONV', OBFonction_l.Detail[0].GetString('JFT_CONV') = 'X');
         SetControlVisible('PAGECONV', OBFonction_l.Detail[0].GetString('JFT_CONV') = 'X');
         SetControlVisible('PANELTRAV', OBFonction_l.Detail[0].GetString('JFT_TRAV') = 'X');
         SetControlVisible('PANELREM', //bNew_c and
//                                         (sFonction_p <> 'ADM') and (sFonction_p <> 'AGI') and
                                         (OBFonction_l.Detail[0].GetString('JFT_TRAV') = 'X'));
      end
      else if nNumOnglet_p = 4 then
      begin
         SetControlVisible('PANELCONV', OBFonction_l.Detail[0].GetString('JFT_CONV') = 'X');
         SetControlVisible('PAGECONV', OBFonction_l.Detail[0].GetString('JFT_CONV') = 'X');
         SetControlVisible('PANELTRAV', OBFonction_l.Detail[0].GetString('JFT_TRAV') = 'X');
         SetControlVisible('PANELREM', //bNew_c and
//                                         (sFonction_p <> 'ADM') and (sFonction_p <> 'AGI') and
                                         (OBFonction_l.Detail[0].GetString('JFT_TRAV') = 'X'));
         SetControlVisible('PANELREP', OBFonction_l.Detail[0].GetString('JFT_REP') = 'X');
         SetControlVisible('PANELCOOP', OBFonction_l.Detail[0].GetString('JFT_COOP') = 'X');
      end;
   end;
   OBFonction_l.Free;
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 14/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.RacineDetermine( nNumOnglet_p : integer; var sPrefixe_p, sRacine_p : string);
begin
   sRacine_p := '';
   sPrefixe_p := '';
   case nNumOnglet_p of
      0 : ;
      1 : begin sPrefixe_p := 'T'; end;
      2 : begin sRacine_p := '_P1'; sPrefixe_p := 'T'; end;
      3 : begin sRacine_p := '_P2'; sPrefixe_p := 'T'; end;
      4 : begin sRacine_p := '_P3'; sPrefixe_p := 'T'; end;
      5 : begin sRacine_p := '_P4'; sPrefixe_p := 'T'; end;
      6 : begin sRacine_p := '_P5'; sPrefixe_p := 'T'; end;
      7 : begin sRacine_p := '_P6'; sPrefixe_p := 'T'; end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 14/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OngletAfficheDP( nNumOnglet_p : integer; sFonction_p : string );
var
   bVisible_l,bTxDetNul : boolean;
begin
   // divers, filiale ou membre d'un groupe dans l'intégration fiscale
   bVisible_l := (sFonction_p = 'DIV') or (sFonction_p = 'FIL') or (sFonction_p = 'MIF');

   SetControlVisible ('ANL_GRPFISCAL', bVisible_l and ExisteSQL('SELECT 1 FROM DPFISCAL WHERE DFI_GUIDPER="'+GetField('ANL_GUIDPERDOS')+'" AND DFI_TETEGROUPE="X"') ); // Case à cocher "Intégration fiscale"

   {+GHA - 12/2007}
   // En mode modification,
   // Si Intégration fiscale est cochée et les taux de détention ne sont pas
   // renseignés, on force une MàJ pour obliger l'utilisateur à les saisir.
   if (sFonction_p = 'FIL') then
   begin
     bTxDetNul := (GetField('ANL_TXDETDIRECT') = 0) and (GetField ('ANL_TXDETINDIRECT') = 0);
     if (DS.State in [dsBrowse]) and (GetCheckBoxState('ANL_GRPFISCAL') = cbChecked) and (bTxDetNul) then
       DS.Edit;
   end;
   {-GHA - 12/2007}

   SetControlVisible ('TSPAIE', not bVisible_l);    // Onglet "Informations de l'organisme"
   SetControlVisible ('GRPBOX1', (sFonction_p = 'FIL'));        // Détail de la participation
   SetControlVisible ('GBPERIODICITDUC', (sFonction_p = 'SOC')); // Périodicité édition

   // si pas DIV/FIL/MIF
   if not bVisible_l then
   begin
      // Récupère les N° administratifs de l'organisme
      FLien_c.Organisme(not bVisible_l);
      if FLien_c.sFamPer_c = 'OGA' then
      begin
         SetControlCaption('GBFRP', '');

         SetControlCaption('GBNOADMIN', 'N° d''agrément du CGA');
         SetControlProperty('ANN_NOADMIN', 'MaxLength', 6);
         // SetControlProperty('ANN_NOADMIN', 'EditMask', '! |  a  |  a  |  a  |  a  |  a  |  a  |;0;_');

         SetControlVisible('GBNOIDENTIF', false);
         SetControlVisible('GBNODHCOMPL', false);

         SetControlCaption('GBCOMPLTNOADMIN', 'Identification du cabinet');
         SetControlProperty('GBCOMPLTNOADMIN', 'Left', 8);
         SetControlProperty('GBCOMPLTNOADMIN', 'Width', 310);

         SetControlProperty('ANN_COMPLTNOADMIN', 'Left', 27);
         SetControlProperty('ANN_COMPLTNOADMIN', 'Width', 256);
         SetControlProperty('ANN_COMPLTNOADMIN', 'MaxLength', 12);
         // SetControlProperty('ANN_COMPLTNOADMIN', 'EditMask', '! |  9  |  9  |  9  |  9  |  9  |  9  |  9  |  9  |  9  |  9  |  9  |  9  |;0;_');
      end
      else if FLien_c.sFamPer_c = 'FIS' then
      begin
         SetControlCaption('GBNOADMIN', 'Recette / CDIR');
         SetControlProperty('ANN_NOADMIN', 'MaxLength', 7);
         // SetControlProperty('ANN_NOADMIN', 'EditMask', '! |  a  |  a  |  a  |  a  |  a  |  a  |  a  |;0;_');

         SetControlVisible('GBNOIDENTIF', true);

         SetControlCaption('GBCOMPLTNOADMIN', 'Inspection / IFU');
         SetControlProperty('GBCOMPLTNOADMIN', 'Left', 118);
         SetControlProperty('GBCOMPLTNOADMIN', 'Width', 128);

         SetControlProperty('ANN_COMPLTNOADMIN', 'Left', 27);
         SetControlProperty('ANN_COMPLTNOADMIN', 'Width', 74);
         SetControlProperty('ANN_COMPLTNOADMIN', 'MaxLength', 3);
         // SetControlProperty('ANN_COMPLTNOADMIN', 'EditMask', '! |  9  |  9  |  9  |;0;_');

         FLien_c.sCompltNoAdmin_c := Copy(FLien_c.sCompltNoAdmin_c, 1, 3);
      end
      // Liens sociaux, etc...
      else
      begin
         SetControlCaption('GBFRP', 'Informations administratives');
         SetControlProperty('GNNOADHESION', 'Left', 25);
         SetControlProperty('ANL_NOADHESION', 'MaxLength', 20);

         SetControlVisible('GBNOADMIN',       False);
         SetControlVisible('GBNODHCOMPL',     False);
         SetControlVisible('GBNOIDENTIF',     False);
         SetControlVisible('GBCOMPLTNOADMIN', False);
      end;
      // Affiche les N° administratifs récupérés
      SetControlProperty('ANN_NOIDENTIF', 'Width', 54);
      SetControlText('ANN_NOADMIN', FLien_c.sNoAdmin_c );
      SetControlText('ANN_NOIDENTIF', FLien_c.sNoIdentif_c );
      SetControlText('ANN_COMPLTNOADMIN', FLien_c.sCompltNoAdmin_c );
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 28/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.OngletAfficheSTE( nNumOnglet_p : integer; sFonction_p, sForme_p : string);
var
   sNomOnglet_l, sRequete_l : string;
   bRem_l, bTitre_l, bConv_l, bRep_l, bCoop_l : boolean;
   OBFonction_l : TOB;
begin
   OBFonction_l := TOB.Create('JUFONCTION', nil, -1);

   if sForme_p <> '' then
   begin
      sRequete_l := 'select JFT_TITRE, JFT_CONV, JFT_TRAV, JFT_REP, JFT_COOP ' +
                    'from JUFONCTION ' +
                    'where JFT_FONCTION = "' + sFonction_p + '"' +
                    'and JFT_FORME = "' + sForme_p + '"';

      OBFonction_l.LoadDetailFromSQL(sRequete_l);
   end;

   bRem_l := false;
   bTitre_l := false;
   bConv_l := false;
   bRep_l := false;
   bCoop_l := false;

   if OBFonction_l.Detail.Count > 0 then
   begin
      bRem_l   := OBFonction_l.Detail[0].Getstring('JFT_TRAV') = 'X';
      bTitre_l := OBFonction_l.Detail[0].GetString('JFT_TITRE') = 'X';
      bConv_l  := OBFonction_l.Detail[0].GetString('JFT_CONV') = 'X';
      bRep_l   := OBFonction_l.Detail[0].GetString('JFT_REP') = 'X';
      bCoop_l  := OBFonction_l.Detail[0].GetString('JFT_COOP') = 'X';
   end;

   case nNumOnglet_p of
      0 : ;
      1 : ;// appel de la fenetre delphi : INTERV...
      2 :
      begin
         sNomOnglet_l := 'PAGETITRES';
         bConv_l := true;
      end;
      3 : sNomOnglet_l := 'PAGEINDIVI';
      4 :
      begin
         sNomOnglet_l := 'PAGEMANDAT';
         bConv_l := true;
      end;
      5 : sNomOnglet_l := 'PAGECE';
      6 : sNomOnglet_l := 'PAGETIERS';
      7 : sNomOnglet_l := 'PAGEFIL';
//      8 : sNomOnglet_l := 'PAGEREMUNERATION';
   end;


   SetControlVisible('PAGETITRES', false);
   SetControlVisible('PAGEINDIVI', false);
   SetControlVisible('PAGEMANDAT', false);
   SetControlVisible('PAGECE', false);
   SetControlVisible('PAGETIERS', false);
   SetControlVisible('PAGEFIL', false);
   SetControlVisible('PAGEREMUNERATION', bRem_l);

   SetControlVisible('BTREM', bRem_l);
   SetControlVisible('BTHISTOREM', bRem_l);
//   SetControlVisible('PAGECONVENTIONS', bConventions_l);

   SetControlVisible('PAGECONV', bConv_l);
   SetControlVisible('PANELTITRE', bTitre_l);

   SetControlVisible('PANELREP', bRep_l);
   SetControlVisible('PANELCOOP', bCoop_l);

   SetControlVisible(sNomOnglet_l, true);
   SetActiveTabSheet(sNomOnglet_l);

   SetControlVisible('ANL_APPINDUS', sFonction_p = 'ASS');
   SetControlVisible('PLIMPOUVOIR', (sFonction_p = 'DGD') or (sFonction_p = 'DGA'));

   OBFonction_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 13/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.EllipsisSelonOngletChamp( sChamp_p : string);
var
   sRac_l, sPref_l : string;
begin
   RacineDetermine(FLien_c.nOnglet_c, sPref_l, sRac_l);
   if FLien_c.nOnglet_c = 0 then   //DP
   begin
      if GetControl(sPref_l + sChamp_p + sRac_l) <> nil then
         THDBEdit(GetControl(sPref_l + sChamp_p + sRac_l)).OnElipsisClick := OnElipsisClick_LIEN;
   end
   else
   begin
      if GetControl(sPref_l + sChamp_p + sRac_l) <> nil then
         THEdit(GetControl(sPref_l + sChamp_p + sRac_l)).OnElipsisClick := OnElipsisClick_LIEN;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 03/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.BoutonSelonOngletChamp(sChamp_p : string; OnClickBT_p : TNotifyEvent);
var
   sRac_l, sPref_l : string;
begin
   RacineDetermine(FLien_c.nOnglet_c, sPref_l, sRac_l);
   if GetControl(sChamp_p + sRac_l) <> nil then
   begin
      SetControlVisible(sChamp_p + sRac_l, true);//bNew_c);
      TToolbarButton97(GetControl(sChamp_p + sRac_l)).OnClick := OnClickBT_p;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 02/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{procedure TOM_YYANNULIEN.CheckSelonOngletChamp(sChamp_p : string; OnClickCB_p : TNotifyEvent);
var
   sRac_l, sPref_l : string;
begin
   RacineDetermine(FLien_c.nOnglet_c, sPref_l, sRac_l);
   if GetControl(sChamp_p + sRac_l) <> nil then
   begin
      TDBCheckBox(GetControl(sChamp_p + sRac_l)).OnClick := OnClickCB_p;
      OnClickCB_p(TObject(GetControl(sChamp_p + sRac_l)));
   end;
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 08/04/2003
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.FiltreListe;
var
   sSep_l, sFiltre_l : string;
   {$IFNDEF EAGLCLIENT}
   taAnnuLien_l : TTable;
   {$ENDIF}
begin
   sSep_l := '';
   if FLien_c.sGuidPerDos_c <> '' then
      sFiltre_l := 'ANL_GUIDPERDOS = ''' + FLien_c.sGuidPerDos_c + '''';
   if sFiltre_l <>'' then
      sSep_l := ' AND ';
   if FLien_c.sTypeDos_c<>'' then
      sFiltre_l := sFiltre_l + sSep_l + 'ANL_TYPEDOS = ''' + FLien_c.sTypeDos_c + '''';
   if sFiltre_l <>'' then
      sSep_l := ' AND ';
   if FLien_c.nNoOrdre_c <> 0 then
      sFiltre_l := sFiltre_l + sSep_l + 'ANL_NOORDRE = ' + IntToStr(FLien_c.nNoOrdre_c);
   if sFiltre_l <>'' then
      sSep_l := ' AND ';
   if FLien_c.sGuidPer_c <> '' then
      sFiltre_l := sFiltre_l + sSep_l + 'ANL_GUIDPER = ''' + FLien_c.sGuidPer_c + '''';

   {$IFDEF EAGLCLIENT}
   TFFicheListe(ecran).SetNewRange('', StringReplace(sFiltre_l, '''', '"', [rfReplaceAll, rfIgnoreCase]));
   {$ELSE}
   taAnnuLien_l := TTable(Ecran.FindComponent('Ta'));
   if taAnnuLien_l = nil then
      exit;
   taAnnuLien_l.Active := False;
   taAnnuLien_l.Filtered := false;
//
   taAnnuLien_l.Filter := sFiltre_l;
   taAnnuLien_l.Filtered := true;
   taAnnuLien_l.Active := true;
   {$ENDIF}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 12/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{function TOM_YYANNULIEN.CalculeCle( sNoAdmin_p, sNoAdhesion_p : string ) : string;
var
   sCodeFRP_l, sCleCalculee_l : string;
begin
   result := '';
   if ( sNoAdmin_p = '' ) or ( sNoAdhesion_p = '' )  then
      exit;

   sCodeFRP_l := sNoAdmin_p + sNoAdhesion_p;
   CalculerCodeFRP( sCodeFRP_l, sCleCalculee_l );
   result := sCleCalculee_l;
end;}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 14/03/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_YYANNULIEN.ControleDetermine(sZoneLibelle_p : string) : string;
var
   sChamp_l, sPref_l, sRac_l : string;
begin
   RacineDetermine(FLien_c.nOnglet_c, sPref_l, sRac_l);

   sChamp_l := 'ANL_PERASS1GUID';
   if sZoneLibelle_p = 'T' + sChamp_l + sRac_l then
   begin
      result := sChamp_l;
      exit;
   end;

   sChamp_l := 'ANL_COOPTGUID';
   if sZoneLibelle_p = 'T'+ sChamp_l + sRac_l then
   begin
      result := sChamp_l;
      exit;
   end;

   sChamp_l := 'ANL_GUIDPER';
   if sZoneLibelle_p = 'T' + sChamp_l then
   begin
      result := sChamp_l;
      exit;
   end;

   sChamp_l := 'ANL_GUIDPER';
   if sZoneLibelle_p = 'ANL_INTERLOCUTEUR' then
      result := sChamp_l;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/09/2004
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.RemChargeEtAffiche(iNumOnglet_p : integer);
begin
//   if not bNew_c then
//      Exit;
   if GetControlEnabled('ANL_TRAVCONT') then
      OnCheck_TRAVCONT(TObject(GetControl('ANL_TRAVCONT')));
   RemChargeEtAffiche(iNumOnglet_p, OBRemOldN_c, OBRemOldN1_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.RemChargeEtAffiche(iNumOnglet_p : integer; OBRemN_p, OBRemN1_p : TOB);
begin
   bHistoRem_c := false;

//   RacineDetermine(iNumOnglet_p, sPref_l, sRacine_l);
//   bVisible_l := (GetField('ANL_FONCTION') <> 'ADM') and (GetField('ANL_FONCTION') <> 'AGI');

   RemChargeN(OBRemN_p);
   RemAffiche(OBRemN_p, '001', 'SAL', true);  //bVisible_l);
   RemAffiche(OBRemN_p, '002', 'PRI', true);  //bVisible_l);
   RemAffiche(OBRemN_p, '003', 'JET', true);
   RemAffiche(OBRemN_p, '004', '',    true);  //bVisible_l);

   RemChargeN1(OBRemN1_p);
   RemAffiche(OBRemN1_p, '001', 'N1', true);  //bVisible_l);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.RemRechargeEtAffiche(iNumOnglet_p : integer);
var
   OBRemNewN_l, OBRemNewN1_l : TOB;
   fMontant_l, fTmp_l : double;
begin
//   if not bNew_c then
//      Exit;

   OBRemNewN_l := TOB.Create('JUMVTREM', nil, -1);
   OBRemNewN1_l := TOB.Create('JUMVTREM', nil, -1);

   RemChargeEtAffiche(iNumOnglet_p, OBRemNewN_l, OBRemNewN1_l);
   RemVerif(OBRemOldN_c, OBRemNewN_l, '001', fMontant_l);
   RemVerif(OBRemOldN_c, OBRemNewN_l, '002', fMontant_l);
   RemVerif(OBRemOldN_c, OBRemNewN_l, '003', fMontant_l);
   RemVerif(OBRemOldN_c, OBRemNewN_l, '004', fMontant_l);

   RemVerif(OBRemOldN_c, OBRemNewN_l, '001', fTmp_l);

{   if bHistoRem_c then
   begin
      if DS.State = dsBrowse then
         DS.Edit;
      if GetField('ANL_MDSALAIRE') <> fMontant_l then
         SetField('ANL_MDSALAIRE', fMontant_l);
   end;}

   OBRemNewN_l.Free;
   OBRemNewN1_l.Free;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_YYANNULIEN.RemChargeN(OBRem_p : TOB) : boolean;
begin
   result := RemCharge(OBRem_p, 'AND JMR_DATEDEB    <= "' + USDATETIME(Date) + '" ' +
                                'AND (JMR_DATEFIN    > "' + USDATETIME(Date) + '" ' +
                                '     OR JMR_DATEFIN = "01/01/1900")');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_YYANNULIEN.RemChargeN1(OBRem_p : TOB) : boolean;
begin
   result := RemCharge(OBRem_p, 'AND JMR_DATEFIN    < "' + USDATETIME(Date) + '" ' +
                                'AND JMR_DATEFIN    <> "01/01/1900"');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_YYANNULIEN.RemCharge(OBRem_p : TOB; sCondition_p : string) : boolean;
var
   sRequete_l : string;
begin
   sRequete_l := 'SELECT * FROM JUMVTREM ' +
                 'WHERE JMR_GUIDPERDOS = "' + GetField('ANL_GUIDPERDOS') + '"' +
                 '  AND JMR_GUIDPER    = "' + GetField('ANL_GUIDPER') + '"' +
                 '  AND JMR_FONCTION   = "' + GetField('ANL_FONCTION') + '" ' +
                 ' ' + sCondition_p + ' ' +
                 'ORDER BY JMR_DATEDEB DESC';

   OBRem_p.ClearDetail;
   OBRem_p.LoadDetailDBFromSQL('JUMVTREM', sRequete_l);
   result := (OBRem_p.Detail.Count <> 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 09/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.RemAffiche(OBRem_p : TOB;
                                    sTypeRem_p, sPref_p : string;
                                    bVisible_p : boolean);
var
   OBDetail_l : TOB;
begin
   RemRazDetail('DATEDEB', sPref_p);
   RemRazDetail('DATEFIN', sPref_p);
   RemRazDetail('MONTANT', sPref_p);
   RemRazDetail('PERIODEREM', sPref_p);
   RemRazTablette('AVANTAGE', sPref_p);
   RemRazDetail('NATURE', sPref_p);

   OBDetail_l := OBRem_p.FindFirst(['JMR_TYPEREM'], [sTypeRem_p], false);
   if OBDetail_l <> nil then
   begin
      RemAfficheDetail(OBDetail_l, 'DATEDEB', sPref_p);
      RemAfficheDetail(OBDetail_l, 'DATEFIN', sPref_p);
      RemAfficheDetail(OBDetail_l, 'MONTANT', sPref_p);
      RemAfficheDetail(OBDetail_l, 'PERIODEREM', sPref_p);
      RemAfficheTablette(OBDetail_l, 'PGAVANTAGENATURE', 'AVANTAGE', sPref_p);
      RemAfficheDetail(OBDetail_l, 'NATURE', sPref_p);
   end;
   RemCacheDetail('DATEDEB', sPref_p, bVisible_p);
   RemCacheDetail('DATEFIN', sPref_p, bVisible_p);
   RemCacheDetail('MONTANT', sPref_p, bVisible_p);
   RemCacheDetail('PERIODEREM', sPref_p, bVisible_p);
   RemCacheDetail('AVANTAGE', sPref_p, bVisible_p);
   RemCacheDetail('NATURE', sPref_p, bVisible_p);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.RemAfficheDetail(OBRem_p : TOB;
                                    sChamp_p, sPref_p : string);
var
   sChamp_l : string;
begin
   sChamp_l := 'JMR_' + sPref_p + sChamp_p;
   if (GetControl(sChamp_l) <> nil) then
   begin
      SetControlText(sChamp_l, OBRem_p.GetValue('JMR_' + sChamp_p));
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 04/01/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.RemAfficheTablette(OBRem_p : TOB;
                                    sTablette_p, sChamp_p, sPref_p : string);
var
   sChamp_l{, sCodes_l, sTmp_l, sLibelles_l} : string;
begin
   sChamp_l := 'JMR_' + sPref_p + sChamp_p;
   if (GetControl(sChamp_l) <> nil) then
   begin
      FCheckGroup_c.OnLoadFromTOB(OBRem_p);
{      sCodes_l := OBRem_p.GetValue('JMR_' + sChamp_p);
      while sCodes_l <> '' do
      begin
         sTmp_l := ReadTokenSt(sCodes_l);
         sTmp_l := RechDom(sTablette_p, sTmp_l, false);
         if sLibelles_l <> '' then sLibelles_l := sLibelles_l + ', ';
         sLibelles_l := sLibelles_l + sTmp_l;
      end;
      if sLibelles_l = '' then sLibelles_l := 'Aucun';
      THLabel(GetControl(sChamp_l)).Caption := sLibelles_l;}
   end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 04/01/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.RemCacheDetail(sChamp_p, sPref_p : string;
                                    bVisible_p : boolean);
var
   sChamp_l : string;
begin
   sChamp_l := 'JMR_' + sPref_p + sChamp_p;
   if (GetControl(sChamp_l) <> nil) then
   begin
      SetControlVisible(sChamp_l, bVisible_p);
      SetControlVisible('T' + sChamp_l, bVisible_p);      
   end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 14/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.RemRazDetail(sChamp_p, sPref_p : string);
var
   sChamp_l : string;
begin
   sChamp_l := 'JMR_' + sPref_p + sChamp_p;
   if (GetControl(sChamp_l) <> nil) then
   begin
      if THEdit(sChamp_l).OpeType = otDate then
         SetControlText(sChamp_l, '01/01/1900')
      else
         SetControlText(sChamp_l, '');
   end;            
end;     

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 14/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.RemRazTablette(sChamp_p, sPref_p : string);
var
   sChamp_l : string;
begin
   sChamp_l := 'JMR_' + sPref_p + sChamp_p;
   if (GetControl(sChamp_l) <> nil) then
      FCheckGroup_c.RazValeurs;
//      THLabel(GetControl(sChamp_l)).Caption := 'Aucun';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 09/09/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_YYANNULIEN.RemVerif(OBRemOld_p, OBRemNew_p : TOB;
                                  sTyperem_p : string; var fMontant_p : double) : boolean;
var
   OBDetailOld_l, OBDetailNew_l : TOB;
begin
   result := false;
   OBDetailOld_l := OBRemOld_p.FindFirst(['JMR_TYPEREM'], [sTypeRem_p], false);
   OBDetailNew_l := OBRemNew_p.FindFirst(['JMR_TYPEREM'], [sTypeRem_p], false);
   if (OBDetailOld_l <> nil) and (OBDetailNew_l <> nil) then
   begin
      if (OBDetailOld_l.GetValue('JMR_DATEDEB') <> OBDetailNew_l.GetValue('JMR_DATEDEB')) then
      begin
         bHistoRem_c := true;
      end;

      if (OBDetailOld_l.GetValue('JMR_DATEFIN') <> OBDetailNew_l.GetValue('JMR_DATEFIN')) then
      begin
         bHistoRem_c := true;
      end;

      if (OBDetailOld_l.GetValue('JMR_MONTANT') <> OBDetailNew_l.GetValue('JMR_MONTANT')) then
      begin
         bHistoRem_c := true;
      end;

      fMontant_p := fMontant_p + OBDetailNew_l.GetValue('JMR_MONTANT');

      if (OBDetailOld_l.GetValue('JMR_PERIODEREM') <> OBDetailNew_l.GetValue('JMR_PERIODEREM')) then
      begin
         bHistoRem_c := true;
      end;

      if (OBDetailOld_l.GetValue('JMR_AVANTAGE') <> OBDetailNew_l.GetValue('JMR_AVANTAGE')) then
      begin
         bHistoRem_c := true;
      end;

      if (OBDetailOld_l.GetValue('JMR_NATURE') <> OBDetailNew_l.GetValue('JMR_NATURE')) then
      begin
         bHistoRem_c := true;
      end;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/10/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_YYANNULIEN.OnChange_TITNATURE(Sender : Tobject);
var
   sCode_l : string;
begin
   if Ecran.Name = 'INTERVENANT' then exit;
   sCode_l := Copy(GetControlText('ANL_TITNATURE'), 1, 2);
   SetControlEnabled('ANL_TITTYPE', (sCode_l = 'AP'));
   if not (DS.State in [dsBrowse]) and (sCode_l <> 'AP') then
      SetField('ANL_TITTYPE', '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/05/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.OnCheck_CBGrpFiscal(Sender : Tobject);
begin
   if bLoadFiche then exit;
   ModeEdition(DS);
   if (GetCheckBoxState('ANL_GRPFISCAL') = cbChecked) then
      SetField('ANL_GRPFISCAL', 'X')
   else
      SetField('ANL_GRPFISCAL', '-');
//   GereGrpFiscal;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/05/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.GereGrpFiscal;
var
   bGrpFiscal_l : boolean;
begin
   bGrpFiscal_l := (GetField('ANL_GRPFISCAL') = 'X');
   SetControlEnabled('ANL_TXDETDIRECT', bGrpFiscal_l);
   SetControlEnabled('ANL_TXDETINDIRECT', bGrpFiscal_l);
//   SetControlEnabled ('ANL_TXDETTOTAL', bGrpFiscal_l);

   if not bGrpFiscal_l and not bLoadFiche and not (DS.State in [dsInsert, dsBrowse]) then
   begin
      SetField('ANL_TXDETDIRECT', 0);
      SetField('ANL_TXDETINDIRECT', 0);
//      SetField('ANL_TXDETTOTAL', 0);
   end;
end;

procedure TOM_YYANNULIEN.handlerChangeField (sender:TObject) ;//LM20070704 contourne la contrainte CWas
var FldName : variant ;
begin
  if (sender=nil) or bLoadFiche then exit ;
  FldName := TControl(sender).Name ;
  FldName:=copy(FldName, 1, length(FldName)-3) ;//On enlève le '_P?' en, fin de zone ce qui donne le nom du champ
  // #### pb : le getcontroltext plante sur une date vide,
  // #### et les zones dont datafield<>name ne sont plus chargées par cbp !!!
  if getControlText(TControl(sender).Name) <> getField(FldName) then
  begin
    ModeEdition(DS);
    setField(FldName, getControlText(TControl(sender).Name) ) ;
  end ;
end ;
////////////////////////////////////////////////////////////////////////////////
//GHA 11/2007 - FQ 11860
procedure TOM_YYANNULIEN.AjouteLienAnnuaire(Guidperdos, Guidper, typedos, fct: String;NoOrdre : integer);
var
  Q : TQuery;
  Sql,s,Pref,cle,Formjur,NoDossierTIF,NoDossierFIL,TitleDlg : string;
  TobAnnulien,TobTmp : TOB;
  NumTable,i : integer;
  ExpreSelect,MsgDlg : WideString;
begin
  if (Guidperdos = '') or (Guidper = '') then
    exit;

  //ctrl si enreg existe avant création
  if fct = 'FIL' then
  begin
    Sql := 'SELECT 1 FROM ANNULIEN'+
        ' WHERE ANL_GUIDPERDOS = "'+Guidper+'"'+
        ' AND ANL_TYPEDOS = "'+typedos+'"'+
        ' AND ANL_NOORDRE = '+intTOstr(NoOrdre)+
        ' AND ANL_FONCTION = "TIF"'+
        ' AND ANL_GUIDPER = "'+Guidperdos+'"';
  end
  else if fct = 'TIF' then
  begin
    Sql := 'SELECT 1 FROM ANNULIEN'+
        ' WHERE ANL_GUIDPERDOS = "'+Guidper+'"'+
        ' AND ANL_TYPEDOS = "'+typedos+'"'+
        ' AND ANL_NOORDRE = '+intTOstr(NoOrdre)+
        ' AND ANL_FONCTION = "FIL"'+
        ' AND ANL_GUIDPER = "'+Guidperdos+'"';
  end
  else
      exit;

  if ExisteSQL(Sql) then exit;

  //récupération des champs de la table annulien pour Ordre "select"
  ExpreSelect := '*';
  Pref := TableToPrefixe('ANNULIEN');
  NumTable := PrefixeToNum(Pref);

  if NumTable >= 1 then
  begin
    ExpreSelect := '';
    for i:=low(V_PGI.DEChamps[NumTable]) to high(V_PGI.DEChamps[NumTable]) do
    begin
      if trim(V_PGI.DEChamps[NumTable,i].Nom) <> '' then
      begin
        if ExpreSelect <> '' then
           ExpreSelect := ExpreSelect + ',' ;
        ExpreSelect := ExpreSelect+V_PGI.DEChamps[NumTable,i].Nom ;
      end;
    end;
  end;

  Q := nil;
  // Màj de la table ANNULIEN
  Sql := 'SELECT '+ExpreSelect+' FROM ANNULIEN'+
        ' WHERE ANL_GUIDPERDOS = "'+Guidperdos+'"'+
        ' AND ANL_TYPEDOS = "'+typedos+'"'+
        ' AND ANL_NOORDRE = '+intTOstr(NoOrdre)+
        ' AND ANL_FONCTION = "'+fct+'"'+
        ' AND ANL_GUIDPER = "'+Guidper+'"';

  TobAnnulien := TOB.Create('ANNULIEN', nil, -1);
  TobAnnulien.LoadDetailDBFromSQL('ANNULIEN',Sql,FALSE);
  TobTmp := TOB.Create('VirtualTob',TobAnnulien,-1);
  try
    // Nom abrégé = clé de recherche
    s := '';
    Q := OpenSQL('SELECT ANN_NOMPER FROM ANNUAIRE WHERE ANN_GUIDPER = "'+Guidperdos+'"',TRUE);
    if not Q.eof then
     s := Q.Fields[0].AsString;
    Ferme(Q);

    TobTmp.PutValue('ANL_GUIDPERDOS',Guidper);
    TobTmp.PutValue('ANL_GUIDPER',Guidperdos);
    TobTmp.PutValue('ANL_NOMPER',s);

    if fct = 'FIL' then
    begin
      TobTmp.PutValue('ANL_FONCTION','TIF');
      TobTmp.PutValue('ANL_RACINE','TIF');
      TobTmp.PutValue('ANL_AFFICHE',RechDom('JUTYPEFONCTAFF', 'TIF', FALSE));
      TobTmp.PutValue('ANL_TXDETDIRECT',0);
      TobTmp.PutValue('ANL_TXDETINDIRECT',0);
      TobTmp.PutValue('ANL_TXDETTOTAL',0);
      TobTmp.PutValue('ANL_GRPFISCAL','-');
    end
    else if fct = 'TIF' then
    begin
      TobTmp.PutValue('ANL_FONCTION','FIL');
      TobTmp.PutValue('ANL_RACINE','FIL');
      TobTmp.PutValue('ANL_AFFICHE',RechDom('JUTYPEFONCTAFF', 'FIL', FALSE));
      TobTmp.PutValue('ANL_TXDETDIRECT',0);
      TobTmp.PutValue('ANL_TXDETINDIRECT',0);
      TobTmp.PutValue('ANL_TXDETTOTAL',0);
      TobTmp.PutValue('ANL_GRPFISCAL','X');
    end;
    //insertion
    TobTmp.InsertDB(nil);

    {+GHA - 12/2007}
    //Si sélection d'une Sté tête de groupe, on affiche la fenêtre
    //pour la saisie des taux de détention de la filiale (dossier synthése en cours)
    if fct = 'TIF' then
    begin
      Cle := Guidper
        +';'+typedos
        +';1' // NoOrdre
        +';FIL'
        +';'+Guidperdos;

      if typedos = 'DP' then
      begin
        Q := nil;
        Q := OpenSQL('select ANN_FORME from ANNUAIRE where ANN_GUIDPER = "'+ GuidPerDos+'"', TRUE);
        if not Q.eof then
          Formjur := Q.Fields[0].AsString;
        Ferme(Q);

        NoDossierFIL := GetNoDossierFromGuidPer(GuidPerDos)+' '+GetNomCompPer(GuidPerDos);
        NoDossierTIF := GetNoDossierFromGuidPer(GuidPer)+' '+GetNomCompPer(GuidPer);

        TitleDlg  := 'Lien d''une personne : ' + GetNomCompPer(GuidPerDos);

        MsgDlg := 'Le choix de la Société "'+NoDossierTIF+'" comme tête de groupe,'+#13#10+
                  'entraîne la création du lien filiale pour le dossier "'+NoDossierFIL+'"'+#13#10+
                  'Veuillez renseigner les taux de détention pour cette filiale';

        PGIBox(MsgDlg,TitleDlg);
        AGLLanceFiche('DP','FICHINTERVENTION', Cle, Cle, 'ACTION=MODIFICATION;'+Cle+';'+FormJur)
      end;
    end;
    {-GHA - 12/2007}
  finally
    TobAnnulien.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GHA
Créé le ...... : 08/02/2008
Modifié le ... :   /  /
Description .. : //GHA 11/2007 - FQ 11860
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.AjouteDPFiscal(Guidperdos, Guidper, fct ,IntegFisc: String);
var
  Q : TQuery;
  Sql,s,Pref : string;
  TobDpFiscal,TobTmp : TOB;
  NumTable,i : integer;
  ExpreSelect : WideString;
begin
  if ((Guidper = '') or (Guidperdos = '')) or ((fct <> 'TIF') and (fct <> 'FIL')) then
    exit;

  Q := nil;

  Sql := 'SELECT 1 FROM DOSSIER WHERE DOS_GUIDPER = "'+Guidper+'"';
  if ExisteSQL(Sql) then
  begin
      Sql := 'SELECT 1 FROM DPFISCAL WHERE DFI_GUIDPER = "'+Guidper+'"';
      if ExisteSQL(Sql) then
      begin
        if fct = 'TIF' then
          exit;
        ExecuteSQL('UPDATE DPFISCAL SET DFI_INTEGRAFISC = "'+IntegFisc+'" WHERE DFI_GUIDPER = "'+Guidper+'"')
      end
      else
      begin
        //récupération des champs de la table annulien pour Ordre "select"
        ExpreSelect := '*';
        Pref := TableToPrefixe('DPFISCAL');
        NumTable := PrefixeToNum(Pref);

        if NumTable >= 1 then
        begin
          ExpreSelect := '';
          for i:=low(V_PGI.DEChamps[NumTable]) to high(V_PGI.DEChamps[NumTable]) do
          begin
            if trim(V_PGI.DEChamps[NumTable,i].Nom) <> '' then
            begin
              if ExpreSelect <> '' then
                 ExpreSelect := ExpreSelect + ',' ;
              ExpreSelect := ExpreSelect+V_PGI.DEChamps[NumTable,i].Nom ;
            end;
          end;
        end;

        Sql := 'SELECT '+ExpreSelect+' FROM DPFISCAL WHERE DFI_GUIDPER = "'+Guidperdos+'"';
        TobDpFiscal := TOB.Create('DPFISCAL', nil, -1);
        TobDpFiscal.LoadDetailDBFromSQL('DPFISCAL',Sql);
        TobTmp := TOB.Create('VirtualTob', TobDpFiscal, -1);
        try
          // No intracommunautaire de la filiale
          s := '';
          Q := OpenSQL('SELECT ANN_SIREN FROM ANNUAIRE WHERE ANN_GUIDPER = "'+Guidper+'"',TRUE);
          if not Q.eof then
            s := Q.Fields[0].AsString;
          Ferme(Q);

          // No intracommunautaire
          s := StringOfChar(' ', 13-length(s))+s;

          TobTmp.PutValue('DFI_GUIDPER',Guidper);
          TobTmp.PutValue('DFI_NOINTRACOMM',s);

          if fct = 'TIF' then
            TobTmp.PutValue('DFI_TETEGROUPE','X')
          else
            TobTmp.PutValue('DFI_TETEGROUPE','-');

          TobTmp.InsertDB(nil);

        finally
          TobDpFiscal.Free;
        end;
      end;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 08/02/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.ConventionsCharge;
var
   OBBlob_l : TOB;
   sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l : string;
   sGuidConv_l : string;
begin
   OBConventions_c.ClearDetail;
   sGuidConv_l := GetField('ANL_CONVTXT');
   if GetControl('OLE_CONVENTION') <> nil then
   begin
      BlobGetCode('OLE_CONVENTION', sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);
      OBBlob_l := BlobChampCreation(sPrefixe_l, sGuidConv_l, sRangBlob_l, sEmploiBlob_l,
                                    sLibelle_l, TRichEdit(GetControl('OLE_CONVENTION'))) ;
      OBBlob_l.ChangeParent(OBConventions_c, -1);
   end;
   if GetControl('OLE_CONVLIBRE') <> nil then
   begin
      BlobGetCode('OLE_CONVLIBRE', sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);
      OBBlob_l := BlobChampCreation(sPrefixe_l, sGuidConv_l, sRangBlob_l, sEmploiBlob_l,
                              sLibelle_l, TRichEdit(GetControl('OLE_CONVLIBRE'))) ;
      OBBlob_l.ChangeParent(OBConventions_c, -1);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : GHA
Créé le ...... : 08/02/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_YYANNULIEN.OnChange_Convention(Sender : TObject);
var
   sControlName_l : string;
   sPrefixe_l, sEmploiBlob_l, sRangBlob_l, sLibelle_l : string;
begin
   sControlName_l := TControl(Sender).Name;
   BlobGetCode(sControlName_l, sEmploiBlob_l, sPrefixe_l, sRangBlob_l, sLibelle_l);

   if BlobChange(OBConventions_c, sPrefixe_l, GetField('ANL_CONVTXT'), sRangBlob_l,
                  TRichEdit(GetControl(sControlName_l)) ) then
   begin
      ModeEdition(DS);
      // Pour forcer la mise à jour des données
      SetField('ANL_GUIDPERDOS', GetField('ANL_GUIDPERDOS'));
   end;

end;


////////////////////////////////////////////////////////////////////////////////////////////////
Initialization
registerclasses([TOM_YYANNULIEN]) ;
end.
