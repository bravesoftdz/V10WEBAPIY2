{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/08/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : JUMVTTITRES ()
Mots clefs ... : TOF;JUMVTTITRES
*****************************************************************}
Unit UTOFMvtTitres;
//////////////////////////////////////////////////////////////////
Interface
//////////////////////////////////////////////////////////////////
Uses
{$IFNDEF EAGLCLIENT}
   mul, FE_Main, EdtREtat,
{$else}
   eMul, MaineAGL, UtileAGL,
{$ENDIF}
   UTOF, UCLASSMyMenu, HDB, HTB97, HCtrls, Menus, sysutils,
   AGLInit, DpJurOutils, HMSGBOX, Classes, UTOB, Controls, 
   Hent1, UJUROutilsMVTTitres, UCLASSKeys, HStatus;       

//////////////////////////////////////////////////////////////////
Type
   TOF_JUMVTTITRES = Class (TOF)
      procedure OnArgument (sArguments_p : String ) ; override ;
      procedure OnLoad; override ;
      procedure OnDisplay; override;
      procedure OnClose; override ;
   private
      FMenuSens_c    : TMyMenu;

      OBMvtCpt_c    : TOB;

      sCaption_c     : string;
      sFonction_c    : string;
      sValNominOuv_c : String; // utile pour la Tom TOM_JUMVTTITRES
      sTypeDos_c     : string;
      sCle_c         : string;
      sNatureCPT_c   : string;

      sGuidPerDos_c  : string;
      sGuidPer_c     : string;
      sFormeJur_c    : string;
      iTTNBPP_c      : integer;

      MyKeys_c       : TKeysPlusMul;

      procedure OnDblClick_FLISTE(Sender : TObject);

      procedure OnClick_BOUVRIR(Sender : TObject);
      procedure OnClick_BINSERT(Sender : TObject);
      procedure OnClick_BIMPRIMER(Sender : TObject);
      procedure OnClick_BSUPPRIMER(Sender : TObject);
      procedure OnClick_BINFOSCPT(Sender : TObject);

      function  InfosCompte : boolean;
      procedure InscriptionEnCompte;
      procedure ReOuvre(sNoOrdre_p : string);

      function SupprimeLesMouvements : boolean;
  end ;


//////////////////////////////////////////////////////////////////
Implementation
//////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_JUMVTTITRES.OnArgument (sArguments_p : String ) ;
var
   sAction_l : String;
begin
   Inherited ;
   THDBGrid(GetControl('FLISTE')).OnDblClick          := OnDblClick_FLISTE;
   TToolbarButton97(GetControl('BOUVRIR')).OnClick    := OnClick_BOUVRIR;
   TToolBarButton97(GetControl('BIMPRIMER')).OnClick := OnClick_BIMPRIMER;   
   TToolbarButton97(GetControl('BINSERT')).OnClick    := OnClick_BINSERT;
   TToolbarButton97(GetControl('BSUPPRIMER')).OnClick := OnClick_BSUPPRIMER;
   TToolbarButton97(GetControl('BINFOSCPT')).OnClick  := OnClick_BINFOSCPT;

   if sArguments_p = '' then
      exit;
   sAction_l      := ReadTokenSt(sArguments_p);
   sGuidPerDos_c  := ReadTokenSt(sArguments_p);
   sTypeDos_c     := ReadTokenSt(sArguments_p);
   sGuidPer_c     := ReadTokenSt(sArguments_p);
   sNatureCPT_c   := ReadTokenSt(sArguments_p);
   sCaption_c     := ReadTokenSt(sArguments_p);
   sFonction_c    := ReadTokenSt(sArguments_p);
   sValNominOuv_c := ReadTokenSt(sArguments_p);
   iTTNBPP_c      := StrToInt(ReadTokenSt(sArguments_p));

   sFormeJur_c := GetValChamp('JURIDIQUE', 'JUR_FORME', 'JUR_GUIDPERDOS = "' + sGuidPerDos_c + '"');

   Ecran.Caption := Ecran.Caption + ' ' + sCaption_c;
   SetControlVisible('PCOMPLEMENT', MontreOnglet(TFMul(Ecran), 'PCOMPLEMENT'));

   FMenuSens_c := TMyMenu.Create(TPopUpMenu(GetControl('POPINSERT')),
                                 'Crédit;Débit',
                                 '+;-',
                                 OnClick_BINSERT);

   MyKeys_c := TKeysPlusMul.Create(TFMul(Ecran),
                     [vk_Sup_c],
                     [TToolbarButton97(GetControl('BSUPPRIMER'))]);
   Ecran.OnKeyDown := MyKeys_c.OnMulKeyDown;

   OBMvtCpt_c := TOB.Create('JUMVTCOMPTES', nil, -1);
   InfosCompte;
   InscriptionEnCompte;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_JUMVTTITRES.OnLoad;
begin
   inherited;
   sCle_c := GetControlText('JMT_GUIDPERDOS') + ';' +
             GetControlText('JMT_TYPEDOS') + ';' +
             GetControlText('JMT_GUIDPER') + ';' +
             GetControlText('JMT_NATURECPT');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 17/02/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_JUMVTTITRES.OnDisplay;
begin
   inherited;
//   TFMul(Ecran).Q.Last;
//   iMaxOrdre_c := TFMul(Ecran).Q.FindField('JMT_NOORDRE').AsInteger;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_JUMVTTITRES.OnClose;
begin
   inherited;
   if FMenuSens_c <> nil then
   begin
      freeandnil(FMenuSens_c);//.OnMenuDetruit;
      //FMenuSens_c.Free;
   end;
   if OBMvtCpt_c <> nil then
      OBMvtCpt_c.free;
   if MyKeys_c <> nil then
      MyKeys_c.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_JUMVTTITRES.OnDblClick_FLISTE(Sender : TObject);
var
   iNoOrdre_l : integer;
   sAction_l : string;
begin
   {$IFDEF EAGLCLIENT}
   GetDataSet; // repositionne le query (semble-t'il) sur l'enreg sélectionné dans la liste !
   {$ENDIF}
   
   iNoOrdre_l := TFMul(Ecran).Q.FindField('JMT_NOORDRE').AsInteger;
   sAction_l := 'ACTION=MODIFICATION';

   sAction_l := AGLLanceFiche('JUR', 'JUMVTTITRES',
                              sCle_c + ';' + IntToStr(iNoOrdre_l), sCle_c + ';' + IntToStr(iNoOrdre_l),
                              sAction_l + ';' + sCle_c + ';' + sCaption_c + ';' + sFonction_c + ';' +
                              sValNominOuv_c);
   ReOuvre(sAction_l);
   AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_JUMVTTITRES.OnClick_BOUVRIR(Sender : TObject);
begin
   OnDblClick_FLISTE(Sender);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_JUMVTTITRES.OnClick_BSUPPRIMER(Sender : TObject);
var
   sMessage_l : string;
//   iRows_l : integer;
begin
   if (not TFMul(Ecran).FListe.AllSelected) and
      (TFMul(Ecran).FListe.NbSelected = 0) then
   begin
      PgiInfo( 'Vous devez sélectionner une ou plusieurs lignes.', Ecran.Caption);

      if TFMul(Ecran).FListe.AllSelected then
         TFMul(Ecran).FListe.AllSelected := False
      else
         TFMul(Ecran).FListe.ClearSelected;
      TFMul(Ecran).bSelectAll.Down := False ;
      exit;
   end;

   sMessage_l := 'Attention : vous supprimez un ou plusieurs mouvements de titres validés.' + #13#10 +
                 'Le nombre de titres des titulaires et bénéficiaires vont être rétablis.' + #13#10 +
                 'Confirmez vous?';

   if PGIAsk(sMessage_l, Ecran.Caption) = mrNo then
   begin
      if TFMul(Ecran).FListe.AllSelected then
         TFMul(Ecran).FListe.AllSelected := False
      else
         TFMul(Ecran).FListe.ClearSelected;
      TFMul(Ecran).bSelectAll.Down := False ;
      exit;
   end;

   SupprimeLesMouvements;

   if TFMul(Ecran).FListe.AllSelected then
      TFMul(Ecran).FListe.AllSelected := False
   else
      TFMul(Ecran).FListe.ClearSelected;
   TFMul(Ecran).bSelectAll.Down := False ;

   AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_JUMVTTITRES.OnClick_BINSERT(Sender : TObject);
var
   sAction_l : string;
begin
   if FMenuSens_c.GetHint = '' then exit; 

   sAction_l := AGLLanceFiche('JUR', 'JUMVTTITRES', '', '', 'ACTION=CREATION;' +
                               sCle_c + ';' + sCaption_c + ';' + sFonction_c + ';' +
                               sValNominOuv_c + ';' + FMenuSens_c.GetHint) ;
   ReOuvre(sAction_l);
   AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 01/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_JUMVTTITRES.OnClick_BIMPRIMER(Sender : TObject);
var
   sCle_l : string;
begin
   sCle_l := 'TIT.JMT_GUIDPERDOS = "' + sGuidPerDos_c + '" AND ' +
             'TIT.JMT_TYPEDOS    = "' + sTypeDos_c + '" AND ' +
             'TIT.JMT_GUIDPER    = "' + sGuidPer_c + '" AND ' +
             'TIT.JMT_NATURECPT  = "' + sNatureCPT_c + '"';  // Compte principal uniquement

   LanceEtat('E', 'JID', 'JMT', true, false, false, nil, sCle_l,
             'Ordres de mouvement', false, 0);

end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 30/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_JUMVTTITRES.OnClick_BINFOSCPT(Sender : TObject);
var
   sCle_l : string;
   nNbTitres_l : integer;
begin
   sCle_l := GetControlText('JMT_GUIDPERDOS') + ';' +
             GetControlText('JMT_GUIDPER') + ';' +
             GetControlText('JMT_NATURECPT');
   nNbTitres_l := TFMul(Ecran).Q.RecordCount;

   sCle_l := AGLLanceFiche('JUR', 'JUMVTCOMPTES', '', sCle_l,
                           'ACTION=MODIFICATION;' + sCle_l + ';' + sCaption_c + ';' +
                           IntToStr(nNbTitres_l));
   InfosCompte;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 31/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
function TOF_JUMVTTITRES.InfosCompte : boolean;
begin
   if not JUMvtCompteVerif(sGuidPerDos_c, sGuidPer_c, sNatureCPT_c, OBMvtCpt_c) then
      JUMvtCompteCreation(sGuidPerDos_c, sGuidPer_c, sNatureCPT_c, OBMvtCpt_c);

   SetControlText('JMCNOCPT', OBMvtCpt_c.Detail[0].GetValue('JMC_NOCPT'));
   SetControlText('JMCINTITULE', OBMvtCpt_c.Detail[0].GetValue('JMC_INTITULE'));
   result := (OBMvtCpt_c.Detail.Count <> 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 30/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_JUMVTTITRES.InscriptionEnCompte;
var
   OBLienTit_l : TOB;
begin
   if JUMvtInscriptionEnCompte(sGuidPerDos_c, sTypeDos_c, sGuidPer_c, sFormeJur_c,
                               sNatureCPT_c, iTTNBPP_c, iTTNBPP_c, Valeur(sValNominOuv_c),
                               'Actionnaire : ' + GetNomAvecCivOuForme(sGuidPer_c)) then
   begin
      OBLienTit_l := TOB.Create('ANNULIEN', nil, -1);
      JUMvtInfosLien(OBLienTit_l, sGuidPerDos_c, sGuidPer_c, 1, sTypeDos_c, sFonction_c);
      JUMvtHistoLien(OBLienTit_l.Detail[0], 'Création auto', '001', sGuidPer_c, '',
                     sGuidPerDos_c, 1, sTypeDos_c, sFonction_c, sValNominOuv_c);
      OBLienTit_l.Free;
      AGLRefreshDB([LongInt(Ecran), 'FListe'], 2);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 31/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_JUMVTTITRES.ReOuvre(sNoOrdre_p : string);
var
   sAction_l : string;
begin
   if sNoOrdre_p = '' then exit;

   sAction_l := AGLLanceFiche('JUR', 'JUMVTTITRES',
                              sCle_c + ';' + sNoOrdre_p, sCle_c + ';' + sNoOrdre_p,
                              'ACTION=MODIFICATION;' + sCle_c + ';' + sCaption_c + ';' +
                              sFonction_c + ';' + sValNominOuv_c);
   ReOuvre(sAction_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 17/02/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :

   sMessage_l := 'JMT_GUIDPERDOS;JMT_TYPEDOS;JMT_GUIDPER;JMT_NATURECPT;JMT_NOORDRE;' +
              '@' + sFonction_c + ';@' + sValNominOuv_c;
   iRows_l := GrilleParcours(TFMul(Ecran),
                            sMessage_l, 7,
                            @InitGetValChamp, @AGLDeleteMouvement, true);

*****************************************************************}
function TOF_JUMVTTITRES.SupprimeLesMouvements : boolean;
var
   bContinue_l : boolean;
   iMaxRow_l, iRow_l : integer;

   iNoOrdre_l : integer;
   sGuidPerDos_l, sGuidPer_l, sTypeDos_l, sNatureCpt_l : string;

   procedure SelectEnreg(var sGuidPerDos_p, sGuidPer_p : string; var iNoOrdre_p : integer;
                         var sTypeDos_p, sNatureCpt_p : string);
   begin
      sGuidPerDos_p := TFMul(Ecran).Q.FindField('JMT_GUIDPERDOS').AsString;
      sTypeDos_p    := TFMul(Ecran).Q.FindField('JMT_TYPEDOS').AsString;
      sGuidPer_p    := TFMul(Ecran).Q.FindField('JMT_GUIDPER').AsString;
      sNatureCpt_p  := TFMul(Ecran).Q.FindField('JMT_NATURECPT').AsString;
      iNoOrdre_p    := TFMul(Ecran).Q.FindField('JMT_NOORDRE').AsInteger;
   end;
begin
   bContinue_l := true;

   if TFMul(Ecran).FListe.AllSelected then
      iMaxRow_l := TFMul(Ecran).Q.RecordCount   // Toutes les lignes sélectionnées
   else
      iMaxRow_l := TFMul(Ecran).FListe.NbSelected;   // Quelques lignes sélectionnées

   if iMaxRow_l = 0 then
   begin
      PGIInfo( 'Aucun mouvement sélectionné', Ecran.Caption);
      Exit;
   end;

   if TFMul(Ecran).FListe.AllSelected then
   begin
      {$IFDEF EAGLCLIENT}
      if not TFMul(Ecran).FetchLesTous then
      begin
         PGIInfo('Impossible de récupérer tous les enregistrements', Ecran.Caption);
	      exit;
      end;
      //TFMul(Ecran).Q.TQ.Last;
		{$ELSE}
      TFMul(Ecran).Q.Last;
      {$ENDIF}
   end;

   iRow_l := iMaxRow_l - 1;
   InitMove(iMaxRow_l, '');

   while (bContinue_l) and (iRow_l >= 0) do
   begin
      MoveCur(False);
      // Le suivant
      if TFMul(Ecran).FListe.AllSelected then
      begin
         SelectEnreg(sGuidPerDos_l, sGuidPer_l, iNoOrdre_l, sTypeDos_l, sNatureCpt_l);
         {$IFDEF EAGLCLIENT}
         TFMul(Ecran).Q.TQ.Prior;
			{$ELSE}
         TFMul(Ecran).Q.Prior;
			{$ENDIF}
      end
      else
      begin
         TFMul(Ecran).FListe.GotoLeBookmark(iRow_l);
         {$IFDEF EAGLCLIENT}
         TFMul(Ecran).Q.TQ.Seek(TFMul(Ecran).FListe.Row - 1) ;
         {$ENDIF}
         SelectEnreg(sGuidPerDos_l, sGuidPer_l, iNoOrdre_l, sTypeDos_l, sNatureCpt_l);
      end;

      bContinue_l := JUMvtDernierMouvement(sGuidPerDos_l, sGuidPer_l, iNoOrdre_l, sTypeDos_l, sNatureCpt_l);
      if bContinue_l then
         JUMvtDeleteMouvement(sGuidPerDos_l, sGuidPer_l, iNoOrdre_l, sTypeDos_l, sFonction_c, sNatureCpt_l, sValNominOuv_c);

      Dec(iRow_l);
   end;
   FiniMove;

   result := bContinue_l;
end;

//////////////////////////////////////////////////////////////////
Initialization
  registerclasses ( [ TOF_JUMVTTITRES ] ) ;
end.

