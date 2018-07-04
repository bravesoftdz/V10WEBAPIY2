{***********UNITE*************************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /    
Description .. :                
Mots clefs ... : 
*****************************************************************}
unit UTOMMvtTitres;
/////////////////////////////////////////////////////////////////
interface
/////////////////////////////////////////////////////////////////
uses
	{$IFDEF EAGLCLIENT}
	{$ELSE}
   db, 
	{$ENDIF}
   UTOM, UTOB, Classes, UCLASSMyMenu;
/////////////////////////////////////////////////////////////////

/////////////////////////////////////////////////////////////////
Type
   TOM_JUMVTTITRES = Class (TOM)
      procedure OnArgument (sArgument_p : String ); override;
      procedure OnNewRecord;                        override;
      procedure OnLoadRecord;                       override;
      procedure OnChangeField(F: TField);           override;
      procedure OnUpdateRecord;                     override;
      procedure OnDeleteRecord;                     override;      
      procedure OnAfterUpdateRecord;                override;
      procedure OnClose;                            override;
   private
      sGuidPerDos_c   : string;
      sGuidPer_c      : string;

      dValNominOuv_c  : double;

      sTypeDos_c      : string;
      sFonction_c     : string;
      sSens_c         : String;
      sCaption_c      : string;
      sNatureCPT_c    : string;
      sForme_c        : string;
      sCodeDos_c      : string;

      bAvecBenef_c    : boolean;
      bForceFerme_c   : boolean;
      bValideMvt_c    : boolean;

      bTitIsAct_c     : boolean;
      bBenIsAct_c     : boolean;
      bTitIsAcq_c     : boolean;
      bBenIsAcq_c     : boolean;

      FMenuSens_c     : TMyMenu;

      OBContrePartie_c : TOB;
      OBLienTit_c      : TOB;
      OBLienBen_c      : TOB;
      OBCptTit_c       : TOB;
      OBCptBen_c       : TOB;

      procedure OnClick_BINSERT(Sender : TObject);
      procedure OnClick_BDELETE(Sender : TObject);
      procedure OnClick_BVALIDER(Sender : TObject);
      procedure OnClick_BIMPRIMER(Sender : TObject);      

      procedure OnClick_BTBENEFICIAIRE(Sender : TObject);
      procedure OnClick_BTVALIDEMVT(Sender : TObject);      

      procedure OnClick_CBAFFECTSOLDE(Sender : TObject);
      procedure OnClick_CMBSENS(Sender : TObject);
      procedure OnClick_CMBNATURE(Sender : TObject);
      procedure OnClick_CMBNATUREOP(Sender : TObject);

      function  AffectMontant : double;      
      function  AffectObservations(sNom_p, sSens_p : string) : string;
      procedure AffectLesSoldes(bAffect_p : boolean; sSens_p : string);
      function  TesteSolde(iSolde_p : integer; sQui_p : string) : boolean;

      procedure InfosCpt(sGuidPer_p, sNatureCpt_p, sPref_p : string;
                         OBCpt_p : TOB);

      procedure InfosLienBen(sGuidPer_p : string);
      procedure InfosLienTit(sGuidPer_p : string);
      
      procedure InfosContrePartie(iNoOrdre_p : integer; sSens_p : string);

      procedure InfosPersonne(sGuidPer_p, sPref_p : string);
      procedure InfosRaz(sPref_p : string);
      procedure InfosRazBen;
      procedure SetOngletVisible(sOnglet_p : string);
      procedure SetOrdreVisible;
      procedure SetOrdreActif;

      function AvecBeneficiaire(sNatureOp_p : string) : boolean;

      procedure HistoLien(OBLien_p : TOB; bTit_p : boolean; sTypeModif_p : string);

      procedure OnFormKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState);
   end ;
/////////////////////////////////////////////////////////////////
implementation
/////////////////////////////////////////////////////////////////
uses
   {$ifdef EAGLCLIENT}
   MaineAGL, eFiche, UtileAGL,
   {$else}
   FE_Main, Fiche, EdtREtat,
   {$endif}
   HCtrls, sysutils, HEnt1, HTB97, Menus, DpJurOutils, CLASSAnnulien,
   dbctrls, ComCtrls, CLASSHistorisation,
   UJUROutilsMVTTitres, HMSGBOX, windows, Controls;


/////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_JUMVTTITRES.OnArgument(sArgument_p: String);
begin
   Inherited ;
   ReadTokenSt(sArgument_p);
   sGuidPerDos_c  := ReadTokenSt(sArgument_p);
   sTypeDos_c     := ReadTokenSt(sArgument_p);   
   sGuidPer_c     := ReadTokenSt(sArgument_p);
   sNatureCPT_c   := ReadTokenSt(sArgument_p);
   sCaption_c     := ReadTokenSt (sArgument_p);
   sFonction_c    := ReadTokenSt (sArgument_p);
   dValNominOuv_c := Valeur(ReadTokenSt(sArgument_p));
   sSens_c        := ReadTokenSt(sArgument_p);
   if (sSens_c = '') then
      sSens_c := '+';

   sForme_c := GetValsChamps('JURIDIQUE', 'JUR_CODEDOS, JUR_FORME', 'JUR_GUIDPERDOS = "' + sGuidPerDos_c + '"');
   sCodeDos_c := READTOKENST(sForme_c);
   sForme_c := READTOKENST(sForme_c);

   bTitIsAct_c := true;
   bBenIsAct_c := true;

   SetOngletVisible('PGENERAL');

   TToolBarButton97(GetControl('BINSERT')).OnClick := OnClick_BINSERT;
   TToolBarButton97(GetControl('BDELETE')).OnClick := OnClick_BDELETE;
   TToolBarButton97(GetControl('BVALIDER')).OnClick := OnClick_BVALIDER;
   TToolBarButton97(GetControl('BIMPRIMER')).OnClick := OnClick_BIMPRIMER;
   TToolBarButton97(GetControl('BTBENEFICIAIRE')).OnClick := OnClick_BTBENEFICIAIRE;
   TToolBarButton97(GetControl('BTVALIDEMVT')).OnClick := OnClick_BTVALIDEMVT;

   TDBCheckBox(GetControl('JMT_AFFECTSOLDE')).OnClick := OnClick_CBAFFECTSOLDE;
   THValComboBox(GetControl('JMT_SENSOPER')).OnClick := OnClick_CMBSENS;
   THValComboBox(GetControl('JMT_NATUREOP')).OnClick := OnClick_CMBNATUREOP;
   THValComboBox(GetControl('JMT_NATURETIT')).OnClick := OnClick_CMBNATURE;

   Ecran.OnKeyDown := OnFormKeyDown;

   FMenuSens_c := TMyMenu.Create(TPopUpMenu(GetControl('POPINSERT')),
                    THValComboBox(GetControl('JMT_SENSOPER')),
                    OnClick_BINSERT);

   OBContrePartie_c := TOB.Create('JUMVTTITRES', nil, -1);
   OBLienTit_c := TOB.Create('ANNULIEN', nil, -1);
   OBLienBen_c := TOB.Create('ANNULIEN', nil, -1);
   OBCptTit_c := TOB.Create('JUMVTCOMPTES', nil, -1);
   OBCptBen_c := TOB.Create('JUMVTCOMPTES', nil, -1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_JUMVTTITRES.OnNewRecord;
begin
   inherited;
   SetOngletVisible('PGENERAL');

   InfosRaz('TIT');
   SetControlText('JMT_TITNOCPT', '0');
   SetControlText('JMT_TITNOCPT_', '0');
   SetControlText('JMT_TITSOLDE', '0');

   InfosRazBen;

   SetControlEnabled('JMT_GUIDPERDOS', false);
   SetControlEnabled('JMT_GUIDPER', false);
   SetControlEnabled('JMT_NATURECPT', false);
   SetControlEnabled('JMT_NOORDRE', false);

   SetControlVisible('PBENEFICIAIRE', true);
   SetControlVisible('PBENACTIONNAIRE', true);

   SetField('JMT_GUIDPERDOS', sGuidPerDos_c);
   SetField('JMT_TYPEDOS', sTypeDos_c);
   SetField('JMT_AFFECTSOLDE', '-');

   SetField('JMT_GUIDPER', sGuidPer_c);
   SetField('JMT_NATURECPT', sNatureCPT_c);
   SetField('JMT_SENSOPER', sSens_c);
   SetField('JMT_VALNOM', dValNominOuv_c);
   SetField('JMT_DATE', Date);
   bForceFerme_c := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnLoadRecord;
var
   sGuidBen_l : string;
begin
   // Titulaire
   InfosPersonne(sGuidPer_c, 'TIT');
   InfosCpt(sGuidPer_c, sNatureCpt_c, 'TIT', OBCptTit_c);
   InfosLienTit(sGuidPer_c);

   // Beneficiciaire
   if (DS.State = dsInsert) then
   begin
      OBContrePartie_c.ClearDetail;
      OBLienBen_c.ClearDetail;
      OBCptBen_c.ClearDetail;
   end
   else if (DS.State = dsBrowse) then
   begin
      if AvecBeneficiaire(GetField('JMT_NATUREOP')) then
      begin
         InfosContrePartie(GetField('JMT_NOORDRE'), JUMvtOppositeSens(GetField('JMT_SENSOPER')));
         sGuidBen_l := GetControlText('JMT_BENGUIDPER');
         InfosPersonne(sGuidBen_l, 'BEN');
         InfosCpt(sGuidBen_l, OBContrePartie_c.Detail[0].GetValue('JMT_NATURECPT'), 'BEN', OBCptBen_c);
         InfosLienBen(sGuidBen_l);
      end;
   end;

   if GetField('JMT_SENSOPER') = '+' then
   begin
      SetControlCaption('TJMT_TITGUIDPER', 'Acquéreur');
      SetControlCaption('TJMT_BENGUIDPER', 'Cédant');
      bTitIsAcq_c := true;
      bBenIsAcq_c := false;
   end
   else
   begin
      SetControlCaption('TJMT_TITGUIDPER', 'Cédant');
      SetControlCaption('TJMT_BENGUIDPER', 'Acquéreur');
      bTitIsAcq_c := false;
      bBenIsAcq_c := true;      
   end;

   SetControlVisible('BFIRST', false);
   SetControlVisible('BPREV', false);
   SetControlVisible('BNEXT', false);
   SetControlVisible('BLAST', false);

   SetControlEnabled('BTBENEFICIAIRE', (GetField('JMT_VALIDE') = '-'));

   SetOngletVisible('PGENERAL');
   SetOrdreVisible;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnChangeField(F: TField);
begin
   SetOrdreActif;
   if (F.FieldName = 'JMT_NBTITRES') and (DS.State in [dsInsert, dsEdit]) then
   begin
      SetField('JMT_MONTANT', AffectMontant) ;
      AffectLesSoldes((GetField('JMT_AFFECTSOLDE') = 'X'), GetField('JMT_SENSOPER'));
   end;
{   else if (F.FieldName = 'JMT_VALIDE') then
   begin
      sValide_c := GetField('JMT_VALIDE');
   end;}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnUpdateRecord ;
var
   iNoOrdre_l, iBenTitre_l : integer;
   sBenGuidPer_l : string;
begin
   if (GetField('JMT_VALIDE') = '-') then
   begin
      if (GetField('JMT_DATE') = iDate1900) then
      begin
         LastError := 1;
         LastErrorMsg := 'La saisie d''une date d''opération est obligatoire.' ;
         SetFocusControl('JMT_DATE');
         exit;
      end;
      if (GetField('JMT_NBTITRES') = 0) then
      begin
         LastError := 1;
         LastErrorMsg := 'La saisie d''un nombres de titres est obligatoire.' ;
         SetFocusControl('JMT_NBTITRES');
         exit;
      end;
      if bAvecBenef_c and (GetControlText('JMT_BENGUIDPER') = '0') then
      begin
         LastError := 1;
         LastErrorMsg := 'La saisie d''un bénéficiaire est obligatoire.' ;
         SetFocusControl('JMT_BENGUIDPER');
         exit;
      end;
      if (GetField('JMT_NATUREOP') = '') then
      begin
         LastError := 1;
         LastErrorMsg := 'La saisie d''une nature d''opérations est obligatoire.' ;
         SetFocusControl('JMT_NATUREOP');
         exit;
      end;
      if (GetField('JMT_NATURETIT') = '') then
      begin
         LastError := 1;
         LastErrorMsg := 'La saisie d''une nature de titres est obligatoire.' ;
         SetFocusControl('JMT_NATURETIT');
         exit;
      end;
      if bAvecBenef_c and (GetField('JMT_DATEJOUIS') = iDate1900) then
      begin
         LastError := 1;
         LastErrorMsg := 'La saisie d''une date de jouissance est obligatoire.' ;
         SetFocusControl('JMT_DATEJOUIS');
         exit;
      end;
      if (GetField('JMT_NATUREOP') = '001') and
         ExisteSQL('SELECT * FROM JUMVTTITRES ' +
                    'WHERE JMT_GUIDPERDOS = "' + sGuidPerDos_c + '"' +
                    '  AND JMT_TYPEDOS    = "' + sTypeDos_c + '" ' +
                    '  AND JMT_GUIDPER    = "' + sGuidPer_c + '"' +
                    '  AND JMT_NATURECPT  = "' + sNatureCPT_c + '" ' +
                    '  AND JMT_NATUREOP  = "001"') then
      begin
         LastError := 1;
         LastErrorMsg := 'Il existe déjà une inscription en compte.' ;
         SetFocusControl('JMT_NATUREOP');
         exit;
      end;
      //
{      if (StrToInt(GetControlText('JMT_TITSOLDE')) < 0) then
      begin
         LastError := 1;
         LastErrorMsg := 'Le solde du titulaire ne peut être négatif.' ;
         SetFocusControl('JMT_NBTITRES');
         exit;
      end;}

      if bAvecBenef_c then
      begin
         // Création automatique de l'inscription en compte du bénéficiaire
         sBenGuidPer_l := OBLienBen_c.Detail[0].GetValue('ANL_GUIDPER');
         if not bBenIsAct_c and (GetField('JMT_SENSOPER') = '+') then
            iBenTitre_l := GetField('JMT_NBTITRES')
         else
            iBenTitre_l := OBLienBen_c.Detail[0].GetValue('ANL_TTNBPP');

         if JUMvtInscriptionEnCompte(sGuidPerDos_c, sTypeDos_c, sBenGuidPer_l,
                                     sForme_c, sNatureCPT_c, iBenTitre_l,
                                     iBenTitre_l, dValNominOuv_c,
                                     'Nouvel actionnaire : ' + GetNomAvecCivOuForme(sBenGuidPer_l)) then
         begin
            HistoLien(OBLienBen_c.Detail[0], false, 'Création auto');
         end;

{         if (StrToInt(GetControlText('JMT_BENSOLDE')) < 0) then
         begin
            LastError := 1;
            LastErrorMsg := 'Le solde du titulaire ne peut être négatif.' ;
            SetFocusControl('JMT_NBTITRES');
            exit;
         end;}

      end;

      if DS.State = dsInsert then
      begin
         iNoOrdre_l := MaxChampX('JUMVTTITRES', 'JMT_NOORDRE',
                  'JMT_GUIDPERDOS = "'  + sGuidPerDos_c + '" AND ' +
//                  'JMT_GUIDPER = '  + IntToStr(sGuidPer_c) + ' AND ' +
                  'JMT_TYPEDOS    = "' + sTypeDos_c + '"');
         SetField('JMT_NOORDRE', iNoOrdre_l + 1);
      end;
   end;
   SetField('JMT_SOLDE', StrToInt(GetControlText('JMT_TITSOLDE')));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 11/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnDeleteRecord ;
var
   iSolde_l : integer;
begin
   inherited;

   if (GetField('JMT_VALIDE') = 'X') and (GetField('JMT_NATUREOP') <> '001') then
   begin
      iSolde_l := 0;
      if OBLienTit_c.Detail.Count > 0 then
      begin
         iSolde_l := JUMvtUnaffectUnSolde((GetField('JMT_AFFECTSOLDE') = 'X'),
                                   GetField('JMT_SENSOPER'),
                                   OBLienTit_c.Detail[0].GetValue('ANL_TTNBPP'),
                                   GetField('JMT_NBTITRES'));

      // Mise à jour des liens actionnaires
         OBLienTit_c.Detail[0].PutValue('ANL_TTNBPP', iSolde_l);
         InsOrUpdDepuisTOB(OBLienTit_c.Detail[0]);
         HistoLien(OBLienTit_c.Detail[0], true, 'Suppression');
      end;

      if bAvecBenef_c then
      begin
         if OBLienBen_c.Detail.Count > 0 then
         begin
            iSolde_l := JUMvtUnaffectUnSolde((GetField('JMT_AFFECTSOLDE') = 'X'),
                                      JUMvtOppositeSens(GetField('JMT_SENSOPER')),
                                      OBLienBen_c.Detail[0].GetValue('ANL_TTNBPP'),
                                      GetField('JMT_NBTITRES'));

         // Mise à jour des liens actionnaires
            OBLienBen_c.Detail[0].PutValue('ANL_TTNBPP', iSolde_l);
            InsOrUpdDepuisTOB(OBLienBen_c.Detail[0]);
            HistoLien(OBLienBen_c.Detail[0], false, 'Suppression');
         end;
      end;
   end;

   if bAvecBenef_c then
   begin
      if OBContrePartie_c.Detail.Count > 0 then
         OBContrePartie_c.Detail[0].DeleteDB;
      InfosRazBen;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 25/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_JUMVTTITRES.OnAfterUpdateRecord;
var
   iSolde_l : integer;
begin
   if bAvecBenef_c and (OBContrePartie_c.Detail.Count > 0) then
   begin
      // Mise à jour de la contrepartie
      OBContrePartie_c.Detail[0].PutValue('JMT_NOORDRE', GetField('JMT_NOORDRE'));
      OBContrePartie_c.Detail[0].PutValue('JMT_DATE', GetField('JMT_DATE'));
      OBContrePartie_c.Detail[0].PutValue('JMT_NBTITRES', GetField('JMT_NBTITRES'));
      OBContrePartie_c.Detail[0].PutValue('JMT_SENSOPER', JUMvtOppositeSens(GetField('JMT_SENSOPER')));
      OBContrePartie_c.Detail[0].PutValue('JMT_AFFECTSOLDE', GetField('JMT_AFFECTSOLDE'));
      OBContrePartie_c.Detail[0].PutValue('JMT_SOLDE', StrToInt(GetControlText('JMT_BENSOLDE')));
      OBContrePartie_c.Detail[0].PutValue('JMT_VALNOM', GetField('JMT_VALNOM'));
      OBContrePartie_c.Detail[0].PutValue('JMT_MONTANT', GetField('JMT_MONTANT'));
      OBContrePartie_c.Detail[0].PutValue('JMT_NATUREOP', GetField('JMT_NATUREOP')) ;
      OBContrePartie_c.Detail[0].PutValue('JMT_NATURETIT', GetField('JMT_NATURETIT'));
      OBContrePartie_c.Detail[0].PutValue('JMT_DATEJOUIS', GetField('JMT_DATEJOUIS'));
      OBContrePartie_c.Detail[0].PutValue('JMT_VALIDE', GetField('JMT_VALIDE'));
      OBContrePartie_c.Detail[0].PutValue('JMT_DATEVISA', GetField('JMT_DATEVISA'));
      OBContrePartie_c.Detail[0].PutValue('JMT_COMMENT', GetControlText('JMT_BENCOMMENT'));
      if OBContrePartie_c.Detail[0].Modifie then
         OBContrePartie_c.Detail[0].InsertOrUpdateDB;
   end;

   // Mise à jour des comptes
   if OBCptTit_c.Detail.Count > 0 then
   begin
      OBCptTit_c.Detail[0].PutValue('JMC_NOCPT', GetControlText('JMT_TITNOCPT'));
      if OBCptTit_c.Detail[0].Modifie then
         OBCptTit_c.Detail[0].UpdateDB;
   end;

   if bAvecBenef_c and (OBCptBen_c.Detail.count > 0) then
   begin
      OBCptBen_c.Detail[0].PutValue('JMC_NOCPT', GetControlText('JMT_BENNOCPT'));
      if OBCptBen_c.Detail[0].Modifie then
         OBCptBen_c.Detail[0].InsertOrUpdateDB;
   end;

   if (GetField('JMT_VALIDE') = 'X') and (OBLienTit_c.Detail.Count > 0) then
   begin
      // Mise à jour des liens actionnaires
      iSolde_l := StrToInt(GetControlText('JMT_TITSOLDE'));
      OBLienTit_c.Detail[0].PutValue('ANL_TTNBPP', iSolde_l);
      OBLienTit_c.Detail[0].PutValue('ANL_ACTNAT', GetField('JMT_SENSOPER'));
      OBLienTit_c.Detail[0].PutValue('ANL_ACTDATE', GetField('JMT_DATEVISA'));
      OBLienTit_c.Detail[0].PutValue('ANL_ACTNBRE', GetField('JMT_NBTITRES'));
      OBLienTit_c.Detail[0].PutValue('ANL_ACTMONT', GetField('JMT_MONTANT'));

      if iSolde_l = 0 then
      begin
         HistoLien(OBLienTit_c.Detail[0], true, 'Suppression');
         OBLienTit_c.Detail[0].DeleteDB;
      end
      else
      begin
         InsOrUpdDepuisTOB(OBLienTit_c.Detail[0]);
         if bValideMvt_c then
            HistoLien(OBLienTit_c.Detail[0], true, 'Création')
         else
            HistoLien(OBLienTit_c.Detail[0], true, 'Modification');
		end;

      if bAvecBenef_c and bBenIsAcq_c and (OBLienBen_c.Detail.Count > 0) then
      begin
	      iSolde_l := StrToInt(GetControlText('JMT_BENSOLDE'));
         OBLienBen_c.Detail[0].PutValue('ANL_TTNBPP', iSolde_l);
         OBLienBen_c.Detail[0].PutValue('ANL_ACTNAT', JUMvtOppositeSens(GetField('JMT_SENSOPER')));
         OBLienBen_c.Detail[0].PutValue('ANL_ACTDATE', GetField('JMT_DATEVISA'));
         OBLienBen_c.Detail[0].PutValue('ANL_ACTNBRE', GetField('JMT_NBTITRES'));
         OBLienBen_c.Detail[0].PutValue('ANL_ACTMONT', GetField('JMT_MONTANT'));

         if iSolde_l = 0 then
         begin
            HistoLien(OBLienBen_c.Detail[0], true, 'Suppression');
            OBLienBen_c.Detail[0].DeleteDB;
         end
         else
         begin
            InsOrUpdDepuisTOB(OBLienBen_c.Detail[0]);
            if bValideMvt_c then
               HistoLien(OBLienBen_c.Detail[0], false, 'Création')
            else
               HistoLien(OBLienBen_c.Detail[0], false, 'Modification');

            // Lien intervenant
            CreationUnLien(OBLienBen_c.Detail[0].GetValue('ANL_GUIDPERDOS'),
                           OBLienBen_c.Detail[0].GetValue('ANL_NOORDRE'),
                           OBLienBen_c.Detail[0].GetValue('ANL_GUIDPER'),
                           OBLienBen_c.Detail[0].GetValue('ANL_TYPEDOS'),
                           'INT',
                           OBLienBen_c.Detail[0].GetValue('ANL_FORME'),
                           OBLienBen_c.Detail[0].GetValue('ANL_CODEDOS'));
			end;
      end;
      bValideMvt_c := false;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_JUMVTTITRES.OnClose;
begin
   inherited;
   if FMenuSens_c <> nil then
   begin
      FreeAndNil(FMenuSens_c);
   end;
   if OBContrePartie_c <> nil then
      OBContrePartie_c.Free;
   if OBLienTit_c <> nil then
      OBLienTit_c.Free;
   if OBLienBen_c <> nil then
      OBLienBen_c.Free;
   if OBCptTit_c <> nil then
      OBCptTit_c.Free;
   if OBCptBen_c <> nil then
      OBCptBen_c.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnClick_BINSERT(Sender : TObject);
begin
   if FMenuSens_c.GetHint = '' then exit;
   sSens_c := FMenuSens_c.GetHint;
   TFFiche(Ecran).BinsertClick(Sender);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 23/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnClick_BVALIDER(Sender : TObject);
begin
   if GetActiveTabSheet('PAGES') = TTabSheet(GetControl('PGENERAL')) then
   begin
      TFFiche(Ecran).BValiderClick(Sender);
      if LastError = 1 then exit;
      if bForceFerme_c then
      begin
         TFFiche(Ecran).Retour := GetControlText('JMT_NOORDRE');

         bForceFerme_c := false;
         TFFiche(Ecran).BFermeClick(Sender);
      end;
      SetOrdreVisible;
      SetOrdreActif;
   end
   else
   begin
      TFFiche(Ecran).BFermeClick(Sender);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 14/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnClick_BDELETE(Sender : TObject);
var
   sMessage_l : string;
begin
   sMessage_l := 'Attention : vous supprimez un mouvement de titres validé.' + #13#10 +
                 'Le nombre de titres du titulaire ';
   if bAvecBenef_c then
   begin
      sMessage_l := sMessage_l + 'et du bénéficiaire ';
      sMessage_l := sMessage_l + 'vont être rétablis.';
   end
   else
      sMessage_l := sMessage_l + 'va être rétabli.';
   sMessage_l := sMessage_l + 'Confirmez vous?';


   if JUMvtDernierMouvement(sGuidPerDos_c, sGuidPer_c, GetField('JMT_NOORDRE'),
                            sTypeDos_c, sNatureCPT_c) then
   begin
      TFFiche(Ecran).HM.Mess[1] := '1;?caption?;' + sMessage_l + ';Q;YNC;N;C;';
      TFFiche(Ecran).BDeleteClick(Sender);
   end;
//   TFFiche(Ecran).BFermeClick(Sender);   
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 01/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnClick_BIMPRIMER(Sender : TObject);
var
   sCle_l : string;
begin
   sCle_l := 'TIT.JMT_GUIDPERDOS = "' + sGuidPerDos_c + '" AND ' +
             'TIT.JMT_TYPEDOS    = "' + sTypeDos_c + '" AND ' +
             'TIT.JMT_GUIDPER    = "' + sGuidPer_c + '" AND ' +
             'TIT.JMT_NATURECPT  = "' + sNatureCPT_c + '" AND ' +
             'TIT.JMT_NOORDRE    = ' + IntToStr(GetField('JMT_NOORDRE'));  // Compte principal uniquement

   LanceEtat('E', 'JID', 'JMT', true, false, false, nil, sCle_l,
             'Ordre de mouvement', false, 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 23/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnClick_BTVALIDEMVT(Sender : TObject);
begin
   if (GetField('JMT_VALIDE') = 'X') then
   begin
      SetOngletVisible('PVALIDATION');
   end
   else
   begin
      if DS.State = dsBrowse then
         DS.Edit;

      SetField('JMT_DATEVISA', Date);
      SetField('JMT_VALIDE', 'X');
      bValideMvt_c := true;

      TFFiche(Ecran).BValiderClick(Sender);
      if LastError = 1 then exit;
      SetOngletVisible('PVALIDATION');
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/08/2004
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnClick_BTBENEFICIAIRE(Sender : TObject);
var
   sValeur_l, sNatureCpt_l, sGuidBen_l : string;
begin
   sValeur_l := AGLLanceFiche('YY', 'ANNUAIRE_SEL', '', '', '');
   if (sValeur_l = '') or (sValeur_l = '0') or
      (GetControlText('JMT_BENGUIDPER') = sValeur_l) then
      exit;

   if DS.State = dsBrowse then
      DS.Edit;

   // quelle nature de compte?
   sNatureCpt_l := sNatureCpt_c; // Par défaut, à voir par la suite
   
   // Nouveau bénéficiaire
   SetControlText('JMT_BENGUIDPER', sValeur_l);
   sGuidBen_l := sValeur_l;
   InfosPersonne(sGuidBen_l, 'BEN');
   InfosCpt(sGuidBen_l, sNatureCpt_c, 'BEN', OBCptBen_c);
   InfosLienBen(sGuidBen_l);

   // Nouvelle contre-partie
   TOB.Create('JUMVTTITRES', OBContrePartie_c, -1);
   OBContrePartie_c.Detail[0].PutValue('JMT_GUIDPERDOS', sGuidPerDos_c);
   OBContrePartie_c.Detail[0].PutValue('JMT_GUIDPER', sGuidBen_l);
   OBContrePartie_c.Detail[0].PutValue('JMT_TYPEDOS', sTypeDos_c);
   OBContrePartie_c.Detail[0].PutValue('JMT_NATURECPT', sNatureCpt_l);

   // Pour mise à jour du commentaire
   OnClick_CMBNATURE(Sender);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 20/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnClick_CBAFFECTSOLDE(Sender : TObject);
begin
   if (DS.State = dsbrowse) then
      Exit;
   if (GetControlText('JMT_AFFECTSOLDE') = '') or
      (GetControlText('JMT_SENSOPER') = '') or
      (GetControlText('JMT_TITSOLDE') = '') then
      Exit;
   if (DS.State in [dsInsert,dsEdit]) then
   begin
      AffectLesSoldes((GetControlText('JMT_AFFECTSOLDE') = 'X'), GetField('JMT_SENSOPER'));
      SetOrdreActif;
   end;
end;                 
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 20/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnClick_CMBSENS(Sender : TObject);
begin
   SetField('JMT_MONTANT', AffectMontant) ;
   AffectLesSoldes((GetField('JMT_AFFECTSOLDE') = 'X') and (GetField('JMT_VALIDE') = '-'),
                    GetControlText('JMT_SENSOPER'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 20/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnClick_CMBNATURE(Sender : TObject);
var
   sLesValeurs_l : string;
begin
   sLesValeurs_l := GetControlText('JMT_BENTYPECIV') + ' ' + GetControlText('JMT_BENNOMPER');
   SetField('JMT_COMMENT', AffectObservations(sLesValeurs_l, GetField('JMT_SENSOPER')));
   sLesValeurs_l := GetControlText('JMT_TITTYPECIV') + ' ' + GetControlText('JMT_TITNOMPER');
   SetControlText('JMT_BENCOMMENT', AffectObservations(sLesValeurs_l, JUMvtOppositeSens(GetField('JMT_SENSOPER'))));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnClick_CMBNATUREOP(Sender : TObject);
begin
   OnClick_CMBNATURE(Sender);
   AvecBeneficiaire(GetControlText('JMT_NATUREOP'));
   // Inscription en compte créée manuellement pour le titulaire
   if GetControlText('JMT_NATUREOP') = '001' then
      SetField('JMT_NBTITRES', OBLienTit_c.Detail[0].GetValue('ANL_TTNBPP'));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_JUMVTTITRES.AvecBeneficiaire(sNatureOp_p : string) : boolean;
var
   sValeur_l : string;
begin
   sValeur_l := GetValChamp('YNATUREOP', 'YNO_AVECBENEF', 'YNO_CODE = "' + sNatureOp_p + '"');
   bAvecBenef_c := (sValeur_l = 'X');
   SetControlVisible('PBENEFICIAIRE', bAvecBenef_c);
   SetControlVisible('PBENACTIONNAIRE', bAvecBenef_c);
   
   if not bAvecBenef_c then
      InfosRazBen;

   result := bAvecBenef_c;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 25/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.SetOngletVisible(sOnglet_p : string);
var
   bGeneral_l : boolean;
begin
   bGeneral_l := (sOnglet_p = 'PGENERAL');
   TTabSheet(GetControl('PVALIDATION')).Visible := not bGeneral_l;
   TTabSheet(GetControl('PGENERAL')).Visible := bGeneral_l;
   SetActiveTabSheet(sOnglet_p);
   if sOnglet_p = 'PGENERAL' then
      Ecran.Caption := sCaption_c + ' - Ligne compte actionnaire'
   else
      Ecran.Caption := sCaption_c + ' - Ordre de mouvement';

   SetControlVisible('BVALIDER', bGeneral_l);
//   SetControlVisible('BFERME', bGeneral_l);
   SetControlVisible('BDEFAIRE', bGeneral_l);
   SetControlVisible('BINSERT', bGeneral_l);
   SetControlVisible('BDELETE', bGeneral_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 24/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.SetOrdreVisible;
var
   bValide_l : boolean;
begin
   bValide_l := (GetField('JMT_VALIDE') = 'X');
   SetControlVisible('BTVALIDEMVT', (DS.State <> dsInsert));
   SetControlVisible('JMT_NOORDRE', (DS.State <> dsInsert));
   SetControlVisible('TJMT_NOORDRE', (DS.State <> dsInsert));

   if not bValide_l then
      TToolbarButton97(GetControl('BTVALIDEMVT')).Caption := 'Valider l''ordre de mouvement'
   else
      TToolbarButton97(GetControl('BTVALIDEMVT')).Caption := 'Ordre de mouvement';

   SetControlEnabled('JMT_DATE', not bValide_l);
   SetControlEnabled('JMT_NBTITRES', not bValide_l);
   SetControlEnabled('JMT_VALNOM', not bValide_l);
   //SetControlEnabled('JMT_SENSOPER', not bValide_l);
   SetControlEnabled('JMT_MONTANT', not bValide_l);
   SetControlEnabled('JMT_AFFECTSOLDE', not bValide_l);
   SetControlEnabled('JMT_TITSOLDE', not bValide_l);
   SetControlEnabled('JMT_NATUREOP', not bValide_l);
   SetControlEnabled('JMT_NATURETIT', not bValide_l);
   SetControlEnabled('JMT_COMMENT', not bValide_l);
   //SetControlEnabled('JMT_BENTYPECIV', not bValide_l);
   //SetControlEnabled('JMT_BENNOMPER', not bValide_l);
   SetControlEnabled('JMT_BENCOMMENT', not bValide_l);
   SetControlEnabled('JMT_DATEJOUIS', not bValide_l);

   SetControlEnabled('BINSERT', bValide_l);
//   SetControlEnabled('BDELETE', bValide_l);
   SetControlEnabled('BDEFAIRE', not bValide_l);
   SetControlVisible('BVALIDER', not bValide_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 24/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.SetOrdreActif;
var
   bBtOK_l : boolean;
begin
   bBtOK_l := (DS.State = dsBrowse);// and (GetControlText('JMT_VALIDE') <> 'X');
   SetControlEnabled('BTVALIDEMVT', bBtOK_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 20/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_JUMVTTITRES.AffectMontant : double;
var
   dMontantTotal_l, dValNominOuv_l : double;
begin
   dValNominOuv_l := GetField ('JMT_VALNOM');
   if (dValNominOuv_l <> 0.00) then
      dMontantTotal_l := GetField('JMT_NBTITRES') * dValNominOuv_l
   else
      dMontantTotal_l := 0;
   result := dMontantTotal_l;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOM_JUMVTTITRES.AffectObservations(sNom_p, sSens_p : string) : string;
var
   sValeur_l, sValeur1_l, sValeur2_l, sArticle_l : string;
begin
   sValeur1_l := THValComboBox(GetControl('JMT_NATUREOP')).Text;
   sValeur2_l := THValComboBox(GetControl('JMT_NATURETIT')).Text;

   if (sNom_p = '') or (sValeur1_l = '') or (sValeur2_l = '') then
   begin
      result := '';
      exit;
   end;

   if sValeur2_l[1] in ['A','E','U','I','O','Y','a','e','u','i','o','y'] then
      sArticle_l := 'd'''
   else
      sArticle_l := 'de';
   if sSens_p = '-' then
      sSens_p := 'vers le compte'
   else
      sSens_p := 'à partir du compte';
   sValeur_l := sValeur1_l + ' ' + sArticle_l + ' ' + sValeur2_l;
   if bAvecBenef_c then
      sValeur_l := sValeur_l + ' ' + sSens_p + ' de ' + sNom_p;
   result := sValeur_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 25/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.AffectLesSoldes(bAffect_p : boolean; sSens_p : string);
var
   iTitSolde_l, iBenSolde_l : integer;
begin
   if GetField('JMT_VALIDE') = 'X' then
   begin
      iTitSolde_l := GetField('JMT_SOLDE');
   end
   else
   begin
      iTitSolde_l := 0;
      if not bTitIsAct_c and (GetField('JMT_SENSOPER') = '-') then
         iTitSolde_l := GetField('JMT_NBTITRES')
      else if OBLienTit_c.Detail.Count > 0 then
         iTitSolde_l := OBLienTit_c.Detail[0].GetValue('ANL_TTNBPP');

      iTitSolde_l := JUMvtAffectUnSolde(bAffect_p, sSens_p,
                                   GetControlText('JMT_NATUREOP'),
                                   iTitSolde_l, GetField('JMT_NBTITRES'));
   end;

   if GetControlText('JMT_TITSOLDE') <> IntToStr(iTitSolde_l) then
      SetControlText('JMT_TITSOLDE', IntToStr(iTitSolde_l));

   if not TesteSolde(iTitSolde_l, 'titulaire') then
      exit;

//   if bAvecBenef_c then
   if bAvecBenef_c and (GetControlText('JMT_BENGUIDPER') <> '0') then
   begin
      if GetField('JMT_VALIDE') = 'X' then
         iBenSolde_l := OBContrePartie_c.Detail[0].GetValue('JMT_SOLDE')
      else
      begin
         iBenSolde_l := 0;
         if not bBenIsAct_c and (GetField('JMT_SENSOPER') = '+') then
            iBenSolde_l := GetField('JMT_NBTITRES')
         else if OBLienBen_c.Detail.Count > 0 then
            iBenSolde_l := OBLienBen_c.Detail[0].GetValue('ANL_TTNBPP');

         iBenSolde_l := JUMvtAffectUnSolde(bAffect_p, JUMvtOppositeSens(sSens_p),
                                      GetControlText('JMT_NATUREOP'),
                                      iBenSolde_l, GetField('JMT_NBTITRES'));
      end;

      if GetControlText('JMT_BENSOLDE') <> IntToStr(iBenSolde_l) then
         SetControlText('JMT_BENSOLDE', IntToStr(iBenSolde_l));

      if not TesteSolde(iBenSolde_l, 'bénéficiaire') then
         exit;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 15/02/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_JUMVTTITRES.TesteSolde(iSolde_p : integer; sQui_p : string) : boolean;
var
   bPositif_l : boolean;
begin
//   if DS.State = dsBrowse then exit;
   
   bPositif_l := true;
   if (iSolde_p < 0) then
   begin
      PGIError('Le solde du ' + sQui_p + ' ne peut être négatif.', Ecran.Caption);
      SetFocusControl('JMT_NBTITRES');
      bPositif_l := false;
   end;
//   SetControlEnabled('BINSERT', bPositif_l);
//   SetControlEnabled('BDELETE', bPositif_l);
//   SetControlEnabled('BDEFAIRE', bPositif_l);
   SetControlEnabled('BVALIDER', bPositif_l);

   result := bPositif_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 30/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.InfosCpt(sGuidPer_p, sNatureCpt_p, sPref_p : string;
                                    OBCpt_p : TOB);
var
   sCpt_l : string;                                    
begin
   if not JUMvtCompteVerif(sGuidPerDos_c, sGuidPer_p, sNatureCPT_p, OBCpt_p) then
      JUMvtCompteMAJ(sGuidPerDos_c, sGuidPer_p, sNatureCPT_p, OBCpt_p);

   sCpt_l := OBCpt_p.Detail[0].GetValue('JMC_NOCPT');
   SetControlText('JMT_' + sPref_p + 'NOCPT', sCpt_l);
   SetControlText('JMT_' + sPref_p + 'NOCPT_', sCpt_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 30/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.InfosLienTit(sGuidPer_p : string);
var
   iTitSolde_l : integer;

begin
   bTitIsAct_c := JUMvtInfosLien(OBLienTit_c, sGuidPerDos_c, sGuidPer_p, 1, sTypeDos_c, sFonction_c);
   if not bTitIsAct_c then
   begin
      TOB.Create('ANNULIEN', OBLienTit_c, -1);
      OBLienTit_c.Detail[0].PutValue('ANL_GUIDPERDOS', sGuidPerDos_c);
      OBLienTit_c.Detail[0].PutValue('ANL_GUIDPER', sGuidPer_p);
      OBLienTit_c.Detail[0].PutValue('ANL_TYPEDOS', sTypeDos_c);
      OBLienTit_c.Detail[0].PutValue('ANL_NOORDRE', 1);
      OBLienTit_c.Detail[0].PutValue('ANL_FONCTION', sFonction_c);
      OBLienTit_c.Detail[0].PutValue('ANL_FORME', sForme_c);
      OBLienTit_c.Detail[0].PutValue('ANL_CODEDOS', sCodeDos_c);
   end;

   if GetField('JMT_VALIDE') = 'X' then
   begin
      iTitSolde_l := GetField('JMT_SOLDE');
   end
   else
   begin
      iTitSolde_l := 0;
      if not bTitIsAct_c and (GetField('JMT_SENSOPER') = '-') then
         iTitSolde_l := GetField('JMT_NBTITRES')
      else if OBLienTit_c.Detail.Count > 0 then
         iTitSolde_l := OBLienTit_c.Detail[0].GetValue('ANL_TTNBPP');

      iTitSolde_l := JUMvtAffectUnSolde((GetField('JMT_AFFECTSOLDE') = 'X'),
                                   GetField('JMT_SENSOPER'),
                                   GetField('JMT_NATUREOP'),
                                   iTitSolde_l,
                                   GetField('JMT_NBTITRES'));
   end;

   if GetControlText('JMT_TITSOLDE') <> IntToStr(iTitSolde_l) then
      SetControlText('JMT_TITSOLDE', IntToStr(iTitSolde_l));

   if not TesteSolde(iTitSolde_l, 'titulaire') then
      exit;

   if (DS.State <> dsBrowse) then
      SetField('JMT_SOLDE', iTitSolde_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 30/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.InfosLienBen(sGuidPer_p : string);
var
   iBenSolde_l : integer;
begin
   bBenIsAct_c := JUMvtInfosLien(OBLienBen_c, sGuidPerDos_c, sGuidPer_p, 1, sTypeDos_c, sFonction_c);
   if not bBenIsAct_c then
   begin
      TOB.Create('ANNULIEN', OBLienBen_c, -1);
      OBLienBen_c.Detail[0].PutValue('ANL_GUIDPERDOS', sGuidPerDos_c);
      OBLienBen_c.Detail[0].PutValue('ANL_GUIDPER', sGuidPer_p);
      OBLienBen_c.Detail[0].PutValue('ANL_TYPEDOS', sTypeDos_c);
      OBLienBen_c.Detail[0].PutValue('ANL_NOORDRE', 1);
      OBLienBen_c.Detail[0].PutValue('ANL_FONCTION', sFonction_c);
      OBLienBen_c.Detail[0].PutValue('ANL_FORME', sForme_c);
      OBLienBen_c.Detail[0].PutValue('ANL_CODEDOS', sCodeDos_c);
   end;

   if GetField('JMT_VALIDE') = 'X' then
   begin
      iBenSolde_l := GetField('JMT_SOLDE');
   end
   else
   begin
      iBenSolde_l := 0;
      if not bBenIsAct_c and (GetField('JMT_SENSOPER') = '+') then
         iBenSolde_l := GetField('JMT_NBTITRES')
      else if OBLienBen_c.Detail.Count > 0 then
         iBenSolde_l := OBLienBen_c.Detail[0].GetValue('ANL_TTNBPP');

      iBenSolde_l := JUMvtAffectUnSolde((GetField('JMT_AFFECTSOLDE') = 'X'),
                                   JUMvtOppositeSens(GetField('JMT_SENSOPER')),
                                   GetField('JMT_NATUREOP'),
                                   iBenSolde_l, GetField('JMT_NBTITRES'));
   end;

   if GetControlText('JMT_BENSOLDE') <> IntToStr(iBenSolde_l) then
      SetControlText('JMT_BENSOLDE', IntToStr(iBenSolde_l));
   TesteSolde(iBenSolde_l, 'bénéficiaire');
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 24/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.InfosContrePartie(iNoOrdre_p : integer; sSens_p : string);
var
   sGuidPer_l : string;
begin
   OBContrePartie_c.ClearDetail;
   OBContrePartie_c.LoadDetailDBFromSQL('JUMVTTITRES',
                  'SELECT * FROM JUMVTTITRES ' +
                  'WHERE JMT_GUIDPERDOS = "'  + sGuidPerDos_c + '"' +
                  '  AND JMT_TYPEDOS    = "' + sTypeDos_c + '" ' +
                  '  AND JMT_NOORDRE    = ' + IntToStr(iNoOrdre_p) +
                  '  AND JMT_SENSOPER       = "' + sSens_p + '"');

   if OBContrePartie_c.Detail.Count > 0 then
   begin
      sGuidPer_l := OBContrePartie_c.Detail[0].GetValue('JMT_GUIDPER');
      SetControlText('JMT_BENGUIDPER', sGuidPer_l);
//      sNatureCpt_l := OBContrePartie_c.Detail[0].GetValue('JMT_NATURECPT');
//      SetControlText('JMT_BENNATURECPT', sNatureCpt_l);
      SetControlText('JMT_BENSOLDE', OBContrePartie_c.Detail[0].GetValue('JMT_SOLDE'));
      SetControlText('JMT_BENCOMMENT', OBContrePartie_c.Detail[0].GetValue('JMT_COMMENT'));
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 23/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.InfosPersonne(sGuidPer_p, sPref_p : string);
var
   sTypeCiv_l, sNomPer_l, sRue_l, sCp_l, sVille_l : string;
   OBAnnuaire_l : TOB;
begin
   OBAnnuaire_l := TOB.Create('ANNUAIRE', nil, -1);
   if OBAnnuaire_l.SelectDB('"' + sGuidPer_p + '"', nil) then
   begin
      if OBAnnuaire_l.GetValue('ANN_PPPM') = 'PM' then
         sTypeCiv_l := OBAnnuaire_l.GetValue('ANN_FORME')
      else
         sTypeCiv_l := OBAnnuaire_l.GetValue('ANN_CVA');
      sNomPer_l := OBAnnuaire_l.GetValue('ANN_NOMPER');
      sRue_l := OBAnnuaire_l.GetValue('ANN_ALRUE1');
      sCp_l := OBAnnuaire_l.GetValue('ANN_ALCP');
      sVille_l := OBAnnuaire_l.GetValue('ANN_ALVILLE');
   end;
   OBAnnuaire_l.Free ;

   SetControlText('JMT_' + sPref_p + 'TYPECIV', sTypeCiv_l);
   SetControlText('JMT_' + sPref_p + 'NOMPER', sNomPer_l);
   SetControlText('JMT_' + sPref_p + 'ACTIONNAIRE', sTypeCiv_l + ' ' + sNomPer_l);
   SetControlText('JMT_' + sPref_p + 'ADRESSE', sRue_l);
   SetControlText('JMT_' + sPref_p + 'CP', sCp_l);
   SetControlText('JMT_' + sPref_p + 'VILLE', sVille_l);
   SetControlText('JMT_' + sPref_p + 'ADMIN', '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 23/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.InfosRaz(sPref_p : string);
begin
   SetControlText('JMT_' + sPref_p + 'TYPECIV', '');
   SetControlText('JMT_' + sPref_p + 'NOMPER', '');
   SetControlText('JMT_' + sPref_p + 'ACTIONNAIRE', '');
   SetControlText('JMT_' + sPref_p + 'ADRESSE', '');
   SetControlText('JMT_' + sPref_p + 'VILLE', '');
   SetControlText('JMT_' + sPref_p + 'ADMIN', '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 10/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.InfosRazBen;
begin
   SetControlText('JMT_BENGUIDPER', '0');

   SetControlText('JMT_BENNOCPT', '0');
   SetControlText('JMT_BENNOCPT_', '0');
   SetControlText('JMT_BENSOLDE', '0');
   SetControlText('JMT_BENCOMMENT', '');
   InfosRaz('BEN');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 07/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.HistoLien(OBLien_p : TOB; bTit_p : boolean; sTypeModif_p : string);
var
   sCle_l, sComplement_l : string;
begin
   if bTit_p then
   begin
      sComplement_l := 'HNL_NOMPROPTITRE='   + GetControlText('JMT_TITACTIONNAIRE') + ';' +
                       'HNL_CPPROPTITRE='    + GetControlText('JMT_TITCP') + ';' +
                       'HNL_VILLEPROPTITRE=' + GetControlText('JMT_TITVILLE') + ';' +
                       'HNL_NOMACQUERTITRE=' + GetControlText('JMT_BENACTIONNAIRE');
   end
   else
   begin
      sComplement_l := 'HNL_NOMPROPTITRE='   + GetControlText('JMT_BENACTIONNAIRE') + ';' +
                       'HNL_CPPROPTITRE='    + GetControlText('JMT_BENCP') + ';' +
                       'HNL_VILLEPROPTITRE=' + GetControlText('JMT_BENVILLE') + ';' +
                       'HNL_NOMACQUERTITRE=' + GetControlText('JMT_TITACTIONNAIRE');
   end;
   sComplement_l := sComplement_l +
                    'HNL_DATEMVTTITRE='   + GetControlText('JMT_DATEVISA') + ';' +
                    'HNL_VALEURTITRE='    + GetControlText('JMT_VALNOM') + ';' +
                    'HNL_NATURETITRE='    + GetControlText('JMT_NATUREOP');

   sCle_l := ' HNL_GUIDPERDOS = "' + OBLien_p.GetValue('ANL_GUIDPERDOS') + '"' +
             ' AND HNL_TYPEDOS    = "' + OBLien_p.GetValue('ANL_TYPEDOS') + '"' +
             ' AND HNL_GUIDPER    = "' + OBLien_p.GetValue('ANL_GUIDPER') + '"' +
             ' AND HNL_FONCTION   = "' + OBLien_p.GetValue('ANL_FONCTION') + '"';

   Historisation(OBLien_p, 'ANNULIEN', 'HISTOANNULIEN',
                 sCle_l, ' HNL_NOORDRE DESC', 'HNL_NOORDRE',
                 sTypeModif_p, sComplement_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 01/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUMVTTITRES.OnFormKeyDown(MySender_p : TObject; var MyKey_p : Word; MyShift_p : TShiftState);
var
   Point_l : TPoint;
begin
   if (MyKey_p = VK_F5) then  // Sélection ouvrir
   begin
      MyKey_p := 0;
      if (TToolBarButton97(GetControl('BTBENEFICIAIRE')).Visible) and
         (TToolBarButton97(GetControl('BTBENEFICIAIRE')).Enabled) then
         TToolBarButton97(GetControl('BTBENEFICIAIRE')).OnClick(nil);
   end
   else if (MyKey_p = vk_nouveau)  and (MyShift_p = [ssCtrl]) then  // Ctrl + N : nouveau
   begin
      MyKey_p := 0;
      if (TToolBarButton97(GetControl('BINSERT')).Visible) and
         (TToolBarButton97(GetControl('BINSERT')).Enabled)then
      begin
         Point_l.x := TToolBarButton97(GetControl('BINSERT')).Left;
         Point_l.y := TToolBarButton97(GetControl('BINSERT')).Top +
                      TToolBarButton97(GetControl('BINSERT')).Height;
         Point_l := TToolBarButton97(GetControl('BINSERT')).Parent.ClientToScreen(Point_l);
         TToolBarButton97(GetControl('BINSERT')).DropdownMenu.Popup(Point_l.X, Point_l.Y);
      end;
   end
   else
      TFFiche(Ecran).FormKeyDown(MySender_p, MyKey_p, MyShift_p);
end;

/////////////////////////////////////////////////////////////////
Initialization
  registerclasses([TOM_JUMVTTITRES]) ;
end.
