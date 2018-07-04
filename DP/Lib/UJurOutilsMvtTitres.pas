{***********UNITE*************************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 03/05/2004
Modifié le ... : 08/10/2004
Description .. : 
Mots clefs ... : 
*****************************************************************}
unit UJurOutilsMvtTitres;
/////////////////////////////////////////////////////////////////
interface
/////////////////////////////////////////////////////////////////
uses
   {$IFNDEF EAGLCLIENT}
   FE_Main,
   {$else}
   MaineAGL,
   {$ENDIF}
   sysutils, UTOB, DpJurOutils, HCTRLS, Hmsgbox, controls,
   Hent1, CLASSHISTORISATION, CLASSAnnulien;

/////////////////////////////////////////////////////////////////
function  JUMvtCompteVerif(sGuidPerdos_p, sGuidPer_p, sNatureCpt_p : string;
                           OBMvtCpt_p : TOB) : boolean;
procedure JUMvtCompteMAJ(sGuidPerdos_p, sGuidPer_p, sNatureCpt_p : string;
                              var OBMvtCpt_p : TOB);
procedure JUMvtCompteCreation(sGuidPerdos_p, sGuidPer_p, sNatureCpt_p : string;
                              var OBMvtCpt_p : TOB);
function JUMvtAffectUnSolde(bAffect_p : boolean; sSens_p, sNatureOp_p : string;
                       iSolde_p, iNbTitres_p : integer) : integer;

function JUMvtUnaffectUnSolde(bAffect_p : boolean; sSens_p : string;
                         iSolde_p, iNbTitres_p : integer) : integer;

function JUMvtInscriptionEnCompte(sGuidPerDos_p, sTypeDos_p, sGuidPer_p, sFormeJur_p, sNatureCPT_p : string;
                     iTTNBPP_p, iSolde_p : integer; dValNominOuv_p : Double;
                     sTitre_p : string) : boolean;

function JUMvtMessageSup(bAvecBenef_p : boolean) : string;

function JUMvtOppositeSens(sSens_p : string) : string;

function JUMvtInfosLien(OBLien_p : TOB;
                        sGuidPerDos_p, sGuidPer_p : string; iNoOrdre_p : integer;
                        sTypeDos_p, sFonction_p : string) : boolean;

procedure JUMvtHistoLien(OBLien_p : TOB; sTypeModif_p, sNatureOp_p,
                         sGuidTit_p, sGuidBen_p, sGuidPerDos_p : string; iNoOrdre_p : integer;
                         sTypeDos_p, sFonction_p, sValNominOuv_p : string);

procedure JUMvtDeleteMouvement(sGuidPerDos_p, sGuidPer_p : string; iNoOrdre_p : integer;
                          sTypeDos_p, sFonction_p, sNatureCpt_p, sValNominOuv_p : string);

function JUMvtDernierMouvement(sGuidPerDos_p, sGuidPer_p : string; iNoOrdre_p : integer;
                          sTypeDos_p, sNatureCpt_p : string) : boolean;

function AGLDeleteMouvement(aovParams_p : array of variant; iNb_p : integer) : integer;

/////////////////////////////////////////////////////////////////
implementation
/////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 08/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function  JUMvtCompteVerif(sGuidPerdos_p, sGuidPer_p, sNatureCpt_p : string;
                           OBMvtCpt_p : TOB) : boolean;
begin
   OBMvtCpt_p.ClearDetail;
   OBMvtCpt_p.LoadDetailDBFromSQL('JUMVTCOMPTES',
                    'SELECT * FROM JUMVTCOMPTES ' +
                    'WHERE JMC_GUIDPERDOS = "' + sGuidPerdos_p + '" ' +
                    '  AND JMC_GUIDPER    = "' + sGuidPer_p + '" ' +
                    '  AND JMC_NATURECPT  = "' + sNatureCpt_p + '"');

   result := (OBMvtCpt_p.Detail.Count > 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 08/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure JUMvtCompteMAJ(sGuidPerdos_p, sGuidPer_p, sNatureCpt_p : string;
                         var OBMvtCpt_p : TOB);
begin
   with TOB.Create('JUMVTCOMPTES', OBMvtCpt_p, -1) do
   begin
      PutValue('JMC_GUIDPERDOS', sGuidPerDos_p);
      PutValue('JMC_GUIDPER', sGuidPer_p);
      PutValue('JMC_NATURECPT', sNatureCPT_p);
      PutValue('JMC_NOCPT', '1');
      PutValue('JMC_INTITULE', GetNomAvecCivOuForme(sGuidPer_p));
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 08/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure JUMvtCompteCreation(sGuidPerdos_p, sGuidPer_p, sNatureCpt_p : string;
                              var OBMvtCpt_p : TOB);
begin
   JUMvtCompteMAJ(sGuidPerdos_p, sGuidPer_p, sNatureCpt_p, OBMvtCpt_p);
   OBMvtCpt_p.Detail[0].InsertDB(nil);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 11/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function JUMvtAffectUnSolde(bAffect_p : boolean; sSens_p, sNatureOp_p : string;
                            iSolde_p, iNbTitres_p : integer) : integer;
begin
   result := iSolde_p;
   if not bAffect_p then
      exit;

   // Inscription en compte
   if sNatureOp_p = '001' then
   begin
      result := iNbTitres_p;   
   end
   else
   begin
      if sSens_p = '+' then
         result := iSolde_p + iNbTitres_p
      else
         result := iSolde_p - iNbTitres_p;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 11/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function JUMvtUnaffectUnSolde(bAffect_p : boolean; sSens_p : string;
                              iSolde_p, iNbTitres_p : integer) : integer;
begin
   result := iSolde_p;
   if not bAffect_p then
      exit;

   if sSens_p = '+' then
      result := iSolde_p - iNbTitres_p
   else
      result := iSolde_p + iNbTitres_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function JUMvtInscriptionEnCompte(sGuidPerDos_p, sTypeDos_p, sGuidPer_p, sFormeJur_p, sNatureCPT_p : string;
                                  iTTNBPP_p, iSolde_p : integer; dValNominOuv_p : Double;
                                  sTitre_p : string) : boolean;
var
   OBMvtTitres_l : TOB;
   sDate_l : string;
   iNoOrdre_l : integer;
   dtDate_l : TDateTime;
begin
   result := false;
   OBMvtTitres_l := TOB.Create('JUMVTTITRES', nil, -1);
   OBMvtTitres_l.LoadDetailDBFromSQL('JUMVTTITRES',
                    'SELECT * FROM JUMVTTITRES ' +
                    'WHERE JMT_GUIDPERDOS = "' + sGuidPerDos_p + '" ' +
                    '  AND JMT_TYPEDOS    = "' + sTypeDos_p + '" ' +
                    '  AND JMT_GUIDPER    = "' + sGuidPer_p + '" ' +
                    '  AND JMT_NATURECPT  = "' + sNatureCPT_p + '" ' +
                    '  AND JMT_NATUREOP  = "001"');

   if (OBMvtTitres_l.Detail.Count = 0) then
   begin
      if sTitre_p = '' then
         sTitre_p := 'Mouvement des titres';
      if PGIAsk('Voulez-vous créer automatiquement l''inscription en compte?',
                sTitre_p) = mrYes then
      begin
         if sFormeJur_p = 'SA' then
            dtDate_l := iDate1900
         else
            dtDate_l := Date;

         sDate_l := AGLLanceFiche('JUR', 'CHOIXRECUP', '', '', 'JMT;' + DateToStr(dtDate_l));
         if (sDate_l <> '') and (StrToDate(sDate_l) <> iDate1900) then
         begin
            with TOB.Create('JUMVTTITRES', OBMvtTitres_l, -1) do
            begin
               iNoOrdre_l := MaxChampX('JUMVTTITRES', 'JMT_NOORDRE',
                                 'JMT_GUIDPERDOS = "'  + sGuidPerDos_p + '" AND ' +
                                 'JMT_TYPEDOS    = "' + sTypeDos_p + '"');

               PutValue('JMT_GUIDPERDOS', sGuidPerDos_p);
               PutValue('JMT_TYPEDOS', sTypeDos_p);
               PutValue('JMT_GUIDPER', sGuidPer_p);
               PutValue('JMT_NATURECPT', sNatureCPT_p);
               PutValue('JMT_NATUREOP', '001');
               PutValue('JMT_NOORDRE', iNoOrdre_l + 1);
               PutValue('JMT_DATE', StrToDate(sDate_l));
               PutValue('JMT_NBTITRES', iTTNBPP_p);
               PutValue('JMT_SENSOPER', '+');
               PutValue('JMT_AFFECTSOLDE', 'X');
               PutValue('JMT_SOLDE', iSolde_p);
               PutValue('JMT_VALNOM', dValNominOuv_p);
               PutValue('JMT_MONTANT', iTTNBPP_p * dValNominOuv_p);
               PutValue('JMT_NATUREOP', '001') ;
               PutValue('JMT_NATURETIT', '');
               PutValue('JMT_DATEJOUIS', Date);
               PutValue('JMT_VALIDE', 'X');
               PutValue('JMT_DATEVISA', Date);
               PutValue('JMT_COMMENT', RechDom('YNATUREOP', '001', false));
               InsertDB(nil);
               result := true;
            end;
         end;
      end;
   end;
   OBMvtTitres_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 21/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function JUMvtMessageSup(bAvecBenef_p : boolean) : string;
var
   sMessage_l : string;
begin
   sMessage_l := 'Attention : vous supprimez un mouvement de titres validé.' + #13#10 +
                 'Le nombre de titres du titulaire ';
   if bAvecBenef_p then
      sMessage_l := sMessage_l + 'et du bénéficiaire ';
   sMessage_l := sMessage_l + 'va être rétablit.' + #13#10 + 'Confirmez vous?';
   result := sMessage_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 21/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function JUMvtOppositeSens(sSens_p : string) : string;
begin
   if sSens_p = '+' then
      result := '-'
   else
      result := '+';
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 21/12/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function JUMvtInfosLien(OBLien_p : TOB; sGuidPerDos_p, sGuidPer_p : string;
                        iNoOrdre_p : integer;
                        sTypeDos_p, sFonction_p : string) : boolean;
begin
   OBLien_p.ClearDetail;
   OBLien_p.LoadDetailDBFromSQL('ANNULIEN',
                    'SELECT * FROM ANNULIEN ' +
                    'WHERE ANL_GUIDPERDOS = "' + sGuidPerdos_p + '" ' +
                    '  AND ANL_GUIDPER    = "' + sGuidPer_p + '" ' +
                    '  AND ANL_NOORDRE    = ' + IntToStr(iNoOrdre_p) +
                    '  AND ANL_TYPEDOS    = "' + sTypeDos_p + '" ' +
                    '  AND ANL_FONCTION   = "' + sFonction_p + '"');
   result := (OBLien_p.Detail.Count > 0);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 08/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : TOF_JUMVTTITRES. : integer
*****************************************************************}
procedure JUMvtHistoLien(OBLien_p : TOB; sTypeModif_p, sNatureOp_p, sGuidTit_p, sGuidBen_p, sGuidPerDos_p : string;
                         iNoOrdre_p : integer;
                         sTypeDos_p, sFonction_p, sValNominOuv_p : string);


var
   sCle_l, sComplement_l : string;
   OBAnnuaire_l : TOB;
begin
   OBAnnuaire_l := TOB.Create('ANNUAIRE', nil, -1);
   OBAnnuaire_l.LoadDetailFromSQL('SELECT ANN_GUIDPER, ANN_NOMPER, ANN_ALCP, ANN_ALVILLE ' +
                                  'FROM ANNUAIRE ' +
                                  'WHERE ANN_GUIDPER = "' + sGuidTit_p + '"');

   sComplement_l := 'HNL_NOMPROPTITRE='   + OBAnnuaire_l.Detail[0].GetValue('ANN_NOMPER') + ';' +
                    'HNL_CPPROPTITRE='    + OBAnnuaire_l.Detail[0].GetValue('ANN_ALCP') + ';' +
                    'HNL_VILLEPROPTITRE=' + OBAnnuaire_l.Detail[0].GetValue('ANN_ALVILLE') + ';' +
                    'HNL_DATEMVTTITRE='   + DateToStr(Date) + ';' +
                    'HNL_NOMACQUERTITRE=' + GetValChamp('ANNUAIRE', 'ANN_NOMPER', 'ANN_GUIDPER = "' + sGuidBen_p + '"') + ';' +
                    'HNL_VALEURTITRE='    + sValNominOuv_p + ';' +
                    'HNL_NATURETITRE='    + RechDom('YNATUREOP', sNatureOp_p, false);
   OBAnnuaire_l.Free;
   sCle_l := ' HNL_GUIDPERDOS     = "' + sGuidPerDos_p + '"' +
             ' AND HNL_TYPEDOS    = "' + sTypeDos_p + '"' +
             ' AND HNL_GUIDPER    = "' + sGuidTit_p + '"' +
             ' AND HNL_FONCTION   = "' + sFonction_p + '"';

   Historisation(OBLien_p, 'ANNULIEN', 'HISTOANNULIEN',
                 sCle_l, ' HNL_NOORDRE DESC', 'HNL_NOORDRE',
                 sTypeModif_p, sComplement_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 21/12/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :       TOF_JUMVTTITRES.
*****************************************************************}
function AGLDeleteMouvement(aovParams_p : array of variant; iNb_p : integer) : integer;
var
   sTypeDos_l, sFonction_l, sValNominOuv_l, sNatureCpt_l, sGuidPerDos_l, sGuidPer_l : string;
   iNoOrdre_l : integer;
begin
   //'JMT_GUIDPERDOS;JMT_TYPEDOS;JMT_GUIDPER;JMT_NATURECPT;JMT_NOORDRE;fonction;valnominouv';
   sGuidPerDos_l  := String(aovParams_p[0]);
   sTypeDos_l     := String(aovParams_p[1]);
   sGuidPer_l     := String(aovParams_p[2]);
   sNatureCpt_l   := String(aovParams_p[3]);
   iNoOrdre_l     := StrToInt(String(aovParams_p[4]));
   sFonction_l    := String(aovParams_p[5]);
   sValNominOuv_l := String(aovParams_p[6]);
   try
      JUMvtDeleteMouvement(sGuidPerDos_l, sGuidPer_l, iNoOrdre_l,
                      sTypeDos_l, sFonction_l, sNatureCpt_l, sValNominOuv_l);
   except
      result := 0;
      exit;
   end;
   result := 1;   
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 21/12/2004
Modifié le ... :   /  /    
Description .. : TOF_JUMVTTITRES.
Mots clefs ... :
*****************************************************************}
procedure JUMvtDeleteMouvement(sGuidPerDos_p, sGuidPer_p : string; iNoOrdre_p : integer;
                               sTypeDos_p, sFonction_p, sNatureCpt_p, sValNominOuv_p : string);

var
   sBenGuidPer_l, sSensOper_l, sNatureOp_l, sValide_l : string;
   iSolde_l, iNbTitres_l : integer;
   OBMvt_l, OBContrePartie_l, OBLienTit_l, OBLienBen_l : TOB;
   bAffectSolde_l : boolean;
begin
   OBMvt_l := TOB.Create('JUMVTTITRES', nil, -1);
   OBMvt_l.LoadDetailDBFromSQL('JUMVTTITRES',
            'SELECT * FROM JUMVTTITRES ' +
            'WHERE JMT_GUIDPERDOS = "' + sGuidPerDos_p + '" ' +
            '  AND JMT_TYPEDOS    = "' + sTypeDos_p + '"' +
            '  AND JMT_GUIDPER    = "' + sGuidPer_p + '" ' +
            '  AND JMT_NATURECPT  = "' + sNatureCpt_p + '"' +
            '  AND JMT_NOORDRE    = ' + IntToStr(iNoOrdre_p));

   bAffectSolde_l := (OBMvt_l.Detail[0].GetValue('JMT_AFFECTSOLDE') = 'X');
   sSensOper_l := OBMvt_l.Detail[0].GetValue('JMT_SENSOPER');
   iNbTitres_l := OBMvt_l.Detail[0].GetValue('JMT_NBTITRES');
   sNatureOp_l := OBMvt_l.Detail[0].GetValue('JMT_NATUREOP');
   sValide_l := OBMvt_l.Detail[0].GetValue('JMT_VALIDE');

   OBContrePartie_l := TOB.Create('JUMVTTITRES', nil, -1);
   OBContrePartie_l.LoadDetailDBFromSQL('JUMVTTITRES',
            'SELECT * FROM JUMVTTITRES ' +
            'WHERE JMT_GUIDPERDOS = "' + sGuidPerDos_p + '" ' +
            '  AND JMT_TYPEDOS    = "' + sTypeDos_p + '"' +
            '  AND JMT_GUIDPER    <> "' + sGuidPer_p + '" ' +
            '  AND JMT_NATURECPT  = "' + sNatureCpt_p + '"' +
            '  AND JMT_NOORDRE    = ' + IntToStr(iNoOrdre_p));

   if (sValide_l = 'X') and (sNatureOp_l <> '001') then
   begin
      sBenGuidPer_l := '';
      if OBContrePartie_l.Detail.Count > 0 then
         sBenGuidPer_l := OBContrePartie_l.Detail[0].GetValue('JMT_GUIDPER');

      OBLienTit_l := TOB.Create('ANNULIEN', nil, -1);
      JUMvtInfosLien(OBLienTit_l, sGuidPerDos_p, sGuidPer_p, 1, sTypeDos_p, sFonction_p);
      if OBLienTit_l.Detail.Count > 0 then
      begin
         iSolde_l := JUMvtUnaffectUnSolde(bAffectSolde_l, sSensOper_l,
                                     OBLienTit_l.Detail[0].GetValue('ANL_TTNBPP'), iNbTitres_l);

      // Mise à jour des liens actionnaires
         OBLienTit_l.Detail[0].PutValue('ANL_TTNBPP', iSolde_l);
         InsOrUpdDepuisTOB(OBLienTit_l.Detail[0]);
         JUMvtHistoLien(OBLienTit_l.Detail[0], 'Suppression', sNatureOp_l,
                        sGuidPer_p, sBenGuidPer_l, sGuidPerDos_p, iNoOrdre_p,
                        sTypeDos_p, sFonction_p, sValNominOuv_p);
      end;
      OBLienTit_l.Free;

      if OBContrePartie_l.Detail.Count > 0 then
      begin
         OBLienBen_l := TOB.Create('ANNULIEN', nil, -1);
         JUMvtInfosLien(OBLienBen_l, sGuidPerDos_p, sBenGuidPer_l, 1, sTypeDos_p, sFonction_p);
         if OBLienBen_l.Detail.Count > 0 then
         begin
            iSolde_l := JUMvtUnaffectUnSolde(bAffectSolde_l,
                                        JUMvtOppositeSens(sSensOper_l),
                                        OBLienBen_l.Detail[0].GetValue('ANL_TTNBPP'),
                                        iNbTitres_l);

         // Mise à jour des liens actionnaires
            OBLienBen_l.Detail[0].PutValue('ANL_TTNBPP', iSolde_l);
            InsOrUpdDepuisTOB(OBLienBen_l);
            JUMvtHistoLien(OBLienBen_l.Detail[0], 'Suppression', sNatureOp_l,
                           sBenGuidPer_l, sGuidPer_p, sGuidPerDos_p, iNoOrdre_p,
                           sTypeDos_p, sFonction_p, sValNominOuv_p);
         end;
         OBLienBen_l.Free;
      end;
   end;

   BEGINTRANS;
   OBMvt_l.Detail[0].DeleteDB(true);
   if OBContrePartie_l.Detail.Count > 0 then
      OBContrePartie_l.Detail[0].DeleteDB;
   COMMITTRANS;

   OBMvt_l.Free;
   OBContrePartie_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 17/02/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function JUMvtDernierMouvement(sGuidPerDos_p, sGuidPer_p : string; iNoOrdre_p : integer;
                               sTypeDos_p, sNatureCpt_p : string) : boolean;

var
   OBQui_l : TOB;
   sQui_l, sGuidPer_l : string;
   iMaxOrdre_l, iInd_l : integer;
   bContinue_l : boolean;
begin
   result := true;
   OBQui_l := TOB.Create('QUI', nil, -1);
   OBQui_l.LoadDetailFromSQL(
            'SELECT JMT_GUIDPER FROM JUMVTTITRES ' +
            'WHERE JMT_GUIDPERDOS = "' + sGuidPerDos_p + '" ' +
            '  AND JMT_TYPEDOS    = "' + sTypeDos_p + '"' +
            '  AND JMT_NOORDRE    = ' + IntToStr(iNoOrdre_p) +
            '  AND JMT_NATURECPT  = "' + sNatureCpt_p + '"');

   for iInd_l := 0 to OBQui_l.Detail.Count - 1 do
   begin
      if sQui_l <> '' then sQui_l := sQui_l + ' OR ';
      sQui_l := sQui_l + 'JMT_GUIDPER = "' + OBQui_l.Detail[iInd_l].GetString('JMT_GUIDPER') + '"';
   end;
   OBQui_l.Free;

   OBQui_l := TOB.Create('JUMVTTITRES', nil, -1);
   OBQui_l.LoadDetailDBFromSQL('JUMVTTITRES',
            'SELECT MAX(JMT_NOORDRE) AS JMT_NOORDRE, JMT_GUIDPER FROM JUMVTTITRES ' +
            'WHERE JMT_GUIDPERDOS = "' + sGuidPerDos_p + '" ' +
            '  AND JMT_TYPEDOS    = "' + sTypeDos_p + '"' +
            '  AND (' + sQui_l + ')' +
            '  AND JMT_NATURECPT  = "' + sNatureCpt_p + '" ' +
            'GROUP BY JMT_GUIDPER');

   iInd_l := 0;
   bContinue_l := true;
   while bContinue_l and (iInd_l < OBQui_l.Detail.Count) do
   begin
      iMaxOrdre_l := OBQui_l.Detail[iInd_l].GetInteger('JMT_NOORDRE');
      sGuidPer_l := OBQui_l.Detail[iInd_l].GetString('JMT_GUIDPER');
      bContinue_l := (iNoOrdre_p = iMaxOrdre_l);
      Inc(iInd_l);
   end;
   OBQui_l.Free;

   if not bContinue_l then
   begin
      sQui_l := 'Ordre de mouvement n° ' + IntToStr(iNoOrdre_p) + ' - ';
      if sGuidPer_l <> sGuidPer_p then
         sQui_l := sQui_l + 'Contrepartie ';
      sQui_l := sQui_l + GetValChamp('ANNUAIRE', 'ANN_NOMPER', 'ANN_GUIDPER = "' + sGuidPer_l + '"');
      PGIInfo('Vous ne pouvez supprimer que le dernier ordre de mouvement'#13#10 +
              '(Ordre de mouvement n°' + IntToStr(iMaxOrdre_l) + ').', sQui_l);
   end;

   Result := bContinue_l;
end;
/////////////////////////////////////////////////////////////////
end.
