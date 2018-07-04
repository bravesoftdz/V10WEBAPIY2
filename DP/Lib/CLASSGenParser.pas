{***********UNITE*************************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
unit CLASSGenParser;
//////////////////////////////////////////////////////////////////
interface
//////////////////////////////////////////////////////////////////
uses
   classes, UTOB,
{$IFDEF EAGLCLIENT}
{$ELSE}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
   HCtrls, SysUtils, HMsgBox;

//////////////////////////////////////////////////////////////////
type
   TGenParser = class
         _S : TStrings; // DEBUG

         sModeleDoc_c, sDocument_c, sExtModele_c, sDomaineAct_c : string;
         bModif_c, bDone_c : boolean;

         procedure Trace (debug : string);

         function  ChargeDossier        (var OBDossier_p    : TOB; sGuidPer_p, sCle_p : string) : boolean;
         function  ChargeJuridique      (var OBJuridique_p  : TOB; sJuCodeDos_p, sCodeOpe_p, sCle_p : string) : boolean;
         function  ChargeAnnuaire       (var OBAnnuaire_p   : TOB; sGuidPer_p : string) : boolean;
         procedure ChargeAnnuLien       (var OBAnnuLien_p   : TOB; sGuidPer_p, sTypeDos_p : string);
         procedure ChargeDetailLien     (var OBDetailLien_p : TOB);
         procedure ChargeJuEvenement    (var OBEvenement_p  : TOB; sGuidPer_p : string);
         procedure ChargeJuDosExe       (var OBDosExe_p     : TOB; sJuCodeDos_p : string);
         procedure ChargeJuDosExeDetails(var OBExeDetail_p  : TOB);
         procedure ChargeJuDosOper      (var OBDosOper_p    : TOB; sJuCodeDos_p, sNoOperation_p : string);
         procedure ChargeJuDosOpAct     (var OBDosAct_p     : TOB; sJuCodeDos_p, sNoOperation_p, sCodeAct_p : string);
         procedure ChargeJuDosOpRub     (var OBDosRub_p     : TOB; sJuCodeDos_p, sNoOperation_p, sCodeRub_p : string);
         procedure ChargeJuDosInfo      (var OBDosInfo_p    : TOB; sJuCodeDos_p : string);
         procedure ChargeDetailInfo     (var OBDetailInfo_p : TOB);         
         procedure ChargeJuBauxFonds    (var OBBauxFonds_p  : TOB; sGuidPerDos_p : string);

         procedure ChargeHistoOpRub     (var OBHistoRub_p   : TOB; sJuCodeDos_p : string);
         procedure ChargeHistoOpAct     (var OBHistoAct_p   : TOB; sJuCodeDos_p : string);

         procedure ChargeAnnuBis         (var OBAnnuBis_p      : TOB; sGuidPer_p : string);
         procedure ChargeDPFiscal       (var OBDPFiscal_p     : TOB; sGuidPer_p : string);
         procedure ChargeDPSocial       (var OBDPSocial_p     : TOB; sGuidPer_p : string);
         procedure ChargeDPOrga         (var OBDPOrga_p       : TOB; sGuidPer_p : string);
         procedure ChargeDPTabCompta    (var OBDPTabCompta_p  : TOB; sNoDossier_p : string);
         procedure ChargeDPTabPaie      (var OBDPTabPaie_p    : TOB; sNoDossier_p : string);
         procedure ChargeDPTabGenPaie   (var OBDPTabGenPaie_p : TOB; sNoDossier_p : string);
         procedure ChargeDPControl      (var OBDPControl_p    : TOB; sGuidPer_p : string);
         procedure ChargeTiers            (var OBTiers_p        : TOB; sTiers_p : string);
         procedure ChargeTiersCompl     (var OBTiersCompl_p   : TOB; sTiers_p : string);
         procedure ChargeProspects      (var OBProspects_p    : TOB; sAuxiliaire_p : string);
         procedure ChargeContact          (var OBContact_p      : TOB; sGuidPer_p : string);

         procedure GetListeLienRac  (OBAnnuLien_p : TOB; sRacine_p, sNomChamp_p : string;
                                     var tsListe_p: TStringList);
         procedure GetListeLien     (OBAnnuLien_p : TOB; sNomChamp_p : string;
                                     var tsListe_p: TStringList);
         procedure GetListe         (OBLaTob_p    : TOB; sNomChamp_p : string;
                                     var tsListe_p : TStringList);
         procedure GetListeHisto    (OBLaTob_p : TOB; sNomChamp_p : string;
                                     aosChamps_p, aosValeur_p : array of string;
                                     var tsListe_p : TStringList);


         function GetVarDossier     (OBDossier_p   : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : variant;
         function GetVarSociete     (OBAnnuaire_p  : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : Variant;
         function GetVarJuridique   (OBJuridique_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : variant;
         function GetVarLienRac     (OBAnnuLien_p  : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : Variant;
         function GetVarLien        (OBAnnuLien_p  : TOB; sSuffixe_p : string; nIndex_p : integer) : Variant;
         function GetVarEvenement   (OBEvenement_p : TOB; sSuffixe_p : string; nIndex_p : integer) : Variant;
         function GetVarExeRubrique (OBDosExe_p    : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : Variant;
         function GetVarAutreInfo   (OBDosInfo_p   : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : Variant;
         function GetVarOperation   (OBDosOpe_p    : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : Variant;
         function GetVarRubrique    (OBDosOpRub_p, OBBibRubrique_p : TOB; sRacine_p, sSuffixe_p, sCodeOp_p, sModule_p : string; nIndex_p : integer) : Variant;
         function GetVarActe        (OBDosOpAct_p  : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : Variant;

         function GetVarHistoRub    (OBHistoOpRub_p : TOB; sSuffixe_p, sCodeOp_p, sModule_p : string; nIndex_p : integer) : Variant;
         function GetVarHistoAct    (OBHistoOpAct_p : TOB; sSuffixe_p, sCodeOp_p, sModule_p : string; nIndex_p : integer) : Variant;
         function GetVarBauxFonds   (OBBauxFonds_p  : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : Variant;

         function GetVarAnnuBis      (OBAnnuBis_p      : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
         function GetVarDPFiscal     (OBDPFiscal_p     : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
         function GetVarDPSocial     (OBDPSocial_p     : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
         function GetVarDPOrga       (OBDPOrga_p       : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
         function GetVarTiers        (OBTiers_p        : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
         function GetVarTiersCompl   (OBTiersCompl_p   : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
         function GetVarDPTabCompta  (OBDPTabCompta_p  : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
         function GetVarDPTabPaie    (OBDPTabPaie_p    : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
         function GetVarDPTabGenPaie (OBDPTabGenPaie_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
         function GetVarDPControl    (OBDPControl_p    : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
         function GetVarProspects    (OBProspects_p    : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
         function GetVarContact       (OBContact_p        : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : Variant;

         procedure SetVarDossier    (var OBDossier_p   : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer; vValeur_p : variant);
         procedure SetVarExeRubrique(var OBDosExe_p    : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer; vValeur_p : Variant);
         procedure SetVarJuridique  (var OBJuridique_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer; vValeur_p : variant);
         procedure SetVarLienRac    (var OBAnnuLien_p  : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer; vValeur_p : Variant);
         procedure SetVarLien       (var OBAnnuLien_p  : TOB; sSuffixe_p : string; nIndex_p : integer; vValeur_p : Variant);
         procedure SetVarEvenement  (var OBEvenement_p : TOB; sSuffixe_p : string; nIndex_p : integer; vValeur_p : Variant);
         procedure SetVarSociete    (var OBAnnuaire_p  : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer; vValeur_p : variant);
         procedure SetVarAutreInfo  (var OBInfos_p     : TOB; sRacine_p, sSuffixe_p: string; nIndex_p: integer; vValeur_p: Variant);
         procedure SetVarRubrique   (var OBDosOpRub_p, OBBibRubrique_p : TOB; sRacine_p, sSuffixe_p: string; nIndex_p: integer; vValeur_p: Variant);
         procedure SetVarBauxFonds  (var OBBauxFonds_p : TOB; sSuffixe_p : string; nIndex_p : integer; vValeur_p : Variant);

      public
      private

   end;
//////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////
implementation
//////////////////////////////////////////////////////////////////
uses
     dpJurOutilsGen,

     {$IFDEF VER150}
     Variants,
     {$ENDIF}

     DpJurOutils;

                                          
//////////////////////////////////////////////////////////////////
// CLASSE TGENPARSER                                            //
//////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2003
Modifié le ... :   /  /
Description .. : CHARGEMENT DES DONNEES DU DOSSIER
Suite ........ : Table DOSSIER
Mots clefs ... :
*****************************************************************}
function TGenParser.ChargeDossier(var OBDossier_p : TOB; sGuidPer_p, sCle_p : string) : boolean;
var
   QRYRequete_l : TQuery;
   bResult_l : boolean;
begin
   if sCle_p = '' then
   begin
      QRYRequete_l := OpenSql('Select * FROM DOSSIER ' +
                              'WHERE DOS_GUIDPER = "' + sGuidPer_p + '"',
                              true);
      bResult_l := OBDossier_p.SelectDB(sCle_p, QRYRequete_l,FALSE);
      ferme(QRYRequete_l);
   end
   else
      bResult_l := OBDossier_p.SelectDB(sCle_p, nil,FALSE);

   if bResult_l then
      TraiteUserCollDos(OBDossier_p);
   result := bResult_l;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 22/12/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.ChargeJuEvenement(var OBEvenement_p  : TOB; sGuidPer_p : string);
var
     nInd_l : integer;
     sLibFamEvt_l : string;
     QRYEvenement_l : TQuery;
begin
   Trace('Start : ChargeEvenement');
   // on charge tous les liens du dossier
   QRYEvenement_l := OpenSQL ('select * from JUEVENEMENT ' +
                           'where JEV_GUIDPER = "' + sGuidPer_p + '" ' +
                           '  and JEV_DOMAINEACT = "' + sDomaineAct_c +'" ' +
                           'order by JEV_FAMEVT, JEV_CODEEVT', true);
   OBEvenement_p.LoadDetailDB('JUEVENEMENT', '', '', QRYEvenement_l, false, FALSE);
   Ferme (QRYEvenement_l);

   for nInd_l := 0 to OBEvenement_p.Detail.Count - 1 do
   begin
    // on ajoute des champs supplementaires
     OBEvenement_p.Detail[nInd_l].AddChampSup('JEV_LIBFAMEVT',true);
    sLibFamEvt_l := RechDom('JUFAMEVT', OBEvenement_p.Detail[nInd_l].GetValue('JEV_FAMEVT'), false);
     OBEvenement_p.Detail[nInd_l].PutValue('JEV_LIBFAMEVT', sLibFamEvt_l);
   end;
   Trace('End : ChargeEvenement');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2003
Modifié le ... :
Description .. : CHARGEMENT DES DONNEES DE LA FICHE JURIDIQUE
Suite ........ : Tables : JURIDIQUE et LIENSOLE
Mots clefs ... :
*****************************************************************}
function TGenParser.ChargeJuridique(var OBJuridique_p  : TOB; sJuCodeDos_p, sCodeOpe_p, sCle_p : string) : boolean;
var
   sRequete_l : string;
   QRYRequete_l : TQuery;
   OBBlobJuridique_l : TOB;
   bResult_l : boolean;
begin
   if sCle_p = '' then
   begin
      sRequete_l := 'Select * FROM JURIDIQUE ' +
                    'WHERE JUR_CODEDOS ="' + sJuCodeDos_p + '"';
//      if sTypeDos_p <> '' then
//         sRequete_l := sRequete_l + '  AND ANL_TYPEDOS = "' + sTypeDos_p + '" ';

      QRYRequete_l := OpenSql(sRequete_l, true);
      bResult_l := OBJuridique_p.SelectDB(sCle_p, QRYRequete_l, FALSE);
      ferme(QRYRequete_l);
   end
   else
      bResult_l := OBJuridique_p.SelectDB(sCle_p, nil, FALSE);

   if bResult_l then
   begin
      TraiteUserCollJur(OBJuridique_p, sCodeOpe_p);
      OBBlobJuridique_l := TOB.Create('les blobs juridique', OBJuridique_p, -1);
      OBBlobJuridique_l.LoadDetailDB('LIENSOLE', '"JUR";"' + OBJuridique_p.GetValue('JUR_GUIDPERDOS') + '";', '', nil, false);
   end;
   result := bResult_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2003
Modifié le ... :   /  /
Description .. : CHARGEMENT DES DONNEES DE LA FICHE ANNUAIRE LIEE
Suite ........ : A LA PERSONNE SELECTIONNEE
Suite ........ : Tables : ANNUAIRE, LIENSOLE,
Mots clefs ... :
*****************************************************************}

function TGenParser.ChargeAnnuaire(var OBAnnuaire_p : TOB; sGuidPer_p : string) : boolean;
var
   OBBlobAnnuaire_l : TOB;
   bResult_l : boolean;
begin
   bResult_l := OBAnnuaire_p.SelectDB('"' + sGuidPer_p + '"', nil, FALSE);

   if bResult_l then
   begin
      InitChampsSupPer(OBAnnuaire_p, sGuidPer_p);
      OBBlobAnnuaire_l := TOB.Create('LIENSOLE', OBAnnuaire_p, -1);
      OBBlobAnnuaire_l.LoadDetailDB('LIENSOLE', '"ANN";"' + sGuidPer_p + '";', '', nil, false, false);
   end;
   result := bResult_l;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.ChargeAnnuLien(var OBAnnuLien_p : TOB; sGuidPer_p, sTypeDos_p : string);
var
   nLienInd_l : integer;
   sRequete_l : string;
   QRYAnnuLien_l : TQuery;
   OBAnnuaire_l : TOB;
begin
   Trace('Start : ChargeAnnuLien');
   // on charge tous les liens du dossier
   sRequete_l := 'select * from ANNULIEN ' +
                 'where ANL_GUIDPERDOS = "' + sGuidPer_p + '" ';
   if sTypeDos_p <> '' then
      sRequete_l := sRequete_l + '  AND ANL_TYPEDOS = "' + sTypeDos_p + '" ';
   sRequete_l := sRequete_l + 'order by ANL_TRI, ANL_FONCTION, ANL_NOMPER ASC';
   QRYAnnuLien_l := OpenSQL (sRequete_l, true);
   OBAnnuLien_p.LoadDetailDB('ANNULIEN', '', '', QRYAnnuLien_l, false, FALSE);
   Ferme (QRYAnnuLien_l);

   OBAnnuaire_l := TOB.Create('ANNUAIRE',nil,-1);
   // pour chaque lien,
   for nLienInd_l := 0 to OBAnnuLien_p.Detail.Count - 1 do
   begin
    // on charge les infos de la personne associée au dossier
     OBAnnuaire_l.SelectDB('"' + OBAnnuLien_p.Detail[nLienInd_l].GetValue('ANL_GUIDPER') + '"', nil, false);
    // on ajoute un champ supplementaire "nom de la personne"
     OBAnnuLien_p.Detail[nLienInd_l].AddChampSup('ANL_NOMPER',true);
     OBAnnuLien_p.Detail[nLienInd_l].PutValue('ANL_NOMPER', OBAnnuaire_l.GetValue('ANN_NOMPER'));
    // on ajoute des champs supplementaires pour savoir si le detail de lien est charge
     OBAnnuLien_p.Detail[nLienInd_l].AddChampSup('DETAIL_AFFECTE',true);
     OBAnnuLien_p.Detail[nLienInd_l].PutValue('DETAIL_AFFECTE','0');
   end;
   OBAnnuaire_l.Free;
   Trace('End : ChargeAnnuLien');
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.ChargeDetailLien(var OBDetailLien_p : TOB);
var
     sGuidPer_l : string;
     OBAnnuaire_l, OBAss_l, OBCoopt_l : TOB;
     nIndFils_l : integer;
begin
   Trace('Start : ChargeDetailLien');
   // on affecte true au champ DETAIL_AFFECTE
     OBDetailLien_p.PutValue('DETAIL_AFFECTE','1');
   // on récupère le code de la personne concernée par le lien
     sGuidPer_l := OBDetailLien_p.GetValue('ANL_GUIDPER');
   //on crée une tob OBAnnuaire_l, fille de la tob du lien
   // ATTENTION : on affecte la 0ème position
     OBAnnuaire_l := TOB.Create('ANNUAIRE',OBDetailLien_p,0);
   // dans laquelle on charge les infos ANNUAIRE de sGuidPer_l
     OBAnnuaire_l.SelectDB('"' + sGuidPer_l + '"', nil, false);
     InitChampsSupPer(OBAnnuaire_l, sGuidPer_l, '');
   // création d'autres champs....
     OBAnnuaire_l.AddChampSup('TYPEPERLIB',true);
     OBAnnuaire_l.PutValue('TYPEPERLIB',RechDom('JUTYPEPER',OBDetailLien_p.GetValue('ANL_TYPEPER'),false));
     OBAnnuaire_l.AddChampSup('TYPEPERDOSLIB',true);
     OBAnnuaire_l.PutValue('TYPEPERDOSLIB',RechDom('JUTYPEPER',OBDetailLien_p.GetValue('ANL_TYPEPERDOS'),false));
   // on crée des champs supp dans lesquels on renseignera la position des indices des tobs filles
   // contenant les informations des personnes Liées et Cooptées
   // permet de retrouver l'indice de LIE1 et COOPT cf : ExtraitSufPersonne()
     OBAnnuaire_l.AddChampSup('INDICELIE1',true);
     OBAnnuaire_l.PutValue('INDICELIE1',-1);
     OBAnnuaire_l.AddChampSup('INDICECOOPT',true);
     OBAnnuaire_l.PutValue('INDICECOOPT',-1);
   //gestion de la personne associee
   // on récupère le code de la personne associée
     sGuidPer_l := OBDetailLien_p.GetValue('ANL_PERASS1GUID');
     nIndFils_l:=0;
   // si on a une personne associée
     if sGuidPer_l<>'0' then
     begin
       Inc(nIndFils_l);
      // affectation de l'indice de la tob de la personne associée à la tob OBAnnuaire_l
       OBAnnuaire_l.PutValue('INDICELIE1',nIndFils_l);
      //on crée une tob OBAss_l, fille de la tob du lien
      // ATTENTION : on affecte la nIndFils_l ème position
       OBAss_l := TOB.Create('ANNUAIRE',OBDetailLien_p,nIndFils_l);
      // dans laquelle on charge les infos ANNUAIRE de sGuidPer_l
       OBAss_l.SelectDB('"' + sGuidPer_l + '"', nil, false);
      // on créee les champs supplémentaires de l'ANNUAIRE
       InitChampsSupPer(OBAss_l,sGuidPer_l,OBDetailLien_p.GetValue('ANL_PERASS1QUAL'));
     end;
   //gestion de la personne cooptee
   // on récupère le code de la personne associée
     sGuidPer_l := OBDetailLien_p.GetValue('ANL_COOPTGUID');
   // si on a une personne cooptee
     if sGuidPer_l<>'0' then
     begin
       Inc(nIndFils_l);
      // affectation de l'indice de la tob de la personne associée à la tob OBAnnuaire_l
       OBAnnuaire_l.PutValue('INDICECOOPT',nIndFils_l);
      //on crée une tob OBCoopt_l, fille de la tob du lien
      // ATTENTION : on affecte la nIndFils_l ème position
       OBCoopt_l := TOB.Create('ANNUAIRE',OBDetailLien_p,nIndFils_l);
      // dans laquelle on charge les infos ANNUAIRE de sGuidPer_l
       OBCoopt_l.SelectDB('"' + sGuidPer_l + '"', nil, false);
      // on créee les champs supplémentaires de l'ANNUAIRE
       InitChampsSupPer(OBCoopt_l,sGuidPer_l,'');
     end;
   Trace('End : ChargeDetailLien');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TGenParser.ChargeJuDosExe(var OBDosExe_p : TOB; sJuCodeDos_p : string);
var
  nExeInd_l : integer;
  sCodeDos_l, sCodeExe_l : string;
begin
   Trace('Start : ChargeJuDosExe');
   OBDosExe_p.LoadDetailDB('JUDOSEX', '"' + sJuCodeDos_p + '";', 'JDE_FINEX DESC', nil, true, false);

   for nExeInd_l := 0 to OBDosExe_p.Detail.Count - 1 do
   begin
     sCodeDos_l := OBDosExe_p.Detail[nExeInd_l].GetValue('JDE_CODEDOS');
     sCodeExe_l := OBDosExe_p.Detail[nExeInd_l].GetValue('JDE_CODEEX');
     OBDosExe_p.Detail[nExeInd_l].AddChampSup('JDE_DL',true);
     OBDosExe_p.Detail[nExeInd_l].PutValue('JDE_DL',RechDom('JUDEVISE', OBDosExe_p.Detail[nExeInd_l].GetValue('JDE_DEVISE'),false));
     OBDosExe_p.Detail[nExeInd_l].AddChampSup('JDE_DS',true);
     OBDosExe_p.Detail[nExeInd_l].PutValue('JDE_DS',RechDom('JUDEVISESYM', OBDosExe_p.Detail[nExeInd_l].GetValue('JDE_DEVISE'),false));
     OBDosExe_p.Detail[nExeInd_l].AddChampSup('INIT_DETAIL',true);
     OBDosExe_p.Detail[nExeInd_l].PutValue('INIT_DETAIL',0);
   end;
   Trace('End : ChargeJuDosExe');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.ChargeJuDosExeDetails(var OBExeDetail_p : TOB);
var
     nRubInd_l : integer;
     OBRub_l : TOB;
     sCodeDos_l, sCodeExe_l, sCodeRub_l : string;
begin
   Trace('Start : ChargeJuDosExeDetails');
   OBExeDetail_p.PutValue('INIT_DETAIL', 1);
   sCodeDos_l := OBExeDetail_p.GetValue('JDE_CODEDOS');
   sCodeExe_l := OBExeDetail_p.GetValue('JDE_CODEEX');
   OBRub_l := TOB.Create('Detail Exe', OBExeDetail_p,-1);
   OBRub_l.LoadDetailDB('JUDOSEXRUB', '"' + sCodeDos_l + '";"' + sCodeExe_l + '"', 'JER_CODERUB', nil, true, false);

   for nRubInd_l := 0 to OBRub_l.Detail.Count-1 do
   begin
     sCodeRub_l := OBRub_l.Detail[nRubInd_l].GetValue('JER_CODERUB');
     OBRub_l.Detail[nRubInd_l].AddChampSup('JDE_LIBELLERUB',true);
     OBRub_l.Detail[nRubInd_l].PutValue('JDE_LIBELLERUB', GetValChamp('JUTYPEEX','JTX_CODERUB','JTX_CODERUB="' + sCodeRub_l + '"'));
   end;
   Trace('End : ChargeJuDosExeDetails');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.ChargeJuDosOper (var OBDosOper_p : TOB;
            sJuCodeDos_p, sNoOperation_p : string);
var
   sCle_l : string;
begin
   Trace('Start : ChargeJuDosOper');
   sCle_l := '"' + sJuCodeDos_p + '";';
   if sNoOperation_p <> '' then
      sCle_l := sCle_l + '"' + sNoOperation_p + '"';

   OBDosOper_p.LoadDetailDB('JUDOSOPER', sCle_l, 'JOP_CODEOP,JOP_MODULE', nil, false);
   Trace('End : ChargeJuDosOper');

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.ChargeJuDosOpAct (var OBDosAct_p : TOB;
                        sJuCodeDos_p, sNoOperation_p, sCodeAct_p : string);
var
   sCle_l : string;
begin
   Trace('Start : ChargeJuDosOpAct');
   sCle_l := '"' + sJuCodeDos_p + '";';
   if sNoOperation_p <> '' then
      sCle_l := sCle_l + '"' + sNoOperation_p + '";';
   if sCodeAct_p <> '' then
      sCle_l := sCle_l + '"' + sCodeAct_p + '"';

   OBDosAct_p.LoadDetailDB('JUDOSOPACT', sCle_l, 'JOA_TRI', nil, false);
   Trace('End : ChargeJuDosOpAct');

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.ChargeHistoOpRub(var OBHistoRub_p : TOB; sJuCodeDos_p : string);
var
   QRYRub_l : TQuery;
   sRequete_l : string;
begin
   Trace('Start : ChargeHistoOpRub');
   sRequete_l := 'SELECT JUDOSOPRUB.*, JBR_TYPERUB AS JOR_TYPERUB, JBR_LIBRUB AS JOR_LIBRUB ' +
                 'FROM JUDOSOPRUB, JUBIBRUBRIQUE ' +
                 'WHERE JBR_CODERUB = JOR_CODERUB AND JBR_MODULE = JOR_MODULE ' +
                 '  and JBR_HISTO = "X" ' +
                 '  and JOR_CODEDOS = "' + sJuCodeDos_p + '" ';

   QRYRub_l := OpenSQL(sRequete_l, true);
   OBHistoRub_p.LoadDetailDB('JUDOSOPRUB', '', 'JOR_TRI', QRYRub_l, false);
   Ferme(QRYRub_l);
   Trace('End : ChargeHistoOpRub');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 25/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.ChargeHistoOpAct(var OBHistoAct_p : TOB; sJuCodeDos_p : string);
var
   QRYAct_l : TQuery;
   sRequete_l : string;
begin
   Trace('Start : ChargeHistoOpAct');
   sRequete_l := 'SELECT JUDOSOPACT.*, JBA_LIBACT AS JOA_LIBACT ' +
                 'FROM JUDOSOPACT, JUBIBACTION ' +
                 'WHERE JBA_CODEACT = JOA_CODEACT AND JBA_MODULE = JOA_MODULE ' +
                 '  and JOA_DOCFAIT = "X" ' +
                 '  and JOA_CODEDOS = "' + sJuCodeDos_p + '" ';

   QRYAct_l := OpenSQL(sRequete_l, true);
   OBHistoAct_p.LoadDetailDB('JUDOSOPACT', '', 'JOA_TRI', QRYAct_l, false);
   Ferme(QRYAct_l);
   Trace('End : ChargeHistoOpAct');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.ChargeJuDosOpRub(var OBDosRub_p : TOB;
                        sJuCodeDos_p, sNoOperation_p, sCodeRub_p : string);
var
   sCle_l : string;
begin
   Trace('Start : ChargeJuDosOpRub');
   sCle_l := '"' + sJuCodeDos_p + '";';
   if sNoOperation_p <> '' then
      sCle_l := sCle_l + '"' + sNoOperation_p + '";';
   if sCodeRub_p <> '' then
      sCle_l := sCle_l + '"' + sCodeRub_p + '"';

   OBDosRub_p.LoadDetailDB('JUDOSOPRUB', sCle_l, 'JOR_TRI', nil, false);
   Trace('End : ChargeJuDosOpRub');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 03/11/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.ChargeJuDosInfo( var OBDosInfo_p : TOB; sJuCodeDos_p : string);
var
//     OBAnnuaire_l : TOB;
   nInd_l : integer;
   {sGuidPer_l,} sRequete_l : string;
begin
   Trace('Start : ChargeJuDosInfo');
   sRequete_l := 'select * from JUDOSINFO ' +
                 'where JDI_CODEDOS = "' + sJuCodeDos_p + '" ' +
                 'order by JDI_CODEINFO';

   OBDosInfo_p.LoadDetailDBFromSQL('JUDOSINFO', sRequete_l);

   for nInd_l := 0 to OBDosInfo_p.Detail.Count - 1 do
   begin
     OBDosInfo_p.Detail[nInd_l].AddChampSup('DETAIL_AFFECTE', true);
     OBDosInfo_p.Detail[nInd_l].PutValue('DETAIL_AFFECTE','0');

{     OBAnnuaire_l := TOB.Create('ANNUAIRE',OBDosInfo_p.Detail[nInd_l], -1);
     sGuidPer_l := OBDosInfo_p.Detail[nInd_l].GetValue('JDI_GUIDPER');
     OBAnnuaire_l.SelectDB('"' + sGuidPer_l + '"', nil, false);
     InitChampsSupPer(OBAnnuaire_l, '');}
   end;
   Trace('End : ChargeJuDosInfo');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/12/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.ChargeContact(var OBContact_p : TOB; sGuidPer_p : string);
var
   sRequete_l : string;
begin
   Trace('Start : ChargeContact');
   sRequete_l := 'select * from CONTACT ';
   if sGuidPer_p <> '' then
      sRequete_l := sRequete_l + 'where C_GUIDPER = "' + sGuidPer_p + '"';
   sRequete_l := sRequete_l + ' order by C_GUIDPER, C_NUMEROCONTACT';

   OBContact_p.LoadDetailDBFromSQL('CONTACT', sRequete_l);
   Trace('End : ChargeContact');
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 22/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.ChargeDetailInfo(var OBDetailInfo_p : TOB);
var
     sGuidPer_l : string;
     OBAnnuaire_l : TOB;
begin
   Trace('Start : ChargeDetailInfo');
   // on affecte true au champ DETAIL_AFFECTE
   OBDetailInfo_p.PutValue('DETAIL_AFFECTE','1');
   // on récupère le code de la personne concernée par le lien
   sGuidPer_l := OBDetailInfo_p.GetValue('JDI_GUIDPER');
   //on crée une tob OBAnnuaire_l, fille de la tob du lien
   // ATTENTION : on affecte la 0ème position
   OBAnnuaire_l := TOB.Create('ANNUAIRE', OBDetailInfo_p, 0);
   // dans laquelle on charge les infos ANNUAIRE de sGuidPer_l
   OBAnnuaire_l.SelectDB('"' + sGuidPer_l + '"', nil, false);
   InitChampsSupPer(OBAnnuaire_l, sGuidPer_l, '');
   // création d'autres champs....
   OBAnnuaire_l.AddChampSup('TYPEPERLIB',true);
   OBAnnuaire_l.PutValue('TYPEPERLIB', RechDom('JUTYPEPER', OBAnnuaire_l.GetValue('ANN_TYPEPER'), false));
   OBAnnuaire_l.AddChampSup('TYPEPERDOSLIB', true);
   // on crée des champs supp dans lesquels on renseignera la position des indices des tobs filles
   // contenant les informations des personnes Liées et Cooptées
   // permet de retrouver l'indice de LIE1 et COOPT cf : ExtraitSufPersonne()
   OBAnnuaire_l.AddChampSup('INDICELIE1',true);
   OBAnnuaire_l.PutValue('INDICELIE1',-1);
   OBAnnuaire_l.AddChampSup('INDICECOOPT',true);
   OBAnnuaire_l.PutValue('INDICECOOPT',-1);
   Trace('End : ChargeDetailInfo');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.ChargeJuBauxFonds(var OBBauxFonds_p : TOB; sGuidPerDos_p : string);
var
   sRequete_l : string;
begin
   Trace('Start : ChargeJuBauxFonds');
   sRequete_l := 'select * from JUBAUXFONDS ' +
                 'where JBF_GUIDPERDOS = "' + sGuidPerDos_p + '" ' +
                 'order by JDI_CODEINFO';

   OBBauxFonds_p.LoadDetailDBFromSQL('JUBAUXFONDS', sRequete_l);

   Trace('End : ChargeJuBauxFonds');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 26/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.ChargeAnnuBis(var OBAnnuBis_p : TOB; sGuidPer_p : string);
begin
   Trace('Start : ChargeAnnuBis');
   OBAnnuBis_p.LoadDetailDBFromSQL('ANNUBIS', 'SELECT * FROM ANNUBIS WHERE ANB_GUIDPER = "' + sGuidPer_p + '"');
   Trace('End : ChargeAnnuBis');   
end;

procedure TGenParser.ChargeDPFiscal(var OBDPFiscal_p : TOB; sGuidPer_p : string);
begin
   Trace('Start : ChargeDPFiscal');
   OBDPFiscal_p.LoadDetailDBFromSQL('DPFISCAL', 'SELECT * FROM DPFISCAL WHERE DFI_GUIDPER = "' + sGuidPer_p + '"');
   Trace('End : ChargeDPFiscal');
end;

procedure TGenParser.ChargeDPSocial(var OBDPSocial_p : TOB; sGuidPer_p : string);
begin
   Trace('Start : ChargeDPSocial');
   OBDPSocial_p.LoadDetailDBFromSQL('DPSOCIAL', 'SELECT * FROM DPSOCIAL WHERE DSO_GUIDPER = "' + sGuidPer_p + '"');
   Trace('End : ChargeDPSocial');
end;

procedure TGenParser.ChargeDPOrga(var OBDPOrga_p : TOB; sGuidPer_p : string);
begin
   Trace('Start : ChargeDPOrga');
   OBDPOrga_p.LoadDetailDBFromSQL('DPORGA', 'SELECT * FROM DPORGA WHERE DOR_GUIDPER = "' + sGuidPer_p + '"');
   Trace('End : ChargeDPOrga');
end;

procedure TGenParser.ChargeDPControl(var OBDPControl_p : TOB; sGuidPer_p : string);
begin
   Trace('Start : ChargeDPControl');
   OBDPControl_p.LoadDetailDBFromSQL('DPCONTROLE', 'SELECT * FROM DPCONTROLE WHERE DCL_GUIDPER = "' + sGuidPer_p + '"');
   Trace('End : ChargeDPControl');
end;

procedure TGenParser.ChargeDPTabCompta(var OBDPTabCompta_p : TOB; sNoDossier_p : string);
var
   sRequete_l : string;
begin
   Trace('Start : ChargeDPTabCompta');
   sRequete_l := 'SELECT * FROM DPTABCOMPTA ' +
                 'WHERE DTC_NODOSSIER = "' + sNoDossier_p + '"';
   OBDPTabCompta_p.LoadDetailDBFromSQL('DPTABCOMPTA', sRequete_l);
   Trace('End : Chargev');
end;

procedure TGenParser.ChargeDPTabPaie(var OBDPTabPaie_p : TOB; sNoDossier_p : string);
var
   sRequete_l : string;
begin
   Trace('Start : ChargeDPTabPaie');
   sRequete_l := 'SELECT * FROM DPTABPAIE ' +
                 'WHERE DTP_NODOSSIER = "' + sNoDossier_p + '"';
   OBDPTabPaie_p.LoadDetailDBFromSQL('DPTABPAIE', sRequete_l);
   Trace('End : ChargeDPTabPaie');
end;

procedure TGenParser.ChargeDPTabGenPaie(var OBDPTabGenPaie_p : TOB; sNoDossier_p : string);
var
   sRequete_l : string;
begin
   Trace('Start : ChargeDPTabGenPaie');
   sRequete_l := 'SELECT * FROM DPTABGENPAIE ' +
                 'WHERE DT1_NODOSSIER = "' + sNoDossier_p + '"';
   OBDPTabGenPaie_p.LoadDetailDBFromSQL('DPTABGENPAIE', sRequete_l);

   Trace('End : ChargeDPTabGenPaie');
end;

procedure TGenParser.ChargeTiers(var OBTiers_p : TOB; sTiers_p : string);
begin
   Trace('Start : ChargeTiers');
   OBTiers_p.LoadDetailDBFromSQL('TIERS', 'SELECT * FROM TIERS WHERE T_TIERS = "' + sTiers_p + '"');
   Trace('End : ChargeTiers');
end;

procedure TGenParser.ChargeTiersCompl(var OBTiersCompl_p : TOB; sTiers_p : string);
begin
   Trace('Start : ChargeTiersCompl');
   OBTiersCompl_p.LoadDetailDBFromSQL('TIERSCOMPL', 'SELECT * FROM TIERSCOMPL WHERE YTC_TIERS = "' + sTiers_p + '"');
   Trace('End : ChargeTiersCompl');
end;

procedure TGenParser.ChargeProspects(var OBProspects_p : TOB; sAuxiliaire_p : string);
begin
   Trace('Start : ChargeProspects');
   OBProspects_p.LoadDetailDBFromSQL('PROSPECTS', 'SELECT * FROM PROSPECTS WHERE RPR_AUXILIAIRE = "' + sAuxiliaire_p + '"');
   Trace('End : ChargeProspects');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.GetListe(OBLaTob_p : TOB; sNomChamp_p : string; var tsListe_p : TStringList);
var
   nInd_l : integer;
begin
   for nInd_l := 0 to OBLaTob_p.Detail.Count-1 do
      tsListe_p.Add(OBLaTob_p.Detail[nInd_l].GetValue(sNomChamp_p));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 08/04/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.GetListeHisto(OBLaTob_p : TOB; sNomChamp_p : string;
            aosChamps_p, aosValeur_p : array of string;
            var tsListe_p : TStringList);
var
   nInd_l, nAos_l : integer;
   sValTob_l, sValParam_l : string;
   bOK_l : boolean;
begin
   for nInd_l := 0 to OBLaTob_p.Detail.Count-1 do
   begin
      bOK_l := true;
      nAos_l := 0;
      while bOk_l and (nAos_l < Length(aosChamps_p)) do
      begin
         sValTob_l := OBLaTob_p.Detail[nInd_l].GetValue(aosChamps_p[nAos_l]);
         sValParam_l := aosValeur_p[nAos_l];
         bOK_l := bOK_l and (sValTob_l = sValParam_l);
         Inc(nAos_l);
      end;
      if bOK_l then
         tsListe_p.Add(OBLaTob_p.Detail[nInd_l].GetValue(sNomChamp_p));
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/12/2003
Modifié le ... :   /  /
Description .. : Racine = ANL
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.GetListeLien(OBAnnuLien_p : TOB; sNomChamp_p : string; var tsListe_p : TStringList);
var
  nInd_l : integer;
  OBDetail_l : TOB;
begin
   Trace('Start : GetListeLien');
   nInd_l := 0;
   while nInd_l < OBAnnuLien_p.Detail.Count do
   begin
    OBDetail_l := OBAnnuLien_p.Detail[nInd_l];
     if OBDetail_l.GetValue('DETAIL_AFFECTE') = '0' then
       ChargeDetailLien(OBDetail_l);
     tsListe_p.Add(OBDetail_l.Detail[0].GetValue(sNomChamp_p));
    Inc(nInd_l);
   end;
   Trace('End : GetListeLien');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/09/2003
Modifié le ... :   /  /
Description .. : Liste des liens en fonction de la racine
Mots clefs ... :
*****************************************************************}

procedure TGenParser.GetListeLienRac(OBAnnuLien_p : TOB; sRacine_p,sNomChamp_p : string; var tsListe_p : TStringList);
var
  OBDetail_l : TOB;
begin
   Trace('Start : GetListeLienRac');
   OBDetail_l := OBAnnuLien_p.FindFirst(['ANL_RACINE'],[sRacine_p],true);
   while OBDetail_l <> nil do
   begin
     if OBDetail_l.GetValue('DETAIL_AFFECTE') = '0' then
       ChargeDetailLien(OBDetail_l);
     tsListe_p.Add(OBDetail_l.Detail[0].GetValue(sNomChamp_p));
     OBDetail_l := OBAnnuLien_p.FindNext(['ANL_RACINE'],[sRacine_p],true);
   end;
   Trace('End : GetListeLienRac');
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/09/2003
Modifié le ... :   /  /
Description .. : Récupération d'une variable AnnuLien en fonction de la racine
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarLienRac(OBAnnuLien_p : TOB; sRacine_p,sSuffixe_p : string;nIndex_p: integer) : Variant;
var
     OBDetail_l : TOB;
     nLienNb_l,nInd_l : integer;
     bAffecte_l : boolean;
     vValeur_l : Variant;
     sTypeChamp_l : string;
begin
   Trace('Start : GetVarLienRac');
   result := 0;
   nLienNb_l := 0;
   bAffecte_l := false;
   vValeur_l := '';

   // traite la demande du nombre total des liens d'un type particulier (ASS, IND, ACT, ....)
   if (sSuffixe_p = 'NBTOT') then
   begin
       // accede au premier lien de type recherché
       OBDetail_l := OBAnnuLien_p.FindFirst(['ANL_RACINE'], [sRacine_p], true);
       // boucle tantque jusqu'à ne plus trouver de lien du type recherché ; incrément d'un compteur
       while OBDetail_l <> nil do
       begin
          Inc(nLienNb_l);
          OBDetail_l := OBAnnuLien_p.FindNext(['ANL_RACINE'], [sRacine_p], true);
       end;
       // affectation de la valeur au resultat de la fonction
        result := nLienNb_l;
       Trace('Exit : GetVarLienRac = NBTOT');
        exit;
   end;

   // sinon, c'est un champ d'un lien
   // accede au premier lien de type recherché
   OBDetail_l := OBAnnuLien_p.FindFirst(['ANL_RACINE'],[sRacine_p],true);

   // tant qu'on récupère des liens du type recherché,
   // et tant que le champ recherché n'a pas été affecté
   nInd_l := 1;
   while (OBDetail_l <> nil) and (not bAffecte_l) do
   begin
      // on recherche le Ième lien correspondant à l'indice passé en paramètre par le parser
      if (nInd_l = nIndex_p) then
      begin
         // on est sur le lien recherché
         // si le detail du lien n'a pas été encore chargé, on le charge
         if OBDetail_l.GetValue('DETAIL_AFFECTE') = '0' then
            ChargeDetailLien(OBDetail_l);
         // si le champ fait partie de la table ANNULIEN (prefixe ANL)
         // on récupère le champ que l'on préfixe par 'ANL_' et son type
         if IsChampType('ANL', sSuffixe_p, sTypeChamp_l) then
            vValeur_l := OBDetail_l.GetValue('ANL_' + sSuffixe_p)
         // sinon, il s'agit d'un champ venant soit de la table ANNUAIRE, soit d'un champ supplémentaire
         // on récupère sa valeur et son type
         else
            vValeur_l := GetInfoPersonne(OBDetail_l, sSuffixe_p, sTypeChamp_l);
         // le champ est affecté, on arrète (bAffecte=true permet de sortir du While)
         bAffecte_l := true;
      end;
      // si champ pas affecte (bAffecte=false) on accède au lien du type recherche suivant
      if not bAffecte_l then
      begin
         OBDetail_l := OBAnnuLien_p.FindNext(['ANL_RACINE'],[sRacine_p],true);
      end;
      // on incrémente le compteur du Ième lien
      Inc(nInd_l);
   end;

   vValeur_l := AffecteChampSelonType( sTypeChamp_l, vValeur_l );
   result := vValeur_l;
   Trace('End : GetVarLienRac');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/12/2003
Modifié le ... :   /  /
Description .. : Racine = ANL
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarLien(OBAnnuLien_p : TOB; sSuffixe_p : string;nIndex_p: integer) : Variant;
var
     OBDetail_l : TOB;
     vValeur_l : Variant;
     sTypeChamp_l : string;
begin
   Trace('Start : GetVarLien');
   result := 0;
   vValeur_l := '';

   if (sSuffixe_p = 'NBTOT') then
   begin
     result := OBAnnuLien_p.Detail.Count;
     Trace('Exit : GetVarLien = NBTOT');
     exit;
   end;

   if nIndex_p <= OBAnnuLien_p.Detail.Count then
   begin
      OBDetail_l := OBAnnuLien_p.Detail[nIndex_p - 1];
      if OBDetail_l.GetValue('DETAIL_AFFECTE') = '0' then
         ChargeDetailLien(OBDetail_l);
      if IsChampType('ANL', sSuffixe_p, sTypeChamp_l) then
         vValeur_l := OBDetail_l.GetValue('ANL_' + sSuffixe_p)
      else
         vValeur_l := GetInfoPersonne(OBDetail_l, sSuffixe_p, sTypeChamp_l);
   end;
   vValeur_l := AffecteChampSelonType( sTypeChamp_l, vValeur_l );
   result := vValeur_l;
   Trace('End : GetVarLien');
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarDossier(OBDossier_p : TOB; sRacine_p,sSuffixe_p : string; nIndex_p: integer) : variant;
var
     sTypeChamp_l, sRacChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarDossier');
   sTypeChamp_l := '';
   vValeur_l := '';
   if ChampBlob(sSuffixe_p, sRacChamp_l ) then
      vValeur_l := GetValChampLiensOle(OBDossier_p.Detail[0], 'OLE_' + sSuffixe_p)
   else
   begin
        sTypeChamp_l := ChampToType('DOS_' + sSuffixe_p);
        vValeur_l := OBDossier_p.GetValue('DOS_' + sSuffixe_p);
   end;
   vValeur_l := AffecteChampSelonType( sTypeChamp_l, vValeur_l );
   if vValeur_l = Null then
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarDossier');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarEvenement(OBEvenement_p : TOB; sSuffixe_p : string; nIndex_p : integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarEvenement');
   sTypeChamp_l := '';
   vValeur_l := '';
   if (sSuffixe_p = 'NBTOT') then
   begin
      result := OBEvenement_p.Detail.Count;
      Trace('Exit : GetVarEvenement = NBTOT');
      exit;
   end;

   sTypeChamp_l := ChampToType('JEV_' + sSuffixe_p);
//   vValeur_l := GetVarGuid(OBEvenement_p.Detail[nIndex_p - 1], 'JEV_', sSuffixe_p);
   vValeur_l := OBEvenement_p.Detail[nIndex_p - 1].GetValue('JEV_' + sSuffixe_p);

   vValeur_l := AffecteChampSelonType( sTypeChamp_l, vValeur_l );
   if vValeur_l = Null then
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarEvenement');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarExeRubrique(OBDosExe_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : Variant;
var
     OBDetail_l, OBRub_l : TOB;
     vValeur_l : Variant;
     sTypeChamp_l : string;
begin
   Trace('Start : GetVarExeRubrique');
   vValeur_l := '';
   result := 0;
   // si pas de detail , pas d'exercice
   if OBDosExe_p.Detail.Count = 0 then
   begin
      HShowMessage('0;Questionnaire;Pas d''exercice pour ce dossier;W;O;O;O;','','');
      Trace('Exit : GetVarExeRubrique = Count');
      exit;
   end;
   // si Suf=NBTOT, renvoi le nombre d'exercies
   if (sSuffixe_p = 'NBTOT') then
   begin
      result := OBDosExe_p.Detail.Count;
      Trace('Exit : GetVarExeRubrique = NBTOT');
      exit;
   end;
   // on récupère le nIndex_p-1 ème exercice de la TOB car nIndex_p débute à l'indice 1
   if ( nIndex_p <= 0 ) or ( nIndex_p > OBDosExe_p.Detail.Count ) then
   begin
      Trace('Exit : GetVarExeRubrique = nIndex_p-1=-1');
      exit;
   end;

   OBDetail_l := OBDosExe_p.Detail[nIndex_p - 1];
   if OBDetail_l = nil then
   begin
      Trace('Exit : GetVarExeRubrique = OBDetail_l');
      exit;
   end;
   // si detail pas chargé, on le charge
   if OBDetail_l.GetValue('INIT_DETAIL') = 0 then
      ChargeJuDosExeDetails(OBDetail_l);
   // si le champ est de type 'JDE', c'est qu'on est au niveau Exercice (différent de detail)
   if IsChampType('JDE', sSuffixe_p, sTypeChamp_l) then
   begin
      vValeur_l := OBDetail_l.GetValue('JDE_' + sSuffixe_p);
   end
   else
   begin
      // sinon, c'est qu'on est au niveau du detail d'un exercice
      // on récupère le premier detail avec CodeRub=sSuffixe_p
       OBRub_l := OBDetail_l.FindFirst(['JER_CODERUB'],[sSuffixe_p],true);
      // si pas de detail
      if OBRub_l=nil then
      begin
         Trace('Exit : GetVarExeRubrique = OBRub_l');
         exit; // la rubrique n'existe pas !...?
      end;
      // recupere la vValeur_l du champ
       vValeur_l := OBRub_l.GetValue('JER_VALEUR');
   end;
   // converti si booleen le 'X' ou '-' en '1' ou '0'
   vValeur_l := AffecteChampSelonType( sTypeChamp_l, vValeur_l );
   if vValeur_l = Null then
      vValeur_l := 0;
   result := vValeur_l;
   Trace('End : GetVarExeRubrique');
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 03/11/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarAutreInfo(OBDosInfo_p : TOB; sRacine_p,sSuffixe_p : string;nIndex_p: integer) : Variant;
var
     OBDetail_l : TOB;
     nNb_l, nInd_l : integer;
     bAffecte_l : boolean;
     vValeur_l : Variant;
     sTypeChamp_l : string;
begin
   if (OBDosInfo_p = nil) then
      exit;

   Trace('Start : GetVarAutreInfo');
   nInd_l :=0;
   Result := 0;
   vValeur_l := '';
   nNb_l := 0;
   bAffecte_l := false;

   if (sSuffixe_p = 'NBTOT') then
   begin
     OBDetail_l := OBDosInfo_p.FindFirst(['JDI_RACINE'], [sRacine_p], true);
     while OBDetail_l <> nil do
     begin
        Inc(nNb_l);
        OBDetail_l := OBDosInfo_p.FindNext(['JDI_RACINE'], [sRacine_p], true);
     end;
     Result := nNb_l;
     exit;
   end;

   OBDetail_l := OBDosInfo_p.FindFirst(['JDI_RACINE'], [sRacine_p], true);
   while (OBDetail_l <> nil) and (not bAffecte_l) do
   begin
        if (nInd_l = nIndex_p - 1) then
       begin
          if OBDetail_l.GetValue('DETAIL_AFFECTE') = '0' then
             ChargeDetailInfo(OBDetail_l);

          if IsChampType('JDI', sSuffixe_p, sTypeChamp_l) then
//             vValeur_l := GetVarGuid(OBDetail_l, 'JDI_', sSuffixe_p)
             vValeur_l := OBDetail_l.GetValue('JDI_' + sSuffixe_p)
          else
             vValeur_l := GetInfoPersonne(OBDetail_l, sSuffixe_p, sTypeChamp_l);
          bAffecte_l := true;
        end;
      if not bAffecte_l then
           OBDetail_l := OBDosInfo_p.FindNext(['JDI_RACINE'], [sRacine_p], true);
        Inc(nInd_l);
   end;

   vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);

   Result := vValeur_l;
   Trace('End : GetVarAutreInfo');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarBauxFonds(OBBauxFonds_p : TOB; sRacine_p, sSuffixe_p : string;
                                    nIndex_p : integer) : Variant;
var
     OBDetail_l : TOB;
     vValeur_l : Variant;
     sTypeChamp_l : string;
begin
   if OBBauxFonds_p.Detail.Count = 0 then
   begin
      HShowMessage('0;Questionnaire;Pas de baux ou de fonds de commerce pour ce dossier;W;O;O;O;','','');
      Trace('Exit : GetVarBauxFonds = Count');
      exit;
   end;
   // si Suf=NBTOT, renvoi le nombre d'exercies
   if (sSuffixe_p = 'NBTOT') then
   begin
      result := OBBauxFonds_p.Detail.Count;
      Trace('Exit : GetVarBauxFonds = NBTOT');
      exit;
   end;
   // on récupère le nIndex_p-1 ème exercice de la TOB car nIndex_p débute à l'indice 1
   if ( nIndex_p <= 0 ) or ( nIndex_p > OBBauxFonds_p.Detail.Count ) then
   begin
      Trace('Exit : GetVarBauxFonds = nIndex_p-1=-1');
      exit;
   end;

   OBDetail_l := OBBauxFonds_p.Detail[nIndex_p - 1];
   if OBDetail_l = nil then
   begin
      Trace('Exit : GetVarExeRubrique = OBDetail_l');
      exit;
   end;
   // si detail pas chargé, on le charge
   // si le champ est de type 'JDE', c'est qu'on est au niveau Exercice (différent de detail)
   if IsChampType('JBF', sSuffixe_p, sTypeChamp_l) then
//      vValeur_l := GetVarGuid(OBDetail_l, 'JBF_', sSuffixe_p)
      vValeur_l := OBDetail_l.GetValue('JBF_' + sSuffixe_p)
   else
   begin
      exit; // n'existe pas !...?
   end;
   // converti si booleen le 'X' ou '-' en '1' ou '0'
   vValeur_l := AffecteChampSelonType( sTypeChamp_l, vValeur_l );
   if vValeur_l = Null then
      vValeur_l := 0;
   result := vValeur_l;
   Trace('End : GetVarBauxFonds');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 19/12/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarContact(OBContact_p : TOB; sRacine_p, sSuffixe_p : string;
                                    nIndex_p : integer) : Variant;
var
     OBDetail_l : TOB;
     vValeur_l : Variant;
     sTypeChamp_l : string;
begin
   if OBContact_p.Detail.Count = 0 then
   begin
      HShowMessage('0;Questionnaire;Pas de contacts pour cette personne;W;O;O;O;','','');
      Trace('Exit : GetVarContact = Count');
      exit;
   end;
   // si Suf=NBTOT, renvoi le nombre d'exercies
   if (sSuffixe_p = 'NBTOT') then
   begin
      result := OBContact_p.Detail.Count;
      Trace('Exit : GetVarContact = NBTOT');
      exit;
   end;
   // on récupère le nIndex_p-1 ème exercice de la TOB car nIndex_p débute à l'indice 1
   if ( nIndex_p <= 0 ) or ( nIndex_p > OBContact_p.Detail.Count ) then
   begin
      Trace('Exit : GetVarContact = nIndex_p-1=-1');
      exit;
   end;

   OBDetail_l := OBContact_p.Detail[nIndex_p - 1];
   if OBDetail_l = nil then
   begin
      Trace('Exit : GetVarContact = OBDetail_l');
      exit;
   end;

   if sRacine_p = 'ANI' then
   begin
      if sSuffixe_p = 'ASSIST'             then sSuffixe_p := 'TEXTELIBRE1'
      else if sSuffixe_p = 'CLETELEPHONE'  then sSuffixe_p := 'CLETELEPHONE'
      else if sSuffixe_p = 'GUIDPER'       then sSuffixe_p := 'AUXILIAIRE'
      else if sSuffixe_p = 'CV'            then sSuffixe_p := 'CIVILITE'
      else if sSuffixe_p = 'DATECREATION'  then sSuffixe_p := 'DATECREATION'
      else if sSuffixe_p = 'DATEMODIF'     then sSuffixe_p := 'DATEMODIF'
      else if sSuffixe_p = 'EMAIL'         then sSuffixe_p := 'EMAILING'
      else if sSuffixe_p = 'FAX'           then sSuffixe_p := 'FAX'
      else if sSuffixe_p = 'FONCTIONJ'     then sSuffixe_p := 'FONCTION'
      else if sSuffixe_p = 'NOCONT'        then sSuffixe_p := 'NUMEROCONTACT'
      else if sSuffixe_p = 'NOM'           then sSuffixe_p := 'NOM'
//      else if sSuffixe_p = 'NOMPER' then sSuffixe_p := ''
      else if sSuffixe_p = 'NOTECON'       then sSuffixe_p := 'BLOCNOTE'
      else if sSuffixe_p = 'PRENOM'        then sSuffixe_p := 'PRENOM'
      else if sSuffixe_p = 'PRINCIPAL'     then sSuffixe_p := 'PRINCIPAL'
      else if sSuffixe_p = 'SERVICE'       then sSuffixe_p := 'SERVICE'
      else if sSuffixe_p = 'TEL1'          then sSuffixe_p := 'TELEPHONE'
      else if sSuffixe_p = 'TEL2'          then sSuffixe_p := 'TELEX'
      else if sSuffixe_p = 'TELASSIST'     then sSuffixe_p := 'TEXTELIBRE2'
      else if sSuffixe_p = 'UTILISATEUR'   then sSuffixe_p := 'UTILISATEUR'
      else exit;
   end;

   if IsChampType('C', sSuffixe_p, sTypeChamp_l) then
//      vValeur_l := GetVarGuid(OBDetail_l, 'C_', sSuffixe_p)
      vValeur_l := OBDetail_l.GetValue('C_' + sSuffixe_p)
   else
   begin
      exit; // n'existe pas !...?
   end;
   // converti si booleen le 'X' ou '-' en '1' ou '0'
   vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);
   if vValeur_l = Null then
      vValeur_l := 0;
   result := vValeur_l;
   Trace('End : GetVarContact');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarHistoAct(OBHistoOpAct_p : TOB; sSuffixe_p, sCodeOp_p, sModule_p : string; nIndex_p : integer) : Variant;
var
   sTypeChamp_l : string;
   vValeur_l : variant;
   OBHistoDetail_l : TOB;
   nInd_l : integer;
begin
   Result := 0;
   Trace('Start : GetVarHistoAct');

   if (sSuffixe_p = 'NBTOT') then
   begin
    if (sCodeOp_p <> '') and (sModule_p <> '') then
    begin
       nInd_l := 0;
       OBHistoDetail_l := OBHistoOpAct_p.FindFirst(['JOA_CODEOP', 'JOA_MODULE'], [sCodeOp_p, sModule_p], false);
       while OBHistoDetail_l <> nil do
       begin
          Inc(nInd_l);
            OBHistoDetail_l := OBHistoOpAct_p.FindNext(['JOA_CODEOP', 'JOA_MODULE'], [sCodeOp_p, sModule_p], false);
       end;
       result := nInd_l;
    end
    else
    begin
        result := OBHistoOpAct_p.Detail.Count;
    end;
    Trace('Exit : GetVarHistoAct = NBTOT');
     exit;
   end
   else if OBHistoOpAct_p.Detail.Count = 0 then
   begin
      result := Unassigned;
      Trace('Exit : GetVarHistoAct = Unassigned');
      Exit;
   end;

   OBHistoDetail_l := OBHistoOpAct_p.FindFirst(['JOA_CODEOP', 'JOA_MODULE'], [sCodeOp_p, sModule_p], true);
   for nInd_l := 2 to nIndex_p do
   begin
      OBHistoDetail_l := OBHistoOpAct_p.FindNext(['JOA_CODEOP', 'JOA_MODULE'], [sCodeOp_p, sModule_p], true);
   end;

   if (OBHistoDetail_l = nil) then
   begin
      result := Unassigned;
      Trace('Exit : GetVarHistoAct = Unassigned');
      Exit;
   end;

   sTypeChamp_l := ChampToType('JOA_' + sSuffixe_p);
//   vValeur_l := GetVarGuid(OBHistoDetail_l, 'JOA_', sSuffixe_p);
   vValeur_l := OBHistoDetail_l.GetValue('JOA_' + sSuffixe_p);

   vValeur_l := AffecteChampSelonType( sTypeChamp_l, vValeur_l );
   if vValeur_l = Null then
      vValeur_l := '';
   result := vValeur_l;

   Trace('End : GetVarHistoAct');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarHistoRub(OBHistoOpRub_p : TOB; sSuffixe_p, sCodeOp_p, sModule_p : string; nIndex_p : integer) : Variant;
var
   OBHistoDetail_l : TOB;
   sTypeChamp_l : string;
   vValeur_l : variant;
   nInd_l : integer;
begin
   Result := 0;
   Trace('Start : GetVarHistoRub');

   if (sSuffixe_p = 'NBTOT') then
   begin
    if (sCodeOp_p <> '') and (sModule_p <> '') then
    begin
       nInd_l := 0;
       OBHistoDetail_l := OBHistoOpRub_p.FindFirst(['JOR_CODEOP', 'JOR_MODULE'], [sCodeOp_p, sModule_p], false);
       while OBHistoDetail_l <> nil do
       begin
          Inc(nInd_l);
          OBHistoDetail_l := OBHistoOpRub_p.FindNext(['JOR_CODEOP', 'JOR_MODULE'], [sCodeOp_p, sModule_p], false);
       end;
       result := nInd_l;
    end
    else
    begin
        result := OBHistoOpRub_p.Detail.Count;
    end;
    Trace('Exit : GetVarHistoRub = NBTOT');
     exit;
   end
   else if OBHistoOpRub_p.Detail.Count = 0 then
   begin
      result := Unassigned;
      Trace('Exit : GetVarHistoRub = Unassigned');
      Exit;
   end;

   OBHistoDetail_l := OBHistoOpRub_p.FindFirst(['JOR_CODEOP', 'JOR_MODULE'], [sCodeOp_p, sModule_p], true);
   for nInd_l := 2 to nIndex_p do
   begin
      OBHistoDetail_l := OBHistoOpRub_p.FindNext(['JOR_CODEOP', 'JOR_MODULE'], [sCodeOp_p, sModule_p], true);
   end;

   if (OBHistoDetail_l = nil) then
   begin
      result := Unassigned;
      Trace('Exit : GetVarHistoRub = Unassigned');
      Exit;
   end;

   if sSuffixe_p = 'VALEUR' then
   begin
      Result := GetBlobFusion(OBHistoDetail_l.GetValue('JOR_' + sSuffixe_p), OBHistoDetail_l.GetValue('JOR_TYPERUB'));
   end
   else
   begin
      sTypeChamp_l := ChampToType('JOR_' + sSuffixe_p);
      vValeur_l := OBHistoDetail_l.GetValue('JOR_' + sSuffixe_p);
      vValeur_l := AffecteChampSelonType( sTypeChamp_l, vValeur_l );
      if vValeur_l = Null then
         vValeur_l := '';
      result := vValeur_l;
   end;
   Trace('End : GetVarHistoRub');
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarOperation(OBDosOpe_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : Variant;
var
     vValeur_l : Variant;
     sTypeChamp_l : string;

begin
   Trace('Start : GetVarOperation');
   sTypeChamp_l := '';
   vValeur_l := '';
   if (sSuffixe_p = 'NBTOT') then
   begin
      result := OBDosOpe_p.Detail.Count;
      Trace('Exit : GetVarOperation = NBTOT');
      exit;
   end
   else if OBDosOpe_p.Detail.Count = 0 then
   begin
      result := Unassigned;
      Trace('Exit : GetVarOperation = Unassigned');
      Exit;
   end;

   sTypeChamp_l := ChampToType('JOP_' + sSuffixe_p);
   vValeur_l := OBDosOpe_p.Detail[nIndex_p - 1].GetValue('JOP_' + sSuffixe_p);
   vValeur_l := AffecteChampSelonType( sTypeChamp_l, vValeur_l );
   if vValeur_l = Null then
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarOperation');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 03/11/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarRubrique(OBDosOpRub_p, OBBibRubrique_p : TOB;
            sRacine_p, sSuffixe_p, sCodeOp_p, sModule_p : string; nIndex_p : integer) : Variant;
var
   OBBibDetail_l, OBRubDetail_l : TOB;
   nInd_l : integer;
begin
   Result := 0;
   Trace('Start : GetVarRubrique');

   if (sSuffixe_p = 'NBTOT') then
   begin
      if (sCodeOp_p <> '') and (sModule_p <> '') then
      begin
         nInd_l := 0;
         OBRubDetail_l := OBDosOpRub_p.FindFirst(['JOR_CODEOP', 'JOR_MODULE'], [sCodeOp_p, sModule_p], false);
         while OBRubDetail_l <> nil do
         begin
            Inc(nInd_l);
              OBRubDetail_l := OBDosOpRub_p.FindNext(['JOR_CODEOP', 'JOR_MODULE'], [sCodeOp_p, sModule_p], false);
         end;
         result := nInd_l;
      end
      else
      begin
          result := OBDosOpRub_p.Detail.Count;
      end;
      Trace('Exit : GetVarRubrique = NBTOT');
       exit;
   end
   else if OBDosOpRub_p.Detail.Count = 0 then
   begin
      result := Unassigned;
      Trace('Exit : GetVarRubrique = Unassigned');
      Exit;
   end;

   if (sCodeOp_p <> '') and (sModule_p <> '') then
   begin
        OBBibDetail_l := OBBibRubrique_p.FindFirst(['JBR_MODULE', 'JBR_CODERUB'], [sModule_p, sSuffixe_p], false);
        OBRubDetail_l := OBDosOpRub_p.FindFirst(['JOR_CODEOP', 'JOR_MODULE', 'JOR_CODERUB'], [sCodeOp_p, sModule_p, sSuffixe_p], false);
   end
   else
   begin
        OBBibDetail_l := OBBibRubrique_p.FindFirst(['JBR_CODERUB'], [sSuffixe_p], false);
        OBRubDetail_l := OBDosOpRub_p.FindFirst(['JOR_CODERUB'], [sSuffixe_p], false);
   end;

   Trace('Exit : GetVarRubrique : NIL');
   if (OBRubDetail_l = nil) or (OBBibDetail_l = nil) then
      exit;

   Result := GetBlobFusion(OBRubDetail_l.GetValue('JOR_VALEUR'), OBBibDetail_l.GetValue('JBR_TYPERUB'));
   Trace('End : GetVarRubrique');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 24/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarActe(OBDosOpAct_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer) : Variant;
var
     vValeur_l : Variant;
     sTypeChamp_l : string;
begin
   Trace('Start : GetVarActe');
   sTypeChamp_l := '';
   vValeur_l := '';
   if (sSuffixe_p = 'NBTOT') then
   begin
      result := OBDosOpAct_p.Detail.Count;
      Trace('Exit : GetVarActe = NBTOT');
      exit;
   end
   else if OBDosOpAct_p.Detail.Count = 0 then
   begin
      result := Unassigned;
      Trace('Exit : GetVarActe = Unassigned');
      exit;
   end;

   sTypeChamp_l := ChampToType('JOA_' + sSuffixe_p);
//   vValeur_l := GetVarGuid(OBDosOpAct_p.Detail[nIndex_p - 1], 'JOA_', sSuffixe_p);
   vValeur_l := OBDosOpAct_p.Detail[nIndex_p - 1].GetValue('JOA_' + sSuffixe_p);

   vValeur_l := AffecteChampSelonType( sTypeChamp_l, vValeur_l );
   if vValeur_l = Null then
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarActe');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarJuridique(OBJuridique_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : variant;
var
     sTypeChamp_l, sRacChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarJuridique');
   sTypeChamp_l := '';
   vValeur_l := '';
   if ChampBlob(sSuffixe_p, sRacChamp_l) then
      vValeur_l := GetValChampLiensOle(OBJuridique_p.Detail[0], 'OLE_' + sSuffixe_p)
   else
   begin
      sTypeChamp_l := ChampToType('JUR_' + sSuffixe_p);
//      vValeur_l := GetVarGuid(OBJuridique_p, 'JUR_', sSuffixe_p);
      vValeur_l := OBJuridique_p.GetValue('JUR_' + sSuffixe_p);
   end;
   vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);
   if vValeur_l = Null then
        vValeur_l := '';
   result := vValeur_l;

   Trace('End : GetVarJuridique');
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TGenParser.GetVarSociete(OBAnnuaire_p : TOB; sRacine_p,sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l, sRacChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarSociete');
   if ChampBlob(sSuffixe_p, sRacChamp_l) then
   begin
      vValeur_l := GetValChampLiensOle(OBAnnuaire_p.Detail[0], 'OLE_' + sSuffixe_p);
   end
   else if IsChampType('ANN', sSuffixe_p, sTypeChamp_l) then
   begin
//      vValeur_l := GetVarGuid(OBAnnuaire_p, 'ANN_', sSuffixe_p);
      vValeur_l := OBAnnuaire_p.GetValue('ANN_' + sSuffixe_p);
   end
   else
   begin
      vValeur_l := OBAnnuaire_p.GetValue(sSuffixe_p);
   end;

   vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);
   result := vValeur_l;
   Trace('End : GetVarSociete');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarAnnuBis(OBAnnuBis_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarAnnuBis');
   if OBAnnuBis_p.Detail.Count > 0 then
   begin
      if IsChampType('ANB', sSuffixe_p, sTypeChamp_l) then
//         vValeur_l := GetVarGuid(OBAnnuBis_p.Detail[0], 'ANB_', sSuffixe_p)
         vValeur_l := OBAnnuBis_p.Detail[0].GetValue('ANB_' + sSuffixe_p)
      else
         vValeur_l := OBAnnuBis_p.Detail[0].GetValue(sSuffixe_p);
      vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);
   end
   else
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarAnnuBis');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarDPFiscal(OBDPFiscal_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarDPFiscal');
   if OBDPFiscal_p.Detail.Count > 0 then
   begin
      if IsChampType('DFI', sSuffixe_p, sTypeChamp_l) then
         vValeur_l := OBDPFiscal_p.Detail[0].GetValue('DFI_' + sSuffixe_p)
      else
         vValeur_l := OBDPFiscal_p.Detail[0].GetValue(sSuffixe_p);
      vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);
   end
   else
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarDPFiscal');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarDPSocial(OBDPSocial_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarDPSocial');
   if OBDPSocial_p.Detail.Count > 0 then
   begin
      if IsChampType('DSO', sSuffixe_p, sTypeChamp_l) then
         vValeur_l := OBDPSocial_p.Detail[0].GetValue('DSO_' + sSuffixe_p)
      else
         vValeur_l := OBDPSocial_p.Detail[0].GetValue(sSuffixe_p);
      vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);
   end
   else
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarDPSocial');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarDPOrga(OBDPOrga_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarDPOrga');
   if OBDPOrga_p.Detail.Count > 0 then
   begin
      if IsChampType('DOR', sSuffixe_p, sTypeChamp_l) then
         vValeur_l := OBDPOrga_p.Detail[0].GetValue('DOR_' + sSuffixe_p)
      else
         vValeur_l := OBDPOrga_p.Detail[0].GetValue(sSuffixe_p);
      vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);
   end
   else
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarDPOrga');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarDPControl(OBDPControl_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarDPControl');
   if OBDPControl_p.Detail.Count > 0 then
   begin
      if IsChampType('DCL', sSuffixe_p, sTypeChamp_l) then
         vValeur_l := OBDPControl_p.Detail[0].GetValue('DCL_' + sSuffixe_p)
      else
         vValeur_l := OBDPControl_p.Detail[0].GetValue(sSuffixe_p);
      vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);
   end
   else
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarDPControl');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarTiers(OBTiers_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarTiers');
   if OBTiers_p.Detail.Count > 0 then
   begin
      if IsChampType('T', sSuffixe_p, sTypeChamp_l) then
         vValeur_l := OBTiers_p.Detail[0].GetValue('T_' + sSuffixe_p)
      else
         vValeur_l := OBTiers_p.Detail[0].GetValue(sSuffixe_p);
      vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);
   end
   else
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarTiers');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarTiersCompl(OBTiersCompl_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarTiersCompl');
   if OBTiersCompl_p.Detail.Count > 0 then
   begin
      if IsChampType('YTC', sSuffixe_p, sTypeChamp_l) then
         vValeur_l := OBTiersCompl_p.Detail[0].GetValue('YTC_' + sSuffixe_p)
      else
         vValeur_l := OBTiersCompl_p.Detail[0].GetValue(sSuffixe_p);
      vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);
   end
   else
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarTiersCompl');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarProspects(OBProspects_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarProspects');
   if OBProspects_p.Detail.Count > 0 then
   begin
      if IsChampType('RPR', sSuffixe_p, sTypeChamp_l) then
         vValeur_l := OBProspects_p.Detail[0].GetValue('RPR_' + sSuffixe_p)
      else
         vValeur_l := OBProspects_p.Detail[0].GetValue(sSuffixe_p);
      vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);
   end
   else
      vValeur_l := '';
   result := vValeur_l;
   Trace('End : GetVarProspects');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarDPTabCompta(OBDPTabCompta_p  : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarDPTabCompta');
   // si Suf=NBTOT, renvoi le nombre d'exercies
   if (sSuffixe_p = 'NBTOT') then
   begin
      result := OBDPTabCompta_p.Detail.Count;
      Trace('Exit : GetVarDPTabCompta = NBTOT');
      exit;
   end;
   if OBDPTabCompta_p.Detail.Count = 0 then
   begin
//      HShowMessage('0;Questionnaire;Pas d''exercice pour ce dossier;W;O;O;O;','','');
      vValeur_l := '';
      Trace('Exit : GetVarDPTabCompta = Count');
      exit;
   end;
   // on récupère le nIndex_p-1 ème exercice de la TOB car nIndex_p débute à l'indice 1
   if ( nIndex_p <= 0 ) or ( nIndex_p > OBDPTabCompta_p.Detail.Count ) then
   begin
      Trace('Exit : GetVarDPTabCompta = nIndex_p-1=-1');
      exit;
   end;

   // si detail pas chargé, on le charge
   //if OBDetail_l.GetValue('INIT_DETAIL') = 0 then
      //ChargeJuDPTabComptaDetails(OBDetail_l);

   if IsChampType('DTC', sSuffixe_p, sTypeChamp_l) then
      vValeur_l := OBDPTabCompta_p.Detail[nIndex_p - 1].GetValue('DTC_' + sSuffixe_p)
   else
      vValeur_l := OBDPTabCompta_p.Detail[nIndex_p - 1].GetValue(sSuffixe_p);
   // converti si booleen le 'X' ou '-' en '1' ou '0'
   vValeur_l := AffecteChampSelonType( sTypeChamp_l, vValeur_l );
   result := vValeur_l;
   Trace('End : GetVarDPTabCompta');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarDPTabPaie(OBDPTabPaie_p    : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarDPTabPaie');
   // si Suf=NBTOT, renvoi le nombre d'exercies
   if (sSuffixe_p = 'NBTOT') then
   begin
      result := OBDPTabPaie_p.Detail.Count;
      Trace('Exit : GetVarDPTabPaie = NBTOT');
      exit;
   end;
   if OBDPTabPaie_p.Detail.Count = 0 then
   begin
//      HShowMessage('0;Questionnaire;Pas d''exercice pour ce dossier;W;O;O;O;','','');
      vValeur_l := '';
      Trace('Exit : GetVarDPTabPaie = Count');
      exit;
   end;
   // on récupère le nIndex_p-1 ème exercice de la TOB car nIndex_p débute à l'indice 1
   if (nIndex_p <= 0) or (nIndex_p > OBDPTabPaie_p.Detail.Count) then
   begin
      Trace('Exit : GetVarDPTabPaie = nIndex_p-1=-1');
      exit;
   end;

   // si detail pas chargé, on le charge
   //if OBDetail_l.GetValue('INIT_DETAIL') = 0 then
      //ChargeJuDPTabPaieDetails(OBDetail_l);

   if IsChampType('DTP', sSuffixe_p, sTypeChamp_l) then
      vValeur_l := OBDPTabPaie_p.Detail[nIndex_p - 1].GetValue('DTP_' + sSuffixe_p)
   else
      vValeur_l := OBDPTabPaie_p.Detail[nIndex_p - 1].GetValue(sSuffixe_p);
   // converti si booleen le 'X' ou '-' en '1' ou '0'
   vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);

   result := vValeur_l;
   Trace('End : GetVarDPTabPaie');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 27/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TGenParser.GetVarDPTabGenPaie (OBDPTabGenPaie_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer) : Variant;
var
     sTypeChamp_l : string;
     vValeur_l : Variant;
begin
   Trace('Start : GetVarDPTabGenPaie');
   // si Suf=NBTOT, renvoi le nombre d'exercies
   if (sSuffixe_p = 'NBTOT') then
   begin
      result := OBDPTabGenPaie_p.Detail.Count;
      Trace('Exit : GetVarDPTabGenPaie = NBTOT');
      exit;
   end;
   if OBDPTabGenPaie_p.Detail.Count = 0 then
   begin
//      HShowMessage('0;Questionnaire;Pas d''exercice pour ce dossier;W;O;O;O;','','');
      vValeur_l := '';
      Trace('Exit : GetVarDPTabGenPaie = Count');
      exit;
   end;
   // on récupère le nIndex_p-1 ème exercice de la TOB car nIndex_p débute à l'indice 1
   if (nIndex_p <= 0) or (nIndex_p > OBDPTabGenPaie_p.Detail.Count) then
   begin
      Trace('Exit : GetVarDPTabGenPaie = nIndex_p-1=-1');
      exit;
   end;

   // si detail pas chargé, on le charge
   //if OBDetail_l.GetValue('INIT_DETAIL') = 0 then
      //ChargeJuDPTabGenPaieDetails(OBDetail_l);

   if IsChampType('DT1', sSuffixe_p, sTypeChamp_l) then
      vValeur_l := OBDPTabGenPaie_p.Detail[nIndex_p - 1].GetValue('DT1_' + sSuffixe_p)
   else
      vValeur_l := OBDPTabGenPaie_p.Detail[nIndex_p - 1].GetValue(sSuffixe_p);
   // converti si booleen le 'X' ou '-' en '1' ou '0'
   vValeur_l := AffecteChampSelonType(sTypeChamp_l, vValeur_l);

   result := vValeur_l;
   Trace('End : GetVarDPTabGenPaie');
end;






{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.SetVarDossier(var OBDossier_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer; vValeur_p : Variant);
var
     sTypeChamp_l : string;
begin
   Trace('Start : SetVarDossier');
   if IsChampType('DOS', sSuffixe_p, sTypeChamp_l) then
       OBDossier_p.PutValue('DOS_'+ sSuffixe_p, AffecteChampSelonType( sTypeChamp_l, vValeur_p ));
   Trace('End : SetVarDossier');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.SetVarEvenement  (var OBEvenement_p : TOB; sSuffixe_p : string; nIndex_p : integer; vValeur_p : Variant);
var
     sTypeChamp_l : string;
begin
   Trace('Start : SetVarEvenement');
   if IsChampType('JEV', sSuffixe_p, sTypeChamp_l) then
       OBEvenement_p.PutValue('JEV_'+ sSuffixe_p, AffecteChampSelonType( sTypeChamp_l, vValeur_p ));
   Trace('End : SetVarEvenement');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 12/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.SetVarBauxFonds(var OBBauxFonds_p : TOB; sSuffixe_p : string;
                                     nIndex_p : integer; vValeur_p : Variant);
var
     sTypeChamp_l : string;
begin
   Trace('Start : SetVarBauxFonds');
   if IsChampType('JBF', sSuffixe_p, sTypeChamp_l) then
       OBBauxFonds_p.PutValue('JBF_'+ sSuffixe_p, AffecteChampSelonType( sTypeChamp_l, vValeur_p ));
   Trace('End : SetVarBauxFonds');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.SetVarExeRubrique(var OBDosExe_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer; vValeur_p : Variant);
var
     OBDetail_l, OBRub_l : TOB;
     sTypeChamp_l : string;
begin
   Trace('Start : SetVarExeRubrique');
   if ( nIndex_p <= 0 ) or ( nIndex_p > OBDosExe_p.Detail.Count ) then
   begin
      Trace('Exit : SetVarExeRubrique = nIndex_p-1=-1');
      exit;
   end;
   // on récupère le nIndex_p-1 ème exercice de la TOB car nIndex_p débute à l'indice 1
   OBDetail_l := OBDosExe_p.Detail[nIndex_p - 1];
   if OBDetail_l = nil then
      exit;

   if IsChampType('JDE', sSuffixe_p, sTypeChamp_l) then
       OBDetail_l.PutValue('JDE_' + sSuffixe_p, AffecteChampSelonType( sTypeChamp_l, vValeur_p ))
   else
   begin
       OBRub_l := OBDetail_l.FindFirst(['JER_CODERUB'], [sSuffixe_p], true);
       if OBRub_l = nil then exit; // probleme : la rubrique n'existe pas
       OBRub_l.PutValue('JER_VALEUR', vValeur_p);
   end;
   Trace('End : SetVarExeRubrique');
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 03/11/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.SetVarAutreInfo(var OBInfos_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer; vValeur_p : Variant);
var
     OBDetail_l : TOB;
     nInd_l : integer;
     bAffecte_l : boolean;
     sTypeChamp_l : string;
begin
   Trace('Start : SetVarAutreInfo');
   nInd_l:=0;
   bAffecte_l := false;
   OBDetail_l := OBInfos_p.FindFirst(['JDI_RACINE'],[sRacine_p],true);
   while (OBDetail_l <> nil) and (not bAffecte_l) do
   begin
       if (nInd_l = nIndex_p - 1) then
       begin
         if IsChampType('JDI',sSuffixe_p,sTypeChamp_l) then
            OBDetail_l.Detail[0].PutValue('JDI_' + sSuffixe_p, vValeur_p)
         else
            SetInfoPersonne(OBDetail_l, sSuffixe_p, sTypeChamp_l, vValeur_p);
         bAffecte_l := true;
       end;
       OBDetail_l := OBInfos_p.FindNext(['JDI_RACINE'],[sRacine_p], true);
       Inc(nInd_l);
   end;
   Trace('End : SetVarAutreInfo');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 03/11/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.SetVarRubrique(var OBDosOpRub_p, OBBibRubrique_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p : integer; vValeur_p : Variant);
var
     OBBibDetail_l, OBRubDetail_l : TOB;
begin
   Trace('Start : SetVarOpeRubrique');
   OBBibDetail_l := OBBibRubrique_p.FindFirst(['JBR_CODERUB'],[sSuffixe_p],false);
   OBRubDetail_l := OBDosOpRub_p.FindFirst(['JOR_CODERUB'],[sSuffixe_p],false);
   if (OBBibDetail_l = nil) or (OBRubDetail_l = nil) then
    exit;

   if TypeDate(OBBibDetail_l.GetValue('JBR_TYPERUB')) then
    OBRubDetail_l.PutValue('JOR_VALEUR', DateTimeToStr(vValeur_p))
   else if TypeNombre(OBBibDetail_l.GetValue('JBR_TYPERUB')) then
    OBRubDetail_l.PutValue('JOR_VALEUR', VarAsType(vValeur_p,varString))
   else
    OBRubDetail_l.PutValue('JOR_VALEUR', VarAsType(vValeur_p,varString));

   OBRubDetail_l.PutValue('JOR_ATTENTE','-');
   Trace('End : SetVarOpeRubrique');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.SetVarJuridique(var OBJuridique_p : TOB; sRacine_p, sSuffixe_p : string ; nIndex_p : integer; vValeur_p : variant);
var
     sTypeChamp_l, sRacChamp_l : string;
begin
   Trace('Start : SetVarJuridique');
   if ChampBlob(sSuffixe_p, sRacChamp_l) then
   begin
      //vValeur_l := GetValChampLiensOle(OBJuridique_p.Detail[0], 'OLE_' + sSuffixe_p);
      SetValChampLiensOle(OBJuridique_p, 'OLE_' + sSuffixe_p, vValeur_p);
   end
   else
   begin
      if IsChampType('JUR', sSuffixe_p, sTypeChamp_l) then
        OBJuridique_p.PutValue('JUR_' + sSuffixe_p, AffecteChampSelonType( sTypeChamp_l, vValeur_p ));
   end;
   Trace('End : SetVarJuridique');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParser.SetVarSociete(var OBAnnuaire_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer; vValeur_p : Variant);
var
     sTypeChamp_l : string;
begin
   Trace('Start : SetVarSociete');
   if IsChampType('ANN',sSuffixe_p,sTypeChamp_l) then
    OBAnnuaire_p.PutValue('ANN_'+sSuffixe_p, AffecteChampSelonType( sTypeChamp_l, vValeur_p ));

   if sSuffixe_p='MOISCLOTURE' then
   begin
     OBAnnuaire_p.Putvalue('EXEDEB', RechDomPersonne(OBAnnuaire_p, 'JUMOISCLOTLIB', 'ANN_MOISCLOTURE'));
     OBAnnuaire_p.Putvalue('EXEFIN', RechDomPersonne(OBAnnuaire_p, 'JUMOISCLOTABR', 'ANN_MOISCLOTURE'));
   end;
   Trace('End : SetVarSociete');
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/09/2003
Modifié le ... :   /  /
Description .. : Mise à jour d'une variable AnnuLien en fonction de la racine
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.SetVarLienRac(var OBAnnuLien_p : TOB; sRacine_p, sSuffixe_p : string; nIndex_p: integer; vValeur_p : Variant);
var
     OBDetail_l : TOB;
     sTypeChamp_l : string;
     nInd_l : integer;
     bAffecte_l : boolean;
begin
   Trace('Start : SetVarLienRac');
   nInd_l:=0;
   bAffecte_l := false;
   OBDetail_l := OBAnnuLien_p.FindFirst(['ANL_RACINE'], [sRacine_p], true);
   while (OBDetail_l <> nil) and (not bAffecte_l) do
   begin
      if (nInd_l = nIndex_p - 1) then
      begin
         if OBDetail_l.GetValue('DETAIL_AFFECTE') = '0' then
            ChargeDetailLien(OBDetail_l);
         if IsChampType('ANL', sSuffixe_p, sTypeChamp_l) then
            OBDetail_l.PutValue('ANL_'+ sSuffixe_p, AffecteChampSelonType( sTypeChamp_l, vValeur_p ))
         else
            SetInfoPersonne(OBDetail_l, sSuffixe_p, sTypeChamp_l, vValeur_p);
         bAffecte_l := true;
      end;
      OBDetail_l := OBAnnuLien_p.FindNext(['ANL_RACINE'], [sRacine_p],true);
      Inc(nInd_l);
   end;
   Trace('End : SetVarLienRac');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 23/12/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.SetVarLien(var OBAnnuLien_p : TOB; sSuffixe_p : string; nIndex_p: integer; vValeur_p : Variant);
var
     OBDetail_l : TOB;
     sTypeChamp_l : string;
begin
   Trace('Start : SetVarLien');
   if nIndex_p - 1 < OBAnnuLien_p.Detail.Count then
   begin
      OBDetail_l := OBAnnuLien_p.Detail[nIndex_p - 1];
      if OBDetail_l.GetValue('DETAIL_AFFECTE') = '0' then
         ChargeDetailLien(OBDetail_l);
      if IsChampType('ANL', sSuffixe_p, sTypeChamp_l) then
         OBDetail_l.PutValue('ANL_'+ sSuffixe_p, AffecteChampSelonType( sTypeChamp_l, vValeur_p ))
      else
         SetInfoPersonne(OBDetail_l, sSuffixe_p, sTypeChamp_l, vValeur_p);
   end;
   Trace('End : SetVarLien');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 15/09/2003
Modifié le ... :   /  /    
Description .. : Débugage
Mots clefs ... : 
*****************************************************************}
procedure TGenParser.Trace(debug: string);
begin
   _S.Add(FormatDateTime('hh:nn:ss:zzz',Time)+ ' = ' + debug);
end;

end.
