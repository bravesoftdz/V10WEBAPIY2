{***********UNITE*************************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 30/10/2003
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
unit UDocumentDepuisModele;
//////////////////////////////////////////////////////////////////
interface
//////////////////////////////////////////////////////////////////
uses
   classes, UTOB, HCtrls,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
   CLASSGenParser;
//////////////////////////////////////////////////////////////////
type
   TGenParserDP = class(TGenParser)

         procedure InitVarParser( sGuidPer_p, sTypeDos_p, sModeleDoc_p, sKey_p : string);
         procedure OnGetVar  ( Sender : TObject; sVarName_p   : String; nIndex_p : Integer; var vValeur_p: variant);
         procedure OnSetVar  ( Sender : TObject; sVarName_p   : String; nIndex_p : Integer; vValeur_p: variant);
         procedure OnGetList ( Sender : TObject; sIdentList_p : String; var tsListe_p : TStringList );
         procedure OnFunction( Sender : TObject; sFuncName_p  : String; aovParams_p : Array of variant; var vValeur_p : Variant );

         procedure OnInitialization( Sender: TObject );
         procedure OnFinalization( Sender: TObject );

      public
      private

         OBDossier_c :TOB;
         OBDPFiscal_c, OBDPSocial_c, OBDPOrga_c, OBTiers_c, OBTiersCompl_c    :TOB;
         OBProspects_c                                                        :TOB;

         OBJuridique_c, OBAnnuaire_c, OBAnnuBis_c, OBAnnuLien_c, OBDosInfo_c, OBEvenement_c, OBDosExe_c: TOB;
         OBDosOper_c, OBDosOpRub_c, OBDosOpAct_c, OBBibRubrique_c : TOB;
         OBBauxFonds_c, OBContact_c : TOB;
         //OBDPTabCompta_c, OBDPTabPaie_c, OBDPTabGenPaie_c, OBDPControl_c                :TOB;


         sJuCodeDos_c, sGuidPer_c, sTypeDos_c, sCle_c : string;
         sNoOperation_c, sModule_c : string;
         sNoDossier_c, sNoDP_c, sTiers_c, sAuxiliaire_c: string;

         // $$$ JP 06/04/06 - FQ 10812, plus de DPTAB... et DPCONTROL

         function TraiteFusion : boolean;

   end;
//////////////////////////////////////////////////////////////////
function FusionneModele( sGuidPer_p, sTypeDos_p, sKey_p, sNomModele_p : string) : string;
//////////////////////////////////////////////////////////////////
implementation
//////////////////////////////////////////////////////////////////
uses
   UtilDossier, DOC_Parser, SysUtils, HMsgBox, windows, ParamSoc,
   DpJurOutils, HEnt1, dpJurOutilsGen;
   
//////////////////////////////////////////////////////////////////

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function FusionneModele(sGuidPer_p, sTypeDos_p, sKey_p, sNomModele_p : string) : string;
var
     GenFusionDP : TGenParserDP;
BEGIN
     GenFusionDP := TGenParserDP.Create ;
     Result := '';
     if (GenFusionDP<>nil)then
     begin
      // Récupération des paramètres passés en argument
       GenFusionDP._S := TStringList.Create;
       GenFusionDP.Trace('=============DEBUT============');
       GenFusionDP.InitVarParser(sGuidPer_p, sTypeDos_p, sNomModele_p, sKey_p );
       try
         // appel à la fonction de traitement du document
         // la variable sDocument_l sera initialisée avec le nom du document temporaire produit
         GenFusionDP.Trace('Call : TraiteFusionModele');
         if GenFusionDP.TraiteFusion then
         begin
            // si la fusion s'est bien passée,
            { FQ 11520 : ne pas voir le résultat fusionné, sinon on ne peut pas l'intégrer en GED
            GenFusionDP.Trace('Call : LanceAppliMaquette');
            LanceAppliMaquette(GenFusionDP.sDocument_c, GenFusionDP.sExtModele_c); }

            // on retourne le nom du document produit
            Result := GenFusionDP.sDocument_c;
         end;
       Finally
          GenFusionDP.Trace('=============FIN============');
          GenFusionDP._S.SaveToFile('DPDebugParser.txt');
          GenFusionDP._S.Free;
          GenFusionDP.Free ;
       End ;
     end;
end;

//////////////////////////////////////////////////////////////////
// CLASSE TGENPARSER                                            //
//////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. : Initialisation des variables
Mots clefs ... :
*****************************************************************}
procedure TGenParserDP.InitVarParser( sGuidPer_p, sTypeDos_p, sModeleDoc_p, sKey_p : string);
var i : Integer;
    sNomFichier : String;
begin
   Trace('Start : InitVarParser');
   sGuidPer_c := sGuidPer_p;
   sTypeDos_c := sTypeDos_p;
   sCle_c := sKey_p;
   sModeleDoc_c := sModeleDoc_p;
   // Extension (avec le point)
   sExtModele_c := UpperCase(ExtractFileExt(sModeleDoc_c));
   // Nom du fichier sans extension
   sNomFichier := ExtractFileName(sModeleDoc_c);
   sNomFichier := Copy(sNomFichier, 1, Length(sNomFichier)-Length(sExtModele_c));
   // sDocument_c := GetParamSocSecur('SO_MDREPDOCUMENTS', 'C:\PGI00\STD\BUREAU') + '\Res_' + ExtractFileName(sModeleDoc_c);
   // MD 09/10/07 : mieux vaut fusionner dans le répertoire temporaire,
   // en gardant un nom qui ressemble au nom du modèle (au lieu de Res_xxxx)
   sDocument_c := ExtractFilePath(sModeleDoc_c) + sNomFichier + '_res' + sExtModele_c  ;
   i := 0;
   While (FileExists(sDocument_c)) and (i<100) do
   begin
     i := i + 1 ;
     sDocument_c := ExtractFilePath(sModeleDoc_c) + sNomFichier + '_res' + IntToStr(i) + sExtModele_c;
   end;
   Trace('End : InitVarParser');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. : Cette fonction appel le PARSER pour traiter le document
Mots clefs ... :
*****************************************************************}
function TGenParserDP.TraiteFusion : boolean;
var
   sModeleDocTmp_l, sDocumentTmp_l : string;
begin
   Trace('Start : TraiteFusion');
   Result := false;

   // On copie le modèle original dans le temporaire qui sera utilisé à la place du modèle...
   // ... et détruit à la fin
   sDocumentTmp_l := TempFileName;
   sModeleDocTmp_l := TempFileName;
   CopyFile(PChar(sModeleDoc_c), PChar(sModeleDocTmp_l), false);

   // initialisation des variables qui permettront d'évaluer si le traitement s'est bien passé
   Result := true;
   bDone_c := true;
   try
      // Appel à la fonction de l'agl (Parser) de fusion
      // passage des paramètres : document initial, document final, liste des fonctions
      // de traitement des informations pour le Parser
      Trace('Call : ConvertDocFile');
      if ((sExtModele_c = cstAppliWord_g) or (sExtModele_c = cstAppliNotePad_g)) then
         ConvertDocFile( sModeleDocTmp_l, sDocumentTmp_l, OnInitialization, OnFinalization, OnFunction, OnGetVar, OnSetVar, OnGetList )
      else
           CopyFile( PChar(sModeleDocTmp_l), PChar(sDocumentTmp_l), false);
   except
      // si echec
      Trace('Except : TraiteFusion');
      MessageAlerte('Fusion du document impossible') ;
      // Suppression des fichiers temporaires
      if FileExists(sModeleDocTmp_l) then
         Deletefile(PChar(sModeleDocTmp_l));
      if FileExists(sDocumentTmp_l) then
         Deletefile(PChar(sDocumentTmp_l));

      Result := false;
   end;

   CopyFile( PChar(sDocumentTmp_l), PChar(sDocument_c), false);
   // Suppression des fichiers temporaires
   if FileExists(sModeleDocTmp_l) then
      Deletefile(PChar(sModeleDocTmp_l));
   if FileExists(sDocumentTmp_l) then
      Deletefile(PChar(sDocumentTmp_l));

   Trace('End : TraiteFusion');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TGenParserDP.OnInitialization( Sender: TObject );
var Q : TQuery;
Begin
   Trace('Start : OnInitialization');
   Try
   begin
      Trace('Table DOSSIER');
      OBDossier_c := TOB.Create('DOSSIER',nil,-1);
      if not ChargeDossier(OBDossier_c, sGuidPer_c, '') then exit;
      sNoDossier_c := OBDossier_c.GetValue('DOS_NODOSSIER');
      sNoDP_c := OBDossier_c.GetString('DOS_GUIDPER');

      Trace('Table JURIDIQUE');
      OBJuridique_c := TOB.Create('JURIDIQUE',nil,-1);

      Trace('Table ANNUAIRE');
      OBAnnuaire_c := TOB.Create('ANNUAIRE', nil, -1);
      if not ChargeAnnuaire(OBAnnuaire_c, sGuidPer_c) then exit;
      sTiers_c := OBAnnuaire_c.GetString('ANN_TIERS');
      // FQ 11883
      sAuxiliaire_c := '';
      if sTiers_c<>'' then
      begin
        Q := OpenSQL('SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+sTiers_c+'"', True);
        if Not Q.Eof then sAuxiliaire_c := Q.FindField('T_AUXILIAIRE').AsString;
        Ferme (Q);
      end;

      {*******
      // CHARGEMENT DES DONNEES DU DOSSIER
      // Tables : ANNULIEN, JUDOSINFO, JUDOSEX, JUBIBRUBRIQUE,JUDOSOPRUB,JUDOSOPER
      *******}
      // Création de la tob qui contiendra les infos
      // chargé lors du premier appel
      Trace('Table ANNUBIS');
      OBAnnuBis_c := TOB.Create('ANNUBIS ', nil, -1);
      Trace('Table ANNULIEN');
      OBAnnuLien_c := TOB.Create('les liens', nil, -1);
      Trace('Table JUDOSEX');
      OBDosExe_c := TOB.Create ('JUDOSEX', nil, -1); // $$$ JP 21/04/06 - d'après BM - OBDosExe_c := TOB.Create('ANNUAIRE', nil, -1);
      Trace('Table JUDOSINFO');
      OBDosInfo_c := TOB.Create('JUDOSINFO', nil, -1);
      Trace('Table JUEVENEMENT');
      OBEvenement_c := TOB.Create('JUEVENEMENT', nil, -1);
      Trace('Table JUDOSOPER');
      OBDosOper_c := TOB.Create('JUDOSOPER', nil, -1);
      Trace('Table JUDOSOPRUB');
      OBDosOpRub_c := TOB.Create('JUDOSOPRUB', nil, -1);
      Trace('Table JUDOSOPACT');
      OBDosOpAct_c := TOB.Create('JUDOSOPACT', nil, -1);
      Trace('Table JUBIBRUBRIQUE');
      OBBibRubrique_c := TOB.Create('JUBIBRUBRIQUE', nil, -1);
      OBBibRubrique_c.LoadDetailDB('JUBIBRUBRIQUE', '', 'JBR_TRI', nil, false);

      OBDPFiscal_c := TOB.Create('DPFISCAL', nil, -1);
      OBDPSocial_c := TOB.Create('DPSOCIAL', nil, -1);
      OBDPOrga_c := TOB.Create('DPORGA', nil, -1);
      OBTiers_c := TOB.Create('TIERS', nil, -1);
      OBTiersCompl_c := TOB.Create('TIERSCOMPL', nil, -1);
      // $$$ JP 06/04/06 - FQ 10812 corrigée - OBDPTabCompta_c := TOB.Create('TABCOMPTA', nil, -1);
      // $$$ JP 06/04/06 - FQ 10812 corrigée - OBDPTabPaie_c := TOB.Create('TABPAIE', nil, -1);
      // $$$ JP 06/04/06 - FQ 10812 corrigée - OBDPTabGenPaie_c := TOB.Create('TABGENPAIE', nil, -1);
      // $$$ JP 06/04/06 - FQ 10812 corrigée - OBDPControl_c := TOB.Create('DPCONTROL', nil, -1);
      OBProspects_c := TOB.Create('PROSPECTS', nil, -1);
      OBContact_c := TOB.Create('CONTACT', nil, -1);

   end;
   except
      MessageAlerte('L''application ne parvient pas à initialiser les informations utiles à la fusion ( Initialisation )');
   end;
   Trace('End : OnInitialization');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. : Fonction de gestion des listes, appelée par le Parser dans le
Suite ........ : cas de demande de liste dans le document : [LISTEVAR,...
Mots clefs ... :
*****************************************************************}
procedure TGenParserDP.OnGetList( Sender: TObject; sIdentList_p: String;var tsListe_p: TStringList );
var
   sRacine_l, sSuffixe_l : string;
begin
   Trace('Start : OnGetList');
   Try
   begin
      // découpage du nom de la liste pour connaitre la table sur laquelle porte la liste
        RecupereRacineSuffixe(sIdentList_p, sRacine_l, sSuffixe_l);
      // sRacine_p : racine (prefixe) / sSuffixe_p : nom de la liste à proprement parler
      tsListe_p.Clear;

      if VarDossier(sRacine_l) then   // (sRacine_l=DOS)
      begin
      end
      else if VarJuridique(sRacine_l) then    //(sRacine_l=JUR)
      begin
      end
      else if VarOperation(sRacine_l) then    // DOP
      begin
         if OBDosOper_c.Detail.Count = 0 then  //On n'en charge qu'une seule : sNoOperation_c
         begin
            ChargeJuDosOper(OBDosOper_c, sJuCodeDos_c, sNoOperation_c);
            sModule_c := OBDosOper_c.Detail[0].GetValue('JOP_MODULE');
         end;

         GetListe(OBDosOper_c, 'JOP_' + sSuffixe_l, tsListe_p);
      end
      else if VarRubrique(sRacine_l) then  // sRacine_l = DOR
      begin
      // si besoin chargement du detail
         if OBDosOpRub_c.Detail.Count = 0 then
            ChargeJuDosOpRub(OBDosOpRub_c, sJuCodeDos_c, sNoOperation_c, '');

         GetListe(OBDosOpRub_c, 'JOR_' + sSuffixe_l, tsListe_p);
      end
      else if VarActe(sRacine_l) then  // sRacine_l = DOR
      begin
      // si besoin chargement du detail
         if OBDosOpAct_c.Detail.Count = 0 then
            ChargeJuDosOpAct(OBDosOpAct_c, sJuCodeDos_c, sNoOperation_c, '');

         GetListe(OBDosOpAct_c, 'JOA_' + sSuffixe_l, tsListe_p);
      end

      else if VarBauxFonds(sRacine_l) then // cherche si sRacine_l dans JUBAUXFONDS
      begin
         if OBBauxFonds_c.Detail.Count = 0 then
            ChargeJuBauxFonds(OBBauxFonds_c, sJuCodeDos_c);

         GetListe(OBBauxFonds_c, 'JBF_' + sSuffixe_l, tsListe_p);
      end

      else if VarAutreInfos(sRacine_l) then    // (on cherche sRacine_l dans la table JUTYPEINFO)
      begin
         if OBDosInfo_c.Detail.Count = 0 then
            ChargeJuDosInfo(OBDosInfo_c, sJuCodeDos_c);

         GetListeHisto(OBDosInfo_c, 'JDI_' + sSuffixe_l, ['JDI_RACINE'], [sRacine_l], tsListe_p);
      end

      else if VarEvenement(sRacine_l) then
      begin
         if OBEvenement_c.Detail.Count = 0 then
            ChargeJuEvenement(OBEvenement_c, sGuidPer_c);
         GetListe(OBEvenement_c, 'JEV_' + sSuffixe_l, tsListe_p);
      end

      else if VarExercice(sRacine_l) then     //(sRacine_l=EXE)
      begin
         // si les exercices n'ont pas été encore chargés, on les charge
         if OBDosExe_c.Detail.Count = 0 then
            ChargeJuDosExe(OBDosExe_c, sJuCodeDos_c);
         // on appel la fonction d'initialisation de la liste
         GetListe(OBDosExe_c, 'JDE_' + sSuffixe_l, tsListe_p);
      end

      // $$$ JP 06/04/06 - FQ 10812 corrigée
      {else if VarDPTabCompta(sRacine_l) then
      begin
           if OBDPTabCompta_c.Detail.Count = 0 then
            ChargeDPTabCompta(OBDPTabCompta_c, sNoDossier_c);
          GetListe(OBDPTabCompta_c, 'DTC_' + sSuffixe_l, tsListe_p);
      end
      else if VarDPTabGenPaie(sRacine_l) then
      begin
         if OBDPTabGenPaie_c.Detail.Count = 0 then
            ChargeDPTabGenPaie(OBDPTabGenPaie_c, sNoDossier_c);
         GetListe(OBDPTabGenPaie_c, 'DT1_' + sSuffixe_l, tsListe_p);
      end
      else if VarDPTabPaie(sRacine_l) then
      begin
         if OBDPTabPaie_c.Detail.Count = 0 then
            ChargeDPTabPaie(OBDPTabPaie_c, sNoDossier_c);
         GetListe(OBDPTabPaie_c, 'DTP_' + sSuffixe_l, tsListe_p);
      end}

      else
      begin
         if OBAnnuLien_c.Detail.Count = 0 then
            ChargeAnnuLien(OBAnnuLien_c, sGuidPer_c, sTypeDos_c);
         // on appel la fonction d'initialisation de la liste
         if VarLien(sRacine_l) then
            GetListeLien(OBAnnuLien_c, sSuffixe_l, tsListe_p)
         else
            GetListeLienRac(OBAnnuLien_c, sRacine_l, sSuffixe_l, tsListe_p);
      end;
   end;
   except
      MessageAlerte('L''application ne parvient pas à établir la liste des informations qui lui sont utiles pour la fusion'+
                      #13#10'( Liste : ' + sIdentList_p + ' )');
   end;
   Trace('End : OnGetList');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 15/09/2003
Modifié le ... : 15/09/2003
Description .. : Fonction de gestion des fonctions, appelée par le Parser
Suite ........ : dans le cas de demande de fonction dans le document :
Suite ........ : @NOMFONCTION
Suite ........ : ATTENTION : aovParams_p[0] contient le nombre de paramètres
Suite ........ : dans aovParams_p : par exemple @EXISTELIEN(codeper,'ASS')
Suite ........ : - aovParams_p[0]=2 (on a 2 paramètres)
Suite ........ : - aovParams_p[1]=codeper
Suite ........ : - aovParams_p[2]='ASS'
Mots clefs ... :
*****************************************************************}
procedure TGenParserDP.OnFunction( Sender: TObject; sFuncName_p: String; aovParams_p: Array of variant ; var vValeur_p : Variant );
var
   tvParChamp_l : Array [0..2] of variant;
   tvParVal_l : Array [0..2] of variant;
   nCpt_l, nCpt2_l : integer;
begin
   Trace('Start : OnFunction');
   vValeur_p := true;
   // on teste le nom de la fonction

   if sFuncName_p='@EXISTELIEN' then
      vValeur_p := AppelFonction_ExisteLien(OBAnnuLien_c, String(aovParams_p[1]), String(aovParams_p[2]))
   else if sFuncName_p='@CHAINE' then
      vValeur_p := AppelFonction_Chaine(String(aovParams_p[1]))
   else if sFuncName_p='@PGCD' then  //Plus grand dénominateur commun
      vValeur_p := AppelFonction_PGCD(String(aovParams_p[1]), String(aovParams_p[2]))
   else if sFuncName_p = '@VOIRCHAMP' then
   begin
      nCpt_l := 3;
      nCpt2_l := 1;
      while nCpt_l < aovParams_p[0] do
      begin
         tvParChamp_l[nCpt2_l] := string(aovParams_p[nCpt_l]);
         tvParVal_l[nCpt2_l] := string(aovParams_p[nCpt_l + 1]);
         nCpt2_l := nCpt2_l + 1;
         nCpt_l := nCpt_l + 2;
      end;

      tvParChamp_l[0] := nCpt2_l - 1;
      tvParVal_l[0] := nCpt2_l - 1;

      vValeur_p := AppelFonction_VoirChamp( String(aovParams_p[1]), String(aovParams_p[2]), tvParChamp_l, tvParVal_l);
   end
   else if sFuncName_p = '@GETBIBLE' then
   begin
      vValeur_p := AppelFonction_GetBible(String(aovParams_p[1]), String(aovParams_p[2]));
   end
   else
      MessageAlerte('L''application ne parvient pas à établir la liste des informations qui lui sont utiles pour la fusion'+
            #13#10'( Fonction : ' + sFuncName_p + ' )');
   Trace('End : OnFunction');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. : Fonction de recupération des variables, appelée par le
Suite ........ : parser dans le cas où le document fait appel à des variables
Suite ........ : de la base (variable entre [] sans $)
Suite ........ : passage des paramètres :
Suite ........ : - nom de la variable (sVarName_p),
Suite ........ : - eventuellement l'indice, pour le Ieme
Suite ........ : - la valeur à récupérer
Mots clefs ... :
*****************************************************************}
procedure TGenParserDP.OnGetVar( Sender: TObject; sVarName_p: String; nIndex_p: Integer;var vValeur_p: variant) ;
var
     sRacine_l, sSuffixe_l, sTypeChp_l : string;
begin
   Trace('Start : OnGetVar');
   Trace('OnGetVar : sVarName_p['+IntToStr(nIndex_p)+'] = '+sVarName_p);
   Try
   begin
      // cas du symbole euro :
      if sVarName_p='SYMBOLE_EURO' then
      begin
         vValeur_p:='€';
         exit;
      end;
      sTypeChp_l := '';

      // recupère le prefixe (sRacine_p) et le nom (sSuffixe_l) de la variable
      Trace('OnGetVar : Call RecupereRacineSuffixe');
      RecupereRacineSuffixe(sVarName_p,sRacine_l,sSuffixe_l);
      Trace('OnGetVar : End Call' + ' ' + sRacine_l + ' ' + sSuffixe_l );

      // s'il s'agit d'une variable de paramètres societe (ParamSoc)
      if sRacine_l='SO' then
      begin
          Trace('OnGetVar : GetParamSocSecur');
          vValeur_p := GetParamSocSecur(sVarName_p, '');
          Trace('OnGetVar : End Call GetParamSocSecur');
      end
      ////
      else if VarDossier(sRacine_l) then // Rac  = MDP
      begin
         Trace('OnGetVar : Call GetVarDossier');
         vValeur_p := GetVarDossier(OBDossier_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call');
      end
      ////
      else if VarJuridique(sRacine_l) then // Rac  = DOS
      begin
         if sJuCodeDos_c = '' then
         begin
            Trace('OnGetVar : Call ChargeJuridique');
            // S'il existe un dossier juridique
            if ChargeJuridique(OBJuridique_c, '', '', sGuidPer_c + ';"STE";1') then
               sJuCodeDos_c := OBJuridique_c.GetValue('JUR_CODEDOS');
            Trace('OnGetVar : End Call ChargeJuridique');
         end;
         //récupération de la valeur de la variable
         if sJuCodeDos_c <> '' then  // Il existe un dossier juridique
         begin
            Trace('OnGetVar : Call GetVarJuridique');
            vValeur_p := GetVarJuridique(OBJuridique_c, sRacine_l, sSuffixe_l, nIndex_p);
            Trace('OnGetVar : End Call');
         end;
      end
      ////
      else if VarSociete(sRacine_l) then // Rac = STE
      begin
         Trace('OnGetVar : Call GetVarSociete');
         vValeur_p := GetVarSociete(OBAnnuaire_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarSociete');
      end
      ////
      else if VarAnnuBis(sRacine_l) then
      begin
         if OBAnnuBis_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeAnnuBis');
            ChargeAnnuBis(OBAnnuBis_c, sGuidPer_c);
            Trace('OnGetVar : End Call ChargeAnnuBis');
         end;
         Trace('OnGetVar : Call GetVarAnnuBis');
         vValeur_p := GetVarAnnuBis(OBAnnuBis_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarAnnuBis');
      end
      ////
      else if VarOperation(sRacine_l) then    // DOP
      begin
         if OBDosOper_c.Detail.Count = 0 then  //On n'en charge qu'une seule : sNoOperation_c
         begin
            Trace('OnGetVar : Call ChargeJuDosOper');
            ChargeJuDosOper(OBDosOper_c, sJuCodeDos_c, sNoOperation_c);
            sModule_c := OBDosOper_c.Detail[0].GetValue('JOP_MODULE');
            Trace('OnGetVar : End Call ChargeJuDosOper');
         end;

         Trace('OnGetVar : Call GetVarOperation');
         vValeur_p := GetVarOperation(OBDosOper_c, sRacine_l, sSuffixe_l, 1);
         Trace('OnGetVar : End Call GetVarOperation');
      end
      else if VarRubrique(sRacine_l) then  // sRacine_p = DOR
      begin
      // si besoin chargement du detail
         if OBDosOpRub_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeJuDosOpRub');
            ChargeJuDosOpRub(OBDosOpRub_c, sJuCodeDos_c, sNoOperation_c, '');
            Trace('OnGetVar : End Call ChargeJuDosOpRub');
         end;

         //récupération de la valeur de la variable
         Trace('OnGetVar : Call GetVarRubrique');
         vValeur_p := GetVarRubrique(OBDosOpRub_c, OBBibRubrique_c, sRacine_l, sSuffixe_l,
                                     sNoOperation_c, sModule_c, nIndex_p);
         Trace('OnGetVar : End Call GetVarRubrique');
      end
      else if VarActe(sRacine_l) then  // sRacine_p = DOA
      begin
      // si besoin chargement du detail
         if OBDosOpAct_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeJuDosOpAct');
            ChargeJuDosOpAct(OBDosOpAct_c, sJuCodeDos_c, sNoOperation_c, '');
            Trace('OnGetVar : End Call ChargeJuDosOpAct');
         end;

         //récupération de la valeur de la variable
         Trace('OnGetVar : Call GetVarActe');
         vValeur_p := GetVarActe(OBDosOpAct_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarRubrique');
      end
      else if VarEvenement(sRacine_l) then
      begin
      // si besoin chargement du detail
         if OBEvenement_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeJuEvenement');
            ChargeJuEvenement(OBEvenement_c, sGuidPer_c);
            Trace('OnGetVar : End Call ChargeJuEvenement');
         end;

         //récupération de la valeur de la variable
         Trace('OnGetVar : Call GetVarEvenement');
         vValeur_p := GetVarEvenement(OBEvenement_c, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarEvenement');
      end
      else if VarExercice(sRacine_l) and (sJuCodeDos_c <> '') then  // Rac = EXE et Il existe un dossier juridique
      begin
         if OBDosExe_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeJuDosExe');
            ChargeJuDosExe(OBDosExe_c, sJuCodeDos_c);
            Trace('OnGetVar : End Call ChargeJuDosExe');
         end;
         Trace('OnGetVar : Call GetVarExeRubrique');
         vValeur_p := GetVarExeRubrique(OBDosExe_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarExeRubrique');
      end

      else if VarAutreInfos(sRacine_l) then // cherche si sRacine_p dans JUTYPEINFO
      begin
         // si besoin chargement du detail
         if OBDosInfo_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeJuDosInfo');
            ChargeJuDosInfo(OBDosInfo_c, sJuCodeDos_c);
            Trace('OnGetVar : End Call ChargeJuDosInfo');
         end;
         //récupération de la valeur de la variable
         Trace('OnGetVar : Call GetVarAutreInfo');
         vValeur_p := GetVarAutreInfo(OBDosInfo_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarAutreInfo');
      end
      else if VarBauxFonds(sRacine_l) then // cherche si sRacine_p dans JUBAUXFONDS
      begin
         // si besoin chargement du detail
         if OBBauxFonds_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeJuDosInfo');
            ChargeJuBauxFonds(OBBauxFonds_c, sJuCodeDos_c);
            Trace('OnGetVar : End Call ChargeJuDosInfo');
         end;
         //récupération de la valeur de la variable
         Trace('OnGetVar : Call GetVarAutreInfo');
         vValeur_p := GetVarBauxFonds(OBBauxFonds_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarAutreInfo');
      end
      //
      else if VarDPFiscal(sRacine_l) then
      begin
         if OBDPFiscal_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeDPFiscal');
            ChargeDPFiscal(OBDPFiscal_c, sNoDP_c);
            Trace('OnGetVar : End Call ChargeDPFiscal');
         end;
         Trace('OnGetVar : Call GetVarDPFiscal');
         vValeur_p := GetVarDPFiscal(OBDPFiscal_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarDPFiscal');
      end
      else if VarDPSocial(sRacine_l) then
      begin
         if OBDPSocial_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeDPSocial');
            ChargeDPSocial(OBDPSocial_c, sNoDP_c);
            Trace('OnGetVar : End Call ChargeDPSocial');
         end;
         Trace('OnGetVar : Call GetVarDPSocial');
         vValeur_p := GetVarDPSocial(OBDPSocial_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarDPSocial');
      end
      else if VarDPOrga(sRacine_l) then
      begin
         if OBDPOrga_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeDPOrga');
            ChargeDPOrga(OBDPOrga_c, sNoDP_c);
            Trace('OnGetVar : End Call ChargeDPOrga');
         end;
         Trace('OnGetVar : Call GetVarDPOrga');
         vValeur_p := GetVarDPOrga(OBDPOrga_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarDPOrga');
      end
      else if VarTiers(sRacine_l) then
      begin
         if OBTiers_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeTiers');
            ChargeTiers(OBTiers_c, sTiers_c);
            Trace('OnGetVar : End Call ChargeTiers');
         end;
         Trace('OnGetVar : Call GetVarTiers');
         vValeur_p := GetVarTiers(OBTiers_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarTiers');
      end
      else if VarTiersCompl(sRacine_l) then
      begin
         if OBTiersCompl_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeTiersCompl');
            ChargeTiersCompl(OBTiersCompl_c, sTiers_c);
            Trace('OnGetVar : End Call ChargeTiersCompl');
         end;
         Trace('OnGetVar : Call GetVarTiersCompl');
         vValeur_p := GetVarTiersCompl(OBTiersCompl_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarTiersCompl');
      end
      // $$$ JP 06/04/06 - FQ 10812 corrigée
      {else if VarDPControl(sRacine_l) then
      begin
         if OBDPControl_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeDPControl');
            ChargeDPControl(OBDPControl_c, sNoDP_c);
            Trace('OnGetVar : End Call ChargeDPControl');
         end;
         Trace('OnGetVar : Call GetVarDPControl');
         vValeur_p := GetVarDPControl(OBDPControl_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarDPControl');
      end}
      else if VarProspects(sRacine_l) then
      begin
         if OBProspects_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeProspects');
            ChargeProspects(OBProspects_c, sAuxiliaire_c);
            Trace('OnGetVar : End Call ChargeProspects');
         end;
         Trace('OnGetVar : Call GetVarProspects');
         vValeur_p := GetVarProspects(OBProspects_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarProspects');
      end
      // $$$ JP 06/04/06 - FQ 10812 corrigée
      {else if VarDPTabCompta(sRacine_l) then
      begin
         if OBDPTabCompta_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeDPTabCompta');
            ChargeDPTabCompta(OBDPTabCompta_c, sNoDossier_c);
            Trace('OnGetVar : End Call ChargeDPTabCompta');
         end;
         Trace('OnGetVar : Call GetVarDPTabCompta');
         vValeur_p := GetVarDPTabCompta(OBDPTabCompta_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarDPTabCompta');
      end
      else if VarDPTabPaie(sRacine_l) then
      begin
         if OBDPTabPaie_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeDPTabPaie');
            ChargeDPTabPaie(OBDPTabPaie_c, sNoDossier_c);
            Trace('OnGetVar : End Call ChargeDPTabPaie');
         end;
         Trace('OnGetVar : Call GetVarDPTabPaie');
         vValeur_p := GetVarDPTabPaie(OBDPTabPaie_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarDPTabPaie');
      end
      else if VarDPTabGenPaie(sRacine_l) then
      begin
         if OBDPTabGenPaie_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeDPTabGenPaie');
            ChargeDPTabGenPaie(OBDPTabGenPaie_c, sNoDossier_c);
            Trace('OnGetVar : End Call ChargeDPTabGenPaie');
         end;
         Trace('OnGetVar : Call GetVarDPTabGenPaie');
         vValeur_p := GetVarDPTabGenPaie(OBDPTabGenPaie_c, sRacine_l, sSuffixe_l, nIndex_p);
         Trace('OnGetVar : End Call GetVarDPTabGenPaie');
      end}
      //
      else
      begin
         if OBAnnuLien_c.Detail.Count = 0 then
         begin
            Trace('OnGetVar : Call ChargeAnnuLien');
            ChargeAnnuLien(OBAnnuLien_c, sGuidPer_c, '');   //sTypeDos_c);
            Trace('OnGetVar : End Call ChargeAnnuLien');
         end;
         //récupération de la valeur de la variable
         if VarLien(sRacine_l) then
         begin
            // (sRacine_p = ANL)
            Trace('OnGetVar : Call GetVarLien');
            vValeur_p := GetVarLien(OBAnnuLien_c, sSuffixe_l, nIndex_p);
            Trace('OnGetVar : End Call GetVarLien');
         end
         else
         begin
            // (sRacine_p <> ANL)
            Trace('OnGetVar : Call GetVarLienRac');
            vValeur_p := GetVarLienRac(OBAnnuLien_c, sRacine_l, sSuffixe_l, nIndex_p);
            Trace('OnGetVar : End Call GetVarLienRac');
         end;
      end;
   end
   except
       MessageAlerte('L''application ne parvient pas à récupérer les informations utiles à la fusion'+
                #13#10'( Variable : '+sVarName_p+'['+IntToStr(nIndex_p)+'] )');
   end;
   Trace('End : OnGetVar');
end;

{***********A.G.L.***********************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. : Fonction de mise à jour des variables, appelée par le
Suite ........ : parser dans le cas où le document fait appel à des variables
Suite ........ : de la base (variable entre [] sans $)
Suite ........ : passage des paramètres :
Suite ........ : - nom de la variable (sVarName_p),
Suite ........ : - eventuellement l'indice, pour le Ieme
Suite ........ : - la valeur à mettre à jour
Mots clefs ... :
*****************************************************************}
procedure TGenParserDP.OnSetVar( Sender: TObject; sVarName_p: String; nIndex_p: Integer;vValeur_p: variant) ;
var
     sRacine_l, sSuffixe_l : string;
begin
   Trace('Start : OnSetVar');
   Try
   begin
      bModif_c := true;
      RecupereRacineSuffixe(sVarName_p,sRacine_l, sSuffixe_l);
      if VarDossier(sRacine_l) then
      begin
         SetVarDossier(OBDossier_c, sRacine_l, sSuffixe_l, nIndex_p, vValeur_p);
      end
      else if VarJuridique(sRacine_l) then
      begin
         SetVarJuridique(OBJuridique_c, sRacine_l, sSuffixe_l, nIndex_p, vValeur_p);
      end
      else if VarRubrique(sRacine_l) then
      begin
         if OBDosOpRub_c.Detail.Count = 0 then
            ChargeJuDosOpRub(OBDosOpRub_c, sJuCodeDos_c, sNoOperation_c, '');
         SetVarRubrique(OBDosOpRub_c, OBBibRubrique_c, sRacine_l,sSuffixe_l,nIndex_p,vValeur_p);
      end
      else if VarEvenement(sRacine_l) then
      begin
      // si besoin chargement du detail
         if OBEvenement_c.Detail.Count = 0 then
             ChargeJuEvenement(OBEvenement_c, sGuidPer_c);
         SetVarEvenement(OBEvenement_c, sSuffixe_l, nIndex_p, vValeur_p);
      end
      else if VarExercice(sRacine_l) then     // Rac=EXE
      begin
         if OBDosExe_c.Detail.Count = 0 then
            ChargeJuDosExe(OBDosExe_c, sJuCodeDos_c);
         SetVarExeRubrique(OBDosExe_c, sRacine_l, sSuffixe_l, nIndex_p, vValeur_p);
      end
      else if VarSociete(sRacine_l) then // Rac=STE
      begin
         SetVarSociete(OBAnnuaire_c,sRacine_l, sSuffixe_l, nIndex_p, vValeur_p);
      end
      else if VarBauxFonds(sRacine_l) then // cherche si sRacine_p dans JUBAUXFONDS
      begin
         // si besoin chargement du detail
         if OBBauxFonds_c.Detail.Count = 0 then
            ChargeJuBauxFonds(OBBauxFonds_c, sJuCodeDos_c);
         SetVarBauxFonds(OBBauxFonds_c, sSuffixe_l, nIndex_p, vValeur_p);
      end

      else if VarAutreInfos(sRacine_l) then // cherche sRacine_p dans JUTYPEINFO
      begin
         if OBDosInfo_c.Detail.Count = 0 then
            ChargeJuDosInfo(OBDosInfo_c, sJuCodeDos_c);
         SetVarAutreInfo(OBDosInfo_c,sRacine_l,sSuffixe_l,nIndex_p,vValeur_p);
      end
      else
      begin
         if OBAnnuLien_c.Detail.Count = 0 then
            ChargeAnnuLien(OBAnnuLien_c, sGuidPer_c, sTypeDos_c);
         if VarLien(sRacine_l) then
            SetVarLien(OBAnnuLien_c,sSuffixe_l,nIndex_p,vValeur_p)
         else
            SetVarLienRac(OBAnnuLien_c,sRacine_l, sSuffixe_l, nIndex_p, vValeur_p);
      end;
   end;
   except
       MessageAlerte('L''application ne parvient pas à enregistrer les informations utiles à la fusion'+
                #13#10'( Variable : ' + sVarName_p + '[' + IntToStr(nIndex_p) + '] )');
   end;
   Trace('End : OnSetVar');
end;

{***********A.G.L.***********************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 15/09/2003
Modifié le ... :   /  /
Description .. : Fonction appelée en fin de traitement par le Parser
Suite ........ : Enregistrement et libération
Mots clefs ... :
*****************************************************************}
procedure TGenParserDP.OnFinalization( Sender: TObject );
begin
   Trace('Start : OnFinalization');
   Try
   begin
      // si il y a eu des modifications (bModif_c)
      // et si le traitement s'est bien déroulé (bDone_c
      // on met à jour la base à partir des TOBs
      if (bModif_c) and (bDone_c) then
      begin
         OBDosOpRub_c.InsertOrUpdateDB(true);
         OBDosInfo_c.InsertOrUpdateDB(true);
         OBDosExe_c.InsertOrUpdateDB(true);
         OBDosOper_c.InsertOrUpdateDB(true);
         OBAnnuLien_c.InsertOrUpdateDB(true);
         OBDossier_c.InsertOrUpdateDB(true);
         OBJuridique_c.InsertOrUpdateDB(true);
         OBEvenement_c.InsertOrUpdateDB(true);
         OBBauxFonds_c.InsertOrUpdateDB(true);
         OBContact_c.InsertOrUpdateDB(true);
      // Rétablissement du siren sans blancs
         OBAnnuaire_c.PutValue('ANN_SIREN', DecodeSiren( OBAnnuaire_c.GetValue('ANN_SIREN')));
         OBAnnuaire_c.InsertOrUpdateDB(true);

         OBAnnuBis_c.InsertOrUpdateDB(true);
         OBDPFiscal_c.InsertOrUpdateDB(true);
         OBDPSocial_c.InsertOrUpdateDB(true);
         OBDPOrga_c.InsertOrUpdateDB(true);
         OBTiers_c.InsertOrUpdateDB(true);
         OBTiersCompl_c.InsertOrUpdateDB(true);

         // $$$ JP 06/04/06 - FQ 10812 corrigée
         {OBDPTabCompta_c.InsertOrUpdateDB(true);
         OBDPTabPaie_c.InsertOrUpdateDB(true);
         OBDPTabGenPaie_c.InsertOrUpdateDB(true);
         OBDPControl_c.InsertOrUpdateDB(true);}

         OBProspects_c.InsertOrUpdateDB(true);

      end;

      // libération des TOBs
      if OBDossier_c <> nil then OBDossier_c.Free;
      if OBDosOpRub_c    <> nil then OBDosOpRub_c.Free;
      if OBDosInfo_c     <> nil then OBDosInfo_c.Free;
      if OBDosExe_c      <> nil then OBDosExe_c.Free;
      if OBDosOper_c     <> nil then OBDosOper_c.Free;
      if OBBibRubrique_c <> nil then OBBibRubrique_c.Free;
      if OBDosOpAct_c    <> nil then OBDosOpAct_c.Free;
      if OBAnnuLien_c    <> nil then OBAnnuLien_c.Free;
      if OBJuridique_c   <> nil then OBJuridique_c.Free;
      if OBAnnuaire_c    <> nil then OBAnnuaire_c.Free;
      if OBEvenement_c   <> nil then OBEvenement_c.Free;
      if OBAnnuBis_c <> nil then OBAnnuBis_c.Free;
      if OBDPFiscal_c <> nil then OBDPFiscal_c.Free;
      if OBDPSocial_c <> nil then OBDPSocial_c.Free;
      if OBDPOrga_c <> nil then OBDPOrga_c.Free;
      if OBTiers_c <> nil then OBTiers_c.Free;
      if OBTiersCompl_c <> nil then OBTiersCompl_c.Free;
      if OBBauxFonds_c   <> nil then OBBauxFonds_c.Free;

      // $$$ JP 06/04/06 - FQ 10812 corrigée
      {if OBDPTabCompta_c <> nil then OBDPTabCompta_c.Free;
      if OBDPTabPaie_c <> nil then OBDPTabPaie_c.Free;
      if OBDPTabGenPaie_c <> nil then OBDPTabGenPaie_c.Free;
      if OBDPControl_c <> nil then OBDPControl_c.Free;}

      if OBProspects_c <> nil then OBProspects_c.Free;
   end;
   except
      // si echec
        MessageAlerte('L''application ne parvient pas à terminer le traitement des informations ( Finalisation )');
   end;
   Trace('End : OnFinalization');
end;

end.
