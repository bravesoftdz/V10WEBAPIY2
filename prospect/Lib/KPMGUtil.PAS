Unit KPMGUTIL;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
 //    uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     Utom,
     Ent1,
     EntGc,
     ParamSoc,
     utob,
     Tiersutil,
     UtilGC,        //AttribNewCode
     UtilAction,    //FQ10716  TJA
     utobdebug,
     CPTypeCons,
     UtilsRapport,
     Ed_tools
 ;

      function SuspectVersProspect (CleSuspect : String; Bannuaire : boolean = True ; GereEnseigne : boolean = False ; GereContact : boolean = True; TheRapport : TGestionRapport=nil) : String;
      //Function Ciblage2Operation (TobOperation, TobActionGenerique : TOB; CodeCiblage : String) : Boolean;
      // FQ 10541 : ajout des paramètres whereContact et StrJoinContact
      Function Ciblage2Operation (TobOperation, TobActionGenerique : TOB; CodeCiblage, WhereContact, StrJoinContact : String) : Boolean;
      Function CreationVisa (CodeCiblage, CleCible, NatureCible, TypeCible, NumeroContact : String) : Boolean;
      Function ConstruitRequete (Lcoche, Lcodeciblage : string; F : TFMul) : String;
      Function LibColCiblage (LeChamp, LeLib : String) : String;
      Function TrouveNbElement (CodeCiblage : String) : Integer;
      Function RenvoiTabletteCor (var TobParSuspectCor : Tob ; LeChamp : String) : String;
      procedure ReConstruireListe (F : TFMul);
      function JaiLeDroitCiblage(var bBloque : boolean ; var BAvert : Boolean) : String;
      procedure InitRadioBtModeCiblage (F : TForm ; Bbloque : Boolean = False);

 const
      LibCiblage : Array [1..45,1..2] of String = (
        ('coche'            ,'Ciblé'                      ),
        ('codeciblage'      ,'Code ciblage'               ),
        ('code'             ,'Code'                       ),
        ('numero'           ,'Numéro'                     ),
        ('nature'           ,'Code nature'                ),
        ('lnature'          ,'Nature'                     ),
        ('civilite'         ,'Civilité'                   ),
        ('nom'              ,'Nom'                        ),
        ('prenom'           ,'Prénom'                     ),
        ('fonction'         ,'Fonction'                   ),
        ('libelle'          ,'Raison sociale'             ),
        ('adresse1'         ,'Adresse 1'                  ),
        ('adresse2'         ,'Adresse 2'                  ),
        ('adresse3'         ,'Adresse 3'                  ),
        ('codepostal'       ,'Code postal'                ),
        ('ville'            ,'Ville'                      ),
        ('journais'         ,'Jour de naissance'          ),
        ('moisnais'         ,'Mois de naissance'          ),
        ('anneenais'        ,'Année de naissance'         ),
        ('ape'              ,'Code Naf'                   ),
        ('pays'             ,'Pays'                       ),
        ('principal'        ,'Contact princ.'             ),
        ('region'           ,'Code région'                ),
        ('secteur'          ,'Secteur'                    ),
        ('siret'            ,'Code Siret'                 ),
        ('telephone'        ,'N° téléphone'               ),
        ('zonecom'          ,'Zone comm.'                 ),
        ('visa'             ,'Visa'                       ),
        ('lvisa'            ,'Visa'                       ),
        ('enseigne'         ,'Enseigne'                   ),
        ('representant'     ,'Représentant'               ),
        ('societegroupe'    ,'Société Groupe'             ),
        ('prescripteur'     ,'Préscripteur'               ),
        ('juridique'        ,'Juridique'                  ),
        ('tariftiers'       ,'Tarif'                      ),
        ('facture'          ,'Facture'                    ),
        ('payeur'           ,'Payeur'                     ),
        ('ferme'            ,'Fermé'                      ),
        ('particulier'      ,'Particulier'                ),
        ('ispayeur'         ,'IsPayeur'                   ),
        ('phonetique'       ,'Phonétique'                 ),
        ('moiscloture'      ,'Mois de clôture'            ),
        ('@SI([VCM_CODECIBLAGE]', 'Ciblé'                 ),
        ('@SI([CCM_CODECIBLAGE]', 'Ciblé'                 ),
        ('rva'              ,'Site Web'                   )     //FQ 14453
        );
  KTexteMessage : array [1..70] of string = (
    {1}'Confirmez-vous la suppression du paramétrage de l''import des suspects ?',
    {2}'Création des fiches en Suspect Particulier',
    {3}'Création des fiches en Suspect Entreprise',
    {4}'Critères de sélection des enregistrements du fichier',
    {5}'Pas de critères de sélection des enregistrements du fichier',
    {6}'Sélection sur le champ : %s',
    {7}'Sélection à la position : %s',
    {8}'    sur une longueur de : %s caractères',
    {9}'    sur une longueur de : %s caractère',
   {10}'    Borne inférieure : %s',
   {11}'    Borne supérieure : %s',
   {12}'Traitement en cours...',
   {13}'Nombre de lignes du fichier : %s',
   {14}'Efface la table des suspects',
   {15}'Récupération fichier suspects',
   {16}'Nombre de lignes suspects écrites : %s',
   {17}'Dernier code suspect : %s',
   {18}'%s codes Siret ne sont pas corrects et donc pas renseignés',
   {19}'Nombre de fiches tiers mises à jour : %s',
   {20}'erreur : %s',
   {21}'Efface les suspects existants',
   {22}'Nombre de lignes suspects écrites : %s',
   {23}'Premier code suspect : %s',
   {24}'Interrompue par l''utilisateur',
   {25}'Traitement terminé',
   {26}'Traitement non terminé',
   {27}'Ajout d''un lien sur <%s> avec le libellé',
   {28}'Ajout d''une correspondance pour le champ <%s> sur le libellé',
   {29}'Début champ',
   {30}'Position champ',
   {31}'Début',
   {32}'Position',
   {33}'La formule n''est pas valide',
   {34}'Ciblage validé. Suppression impossible',
   {35}'Ce ciblage contient des cibles sélectionnées, voulez-vous vraiment le supprimer ?',
   {36}'Ciblage validé. Modification impossible',
   {37}'Des adresses risquent d''être enlevées du ciblage',
   {38}'Ce code existe déjà !#10#13 Veuillez essayer un autre code',
   {39}'Ce ciblage n''est pas validé',
   {40}'Ciblage déjà validé',
   {41}'Ce ciblage est déjà validé, vous ne pouvez que le consulter',
   {42}'Vous n''êtes pas en numérotation automatique des tiers',
   {43}'Le caractère de bourrage du compte auxiliaire est numérique',
   {44}'Vous avez activé la gestion des enseignes en tablette dans le module Paramètres#10#13    il n''est donc pas possible de transférer les enseignes présentes #10#13    sur les fiches suspects si celles-ci figurent en tant que texte libre',
   {45}'Aucun élément sélectionné',
   {46}'Impossible de récupérer tous les enregistrements',
   {47}'Il reste des Prévisa',
   {48}'Aucun élément à valider',
   {49}'Transfert suspect : %s',
   {50}'Problèmes rencontrés durant le transfert des suspects',
   {51}'Ciblage validé',
   {52}'Assistant ciblage vers opération',
   {53}'Vous devez saisir un code opération',
   {54}'Vous devez saisir un libellé',
   {55}'Vous devez saisir au moins une action générique',
   {56}'Opération enregistrée',
   {57}'Impossible de créer l''opération',
   {58}'Vous devez saisir un type d''action',
   {59}'Vous devez saisir un libellé',
   {60}'L''opération %s est bien créée et les actions de la première action générique ont été créées sur la cible',
   {61}'Sélection des ciblages : %s',
   {62}'Traitement en série pour actions',
   {63}'Suppression des suspects',
   {64}'Attention : cette commande va supprimer toutes les fiches sélectionnées#10#13 Les fiches complémentaires seront aussi supprimées#10#13 Confirmez-vous le traitement ?',
   {65}'Veuillez patienter',
   {66}'Vous n''avez pas les droits de transfert de suspect en prospect',
   {67}'Ce code de description de fichier existe déjà',
   {68}'Une opération du même nom a déjà été générée',
   {69}'Création de l''opération',
   {70}'Ajout des actions'
    );

    MsgAvertEnseigne : String = '#10#13  - Vous avez activé la gestion des enseignes en tablette dans le module Paramètres#10#13'+
                          '    il n''est donc pas possible de transférer les enseignes présentes #10#13'+
                          '    sur les fiches suspects si celles-ci figurent en tant que texte libre.';

Implementation
uses   CbpMCD,UFonctionsCBP;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 03/07/2006
Modifié le ... : 14/03/2007
Description .. : Transfert de suspect en prospect
Suite ........ : CleSuspect : Rsu_suspect
Suite ........ : Bannuaire : création vers l'annuaire
Suite ........ : GereEnseigne : gestion des enseignes
Suite ........ : GereContact : maj des contacts sur import
Suite ........ : Result : code tiers
Suite ........ : 
Suite ........ : Création du tiers. Si déjà existant alors on met à jour les
Suite ........ : données et ajoute le contact si inexistant
Suite ........ : Reprend les définitions de TomTiers.
Mots clefs ... : SUSPECT, PROSPECT
*****************************************************************}
function SuspectVersProspect (CleSuspect : String; Bannuaire : boolean = True ; GereEnseigne : boolean = False ; GereContact : boolean = True; TheRapport : TGestionRapport=nil) : String;
var
   i, ipos, nbcontact, NumProspect : integer;
   Qr : Tquery;
   Qr2 : Tquery;
   Bcontactexist : boolean;
   TobSuspect, TobProspect, TobContact, TobTiersCompl, TobProspectCompl, TobCor : Tob;
   TobSuspectCompl : tob;
   tobtiers2 : tob;
   stfrom, stto : string;
   codepro, codeauxi: string;
   ChampToTiers, ChampToCompl, ChampFrom : String;
   ErreurTiers : String;
   Alternative : Boolean;

Begin
  Result := '';
  Alternative := (not GetParamSocSecur('SO_GCNUMTIERSAUTO',false));
  //
  TobSuspect := Tob.Create('SUSPECTS', nil, -1);

  TobProspect := Tob.Create('LE TIERS', nil, -1);
  //
  tobtiers2 := tob.create('TIERS', nil, -1);
  tobtiers2.Initvaleurs;

  TobContact := Tob.Create('CONTACT', nil, -1);
  TobContact.Initvaleurs;

  TobTiersCompl := Tob.Create('TIERSCOMPL', nil, -1);
  TobTiersCompl.InitValeurs;

  TobProspectCompl := Tob.create('PROSPECTS', nil, -1);
  TobProspectCompl.InitValeurs;


  if TobSuspect.SelectDB ('"'+clesuspect+'"', nil, False) then
  begin
    // si le tiers est inexistant
    if not (ExisteSql('SELECT * FROM TIERS WHERE T_SIRET="'+tobsuspect.getvalue('RSU_SIRET')+'"'))
    	 or (TobSuspect.GetValue('RSU_SIRET')='') then
    begin
      for i := 1 to TobSuspect.NbChamps do
      begin
        //recherche des champs commun à suspect et prospect
        stfrom := tobsuspect.getnomchamp(i);
        ipos := pos('_', stfrom);
        system.delete(stfrom, 1, ipos-1);
        stto := 'T'+stfrom;
        if tobtiers2.fieldexists(stto) then
        	tobprospect.AddChampSupValeur(stto, tobsuspect.getvalue('RSU'+stfrom));
        system.delete(stfrom, 1, 8);
        stto := 'C_'+stfrom;
        if tobcontact.fieldexists(stto) and TobSuspect.fieldexists('RSU_CONTACT'+StFrom) then
        	tobcontact.Putvalue(stto, tobsuspect.getvalue('RSU_CONTACT'+stfrom));
      end;

      {CRM_20080722_CD_012;10802_DEB}
      if RTTypeProduitCRM = 'COMFI' then tobprospect.addchampsupvaleur('T_NATUREAUXI','CLI')
                                    else tobprospect.addchampsupvaleur('T_NATUREAUXI','PRO');
      {CRM_20080722_CD_012;10802_FIN}
      tobprospect.addchampsupvaleur('T_DATEINTEGR', V_PGI.DateEntree);
      //codepro := GetParamSocSecur('SO_GCPREFIXETIERS','',True) + IntToStr(NumProspect);
      if Alternative then
      begin
      	NumProspect := GetParamSocSecur('SO_BTCOMPTEURTIERS', -1, False) + 1;
      	CodePro := AttribNewcode('TIERS', 'T_TIERS', GetParamsoc('SO_GCLGNUMTIERS'),trim(GetParamsoc('SO_BTPREFIXETIERS')), InttoStr(Numprospect), '',True,true);
        SetParamSoc('SO_BTCOMPTEURTIERS', NumProspect);
      end else
      begin
      	NumProspect := GetParamSocSecur('SO_GCCOMPTEURTIERS', -1, False) + 1;
      	CodePro := AttribNewcode('TIERS', 'T_TIERS', GetParamsoc('SO_GCLGNUMTIERS'),trim(GetParamsoc('SO_GCPREFIXETIERS')), InttoStr(Numprospect), '');
        SetParamSoc('SO_GCCOMPTEURTIERS', NumProspect);
      end;


      //par défaut le code auxiliaire est identique au code tiers
      tobprospect.addchampsupvaleur('T_TIERS',codepro);
      tobprospect.addchampsupvaleur('T_AUXILIAIRE','');
      tobprospect.addchampsupvaleur('T_FACTURE','');
      tobprospect.addchampsupvaleur('T_DATECREATION', V_PGI.DateEntree);
      tobprospect.addchampsupvaleur('T_DATEOUVERTURE', V_PGI.DateEntree);
      tobprospect.addchampsupvaleur('T_CREERPAR', 'SAI');
      tobprospect.addchampsupvaleur('T_DEVISE', V_PGI.devisepivot);
      tobprospect.addchampsupvaleur('T_LETTRABLE', 'X');
      tobprospect.addchampsupvaleur('T_COLLECTIF',VH^.DefautCli);
      tobprospect.addchampsupvaleur('T_REGIMETVA',VH^.RegimeDefaut);
      tobprospect.addchampsupvaleur('T_PAYS', GetParamSoc('SO_GcTiersPays'));
      tobprospect.addchampsupvaleur('T_MODEREGLE', GetParamSoc('SO_GcModeRegleDefaut'));
      tobprospect.addchampsupvaleur('T_TVAENCAISSEMENT', GetParamSoc('SO_TVAENCAISSEMENT'));
      if VH_GC.GCDefFactureHT then tobprospect.addchampsupvaleur('T_FACTUREHT','X');
      tobprospect.addchampsupvaleur('T_CONFIDENTIEL','0');
      tobprospect.addchampsupvaleur('T_PUBLIPOSTAGE','X');
      tobprospect.addchampsupvaleur('T_MODEREGLE',VH_GC.GCModeRegleDefaut);
      TobProspect.AddChampSupValeur('T_MOISCLOTURE', 1);
      {$IFDEF NOMADE}
      tobprospect.addchampsupvaleur('T_REPRESENTANT', VH_GC.PCPRepresentant);
      {$ENDIF}
      //gestion des champs obligatoires du prospect

      //FQ 10479  on vide l'enseigne
      if GereEnseigne then
      TobProspect.PutValue('T_ENSEIGNE', '');

      // appel de la comande tiersutil.createtiersfromtob
      ErreurTiers := CreateTiersFromTob(codepro, TobProspect);

      If ErreurTiers = '' then            //la création du tiers C bien passée
      begin
        // PL le 29/05/07 : on vient de créer le code propsect, on le récupère tel qu'il a été
        // effectivement créé
        if (TobProspect <> nil) then
        codepro := TobProspect.GetString('T_TIERS');
        tobProspect.free;
        TobProspect := tob.create('TIERS', nil, -1);
        Qr2 := OpenSQL('SELECT * FROM TIERS WHERE T_TIERS = "'+codepro+'"', True);
        TobProspect.SelectDB('', Qr2, false);
        ferme(Qr2);

        TobProspect.PutValue('T_AUXILIAIRE', UpperCase_(TobProspect.getvalue('T_AUXILIAIRE')));
        codeauxi := TobProspect.getvalue('T_AUXILIAIRE');

        if Trim(TobContact.GetValue('C_NOM'))<>'' then
        begin
          tobcontact.PutValue('C_AUXILIAIRE',codeauxi);
          tobcontact.putvalue('C_TYPECONTACT','T');
          tobcontact.putvalue('C_NUMEROCONTACT','1');
          {CRM_20080722_CD_012;10802_DEB}
          if RTTypeProduitCRM = 'COMFI' then tobcontact.putvalue('C_NATUREAUXI','CLI')
                                        else tobcontact.putvalue('C_NATUREAUXI','PRO');
          {CRM_20080722_CD_012;10802_FIN}
          tobcontact.putvalue('C_PRINCIPAL','X');
          tobcontact.putvalue('C_FONCTIONCODEE', TobSuspect.GetValue('RSU_CONTACTFONCTION'));
          tobcontact.putvalue('C_TELEPHONE', tobsuspect.getvalue('RSU_CONTACTTELEPH'));
          tobcontact.putvalue('C_SEXE', tobsuspect.getvalue('RSU_SEXE'));
          tobcontact.putvalue('C_JOURNAIS', tobsuspect.getvalue('RSU_JOURNAISSANCE'));
          tobcontact.putvalue('C_MOISNAIS', tobsuspect.getvalue('RSU_MOISNAISSANCE'));
          tobcontact.putvalue('C_ANNEENAIS', tobsuspect.getvalue('RSU_ANNEENAISSANCE'));
          tobcontact.putvalue('C_PUBLIPOSTAGE', tobsuspect.getvalue('RSU_CONTACTPUBLI'));
          tobcontact.putvalue('C_EMAILING', tobsuspect.getvalue('RSU_CONTACTEMLG'));
          tobcontact.putvalue('C_TIERS', tobprospect.getvalue('T_TIERS'));
          tobcontact.insertorupdatedb(false)
        end;

        tobtierscompl.SelectDB('"'+codeauxi+'"', nil, false);
        //      tobtierscompl.putvalue('YTC_AUXILIAIRE', tobprospect.getvalue('T_AUXILIAIRE'));
        //      tobtierscompl.putvalue('YTC_TIERS', tobprospect.getvalue('T_TIERS'));
        tobtierscompl.putvalue('YTC_TIERSLIVRE', codeauxi);

        TobProspectCompl.PutValue('RPR_AUXILIAIRE', codeauxi);


        //recherche des correspondances entre suspects et tiers
        TobCor := Tob.Create('correspondance', nil, -1);
        TobCor.LoadDetailDB('PARSUSPECTCOR', '', '', nil, False, True);
        if TobCor.Detail.Count > 0 then
        begin
          TobSuspectCompl := Tob.Create('SUSPECTSCOMPL', nil, -1);
          TobSuspectCompl.SelectDB('"'+CleSuspect+'"', nil);
          //FQ 10467 ajout des commerciaux
          TobTiersCompl.PutValue('YTC_REPRESENTANT2', TobSuspectCompl.GetValue('RSC_REPRESENTANT2'));
          TobTiersCompl.PutValue('YTC_REPRESENTANT3', TobSuspectCompl.GetValue('RSC_REPRESENTANT3'));
          TobTiersCompl.PutValue('YTC_TAUXREPR1', TobSuspectCompl.GetValue('RSC_TAUXREPR1'));
          TobTiersCompl.PutValue('YTC_TAUXREPR2', TobSuspectCompl.GetValue('RSC_TAUXREPR2'));
          TobTiersCompl.PutValue('YTC_TAUXREPR3', TobSuspectCompl.GetValue('RSC_TAUXREPR3'));

          for i := 0 to TobCor.Detail.Count -1 do
          begin
            ChampFrom := TobCor.Detail[i].GetValue('RSP_CHSUSPECT');
            ChampToTiers := TobCor.Detail[i].Getvalue('RSP_CHTIERS');
            ChampToCompl := TobCor.Detail[i].Getvalue('RSP_CHCOMPL');
            //               TobSuspectCompl := Tob.Create('SUSPECTSCOMPL', nil, -1);
            //               TobSuspectCompl.SelectDB('"'+CleSuspect+'"', nil);

            if ChampToTiers <> '' then
            TobTiersCompl.PutValue(ChampToTiers, TobSuspectCompl.GetValue(ChampFrom));
            if ChampToCompl <> '' then
            TobProspectCompl.PutValue(ChampToCompl, TobSuspectCompl.GetValue(ChampFrom));
            //               TobSuspectCompl.free;
          end;
          TobSuspectCompl.free;
        end;

        TobCor.ClearDetail;
        TobCor.Free;

        TobSuspect.PutValue('RSU_DATESUSPRO', V_PGI.DateEntree);
        tobsuspect.PutValue('RSU_MOTIFFERME', 'PRO');
        tobsuspect.PutValue('RSU_DATEFERMETURE', V_PGI.DateEntree);
        tobsuspect.putvalue('RSU_FERME','X');


        //correction 10421
        TobProspect.PutValue('T_ABREGE', TobProspect.GetValue('T_LIBELLE'));
        TobProspect.PutValue('T_PUBLIPOSTAGE','X');
        TobProspect.PutValue('T_EURODEFAUT','-');
        TobProspect.PutValue('T_CREERPAR','TRL');


        //ajout dans les tables
        if  (TobTiersCompl.InsertOrUpdateDB(False))
            and (TobProspect.UpdateDB(False))
            and (TobProspectCompl.InsertOrUpdateDB(False)) then
        tobsuspect.insertorupdatedb(false);

        //       SetParamSoc('SO_GCCOMPTEURSUSPRO', NumProspect);
        Result := TobProspect.getvalue('T_TIERS');
        if TheRapport <> nil then TheRapport.SauveLigMemo('OK : Le suspect '+clesuspect+' a été transféré en tant que '+Result);

      end else
      begin
    		if TheRapport <> nil then
        begin
        	TheRapport.SauveLigMemo('ERREUR : Le suspect '+clesuspect+' n''a pu être créé en '+codepro+'. Erreur: '+ErreurTiers);
        end else
        begin
        	PGIInfo(ErreurTiers, 'Erreur création prospect');
        end;
      end;

    end else
    begin
    	//si le tiers Existe
      Qr := OpenSql('SELECT * FROM TIERS WHERE T_SIRET="'+tobsuspect.getvalue('RSU_SIRET')+'"', True);
      //chargement du tiers
      //            TobProspect.Free;
      //            TobProspect := tob.create('TIERS', nil, -1);
      codeauxi := Qr.Findfield('T_AUXILIAIRE').AsString;
      codepro := Qr.findfield('T_TIERS').asstring;
      //            tobprospect.SelectDB('"'+codeauxi+'"', nil, false);

      tobtiers2.SelectDB('', Qr, False);

      Ferme(Qr);
      //recherche des données tiers
      for i:=1 to tobtiers2.NbChamps do
      begin
        stfrom := tobtiers2.GetNomChamp(i);
        TobProspect.AddChampSupValeur(stfrom, tobtiers2.getvalue(stfrom));
      end;

      for i := 1 to TobSuspect.NbChamps do
      begin
        //recherche des champs commun à suspect et prospect
        stfrom := tobsuspect.getnomchamp(i);
        ipos := pos('_', stfrom);
        system.delete(stfrom, 1, ipos-1);
        stto := 'T'+stfrom;
        if tobprospect.fieldexists(stto) then
        tobprospect.AddChampSupValeur(stto, tobsuspect.getvalue('RSU'+stfrom));
      end;
      {
      tobprospect.putvalue('T_LIBELLE', tobsuspect.GetValue('RSU_LIBELLE'));
      tobprospect.putvalue('T_ADRESSE1', tobsuspect.GetValue('RSU_ADRESSE1'));
      tobprospect.putvalue('T_ADRESSE2', tobsuspect.GetValue('RSU_ADRESSE2'));
      tobprospect.putvalue('T_ADRESSE3', tobsuspect.GetValue('RSU_ADRESSE3'));
      tobprospect.putvalue('T_CODEPOSTAL', tobsuspect.GetValue('RSU_CODEPOSTAL'));
      tobprospect.putvalue('T_VILLE', tobsuspect.GetValue('RSU_VILLE'));
      tobprospect.putvalue('T_PAYS', tobsuspect.GetValue('RSU_PAYS'));
      tobprospect.putvalue('T_TELEPHONE', tobsuspect.GetValue('RSU_TELEPHONE'));
      tobprospect.putvalue('T_FAX', tobsuspect.GetValue('RSU_FAX'));
      tobprospect.putvalue('T_TELEX', tobsuspect.GetValue('RSU_TELEX'));
      TobProspect.putvalue('T_DATEMODIF', V_PGI.DateEntree);
      }

      //FQ 10479  on vide l'enseigne
      if GereEnseigne then
      TobProspect.PutValue('T_ENSEIGNE', '');

      erreurtiers := CreateOrUpdateTiersFromTob(codepro, TobProspect);

      //vérification des contacts
      // Attention : rajout gestion des contacts ou non
      if GereContact then
      begin
        Qr := OpenSql('SELECT * FROM CONTACT WHERE C_AUXILIAIRE="'+codeauxi+'"', True);
        Qr.First;
        Bcontactexist := false;
        nbcontact := 0;
        While not Qr.Eof do
        begin
          inc(nbcontact) ;

          //si le suspect existe déjà dans la liste des contacts
          if (Qr.FindField('C_NOM').AsString=tobsuspect.getvalue('RSU_CONTACTNOM'))
          and (Qr.FindField('C_PRENOM').AsString=tobsuspect.getvalue('RSU_CONTACTPRENOM')) then
          begin
            Bcontactexist := True;
            tobcontact.SelectDB('"'+Qr.Findfield('C_TYPECONTACT').AsString+'";"'+codeauxi+'";"'+
            Qr.FindField('C_NUMEROCONTACT').AsString+'"', nil);
            break;
          end;

          Qr.Next;
        end;
        Ferme(Qr);

        //mise à jour des contacts pour mettre le nouveau à 'principal'
        ExecuteSql('UPDATE CONTACT SET C_PRINCIPAL="-" WHERE C_AUXILIAIRE="'+codeauxi+'"');

        tobcontact.putvalue('C_PRINCIPAL','X');
        tobcontact.putvalue('C_TELEPHONE', tobsuspect.getvalue('RSU_CONTACTTELEPH'));
        tobcontact.putvalue('C_SEXE', tobsuspect.getvalue('RSU_SEXE'));
        tobcontact.putvalue('C_JOURNAIS', tobsuspect.getvalue('RSU_JOURNAISSANCE'));
        tobcontact.putvalue('C_MOISNAIS', tobsuspect.getvalue('RSU_MOISNAISSANCE'));
        tobcontact.putvalue('C_ANNEENAIS', tobsuspect.getvalue('RSU_ANNEENAISSANCE'));
        tobcontact.putvalue('C_PUBLIPOSTAGE', tobsuspect.getvalue('RSU_CONTACTPUBLI'));
        tobcontact.putvalue('C_EMAILING', tobsuspect.getvalue('RSU_CONTACTEMLG'));
        tobcontact.putvalue('C_RVA', tobsuspect.getvalue('RSU_CONTACTRVA'));


        //si le contact est nouveau
        if not bcontactexist then
        begin
          tobcontact.putvalue('C_FONCTIONCODEE', tobsuspect.getvalue('RSU_CONTACTFONCTION'));
          tobcontact.putvalue('C_CIVILITE', tobsuspect.getvalue('RSU_CONTACTCIVILITE'));
          tobcontact.putvalue('C_PRENOM', tobsuspect.getvalue('RSU_CONTACTPRENOM'));
          tobcontact.putvalue('C_NOM', tobsuspect.getvalue('RSU_CONTACTNOM'));
          tobcontact.putvalue('C_AUXILIAIRE',codeauxi);
          tobcontact.putvalue('C_TYPECONTACT','T');
          tobcontact.putvalue('C_NUMEROCONTACT', nbcontact+1);
          tobcontact.putvalue('C_NATUREAUXI', tobprospect.GetValue('T_NATUREAUXI'));
          tobcontact.putvalue('C_AUXILIAIRE', codeauxi);
          tobcontact.putvalue('C_TIERS', tobprospect.getvalue('T_TIERS'));
        end;
      end;

      //recherche des correspondances entre suspects et tiers
      TobTiersCompl.SelectDB('"'+codeauxi+'"', nil, false);
      TobProspectCompl.selectDb('"'+codeauxi+'"', nil, false);

      TobCor := Tob.Create('correspondance', nil, -1);
      TobCor.LoadDetailDB('PARSUSPECTCOR', '', '', nil, False, True);
      if TobCor.Detail.Count > 0 then
      begin
        TobSuspectCompl := Tob.Create('SUSPECTSCOMPL', nil, -1);
        TobSuspectCompl.SelectDB('"'+CleSuspect+'"', nil);
        for i := 0 to TobCor.Detail.Count -1 do
        begin
            ChampFrom := TobCor.Detail[i].GetValue('RSP_CHSUSPECT');
          ChampToTiers := TobCor.Detail[i].Getvalue('RSP_CHTIERS');
          ChampToCompl := TobCor.Detail[i].Getvalue('RSP_CHCOMPL');
          //            TobSuspectCompl := Tob.Create('SUSPECTSCOMPL', nil, -1);
          //            TobSuspectCompl.SelectDB('"'+CleSuspect+'"', nil);

          if ChampToTiers <> '' then
          begin
            TobTiersCompl.PutValue(ChampToTiers, TobSuspectCompl.GetValue(ChampFrom));
            TobTiersCompl.PutValue('YTC_DATEMODIF', V_PGI.DateEntree);
          end;
          if ChampToCompl <> '' then
          begin
            TobProspectCompl.PutValue(ChampToCompl, TobSuspectCompl.GetValue(ChampFrom));
            TobProspectCompl.PutValue('RPR_DATEMODIF', V_PGI.DateEntree);
          end;
          //            TobSuspectCompl.free;
        end;
        TobSuspectCompl.free;
      end;

      TobCor.ClearDetail;
      TobCor.Free;


      tobsuspect.PutValue('RSU_DATEFERMETURE', V_PGI.DateEntree);
      tobsuspect.putvalue('RSU_FERME','X');

      //mise à jour du tiers et/ou ajout du contact
      if tobprospect.insertorupdatedb(False)
         and (TobTiersCompl.InsertOrUpdateDB(False))
         and (TobProspectCompl.InsertOrUpdateDB(False)) then
      begin
        tobsuspect.insertorupdatedb(false);
        if GereContact then tobcontact.insertorupdatedb(false);
      end;

      Result := tobprospect.getvalue('T_TIERS');
    	if TheRapport <> nil then TheRapport.SauveLigMemo('OK : Le suspect '+clesuspect+' a été transféré en tant que '+Result);
    end;
  end else
  begin
    if TheRapport <> nil then TheRapport.SauveLigMemo('Erreur : Le suspect '+clesuspect+' n''existe pas');
  	Result := '';
  end;

  {$IFDEF DP}
  //Creation de l'annuaire
  if (Result<>'') and Bannuaire then CreeUnAnnuaire (tobprospect.getvalue('T_TIERS'));
  {$ENDIF}

  {$IFDEF EAGLCLIENT}
  AvertirCacheServer('PARAMSOC');
  {$ENDIF}

  tobsuspect.free;
  tobprospect.free;
  tobcontact.free;
  tobtierscompl.free;
  TobProspectCompl.free;
  tobtiers2.free;
end;



{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 19/06/2006
Modifié le ... :   /  /
Description .. : Génération des opérations, Actions génériques et Action
Mots clefs ... :
*****************************************************************}
//function Ciblage2Operation (TobOperation, TobActionGenerique : TOB; CodeCiblage : String) : boolean;
// FQ 10541 : ajout des paramètres whereContact et StrJoinContact
function Ciblage2Operation (TobOperation, TobActionGenerique : TOB; CodeCiblage, WhereContact, StrJoinContact : String) : boolean;
var
  i                     : integer;
  TobListeElement       : Tob;
  StrSql                : String;
  TobNAction            : Tob;
  StrCreatAction        : String;
  LeTiers               : String;
  bRacContact           : boolean;    //TJA 08/09/2008  //FQ 10888
  TobCiblage            : Tob;

begin
  Result                := True;
  bRacContact           := False;

  //céation de l'opération et des actions génériques
  BEGINTRANS;
  Try
    TobOperation.InsertOrUpdateDB(False);
    for i := 0 to TobActionGenerique.Detail.Count -1 do
      TobActionGenerique.Detail[i].InsertOrUpdateDB(False);

    //FQ 10888    TJA 08/09/2008
    StrSql                := 'SELECT * FROM CIBLAGE WHERE RCB_CODECIBLAGE = "' + CodeCiblage + '"';
    TobCiblage            := TOB.Create('CIBLAGE', nil, -1);
    TobCiblage.LoadDetailFromSQL(StrSql);
    if TobCiblage.Detail[0].GetString('RCB_TYPECIBLAGE') = '002' then   // ciblage par contact
      bRacContact       := True;

    //on récupère la liste des tiers concernés
    TobListeElement       := Tob.Create('Les elements cibles', nil, -1);

    //StrSql                := 'SELECT * FROM CIBLAGEELEMENT ';
    // FQ 10541  STR 09/09/2008 : ajout de StrJoinContact et WhereContact
    // et on limite les champs à ceux de ciblageelement sauf t_tiers
    // et si besoin c_numerocontact en +
    if StrJoinContact<>'' then
      StrSql              := 'SELECT CIBLAGEELEMENT.*,T_TIERS,isnull(C_NUMEROCONTACT,0) as C_NUMEROCONTACT FROM CIBLAGEELEMENT '+StrJoinContact
    else
      StrSql              := 'SELECT CIBLAGEELEMENT.*,T_TIERS FROM CIBLAGEELEMENT ';
    //
    if bRacContact then
      StrSql              := StrSql + 'LEFT JOIN TIERS ON T_AUXILIAIRE=RVB_CONTACT '
    else
      StrSql              := StrSql + 'LEFT JOIN TIERS ON T_TIERS=RVB_TIERS ';
    StrSql                := StrSql + 'WHERE RVB_CODECIBLAGE = "' + CodeCiblage + '" ';
    StrSql                := StrSql + 'AND RVB_VISA = "002" ';                // FQ 10915   //TJA 10/09/2008
    StrSql                := strSql + WhereContact; // FQ 10541

    TobListeElement.LoadDetailFromSQL(StrSql);
    InitMoveProgressForm(nil, 'Création de l''opération', '', TobListeElement.Detail.Count -1, True, True);



    for i := 0 to TobListeElement.Detail.Count -1 do
    begin
      LeTiers             := TobListeElement.Detail[i].GetValue('T_TIERS');
      //création de la tob qui va servir à la création des actions.
      TobNAction          := Tob.Create('La tob action', nil, -1);

      //FQ 10888  TJA 08/09/2008
      if StrJoinContact<>'' then // + FQ 10541 STR 09/09/2008
        TobNAction.AddChampSupValeur('RAC_NUMEROCONTACT', TobListeElement.Detail[i].GetString('C_NUMEROCONTACT'))
      else
      TobNAction.AddChampSupValeur('RAC_NUMEROCONTACT', TobListeElement.Detail[i].GetString('RVB_NUMEROCONTACT'));

      TobNAction.AddChampSupValeur('RAC_LIBELLE', TobActionGenerique.Detail[0].GetValue('RAG_LIBELLE'));
      TobNAction.AddChampSupValeur('RAC_TYPEACTION', TobActionGenerique.Detail[0].GetValue('RAG_TYPEACTION'));
      TobNAction.AddChampSupValeur('RAC_INTERVENANT', TobActionGenerique.Detail[0].GetValue('RAG_INTERVENANT'));
      TobNAction.AddChampSupValeur('RAC_OPERATION', TobOperation.GetValue('ROP_OPERATION'));
      TobNAction.AddChampSupValeur('RAC_NUMACTGEN', TobActionGenerique.Detail[0].GetValue('RAG_NUMACTGEN'));
      TobNAction.AddChampSupValeur('RAC_DATEACTION', V_PGI.DateEntree);
      TobNAction.AddChampSupValeur('RAC_ETATACTION', TobActionGenerique.Detail[0].GetValue('RAG_ETATACTION'));
      TobNAction.AddChampSupValeur('RAC_TABLELIBRE1', TobActionGenerique.Detail[0].GetValue('RAG_TABLELIBRE1'));
      TobNAction.AddChampSupValeur('RAC_TABLELIBRE2', TobActionGenerique.Detail[0].GetValue('RAG_TABLELIBRE2'));
      TobNAction.AddChampSupValeur('RAC_TABLELIBRE3', TobActionGenerique.Detail[0].GetValue('RAG_TABLELIBRE3'));
      TobNAction.AddChampSupValeur('RAC_DATEECHEANCE', TobActionGenerique.Detail[0].GetValue('RAG_DATEECHEANCE'));      //FQ 10698  TJA
      // FQ 10871 - STR 20/10/08 - ajout rubrique cout de l'action
      TobNAction.AddChampSupValeur('RAC_COUTACTION', TobActionGenerique.Detail[0].GetValue('RAG_COUTACTION'));

      MoveCurProgressForm('Ajout des actions');

      if TobNAction.GetInteger('RAC_NUMEROCONTACT')<>0 then // FQ 10541
        TobNAction.AddChampSupValeur('RAC_DESTMAIL',';'+IntToStr(TobNAction.GetInteger('RAC_NUMEROCONTACT'))+';');

      StrCreatAction      := CreateOrUpdateAction(True, LeTiers, 0, TobNAction);
      FreeAndNil(TobNAction);

      // si une erreur à la création d'une action
      if StrCreatAction <> '' then
      begin
        PGIInfo(StrCreatAction);
        Result            := False;
        break;
      end;
    end;
    COMMITTRANS
  Except
    on E:Exception do
    begin
      Result              := False;
      ROLLBACK;
    end;  
  end;

  FreeAndNil(TobCiblage);
  FiniMoveProgressForm;

end;



{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 26/06/2006
Modifié le ... : 26/06/2006
Description .. : Ajout d'un élément dans le ciblage
Suite ........ : CodeCiblage : Code du ciblage
Suite ........ : CleCible : code tiers ou code suspect
Suite ........ : NatureCible : Suspect, prospect ou client
Suite ........ : TypeCible : Rs(001) ou Contact(002)
Suite ........ : NumeroContact : num du contact tiers. 1 si suspect
Mots clefs ... :
*****************************************************************}
Function CreationVisa (CodeCiblage, CleCible, NatureCible, TypeCible, NumeroContact : String) : Boolean;
var
   TobVisa : TOB;
begin
Result := False;

TobVisa := Tob.Create ('CIBLAGEELEMENT', nil, -1);
TobVisa.InitValeurs;

TobVisa.PutValue('RVB_CODECIBLAGE', CodeCiblage);

if ExJaiLeDroitConcept(TConcept(gcVisaCiblage), False) then       //FQ10580
  TobVisa.PutValue('RVB_VISA', '002')       //visa
else
  TobVisa.PutValue('RVB_VISA', '001');      //previsa

if NatureCible = 'SUS' then
begin
   TobVisa.PutValue('RVB_SUSPECT', CleCible);
   TobVisa.PutValue('RVB_CONTACT', 'XXX');    //code contact forcé
   TobVisa.PutValue('RVB_TIERS', 'XXX');     //code tiers forcé
end
else
begin
   TobVisa.PutValue('RVB_SUSPECT', 'XXX');     //code suspect forcé
   TobVisa.PutValue('RVB_CONTACT', 'XXX');    //code contact forcé
   TobVisa.PutValue('RVB_TIERS', CleCible);
end;

if TypeCible = '002' then     //suspect
begin
   if NatureCible = 'SUS' then
      TobVisa.PutValue('RVB_NUMEROCONTACT', '1')  //Numéro forcé
   else
   begin
      TobVisa.PutValue('RVB_NUMEROCONTACT', NumeroContact);
      TobVisa.PutValue('RVB_TIERS', 'XXX');
      TobVisa.PutValue('RVB_CONTACT', CleCible);
   end;
end;

if TObVisa.InsertOrUpdateDB(true) then
   Result := True;

TobVisa.Free;
end;     //fin de la function creationvisa




{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 29/06/2006
Modifié le ... :   /  /
Description .. : Permet la construction de la requete pour les topages
Suite ........ : ListeChamp : liste des champs extrait depuis la table LISTE
Suite ........ : Lcoche : nom du champ coche
Suite ........ : Lcodeciblage : nom du champ codeciblage
Mots clefs ... :
*****************************************************************}
Function ConstruitRequete (Lcoche, Lcodeciblage : string; F : TFMul) : String;
var
   x : integer;
   xx : integer;
   nbchamp : integer;
   LChamp : String;
   ListeChamp : String;
   ListeLibChamp : String;
   LibChamp : String;
   StrSelect : String;
   StrGroupby : String;
   Dataliste : Tstringlist;
   Qr : TQuery;
   ListeLib : String;
   Tobliste : Tob;


begin
//recheche dans LISTE pour connaitre la liste des champs
Qr := OpenSQL('SELECT * FROM LISTE WHERE LI_LISTE="'+F.DBListe+'" AND LI_LANGUE="'+V_PGI.LangueDataRef+'"'+
              ' AND LI_UTILISATEUR="---"', True);
Tobliste := tob.Create('LISTE', nil, -1);
Tobliste.SelectDB('',Qr,False);
DataListe := TStringList.Create;
DataListe.SetText(PChar(Qr.FindField('LI_DATA').AsString));
ListeChamp := DataListe.Strings[1];  //récupèration des champs de la liste
ListeLibChamp := Dataliste.Strings[4];

Ferme(Qr);

//la liste des champs retourné par la table LISTE a pour séparateur ";"
x := Pos (';', ListeChamp);
StrSelect := '';
StrGroupby := '';
nbchamp := 0;

while x<>0 do
begin
   LChamp := Copy(ListeChamp, 1, x-1);
   system.delete (ListeChamp, 1, x);
   Inc(nbchamp);

   xx := Pos (';', ListeLibChamp);
   LibChamp := Copy(ListeLibChamp, 1, xx-1);
   System.Delete (ListeLibChamp, 1, xx);


   if LChamp = Lcoche then    //si le champ pour la coche
   begin
      if Length(StrSelect) > 0 then
         StrSelect := StrSelect + ', ';
//      StrSelect := StrSelect + 'IIF((MAX('+Lcodeciblage+') IS NOT NULL), "X", "-") RCB_VALIDATION';
      StrSelect := StrSelect + 'IIF((MAX('+Lcodeciblage+') <> ""), "X", "-") RCB_VALIDATION';
   end
   else
   begin
      if LChamp = Lcodeciblage then // champ du code du ciblage
      begin
         if Length(StrSelect) > 0 then
            StrSelect := StrSelect + ', ';
         StrSelect := StrSelect + 'MAX('+Lcodeciblage+') '+Lcodeciblage;
      end
      else
      begin
         if LChamp[1]<>'(' then      // attention ne pas prendre le champ pour cumul
         begin
            if Length(StrSelect) > 0 then
               StrSelect := StrSelect + ', ';
            StrSelect := StrSelect + LChamp;
            if Length(StrGroupby) > 0 then
               StrGroupby := StrGroupby + ', ';

            StrGroupby := StrGroupby + LChamp;
         end
         else   //cas des cumuls
         begin
            if Length(StrSelect) > 0 then
               StrSelect := StrSelect + ', ';
            StrSelect := StrSelect + LChamp + ' as C'+IntToStr(nbchamp);
//            if Length(StrGroupby) > 0 then
//               StrGroupby := StrGroupby + ', ';
//            StrGroupby := StrGroupby + 'C'+IntToStr(nbchamp);
         end;
      end;
   end;
   x := Pos (';', ListeChamp);
   ListeLib := ListeLib + LibColCiblage(LChamp, LibChamp) + ';';
end;

Dataliste.Strings[4] := ListeLib;
Tobliste.PutValue('LI_DATA', Dataliste.Text);
Tobliste.InsertOrUpdateDB(False);
tobliste.free;
Result := StrSelect + '###' + StrGroupby;



end;



{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 31/08/2006
Modifié le ... :   /  /    
Description .. : Permet de mettre un libellé sur les champs des vues utilisées 
Suite ........ : par le ciblage
Mots clefs ... : LIBELLE; CIBLAGE
*****************************************************************}
Function LibColCiblage (LeChamp, LeLib : String) : String;
var
  i : integer;
  Suffixe : string;
  Libelle : string;

begin
  Libelle := '';
  i := pos('_', LeChamp);
  Suffixe := LowerCase_(copy(LeChamp, i+1, Length(LeChamp)));
  for i:=Low(LibCiblage) to High(LibCiblage) do
  begin
    if LibCiblage[i,1] = Suffixe then
    begin
      Libelle := LibCiblage[i,2];
      break;
    end;
    if pos('@SI([VCM_CODECIBLAGE]', LeChamp) <> 0 then
    begin
      Libelle           := LibCiblage[i, 2];
      break;
    end;
    if pos('@SI([CCM_CODECIBLAGE]', LeChamp) <> 0 then
    begin
      Libelle           := LibCiblage[i, 2];
      break;
    end;
  end;
  if libelle='' then
  begin
//    Libelle := ChampToLibelle(LeChamp);
    Libelle := LeLib;
  end;

  Result := Libelle;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 20/09/2006
Modifié le ... :   /  /
Description .. : Retourne le nombre d'éléments dans un ciblage
Mots clefs ... : CIBLAGE; COMPTE
*****************************************************************}
Function TrouveNbElement (CodeCiblage : String) : Integer;
Var
  Q : Tquery ;
  Nbelement : integer;
Begin
Q := OpenSQL('SELECT COUNT(0) FROM CIBLAGEELEMENT WHERE RVB_CODECIBLAGE="'+CodeCiblage+'"', True);
Nbelement := Q.Fields[0].Asinteger;
Ferme(Q);
result := Nbelement;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/05/2007
Modifié le ... : 22/05/2007
Description .. : Retourne le nom de la tablette 'TIERS' associée au champ
Suite ........ : de la table 'Suspect'
Suite ........ : paramètres :
Suite ........ :     - TobSuspectCor : tob stockant les champs en
Suite ........ : correspondance. Le déchargement de cette tob est à la
Suite ........ : charge de l'appellant
Suite ........ :     - LeChamp : Champ suspect à tester
Mots clefs ... :
*****************************************************************}
Function RenvoiTabletteCor (var TobParSuspectCor : Tob ; LeChamp : String) : String;
var
  TobC                  : Tob;
  TobS                  : Tob;
  NumChamp              : Integer;
  TypeChamp             : String;
  LeLien                : String;
  Q                     : TQuery;
//  ChampSpe              : String;
	Mcd : IMCDServiceCOM;


begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  Result                := '';
  // si la tob est vide on la charge
  if TobParSuspectCor = nil then
    TobParSuspectCor    := Tob.Create('PARSUSPECTCOR', nil, -1);

  if TobParSuspectCor.Detail.Count = 0 then
  begin
    Q                   := OpenSQL('SELECT * FROM PARSUSPECTCOR', True);
    TobParSuspectCor.LoadDetailDB('PARSUSPECTCOR','','',Q,False);
    TobC                := Tob.Create('PARSUSPECTCOR', TobParSuspectCor, -1);
    TobC.PutValue('RSP_CHSUSPECT', 'XXX');    // pour fournir au moins 1 occurence
    Ferme(Q);
  end;

{
  ChampSpe              := '';
  if pos('_6', LeChamp) > 0 then
  begin
    ChampSpe            := LeChamp;
    LeChamp             := StringReplace(LeChamp, '_6', '_', [rfReplaceAll]);
  end;
 }
  // est ce que le champ à une correspondance dans les TIers
  TobC                  := TobParSuspectCor.FindFirst(['RSP_CHSUSPECT'], [LeChamp], False);
  if Assigned(TobC) then
  begin
    if Pos('RSC_', LeChamp) > 0 then
      TobS              := Tob.Create('SUSPECTSCOMPL', nil, -1)
    else
      TobS              := Tob.Create('SUSPECTS', nil, -1);


    TobS.InitValeurs;
{
    if ChampSpe <> '' then
    begin
      NumChamp          := TobS.GetNumChamp(ChampSpe);
      tobc.PutValue('RSP_CHSUSPECT', ChampSpe);
    end
    else
}
    TypeChamp := Mcd.ChampToType(LeChamp);
    LeLien              := TobC.GetValue('RSP_CHTIERS');
    if LeLien = '' then
      LeLien            := TobC.GetValue('RSP_CHCOMPL');

      // si le champ est bien une tablette
//    if Typechamp = 'COMBO' then       //pas de condition à cause des champs particuliers (rsc_6rsclib...., multival)
      Result            := ChampToTT(LeLien);

    FreeAndNil(TobS);
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 06/06/2007
Modifié le ... :   /  /    
Description .. : Modifie la liste du mul
Suite ........ : traduction des libellés de la vue
Mots clefs ... : 
*****************************************************************}
procedure ReConstruireListe (F : TFMul);
var
  Qr                    : TQuery;
  TobC                  : Tob;
  LChamp                : String;
  LibChamp              : String;
  Dataliste             : Tstringlist;
  ListeChamp            : String;
  ListeLibChamp         : String;
  NvListeLib            : String;

begin
  Qr                    := OpenSQL('SELECT * FROM LISTE WHERE LI_LISTE="' + F.DBListe + '" AND LI_LANGUE="' + V_PGI.LangueDataRef +
                                   '"'+ ' AND LI_UTILISATEUR="' + V_PGI.User + '"', True);
  if QR.RecordCount = 0 then          // si pas de liste personnalisée
  Qr                    := OpenSQL('SELECT * FROM LISTE WHERE LI_LISTE="' + F.DBListe + '" AND LI_LANGUE="' + V_PGI.LangueDataRef +
                                   '"'+ ' AND LI_UTILISATEUR="---"', True);

  TobC                  := tob.Create('LISTE', nil, -1);
  TobC.SelectDB('', Qr, False);
  DataListe             := TStringList.Create;
  DataListe.SetText(PChar(Qr.FindField('LI_DATA').AsString));
  ListeChamp            := DataListe.Strings[1];  //récupèration des champs de la liste
  ListeLibChamp         := Dataliste.Strings[4];  //Le libellé en cours
  NvListeLib            := '';
  Ferme(Qr);

  LChamp                := READTOKENST(ListeChamp);
  LibChamp              := READTOKENST(ListeLibChamp);
  while Lchamp <> '' do
  begin
    NvListeLib          := NvListeLib + LibColCiblage(LChamp, LibChamp) + ';';

    LChamp              := READTOKENST(ListeChamp);
    LibChamp            := READTOKENST(ListeLibChamp);
  end;

  Dataliste.Strings[4]  := NvListeLib;
  TobC.PutValue('LI_DATA', Dataliste.Text);
  TobC.InsertOrUpdateDB(False);
  TobC.Free;

end;


{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 18/06/2007
Modifié le ... : 28/08/2007
Description .. : vérifie les paramsoc liés au ciblage :
Suite ........ :  - la numérotation auto des tiers
Suite ........ :  - le caractère de bourrage des auxiliaires
Suite ........ :  - le droit concept suspect vers prospect
Suite ........ :  - la gestion des enseignes
Mots clefs ... :
*****************************************************************}
function JaiLeDroitCiblage(var bBloque : boolean ; var BAvert : Boolean) : String;
var
  StrMsg                : String;
  bNumTiersAuto         : Boolean;
  cBourreAux            : String;
  bConceptSusPro        : Boolean;
  bEnseigne             : Boolean;
  LgTiers               : integer;
  LgAuxiliaire          : integer;
  LgPrefixeAuxiliaire   : integer;
  Alternative : Boolean;
begin
//FQ 10484

  StrMsg                := '';
  bBloque               := False;
  BAvert                := False;
  bNumTiersAuto         := GetParamSocSecur('SO_GCNUMTIERSAUTO', False, True);
  cBourreAux            := UpperCase_(GetParamSocSecur('SO_BOURREAUX', 'X', True));
  bConceptSusPro        := ExJaiLeDroitConcept(TConcept(GcTransfertSusPro), False);
  bEnseigne             := GetParamSocSecur('SO_GCENSEIGNETAB', False, True);
  LgTiers               := GetParamSocSecur('SO_LGCPTEGEN', 0, True);
  LgAuxiliaire          := GetParamSocSecur('SO_LGCPTEAUX', 0, True);
  LgPrefixeAuxiliaire   := Length(GetParamSocSecur('SO_GCPREFIXEAUXI', '', True));
  alternative 				  := not GetParamSocSecur('SO_GCNUMTIERSAUTO',false);

  if (Alternative) and (Trim(GetParamSocSecur('SO_BTPREFIXETIERS',''))='') then
  begin
    StrMsg              := StrMsg + '#10#13  - Vous devez paramétrer le préfixe des codes tiers dans "Gestion de la Relation Clients/Suspect".';
    bBloque             := True;
  end;
  
  if (not Alternative) and (not bNumTiersAuto) then
  begin
    StrMsg              := StrMsg + '#10#13  - Vous devez paramétrer la numérotation automatique des codes tiers.';
    bBloque             := True;
  end;
  (*
  if not (cBourreAux[1] in ['A'..'Z']) then
  begin
    StrMsg              := StrMsg + '#10#13  - Vous devez utiliser un caractère de bourrage alphabétique pour les comptes auxiliaires.';
    bBloque             := True;
  end;
  *)

  if not bConceptSusPro then
  begin
    StrMsg              := StrMsg + '#10#13  - Vous devez autoriser le concept des droits de Transfert par lot des suspects en prospects.';
    bBloque             := True;
  end;

  //FQ 10484
  if LgTiers + LgPrefixeAuxiliaire > LgAuxiliaire then
  Begin
  // FQ 10861   // TJA 11/09/2008 message inutile   // TJA 16/09/2008 correction annulée
    StrMsg              := StrMsg + '#10#13  - La longueur des comptes Auxiliaires doit être égale ou supérieure à celle des codes tiers';
    bBloque             := True;
  End;

  if bBloque then     // FQ 10861  //TJA 16/09/2008
  begin
    StrMsg              := 'Vous ne pourrez pas générer de ciblage par Suspects ou Mixtes :' + StrMsg;
  end;

  if bEnseigne then
  Begin
    StrMsg              := StrMsg + '#10#13 Non Bloquant : ' + MsgAvertEnseigne;
    BAvert              := True;
  end;


//FQ 10861
//  if StrMsg <> '' then
//    StrMsg              := 'Attention ! ' + StrMsg;

  Result                := StrMsg;
end;



procedure InitRadioBtModeCiblage (F : TForm ; Bbloque : Boolean = False);
var
  TobModeCiblage        : Tob;
  StrSql                : String;
  RadioBt               : TRadioButton;
  i                     : integer;

begin
  RadioBt               := TRadioButton(F.FindComponent('RB1'));
  if Assigned(RadioBt) then
  begin
    StrSql                := 'SELECT * FROM COMMUN WHERE CO_TYPE = "MCB"';
    TobModeCiblage        := Tob.Create('Les modes de ciblages', nil, -1);
    TobModeCiblage.LoadDetailFromSQL(StrSql);
    if TobModeCiblage.Detail.Count = 3 then     //3 modes de ciblage
    begin
      for i := 1 to 3 do
      begin
        RadioBt           := TRadioButton(F.FindComponent('RB' + IntToStr(i)));
        RadioBt.Caption   := TobModeCiblage.Detail[i-1].GetString('CO_LIBELLE');

        if i=1 then
          RadioBt.Checked := True;
          
        if bbloque and (i > 1) then
          RadioBt.Enabled := False;
       end;
    end;
    FreeAndNil(TobModeCiblage);
  end;
end;


Initialization


end.
