{***********UNITE*************************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 16/04/2002
Modifié le ... : 22/04/2002
Description .. : Fonctions utilisées pour l'interface comptable
Mots clefs ... : BOS5;INTERCOMPTA
*****************************************************************}
unit UtilInterComsX;

interface

uses
  HEnt1,
  Forms,
  Windows,
  Classes,
  Controls,
  IniFiles,
  HCtrls,
  UTOB,
  ComCtrls,
  StdCtrls,
  {$IFNDEF EAGLCLIENT}
  DB,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ELSE}
  Uhttp,
  {$ENDIF}
  SysUtils,
  ParamSoc,
  ed_tools,
  EntGC,
  FileCtrl,
  LicUtil,
  Ent1;

  Type
  		ChgComsX = RECORD
        FichierEcr		 : HString;
		    GestEtab 			 : HString;
        TransfertVers	 : Hstring;
	      NatureTransfert: HString;
  			Format 				 : HString;
        DateArret			 : HString;
        Exercice			 : HString;
        DateDebut			 : HString;
        DateDeFin			 : HString;
        Societe 			 : HString; //Société courante
			  User 					 : HString; //User courant
        DateFile			 : HString;
        TypeEcr				 : HString;
        Email					 : HString;
        Journal				 : HString;
        TypeEcriture	 : Hstring;
		    TypeEcrRecup	 : HString;
		    TypeEcrGen 	 	 : HString;
			  FichierRapport : HString;
        Envoye				 : HString;
      End;

	function GenereFile(ChargeComsx : ChgComsX; Var BlocNotes : TStringList) : Boolean;
  function ImportFile(ChargeComsx : ChgComsX; Var BlocNotes : TStringList) : Boolean;


const
  ChCodeCpta    = 'CODECPTA';
  ChSeqDateDeb  = 'SEQDATEDEB';
  ChSeqDateFin  = 'SEQDATEFIN';
  ChSociete     = 'SOCIETE';
  ChCaptionEcr = 'DEJAEXTRAIT.Caption';
  ChJournalEvent = 'JOURNALEVENT';

  // DC 04/07/07 - Gestion des retours en négatif ou positif
  COMPTA_RETOUR_POSITIF : string = '?';

  // libellés des messages
  TexteMessage: array[ 1..112 ] of hstring = (
    {1}  'Aucune ligne trouvée pour la période sélectionnée.'
    {2}, 'Aucune pièce ne correspond aux critères sélectionnés.'
    {3}, 'Tiers inconnu.'
    {4}, 'Impossible de trouver le compte HT de l''article pour : '
    {5}, 'Impossible de trouver le compte de taxe pour : '
    {6}, 'Impossible de trouver le compte de remise pour : '
    {7}, 'Impossible de trouver le compte d''escompte pour : '
    {8}, 'Impossible de trouver le compte de règlement pour : '
    {9}, 'Impossible de trouver le compte de ports et frais pour :'
    {10}, 'Traitement annulé par l''utilisateur'
    {11}, 'Erreur lors de la génération des lignes.'
    {12}, 'Veuillez confirmer la comptabilisation des écritures déjà extraites.'
    {13}, 'L''emplacement du fichier d''export n''est pas défini dans le paramétrage.'
    {14}, 'Le nom de fichier est incorrect.'
    {15}, 'Le traitement ne s''est pas terminé correctement.'
    {16}, 'Impossible de trouver le compte de remise en banque pour : '
    {17}, 'Attention : veuillez confirmer l''intégration des écritures dans la base courante.'
    {18}, 'Impossible de trouver le compte de frais sur carte bancaire pour : '
    {19}, 'Nombre de boutiques traitées : %1.0f'
    {20}, 'Nombre de journées traitées : %1.0f'
    {21}, 'Transfert terminé'
    {22}, 'Transfert comptable'
    {23}, 'Confirmez-vous la sélection ci-dessous ? #13#10'
    {24}, 'Rapport de fin : '
    {25}, 'Tâche programmée en process serveur'
    {26}, 'Impossible de trouver le compte de l''opération de caisse pour : '
    {27}, 'L''étape "%s" est terminée.'
    {28}, 'L''étape "%s" ne s''est pas terminée correctement.'
    {29}, 'Vous devez relancer cette étape.'
    {30}, 'Aucune tâche n''est en attente ou en cours de traitement.'
    {31}, 'Tâche terminée.'
    {32}, 'Aucune information n''est disponible, vous devez relancez la tâche de chargement.'
    {33}, 'La sauvegarde des écritures dans la table temporaire n''a pas fonctionnée.'
    {34}, 'Aucunes écritures à générer.'
    {35}, 'Erreur lors de l''enregistrement des écritures.'
    {36}, 'Vous devez installer l''exécutable COMSX, il n''est pas présent dans le répertoire : '
    {37}, 'Le répertoire d''export n''existe pas : '
    {38}, 'Problème lors du chargement des lignes.'
    {39}, 'Problème lors du chargement des pièces.'
    {40}, 'Aucune pièce trouvée pour la période sélectionnée.'
    {41}, 'Problème lors du chargement des ports et frais.'
    {42}, 'Problème lors du chargement des taxes.'
    {43}, 'Problème lors du chargement des acomptes.'
    {44}, 'Problème lors du chargement des échéances pour les pièces négoces.'
    {45}, 'Problème lors du chargement des opérations de caisse client.'
    {46}, 'Problème lors du chargement des opérations de caisse.'
    {47}, 'Problème lors du chargement des modes de paiement.'
    {48}, 'Problème lors du traitement des axes analytiques.'
    {49}, 'Problème lors du chargement des remises en banques du coffre.'
    {50}, 'Le journal n''est pas paramétré pour la nature de pièce : '
    {51}, 'Le paramétrage des tickets est incorrect.'
    {52}, '%0:s souche %1:s numero %2:s de l''établissement %3:s du %4:s'
    {53}, 'Le répertoire d''export n''est pas paramétré.'
    {54}, 'Problème de mise à jour du type compta des pièces.'
    {55}, 'Impossible de trouver le journal de l''opération de caisse pour : '
    {56}, 'Impossible de trouver le journal de réglement pour : '
    {57}, 'Impossible de trouver le journal de remise en banque pour : '
    {58}, 'Vérifier dans CEGIDPGI.INI de la COMSX que la société suivante existe : '
    {59}, 'Impossible de trouver le compte collectif pour : '
    {60}, 'Impossible de trouver le compte d''encaissement pour : '
    {61}, 'Vous devez passer le journal suivant en mode saisie="Piece" : '
    {62}, 'L''exercice comptable correspondant aux pièces à transférer n''existe pas.'
    {63}, 'Le pré-contrôle des pièces est correcte.'
    {64}, 'Le contrôle a révélé des anomalies. Veuillez vous référer à l''onglet rapport pour plus de détails.'
    {65}, 'Les champs "GP_ETABLISSEMENT" et "GL_ETABLISSEMENT" ne sont pas renseignés.'
    {66}, 'Le champ "GP_ETABLISSEMENT" n''est pas renseigné : '
    {67}, 'Le champ "GL_ETABLISSEMENT" n''est pas renseigné : '
    {68}, 'Incohérence sur le régime de taxe qui est différent sur la pièce et les lignes : '
    {69}, 'Impossible de comptabiliser cette pièce sans article mouvementé : '
    {70}, 'Vérifier également que la ComSX se lance manuellement sans anomalie. Si tel n''est pas le cas, elle n''est pas installée correctement.'
    {71}, 'Vous devez installer l''exécutable eCOMSX, il n''est pas présent dans le répertoire : '
    {72}, 'Renseignez le paramétrage dans Paramètres->Documents->Documents->Natures Onglet Généralité :le champ "équivalence" = FFO".'
    {73}, 'Vous devez le renseigner dans les paramètres sociétés->gestion commerciale->Passation comptable->Répertoire d''export.'
    {74}, 'Incohérence sur le tiers qui est différent sur la pièce et les lignes : '
    {75}, 'Incohérence sur le tiers facturé qui est différent sur la pièce et les lignes : '
    {76}, 'Vous devez renseigner un compteur de simulation sur le journal suivant : '
    {77}, 'Problème lors du chargement des litiges.'
    {78}, 'Impossible de trouver le compte du litige : '
    {79}, 'Le nom de la société ne doit pas contenir d''espaces. Veuillez le modifier dans votre fichier CEGIDPGI.INI.'
    {80}, 'Le nom du répertoire d''export ne doit pas contenir d''espaces. Veuillez le modifier dans les paramètres sociétés : gestion commerciale->Passation comptable->Répertoire d''export.'
    {81}, 'Les opérations de caisse suivantes ont un paramétrage incohérent.'
    {82}, 'Le type article financier impose que l''option trésorerie soit décochée :'
    {83}, 'Le type article financier impose que l''option trésorerie soit cochée :'
    {84}, 'L''option trésorerie est indéfinie :'
    {85}, 'Opération de caisse : %s - Type article financier : %s - %s'
    {86}, 'Option opération de caisse cochée'
    {87}, 'Option opération de caisse décochée'
    {88}, 'Option opération de caisse indéfinie'
    {89}, 'Mode de paiement inconnu :'
    {90}, 'Impossible de comptabiliser cette pièce comportant des opérations de caisse trésorerie ainsi que d''autres articles'
    {91}, '%0:s souche %1:s n° %2:s date %3:s établissement %4:s tiers %5:s'
    {92}, 'Le nom du répertoire d''export n''est pas défini. Veuillez le définir dans les paramètres sociétés : gestion commerciale->Passation comptable->Répertoire d''export.'
    {93}, 'Les pièces suivantes, saisies dans une ancienne version, comportent des arrondis de taxes incohérents. Vous devez les recalculer par l''utilitaire "Administration/Maintenance/Recalcul des pièces".'
    {94}, 'Vérifier également que les fichiers suivants soient présents dans le sous-répertoire "Script" de votre serveur CWAS :'
    {95}, 'Erreur lors du chargement mémoire des lignes analytiques'
    {96}, 'Vous devez relancer la génération des écritures.'
    {97}, '%0:s compteurs de la souche %1:s ont été utilisés par un autre utilisateur.'
    {98}, 'Le compteur de la souche %0:s n''est plus positionné à sa valeur initiale %1:s.'
    {99}, 'Les natures de pièces sélectionnées sont incompatibles. Le paramétrage des dates des écritures comptables doit être identique pour toutes les natures sélectionnées.'
   {100}, 'Vérifiez ce paramétrage via le menu paramètres / Documents / Documents / Natures, onglet personnalisation comptable.'
   {101}, 'Les modes de paiement suivants ont un paramétrage incohérent.'
   {102}, 'Le compte non collectif %0:s impose que l''option "Détail sur auxiliaire" du mode de paiement %1:s ne soit pas cochée.'
   {103}, 'Le compte collectif %0:s impose que l''option "Détail sur auxiliaire" du mode de paiement %1:s soit cochée.'
   {104}, 'Les exceptions par établissements des modes de paiement suivants ont un paramétrage incohérent.'
   {105}, 'Le compte non collectif %0:s impose que l''option "Détail sur auxiliaire" du mode de paiement %1:s (exception établissement %2:s) ne soit pas cochée.'
   {106}, 'Le compte collectif %0:s impose que l''option "Détail sur auxiliaire" du mode de paiement %1:s (exception établissement %2:s) soit cochée.'
   {107}, 'Les opérations de caisse suivantes ont un paramétrage incohérent.'
   {108}, 'Le compte non collectif %0:s impose que l''option "Détail sur auxiliaire" de l''opération de caisse 1:s ne soit pas cochée.'
   {109}, 'Le compte collectif %0:s impose que l''option "Détail sur auxiliaire" de l''opération de caisse %1:s soit cochée.'
   {110}, 'Les exceptions par établissements des opérations de caisse suivantes ont un paramétrage incohérent.'
   {111}, 'Le compte non collectif %0:s impose que l''option "Détail sur auxiliaire" de l''opération de caisse %1:s (exception établissement %2:s) ne soit pas cochée.'
   {112}, 'Le compte collectif %0:s impose que l''option "Détail sur auxiliaire" de l''opération de caisse %1:s (exception établissement %2:s) soit cochée.'
    );


  // libellés des taches
  LibTache: array[1..7] of hstring = (
    {1}  'Chargement des informations'
    {2}, 'Affectation des comptes'
    {3}, 'Génération des écritures'
    {4}, 'Export des écritures'
    {5}, 'Transfert comptable'
    {6}, 'Import des écritures'
    {7}, 'Pré-contrôle des pièces avant le chargement'
    ) ;

  // Aide au paramétrage
  TexteParam: array[1..14] of hstring = (
    {1}  'Renseignez le compte dans "Administration/Liaison comptable/Ventilations comptables".'
    {2}, 'Renseignez le compte dans "Paramètres/Gestion/Taxe/Taux de taxe -> Onglet complement".'
    {3}, 'Renseignez le compte dans "Paramètres/Gestion/Mode de paiement/Complément par établissement".'
    {4}, 'Renseignez le compte dans "Administration/Liaison comptable/Ventilations comptables" pour la famille comptable article = FCB.'
    {5}, 'Renseignez le compte dans "Paramètres/Gestion/Opération de caisse/Complément par établissement".'
    {6}, 'Renseignez le journal dans "Paramètres/Documents/Documents/Natures -> Onglet Action/Comptabilité.'
    {7}, 'Renseignez le journal dans "Paramètres/Gestion/Mode de paiement/Complément par établissement".'
    {8}, 'Renseignez le journal dans "Paramètres/Gestion/Opération de caisse/Complément par établissement".'
    {9}, 'Renseignez le compte collectif du tiers dans "Données de base/clients ou fournisseurs".'
    {10}, 'Renseignez le compte d''encaissement de l''établissement dans "Données de base/établissement".'
    {11}, 'Modifier le champ Départ du compteur dans "Administration/Liaison comptable/Compteurs comptables".'
    {12}, 'Modifier le champ Mode saisie du journal dans "Administration/Liaison comptable/Journaux".'
    {13}, 'Créer l''exercice comptable dans "Administration/Liaison comptable/Ouverture d''exercice".'
    {14}, 'Renseignez un compteur de simulation du journal dans "Administration/Liaison comptable/Journaux".'
    ) ;

Var
  // Pour rapport de fin de tache
  NombreBtqTraite, NombreJourneeTraite : Double ;
  // TOB Entete du transfert comptable
  TOBEnteteInterCompta : TOB;
  TobSection : TOB;
  // Evenement avec calcul des temps
  //FLogEvent: TLogEvent;
  // ParamSoc de l'affichage des mesures de temps AUC / SIM / DET
  AfficheMesureTmp: hstring;
  // Bloc notes
  //BlocNotes : TStringList ;
  // Exécution totale ou par étape
  ExecutionTotal : Boolean ;
  //Chemin d'accès
	PathFile				: HString;
  PathComSx				: HString;
  FichierIni 			: TIniFile;


Type
  TTypeTacheInterCpta = ( icPreControl, icPlanifie, icChargement, icAffecteCompte, icGenereEcr, icExport, icImport, icNull );

function NomDuFichierExport( NomFichier, Natures: hstring; var DateFile: hstring ): hstring;

//AC 11/29/05

procedure MAJBlocNoteTache (Debut,TacheOk : Boolean; Tache : TTypeTacheInterCpta; Var BlocNotes : TStringList );
procedure MAJBlocNote (Var BlocNotes : TStringList; sMessage : hstring; NumErreur : integer = 0);

Procedure ChargeRapportOK(NomFichier : Hstring; Var BlocNotes : TStringList);
Procedure ChargeRapportErr(NomFichier : Hstring; Var BlocNotes : TStringList);

function SetLastError( Num: integer; LibEcran : hstring = '' ) : hstring ;
function ControleParametrageExport( RepertoireFichier, RepertoireComSx : hstring) : integer ;

function GetPathFileTransfert : string;
function GetPathFileComSx : string;

Function ControleErreurParametres(ChargeComsX : ChgComsx; Var Blocnotes : TStringList ) : Boolean;
Function ChargeIniExport(ChargeComsX: ChgComsX; FileIni : Hstring; Var BlocNotes :TStringList ) : Boolean;
Function ChargeIniImport(ChargeComsX: ChgComsX; FileIni : Hstring; Var BlocNotes :TStringList ) : Boolean;

implementation

uses
{$IFDEF VER150}
	Variants,
{$ENDIF}

{$IFDEF NOMADE}
  UtilPgi,
{$ENDIF NOMADE}

{$IFDEF EAGLSERVER}
  UtilProcessServer,
{$ENDIF EAGLSERVER}
  HMsgBox,
  CBPPath;
  //TobUtil,
  //cbpPath;

function LanceTraitement( LigneDeCommande: string ): Boolean;
var
  TSI: TStartupInfo;
  TPI: TProcessInformation;
begin

  FillChar( TSI, SizeOf( TStartupInfo ), 0 );

  TSI.cb := SizeOf( TStartupInfo );

  Result := CreateProcess( nil, PCHAR( LigneDeCommande ), nil, nil, False,
                           CREATE_NEW_CONSOLE or NORMAL_PRIORITY_CLASS, nil, nil, TSI, TPI );
  if (Result) then
     begin
     while( WaitForSingleObject(Tpi.hProcess, 1000) <> WAIT_OBJECT_0 ) do
        begin
        if( WaitForSingleObject(Tpi.hProcess, 1000) = WAIT_OBJECT_0 ) then Break;
        Application.ProcessMessages;
        end;
     CloseHandle( TPI.hProcess );
     CloseHandle( TPI.hThread );
     end;

end;

//Génération du fichier INIT DE traitement par COMSX
function GenereFile(ChargeComsx : ChgComsX; Var BlocNotes : TStringList) : Boolean ;//var GenerationFileOk: Boolean; RapportLst: TStrings = nil );
var //
    FileIni 				: HString;
    //
	  NumErreur  			: Integer;
	  //
    GenerationFileOk: Boolean;

begin

	  Result := False ;

  	MAJBlocNoteTache (True, False, icExport, BlocNotes );

    GenerationFileOk := false;

    // Toujour format étendu
    PathFile := IncludeTrailingBackslash(ExtractFileDir(ChargeComsX.FichierEcr));
    PathComSx := GetPathFileComSx;

    FileIni := PathFile + 'COMSX.INI';

    // Controle du paramétrage pour l'export
    NumErreur := ControleParametrageExport( PathFile, PathComSx ) ;
    if NumErreur <> 0 then
       begin
       if ( NumErreur = 37 ) or ( NumErreur = 53 ) then
          MAJBlocNote(BlocNotes, PathFile, NumErreur)
       else
          MAJBlocNote(BlocNotes, PathComSx, NumErreur);
       exit;
       end;

  	Try
  	  if ControleErreurParametres(ChargeComsX, BlocNotes) then
    	   GenerationFileOk := False
    	else
      	 GenerationFileOk :=  ChargeIniExport(ChargeComsX, FileIni, BlocNotes);
	  Finally
      //Tache Export d'ecritures Terminé
    	if (GenerationFileOk) then
    	   begin
         Result := True ;
         // Mise à jour bloc notes de la tache fin + ok
         MAJBlocNoteTache ( False, True, icExport, BlocNotes ) ;
         ChargeRapportOK(ChargeComSx.FichierEcr, BlocNotes);
         end
      else
         begin
         // Mise à jour bloc notes de la tache fin + erreur
         MAJBlocNoteTache ( False, False, icExport, BlocNotes ) ;
         ChargeRapportErr(ChargeComSx.FichierEcr, BlocNotes);
         end;
			   // Suppression du Fichier INI
    		 if FileExists(FileIni) then DeleteFile(FileIni);
      end;

end;

Function ControleErreurParametres(ChargeComsX : ChgComsx; Var Blocnotes : TStringList ) : Boolean;

   //Procedure d'ecriture des erreurs
   procedure ErreurParam( Champ, Msg: hstring);
   begin
       if ( Assigned( BlocNotes ) ) then
          begin
          if ( Champ = '' ) then
            Champ := 'MANQUANT';
          BlocNotes.Add( Msg + Champ );
          end;
   end;

begin

  result := false;

  With ChargeComsX Do
       Begin
	     if ( Societe = '' ) or ( User = '' ) or ( PathFile = '' ) or ( GestEtab = '' ) or
          ( FichierEcr = '' ) or ( DateDebut = '' ) or ( DateDeFin = '' ) or ( TypeEcrGen = '' ) or ( FichierRapport = '' ) then
	  	    begin
      		if (Assigned(BlocNotes)) then
				      begin
              BlocNotes.Add( '***** ' + TraduireMemoire( 'TRAITEMENT ANNULE ' ) + '***************************************************' );
              BlocNotes.Add( '***** ' + TraduireMemoire( 'ERREUR DANS LES PARAMETRES ' ) + '******************************************' );
              end;
          ErreurParam(FichierEcr, TraduireMemoire( 'Fichier des écritures    : ' ) );
          ErreurParam( Societe, TraduireMemoire( 'Nom de la base           : ' ));
          ErreurParam( User, TraduireMemoire( 'Utilisateur              : ' ));
          if (GestEtab = 'TRUE' ) then
             ErreurParam(GestEtab, TraduireMemoire( 'Multi Etablissements     : OUI' ) )
          else
             ErreurParam( GestEtab, TraduireMemoire( 'Multi Etablissements     : NON' ) );
          ErreurParam( PathFile, TraduireMemoire( 'Emplacement des fichiers : ' ) );
          //ErreurParam( FichierTiers, TraduireMemoire( 'Fichier des tiers        : ' ) );
          ErreurParam( FichierRapport, TraduireMemoire( 'Fichier du rapport       : ' ) );
          ErreurParam( DateDebut, TraduireMemoire( 'Date début               : ' ) );
          ErreurParam( DateDeFin, TraduireMemoire( 'Date fin                 : ' ) );
          ErreurParam( TypeEcriture, TraduireMemoire( 'Type écritures           : ' ) );
          ErreurParam( FichierRapport, TraduireMemoire( 'Fichier du rapport       : ' ) );
          result := true;
          end;
       end;

end;

function ImportFile(ChargeComsx : ChgComsX; var BlocNotes : TStringList) : Boolean; //BaseCompta: hstring;
var Fileini					: hstring;
    NomFichier			: hstring;
    GenerationFileOk: Boolean;
    NumErreur  			: Integer;
begin

  	Result := False ;

  	MAJBlocNoteTache ( True, False, icImport, BlocNotes ) ;

    GenerationFileOk := True;
    //
    PathFile :=  ExtractFileDir(ChargeComsX.FichierEcr) + '\'; //GetPathFileTransfert;
    PathComSx := GetPathFileComSx;

    FileIni := PathFile + 'COMSX.INI';

    // Controle du paramétrage pour l'export
    NumErreur := ControleParametrageExport( PathFile, PathComSx ) ;
    if NumErreur <> 0 then
       begin
       if ( NumErreur = 37 ) or ( NumErreur = 53 ) then
          MAJBlocNote(BlocNotes, PathFile, NumErreur)
       else
          MAJBlocNote(BlocNotes, PathComSx, NumErreur);
       exit;
       end;

    //
    NomFichier := ChargeComsX.FichierEcr ;
    NomFichier := ChangeFileExt(NomFichier, '.ARC');
    //
  	Try
       if ControleErreurParametres(ChargeComsX, BlocNotes) then
     		  GenerationFileOk := False
  	   else
          GenerationFileOk := ChargeIniImport(ChargeComsX, FileIni, BlocNotes);
    Finally
   		 if ( GenerationFileOk ) then
    			begin
      		Result := True ;
		      // Mise à jour bloc notes de la tache fin + ok
    		  MAJBlocNoteTache(False, True, icImport, BlocNotes ) ;
          ChargeRapportOK(ChargeComSx.FichierEcr, BlocNotes);
			    end
       else
          Begin
    			// Mise à jour bloc notes de la tache fin + erreur
     			MAJBlocNoteTache ( False, False, icImport, BlocNotes ) ;
          ChargeRapportErr(ChargeComSx.FichierEcr, BlocNotes);
          end;
  	end;

    // Suppression du Fichier INI
    if FileExists(FileIni) then DeleteFile(FileIni);

end;

Procedure ChargeRapportOK(Nomfichier : Hstring ;Var BlocNotes : TStringList);
Begin

    NomFichier := ExtractFileName(Nomfichier);
    NomFichier := ChangeFileExt(NomFichier, '.OK');
    NomFichier := 'ListeCom' + Nomfichier;

    Nomfichier := Pathfile + Nomfichier;

    if not FileExists(Nomfichier) then exit;

    //Chgargement rapport d'erreur dans bloc-note...
    Blocnotes.LoadFromFile(Nomfichier);

    //Suppression du rapport d'erreur
    DeleteFile(NomFichier);

end;

Procedure ChargeRapportErr(Nomfichier : Hstring ;Var BlocNotes : TStringList);
Begin

    NomFichier := ExtractFileName(Nomfichier);
    NomFichier := ChangeFileExt(NomFichier, '.err');
    NomFichier := 'ListeCom' + Nomfichier;

    if not FileExists(Nomfichier) then exit;

    //Chgargement rapport d'erreur dans bloc-note...
    Blocnotes.LoadFromFile(Nomfichier);

    //Suppression du rapport d'erreur
    DeleteFile(NomFichier);

end;

Function ChargeIniExport(ChargeComsX : ChgComsX; FileIni : Hstring; Var BlocNotes :TStringList ) : Boolean;
Var Action				 : HString;
		sExclure,StExt2,sExct1			 : HString;
 	  PwdTraitmtlot	 : HString;
	  UserTraitmtlot : HString;
    Q 						 : TQuery;
    Command				 : HString;
Begin

   ////////// CREATION du .INI
   Result := false;

   Action  := 'COMMANDE';

   if FileExists( Fileini ) then DeleteFile( FileIni );

   //AssignFile( FichierIni, FileIni );

   FichierIni := TIniFile.Create( Fileini );
   //ReWrite( FichierIni );

	 FichierIni.WriteString(Action, 'NOMFICHIER', ChargeComsX.FichierEcr);
   FichierIni.WriteString(Action, 'GESTIONETAB', ChargeComsX.GestEtab);
   FichierIni.WriteString(Action, 'TRANSFERTVERS', ChargeComsX.TransfertVers );
   FichierIni.WriteString(Action, 'NATURETRANSFERT', ChargeComsX.NatureTransfert );
   FichierIni.WriteString(Action, 'FORMAT', ChargeComsX.Format );
   FichierIni.WriteString(Action, 'DATEARRET', ChargeComsX.DateArret);
   FichierIni.WriteString(Action, 'EXERCICE', ChargeComsX.Exercice );
   FichierIni.WriteString(Action, 'DATEECR1', ChargeComsX.DateDebut );
   FichierIni.WriteString(Action, 'DATEECR2', ChargeComsX.DateDeFin );
   FichierIni.WriteString(Action, 'TYPE', ChargeComsX.TypeEcrRecup );
   FichierIni.WriteString(Action, 'TYPEECR', ChargeComsX.TypeEcrGen );
   FichierIni.WriteString(Action, 'MAIL', ChargeComsX.EMail );
   FichierIni.WriteString(Action, 'JOURNAL', ChargeComsX.Journal );
   FichierIni.WriteString(Action, 'RAPPORT', ChargeComsX.FichierRapport );

   // Les paramètres à exclure lors de l'export ComSx sont définis par le paramètre société :
   //Administration PGI / Paramètres de débogage / Paramètres exclus lors de l'export : multi-combo sur la tablette MBOCPTAEXCLPAR
   if ExisteSql( 'SELECT 1 FROM PARAMSOC WHERE SOC_NOM="SO_MBOCPTAEXCLPA1"' ) then // Test pour mise en place avant création du paramsoc dans la socref officielle
      begin
      sExclure := '';
      sExct1 := GetParamSocSecur( 'SO_MBOCPTAEXCLPA1', 'MDP;MDR;DEV;REG;SOU;EXO;TL;SSA;PARAM' ); // Tous par défaut
      if sExct1 = '<<Tous>>' then sExct1 := '';
      StExt2 :=  GetParamSocSecur( 'SO_MBOCPTAEXCLPA2', '' ); // Aucun par défaut
      if StExt2 = '<<Tous>>' then StExt2 := '';
      if sExct1 <> '' then sExclure := sExct1;
      if StExt2 <> '' then if sExclure<>'' then sExclure := sExclure+';'+stExt2 else sExclure := StExt2;
      end
   else
   //sExclure := 'MDP;MDR;DEV;REG;SOU;EXO;TL;SSA;PARAM';
     sExclure := '';

   FichierIni.WriteString( Action, 'EXCLURE', '[' + sExclure + ']' );
   FichierIni.WriteString( Action,'EXPORTE', ChargeComsX.Envoye);
   FichierIni.WriteString( Action, 'CTRS', 'TRUE' );

   FichierIni.Free;

   // Fin FICHIER INI
   PwdTraitmtlot := V_PGI.Password;
   UserTraitmtlot := V_PGI.UserLogin;

   // ... car le mot de passe du jour ne marchera plus à minuit...
	 if V_PGI.Password=CryptageSt(DayPass(Date)) then
     	begin
      Q := OpenSQL('select US_ABREGE, US_PASSWORD from UTILISAT where US_SUPERVISEUR="X"', True);
      if not Q.Eof then
         begin
         UserTraitmtlot := Q.FindField('US_ABREGE').AsString;
         PwdTraitmtlot := Q.FindField('US_PASSWORD').AsString;
         end;
	 	  Ferme(Q);
      End;

   {$IFDEF EAGLCLIENT}
   Command := PathComSx + 'eComSx.exe /SERVER="' + AppServer.HostName + '"';
   {$ELSE}
   Command := PathComSx+'ComSx.exe';
   {$ENDIF EAGLCLIENT}

   Command := Command + ' /USER="' + UserTraitmtlot + '"'+ ' /PASSWORD=' + PwdTraitmtlot
                      + ' /DATE=' + DateToStr( V_PGI.DateEntree ) + ' /DOSSIER=' + ChargeComsx.Societe
                      + ' /INI=' + PathFile + 'ComSx.ini;' + ChargeComsx.TypeEcriture+ ';Minimized' ;

   MAJBlocNote(BlocNotes, Command, 0);

   Result := LanceTraitement( Command );

   //Control existance du fichier final
	 if not FileExists(ChargeComSx.FichierEcr ) then
      begin
      MAJBlocNote (BlocNotes, ChargeComsX.Societe, 58) ;
      MAJBlocNote (BlocNotes, 'PGIComSxAPI.dll ComSx.inf', 94);
      MAJBlocNote (BlocNotes , '', 70);
      Result := False ;
      end;

end;

Function ChargeIniImport(ChargeComsX : ChgComsX; FileIni : Hstring; Var BlocNotes :TStringList ) : Boolean;
Var Action				 : HString;
 	  PwdTraitmtlot	 : HString;
	  UserTraitmtlot : HString;
    Q 						 : TQuery;
    Command				 : HString;
Begin

   ////////// CREATION du .INI
   Action := 'COMMANDE';

   if FileExists( Fileini ) then DeleteFile( FileIni );

   FichierIni := TIniFile.Create( Fileini );

   //Writeln( FichierIni, '[COMMANDE]' );

   FichierIni.WriteString( Action, 'REPERTOIRE', PathFile );
   FichierIni.WriteString( Action, 'NOMFICHIER', ExtractFileName(ChargeComsX.FichierEcr));
   FichierIni.WriteString( Action, 'MAIL', ChargeComsX.EMail );
   FichierIni.WriteString( Action, 'TYPEECR', ChargeComsX.TypeEcrRecup);
   FichierIni.WriteString( Action, 'ECRINTEGE', 'N');

   // AC 09/03/07 SansEchec à FALSE tout le temps sinon blocage de l'import si les sections analytiques n'existes pas
   FichierIni.WriteString( Action, 'SANSECHEC', 'FALSE' );

   FichierIni.free;

   // Fin FICHIER INI
   PwdTraitmtlot := V_PGI.Password;
   UserTraitmtlot := V_PGI.UserLogin;

   // ... car le mot de passe du jour ne marchera plus à minuit...
   if V_PGI.Password=CryptageSt(DayPass(Date)) then
      begin
      Q := OpenSQL('select US_ABREGE, US_PASSWORD from UTILISAT where US_SUPERVISEUR="X"', True);
      if not Q.Eof then
         begin
         UserTraitmtlot := Q.FindField('US_ABREGE').AsString;
         PwdTraitmtlot := Q.FindField('US_PASSWORD').AsString;
         end;
      Ferme(Q);
      end ;

   //if BaseCompta <> '' then Societe := BaseCompta;
   {$IFDEF EAGLCLIENT}
   Command := PathComSx + 'eComSx.exe /SERVER="'+ AppServer.HostName +'"';
   {$ELSE}
   Command := PathComSx + 'ComSx.exe';
   {$ENDIF EAGLCLIENT}
   Command := Command + ' /USER="' + UserTraitmtlot + '"'
                      + ' /PASSWORD='+ PwdTraitmtlot
                      + ' /DATE='+ DateToStr(V_PGI.DateEntree)
                      + ' /DOSSIER=' + ChargeComSx.Societe
                      + ' /INI=' + PathFile + 'ComSx.ini;' + ChargeComsx.TypeEcriture+ ';Minimized' ;

   Result := LanceTraitement( Command );

   if FileExists(ChargeComsx.FichierEcr) then
    	begin
      MAJBlocNote (BlocNotes, ChargeComSx.Societe, 58) ;
      MAJBlocNote (BlocNotes, 'PGIComSxAPI.dll ComSx.inf', 94);
      Result := False;
      end;

end;

function HGetComputerName: hstring;
var S: array[0..255] of char;
    i: {$IFNDEF VER110}LongWord{$ELSE}Integer{$ENDIF};
begin
    i := sizeof(s);
    GetComputerName(@s, i);
    Result := StrPas(s);
end;

function NomDuFichierExport( NomFichier, Natures: hstring; var DateFile: hstring ): hstring;
var Extract: integer;
    DateSyst: TDateTime;
    Annee, Mois, Jour, Hre, Min, Sec, mSec: Word;
    ComplementNom: hstring;

  function TraiteDecode( Num : integer ): hstring;
  begin
    if Num < 10 then
      Result := '0' + IntToStr( Num )
    else
      Result := IntToStr( Num );
  end;

begin

  DateSyst := Now;
  DecodeDate( DateSyst, Annee, Mois, Jour );
  DecodeTime( DateSyst, Hre, Min, Sec, mSec );
  DateFile := IntToStr( Annee ) + TraiteDecode( Mois ) + TraiteDecode( Jour ) + TraiteDecode( Hre ) + TraiteDecode( Min );

  while pos( ';', Natures ) > 0 do
     begin
     Extract := pos( ';', Natures );
     Natures := copy( Natures, 1, Extract - 1 ) + '-' + copy( Natures, Extract + 1, length( Natures ) );
     end;

  if NomFichier <> '' then
     ComplementNom := NomFichier
  else
     ComplementNom := Natures ;

  //SO_TYPECOMSX
  if ( GetParamSocSecur( 'SO_TYPECOMSX', 'S5' ) = 'S1' ) or
     ( GetParamSocSecur( 'SO_TYPECOMSX', 'S5' ) = 'S5' ) then
     Result := ComplementNom + '-' + DateFile + '.TRA'
  else if GetParamSocSecur( 'SO_TYPECOMSX', 'S5' ) = 'SIS' then
     Result := ComplementNom + '-' + DateFile + '.TRF'
  else if GetParamSocSecur( 'SO_TYPECOMSX', 'S5' ) = 'SAG' then
     Result := ComplementNom + '-' + DateFile + '.PNM';

end;


{***********A.G.L.***********************************************
Auteur  ...... : A. Cathelineau
Créé le ...... : 25/04/2006
Modifié le ... : 28/04/2006
Description .. : Mise à jour du bloc note pour chaque tache avec son état
Suite ........ : Debut, Fin et son statut Ok, Err
Mots clefs ... :
*****************************************************************}
procedure MAJBlocNoteTache ( Debut,TacheOk : Boolean; Tache : TTypeTacheInterCpta; Var BlocNotes : TStringList) ;
Var
  sErr : hstring ;
  NumEtape : Integer ;
begin

  NumEtape := 0 ;

  case Tache of
    icChargement :
        NumEtape := 1 ;
    icAffecteCompte :
        NumEtape := 2 ;
    icGenereEcr :
        NumEtape := 3 ;
    icExport :
        NumEtape := 4 ;
    icImport :
        NumEtape := 6 ;
    icNull :
        NumEtape := 5 ;
    icPreControl :
        NumEtape := 7 ;
  end;

  if Debut then
     begin
     if NumEtape > 1 then BlocNotes.Add('');
     BlocNotes.Add( LibTache[NumEtape] )
     end
  else
     begin
     if TacheOk then
        begin
        if NumEtape <> 5 then
           BlocNotes.Add( Format_( TraduireMemoire( TexteMessage[ 27 ] ), [ LibTache[NumEtape] ] ) )
        else
           BlocNotes.Add( TraduireMemoire( 'Le ' )+LibTache[NumEtape]+TraduireMemoire( ' est terminé' ) ) ;
        end
     else
        begin
        sErr := Format_( TraduireMemoire( TexteMessage[ 28 ] ), [ LibTache[NumEtape] ] )  ;
        sErr := sErr + ' ' + TraduireMemoire( TexteMessage[ 29 ] );
        BlocNotes.Add( sErr ) ;
        end ;
     end ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : A. Cathelineau
Créé le ...... : 11/05/2006
Modifié le ... :   /  /
Description .. : Mise à jour du bloc note pour les erreurs
Mots clefs ... :
*****************************************************************}
procedure MAJBlocNote (Var BlocNotes : TStringList; sMessage : hstring; NumErreur : integer = 0)  ;
Var
  sText : hstring ;
begin

  BlocNotes.Add( '' );

  if NumErreur <> 0 then
  begin
    case NumErreur of
      51 : begin
             BlocNotes.Add( TraduireMemoire( TexteMessage [ 51 ] ) );
             BlocNotes.Add( TraduireMemoire( TexteMessage [ 72 ] ) );
           end;
      53 : begin
             BlocNotes.Add( TraduireMemoire( TexteMessage [ 53 ] ) );
             BlocNotes.Add( TraduireMemoire( TexteMessage [ 73 ] ) );
           end;
    else
      BlocNotes.Add( TraduireMemoire( TexteMessage [ NumErreur ] ) );
    end;
  end;

  sText := '  '+ sMessage + #13#10 ;

  if sMessage <> '' then BlocNotes.Add( sText );

end ;

{***********A.G.L.***********************************************
Auteur  ...... : A. Cathelineau
Créé le ...... : 15/05/2006
Modifié le ... :   /  /
Description .. : retourne le message d'erreur suivant le traitement effectué
Suite ........ : et le resultat
Mots clefs ... :
*****************************************************************}
function SetLastError( Num: integer; LibEcran : hstring = '' ) : hstring ;
Var
  sMessage : hstring ;
begin
  case num of
    11 :
      begin
        // Tache chargement ok
        sMessage := Format_( TraduireMemoire( TexteMessage[ 27 ] ), [ LibTache[ 1 ] ]) ;
        {$IFNDEF EAGLSERVER}
          PGIInfo( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    12 :
      begin
        // Tache chargement erreur
        sMessage := Format_( TraduireMemoire( TexteMessage[ 28 ] ), [ LibTache[ 1 ] ] )  ;
        sMessage := sMessage + ' ' + TraduireMemoire( TexteMessage[ 29 ] );
        {$IFNDEF EAGLSERVER}
          PGIError( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    21 :
      begin
        // Tache affectation ok
        sMessage := Format_( TraduireMemoire( TexteMessage[ 27 ] ), [ LibTache[ 2 ] ] ) ;
        {$IFNDEF EAGLSERVER}
          PGIInfo( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    22 :
      begin
        // Tache affectation erreur
        sMessage := Format_( TraduireMemoire( TexteMessage[ 28 ] ), [ LibTache[ 2 ] ] )  ;
        sMessage := sMessage + ' ' + TraduireMemoire( TexteMessage[ 29 ] );
        {$IFNDEF EAGLSERVER}
          PGIError( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    31 :
      begin
        // Tache génération ok
        sMessage := Format_( TraduireMemoire( TexteMessage[ 27 ] ), [ LibTache[ 3 ] ] ) ;
        {$IFNDEF EAGLSERVER}
          PGIInfo( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    32 :
      begin
        // Tache génération erreur
        sMessage := Format_( TraduireMemoire( TexteMessage[ 28 ] ), [ LibTache[ 3 ] ] )  ;
        sMessage := sMessage + ' ' + TraduireMemoire( TexteMessage[ 29 ] );
        {$IFNDEF EAGLSERVER}
          PGIError( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    41 :
      begin
        // Tache exportation ok
        sMessage := Format_( TraduireMemoire( TexteMessage[ 27 ] ), [ LibTache[ 4 ] ] ) ;
        {$IFNDEF EAGLSERVER}
          PGIInfo( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    42 :
      begin
        // Tache exportation erreur
        sMessage := Format_( TraduireMemoire( TexteMessage[ 28 ] ), [ LibTache[ 4 ] ] )  ;
        sMessage := sMessage + ' ' + TraduireMemoire( TexteMessage[ 29 ] );
        {$IFNDEF EAGLSERVER}
          PGIError( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    51 :
      begin
        // Tache importation ok
        sMessage := Format_( TraduireMemoire( TexteMessage[ 27 ] ), [ LibTache[ 6 ] ] ) ;
        {$IFNDEF EAGLSERVER}
          PGIInfo( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    52 :
      begin
        // Tache importation erreur
        sMessage := Format_( TraduireMemoire( TexteMessage[ 28 ] ), [ LibTache[ 6 ] ] )  ;
        sMessage := sMessage + ' ' + TraduireMemoire( TexteMessage[ 29 ] );
        {$IFNDEF EAGLSERVER}
          PGIError( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    100 :
      begin
        // Contrôle de cohérence ok
        sMessage := LibTache[ 7 ] + #13#10  ;
        sMessage := TraduireMemoire( TexteMessage[ 63 ] );
        {$IFNDEF EAGLSERVER}
          PGIError( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    101 :
      begin
        // Contrôle de cohérence non ok
        sMessage := LibTache[ 7 ] + #13#10  ;
        sMessage := TraduireMemoire( TexteMessage[ 64 ] );
        {$IFNDEF EAGLSERVER}
          PGIError( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
    0 :
      begin
        // Transfert comptable terminé
        sMessage := TraduireMemoire( TexteMessage[ 21 ] );
        {$IFNDEF EAGLSERVER}
          PGIInfo( sMessage, LibEcran ) ;
        {$ENDIF EAGLSERVER}
      end ;
  end ;
  Result := sMessage ;
end;

function GetPathFileTransfert : string;
begin

  Result := GetParamSocSecur( 'SO_MBOCHEMINCOMPTA', '' );

  if copy(Result, length( Result ), 1 ) <> '\' then Result := Result + '\';

end;

function GetPathFileComSx : string;
Var V : String;
begin

  {$IFDEF EAGLSERVER}
  if ExisteSql( 'SELECT 1 FROM PARAMSOC WHERE SOC_NOM="SO_MBOCHEMINPSCOMSX"' ) then // Test pour mise en place avant création du paramsoc dans la socref officielle
    Result := GetParamSocSecur('SO_MBOCHEMINPSCOMSX', '' );

  if Result = '' then
     Begin
     V := GetFromRegistry(HKey_Local_Machine, 'Software\CEGID_BTP\PgiService', 'Scripts', V, True);
     Result := V;
     end;
  {$ELSE EAGLSERVER}
  Result := '';
  {$ENDIF EAGLSERVER}

  // Répertoire de la ComSx dans "paramètre de débogage"
  if Result = '' then Result := IncludeTrailingBackslash(GetParamSocSecur('SO_BTCHEMINCOMSX', ''));
  if result = '' then Result := IncludeTrailingBackslash(TcbpPath.GetCegidDistriApp);  
  //
  if Result = '' then Result := ExtractFileDir(Application.ExeName) + '\';

  //
  //Result := IncludeTrailingPathDelimiter( Result );

end;

function ControleParametrageExport ( RepertoireFichier, RepertoireComSx : hstring) : integer ;
begin
  Result := 0 ;
  if ( RepertoireFichier = '' ) then
    Result := 53
  else if not DirectoryExists( RepertoireFichier ) then
    Result := 37
  {$IFDEF EAGLCLIENT}
  else if not FileExists( RepertoireComSx + 'eComSx.exe' ) then
    Result := 71 ;
  {$ELSE EAGLCLIENT} // Version 2 tiers et Process Serveur
  else if not FileExists( RepertoireComSx + 'ComSx.exe' ) then
    Result := 36 ;
  {$ENDIF EAGLCLIENT}
end ;

procedure MajRapport( const sEtablissmt, DateDeb, DateFin, NatureD, sEcritures, FichierEcr: hstring;
                      ComptaExterne, bEcritures: Boolean; BlocNotes : TStringList = nil );
var
  Cpt, QteOptions : integer;
  TousEts, TtesNatPces, NatPce, Ets, RgpPce : hstring;
  TempNotes : TStringList ;

  procedure AddTexte( Texte : hstring );
  begin
    QteOptions := QteOptions + 1;
    TempNotes.Add( Texte );
  end;

begin
  if( BlocNotes = nil ) then
  //if( RapportLst = nil ) then
    Exit;
  TempNotes := TStringList.Create;
  TempNotes.Add( '***** ' + TraduireMemoire( 'PARAMETRES ' ) + '**********************************************************' );
  Cpt := 0;
  TousEts := sEtablissmt;
  while( TousEts <> '' ) do
  begin
    Cpt := Cpt + 1;
    Ets := ReadTokenSt( TousEts );
    if Cpt = 1 then
      AddTexte( TraduireMemoire( 'Etablisement(s) : ' ) + Ets + ' - ' + RechDom( 'TTETABLISSEMENT', Ets, False ) )
    else
      AddTexte( '                : ' + Ets + ' - ' + RechDom( 'TTETABLISSEMENT', Ets, False ) );
  end;
  TempNotes.Add( Format_( TraduireMemoire('Période         : du %s au %s'), [DateDeb, DateFin] ) );
  Cpt := 0;
  TtesNatPces := NatureD;
  while( TtesNatPces <> '' ) do
  begin
    Cpt := Cpt + 1;
    NatPce := ReadTokenSt( TtesNatPces );
    if GetInfoParPiece( NatPce, 'GPP_REGROUPCPTA' ) = 'AUC' then
      RgpPce := TraduireMemoire( ' (aucun regroupement)' )
    else if GetInfoParPiece( NatPce, 'GPP_REGROUPCPTA' ) = 'JOU' then
      RgpPce := TraduireMemoire( ' (regroupement journalier)' )
    else if GetInfoParPiece( NatPce, 'GPP_REGROUPCPTA' ) = 'MOI' then
      RgpPce := TraduireMemoire( ' (regroupement mensuel' )
    else if GetInfoParPiece( NatPce, 'GPP_REGROUPCPTA' ) = 'SEM' then
      RgpPce := TraduireMemoire( ' (regroupement hebdomadaire)' );
    if Cpt = 1 then
      AddTexte( TraduireMemoire( 'Piece(s)        : ' ) + GetInfoParPiece( NatPce, 'GPP_LIBELLE' ) + RgpPce )
    else
      AddTexte( '                  ' + GetInfoParPiece( NatPce, 'GPP_LIBELLE' ) + RgpPce );
  end;
  if ComptaExterne then
  begin
    if bEcritures then
      AddTexte( sEcritures );
    AddTexte( TraduireMemoire( 'Export des écritures en fin de traitement dans :' ) );
    AddTexte( ' ' + FichierEcr );
  end;
  TempNotes.AddStrings ( BlocNotes );
  BlocNotes.Clear;
  BlocNotes.AddStrings ( TempNotes );
end;

end.
