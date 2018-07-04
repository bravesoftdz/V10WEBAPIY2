{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 10/12/2002
Modifié le ... : 28/02/2003 : Compatibilité CWAS
Description .. : Unité commune
Mots clefs ... : DP
*****************************************************************}
unit DpOutilsAgenda;


interface

uses
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
    utob, forms;


const
     // une seconde en tdatetime
     dSecond = 1/86400;

     // Autorisation de modification: aucune, seulement l'état "réalisé" ou toute modification autorisée
     aaNothing  = 0;
     aaDoneOnly = 1;
     aaAll      = 2;

     // Pour lien Outlook
     olFolderCalendar  = 9;
     olAppointment     = 26;
     olAppointmentItem = 1;

     olRecursDaily     = 0;
     olRecursWeekly    = 1;
     olRecursMonthly   = 2;
     olRecursMonthNth  = 3;
     olRecursYearly    = 5;
     olRecursYearNth   = 6;

     olText            = 1;
     olDateTime        = 2;

     olNormal          = 0;
     olPersonal        = 1;
     olPrivate         = 2;
     olConfidential    = 3;

     olFree            = 0;
     olTentative       = 1;
     olBusy            = 2;
     olOutOfOffice     = 3;

     // Type de synchronisation
     ostFull            = 0;
     ostFromOutlook     = 1;
     ostFromCegid       = 2;
     ostNone            = 3;
     ostDesynchro       = 4;
     ostImport          = 5;
     //ostAlready         = -1;
     ostDeleteInOutlook = -2;
     ostDeleteInCegid   = -3;

     // Résolution des conflits
     ocrNone       = 0;
     ocrCegid      = 1;
     ocrOutlook    = 2;
     ocrFull       = 3;
     ocrMostRecent = 4;

     // Résolution de suppression
     odrNone       = 0;
     odrKeep       = 1;
     odrDelete     = 2;
     odrDesync     = 3;


type
     // $$$ JP 25/09/06 - fréquence de l'activité (YYAGENDA_FIC)
     ActFrequence = (ACF_NONE, ACF_DAILY, ACF_WEEKLY, ACF_MONTHLY, ACF_YEARLY);

     // $$$ JP 09/10/09 - portée de la modification (série, l'occurence ou aucune)
     ActPortee = (ACP_NONE, ACP_OCCURENCE, ACP_SERIE, ACP_EXCEPTION);

     // $$$ JP 29/05/07: type d'exception
     ActExcType = (AET_UNKNOWN, AET_NORMAL, AET_DELETED);
                                           

     TOutlookSync = class (TObject)
     private
            m_iTypeSync        :integer;
            m_iConflictSolver  :integer; // Gestion des doubles modification
            m_iSupprSolver     :integer; // Gestion des suppressions
            m_bSupprFull       :boolean; // Gestion des suppressions dans les deux sens? (pour les synchro partielle de type copie)
            m_strGuidBase      :string;
            m_strGuidOutlook   :string;

            // Conteneurs des activités (TOB) et des rendez-vous (variant IDispatch d'outlook: collections de appointmentitem)
            m_ActItems         :TOB;
            m_ApptItems        :variant;
            m_iNbSync          :integer;

            function  WeekDaysToWeekMask (strDaysOfWeek:string):integer;
            function  WeekMaskToWeekDays (iDaysOfWeek:integer):string;
            function  IsRecEqual         (ActItem:TOB;      ApptItem:variant):boolean;
            procedure ActApptLink        (ActItem:TOB;      ApptItem:variant);
            function  ActToAppt          (ActItem:TOB;      ApptItem:variant):variant;
            function  ApptToAct          (ApptItem:variant; ActItem:TOB):TOB;

            procedure CreateOutlookEvtCode;

            procedure SyncAppt;
            procedure CopyAndSaveAct;
            procedure DesyncAppt;
            procedure DesyncAct;

     public
            constructor Create   (iTypeSync:integer; iConflictSolver:integer; iSupprSolver:integer; bSupprFull:boolean);

            function    DoSync   ():integer; // renvoie le nb de rdv synchronisés, ou -1 si erreur
            function    DoImport ():integer; // $$$ JP 07/11/06: import simple depuis outlook
     end;


function  AgendaLanceFiche         (strLequel:string; strArgument:string):string;
procedure AgendaNewActivite        (strUser:string; strDossier:string; bExterne:boolean; bAbsence:boolean);
function  AgendaSendToGi           (strGuidsEvt:string; bUpdateFait:boolean):integer; // $$$ JP - envoi activité agenda vers e-activité gi
procedure AgendaMajDroits_V7; // $$$ JP 16/08/06
procedure AgendaLoadDroits         (TOBDroits:TOB; strGroupes:string='');
function  AgendaMajAutorise        (bExterne, bAbsence:boolean; strDroit:string):boolean;
function  AgendaMajType            (Item:TOB):integer;
function  AgendaGetUsers           (strUser:string; TOBDroits:TOB; bExterne:boolean; bAbsence:boolean; strForceUser:string=''; bWithQuote:boolean=TRUE):string;
function  AgendaGetUsersMaxRights  (strUser:string; TOBDroits:TOB):string;
function  AgendaGetUsersMinRights  (strUser:string; TOBDroits:TOB):string;

// $$$ JP 31/08/07: SQL de sélection des activités de l'agenda déporté ici (avant dans galAgenda)
function AgendaSQLMonoUser         (strUser:string;                         dtMin:TDateTime; dtMax:TDateTime):string;
function AgendaSQLMultiUsers       (strUsers:string; strGroupesConf:string; dtMin:TDateTime; dtMax:TDateTime):string;
function AgendaSQLMonoUserYPL      (strUser:string;                         dtMin:TDateTime; dtMax:TDateTime):string;
function AgendaSQLMultiUsersYPL    (strUsers:string; strGroupesConf:string; dtMin:TDateTime; dtMax:TDateTime):string;

// $$$ JP 09/10/06: création d'une exception (enreg de JUEVENEMENT lié à JURECEXCEPTION)
// $$$ JP 29/05/07: remise en ligne de la création d'une exception à partir d'une série
// $$$ JP 20/08/07: bComplete à FALSE => pas de création d'enreg dans JUEVENEMENT, seulement dans JURECEXCEPTION
function  AgendaCreateException  (Item:TOB; iExceptType:ActExcType=AET_NORMAL; bComplete:boolean=TRUE):string;
procedure AgendaDeleteExceptions (strGuidEvt:string); // strGuidEvt = guid activité d'origine pour laquelle il faut supprimer TOUTES les exceptions

// $$$ JP 05/10/06: calcul des dates des occurences
{$IFNDEF VER150}
function DayOf        (const AValue:TDateTime):Word;
function DaysInAMonth (const AYear, AMonth:Word):Word;
function DaysInMonth  (const AValue:TDateTime):Word;
function MonthOf      (const AValue:TDateTime):Word;
{$ENDIF}

function AgendaGetDateOfDay      (dtStart:TDateTime; iDayOfMonth:integer):TDateTime;
function AgendaGetDateOfInstance (dtStart:TDateTime; strValidWeekDays:string; iInstance:integer):TDateTime;
function AgendaGetInstanceOfDate (dtStart:TDateTime):integer;
function AgendaGetOccurence      (iRecType:ActFrequence; var dtStart:TDateTime; var dtOccurence:TDateTime; var iOccurence:integer; iInterval:integer = 1; iInstance:integer = 0; strDaysOfWeek:string = ''; {iDayOfWeek:integer = 0;} iDayOfMonth:integer = 0; iMonthOfYear:integer = 0):boolean;

// Gestion occurence/exception/série
function AgendaSelectPortee (strTypeOcc:string; strTitle:string; bExceptEnabled:boolean=FALSE):ActPortee;


/////////// IMPLEMENTATION ////////////

implementation

uses
    sysutils, dpjuroutils, hctrls, paramsoc, hent1, hmsgbox, calcolegenericaff,
    entdp, comobj, dpOutils, hStatus, controls, classes,
    galAgendaChoixRec, // $$$ JP 29/05/07
{$IFDEF VER150}
    DateUtils, Variants,
{$ENDIF}

{$IFDEF EAGLCLIENT}
     MaineAGL;
{$ELSE}
     Fe_main;
{$ENDIF}



function AgendaLanceFiche (strLequel:string; strArgument:string):string;
begin
     Result := AGLLanceFiche ('YY', 'YYAGENDA_FIC', '', strLequel, strArgument);
end;

procedure AgendaNewActivite (strUser:string; strDossier:string; bExterne:boolean; bAbsence:boolean);
var
   strUsers           :string;
   strArg             :string;
   TOBDroits          :TOB;
begin
     if VH_DP.SeriaMessagerie = FALSE then
        exit;

     // Chargement des droits agenda
     TOBDroits := TOB.Create ('les droits', nil, -1);
     try
        AgendaLoadDroits (TOBDroits);
        strUsers := AgendaGetUsers (strUser, TOBDroits, bExterne, bAbsence);
     finally
           TOBDroits.Free;
     end;
     if strUsers = '' then
        exit;

     // Création d'une nouvelle activité
     strArg := 'ACTION=CREATION';

     // Utilisateur par défaut: celui connecté
     strArg := strArg + ';JEV_USER1=' + strUser;

     // Utilisateurs autorisés en modification
     strArg := strArg + ';USERS=' + strUsers;

     // Dossier par défaut: celui en sélection par clic droit
     strArg := strArg + ';JEV_NODOSSIER=' + strDossier;

     // Type d'activité: absence, rdv interne, rdv externe
     if bExterne = TRUE then
          strArg := strArg + ';JEV_EXTERNE=X;JEV_ABSENCE=-'
     else if bAbsence = TRUE then
          strArg := strArg + ';JEV_EXTERNE=-;JEV_ABSENCE=X'
     else
          strArg := strArg + ';JEV_EXTERNE=-;JEV_ABSENCE=-';

     // Fiche agenda
     AgendaLanceFiche ('', strArg);
end;


{***********A.G.L.***********************************************
Auteur  ...... : J Pomat
Créé le ...... : 05/05/2004
Modifié le ... :   /  /
Description .. : transfert agenda vers eactivite de la GI
Mots clefs ... :
*****************************************************************}
function AgendaSendOneToGi (TOBUneActivite:TOB):boolean; // $$$ JP 12/04/07 obsolète apparement //; var iCompteur: Integer):boolean;
var
   TOBEAct                             :TOB;
   TOBRess                             :TOB;
   TOBAff                              :TOB;
   TOBArt                              :TOB;
   iDateDeb, iDateFin                  :TDateTime;
   dQteHeure                           :double;
   sAff0, sAff1, sAff2, sAff3, sAff4   :string;
   vNote                               :variant;
   //Q                                   :TQuery;
   TOBNumLig                           :TOB;
   iNumLigne                           :integer;
begin
     // Par défaut, non transféré
     Result := FALSE;

     // Tables liées: ressource, affaire et article
     TOBEAct   := nil;
     TOBRess   := TOB.Create ('la ressource', nil, -1);
     TOBAff    := TOB.Create ('la mission', nil, -1);
     TOBArt    := TOB.Create ('la prestation', nil, -1);
     try
        // Faire recherche de la ressource, du tiers, du code prestation...
        TOBRess.LoadDetailFromSQL ('SELECT ARS_RESSOURCE,ARS_TYPERESSOURCE FROM RESSOURCE WHERE ARS_UTILASSOCIE="' + TOBUneActivite.GetString ('JEV_USER1') + '"');
        if TOBRess.Detail.Count = 0 then
           exit;
        TOBAff.LoadDetailFromSQL ('SELECT AFF_AFFAIRE,AFF_TIERS FROM AFFAIRE WHERE AFF_AFFAIRE="' + TOBUneActivite.GetString ('JEV_AFFAIRE') + '"');
        if TOBAff.Detail.Count = 0 then
           exit;
        TOBArt.LoadDetailFromSQL ('SELECT GA_ARTICLE,GA_CODEARTICLE,GA_ACTIVITEREPRISE FROM ARTICLE JOIN JUTYPEEVT ON GA_CODEARTICLE=JTE_CODEARTICLE AND JTE_CODEEVT="' + TOBUneActivite.GetValue ('JEV_CODEEVT') + '" WHERE GA_TYPEARTICLE="PRE"');
        if TOBArt.Detail.Count = 0 then
           exit;

        // Découpage du code affaire
        CodeAffaireDecoupe (TOBAff.Detail [0].GetValue ('AFF_AFFAIRE'), sAff0, sAff1, sAff2, sAff3, sAff4, taCreat, FALSE);

        // Dates de l'activité
        iDateDeb := TOBUneActivite.GetValue ('JEV_DATE');
        iDateFin := TOBUneActivite.GetValue ('JEV_DATEFIN');

        // $$$ JP 29/03/2005 - Transfert des seules activités sur une même journées (en attendant la gestion des répétitions, et du calendrier GI dans l'agenda)
        if Trunc (iDateDeb) <> Trunc (iDateFin) then
           exit;

        // $$$ JP 03/11/04 - on calcul au 1/4 d'heure près, au plus proche (00-07 min' => 0; 08-22 min' => 1/4 d'heure; 23-37 min' => 2/4 d'heure)
        dQteHeure := Round (96 * (Frac (iDateFin) - Frac (iDateDeb))) / 4;

        // Première initialisation de la clé e-Activité, pour une série d'enregistrements
        iNumLigne := -1;
        TOBNumLig := TOB.Create ('le num ligne', nil, -1);
        try
           TOBNumLig.LoadDetailFromSQL ('SELECT MAX(EAC_NUMLIGNE) AS MAXNUMLIG FROM EACTIVITE WHERE EAC_UTILISATEUR="' + V_PGI.User + '" HAVING COUNT(EAC_NUMLIGNE)>0');
           if TOBNumLig.Detail.Count > 0 then
               iNumLigne := TOBNumLig.Detail [0].GetValue ('MAXNUMLIG')
           else
               iNumLigne := 0;
        finally
               FreeAndNil (TOBNumLig);
        end;
        {if iCompteur=-1 then
        begin
          Q := OpenSQL ('SELECT MAX(EAC_NUMLIGNE) FROM EACTIVITE WHERE EAC_UTILISATEUR="'+V_PGI.User+'"', True);
          if (Q<>nil) and (Not Q.EOF) and (not VarIsNull(Q.Fields[0].Value)) then
            iNumLigne := Q.Fields [0]. iCompteur := Q.Fields[0].Value
          else
            iCompteur := 0;
          Ferme(Q);
          end;}

        // Incrémentation de la clé
        Inc (iNumLigne);

        // Autant de ligne que de nombre de jour de l'activité DP (pas clair, à revoir quand activité répétable actif)
        // Insertion nouvelle ligne activité GI - Clé EACTIVITE: EAC_TYPEACTIVITE,EAC_AFFAIRE,EAC_RESSOURCE, EAC_DATEACTIVITE, EAC_TYPEARTICLE,EAC_NUMLIGNE
        TOBEAct := TOB.Create ('EACTIVITE', nil, -1);

        // MD 23/11/06 - Nouvelle clé de EACTIVITE en v7 + remplacement AddChampSupValeur par PutValue
        TOBEAct.PutValue ('EAC_UTILISATEUR', V_PGI.User);
        TOBEAct.PutValue ('EAC_NUMLIGNE', iNumLigne);

        TOBEAct.PutValue ('EAC_DATEACTIVITE', integer (Trunc (iDateDeb)));
        TOBEAct.PutValue ('EAC_TYPEACTIVITE', 'REA');
        TOBEAct.PutValue ('EAC_ACTORIGINE', 'AGE');
        TOBEAct.PutValue ('EAC_AFFAIRE', TOBAff.Detail [0].GetValue ('AFF_AFFAIRE'));
        TOBEAct.PutValue ('EAC_AFFAIRE0', sAff0);
        TOBEAct.PutValue ('EAC_AFFAIRE1', sAff1);
        TOBEAct.PutValue ('EAC_AFFAIRE2', sAff2);
        TOBEAct.PutValue ('EAC_AFFAIRE3', sAff3);
        TOBEAct.PutValue ('EAC_AVENANT', sAff4);
        TOBEAct.PutValue ('EAC_TIERS', TOBAff.Detail [0].GetValue ('AFF_TIERS'));
        TOBEAct.PutValue ('EAC_TYPERESSOURCE', TOBRess.Detail [0].GetValue ('ARS_TYPERESSOURCE'));
        TOBEAct.PutValue ('EAC_RESSOURCE', TOBRess.Detail [0].GetValue ('ARS_RESSOURCE'));
        TOBEAct.PutValue ('EAC_ARTICLE', TOBArt.Detail [0].GetValue ('GA_ARTICLE'));
        TOBEAct.PutValue ('EAC_CODEARTICLE', TOBArt.Detail [0].GetValue ('GA_CODEARTICLE'));
        TOBEAct.PutValue ('EAC_TYPEARTICLE', 'PRE');
        TOBEact.PutValue ('EAC_ACTIVITEREPRIS', TOBUneActivite.GetValue ('GA_ACTIVITEREPRISE'));
        TOBEAct.PutValue ('EAC_LIBELLE', TOBUneActivite.GetValue ('JEV_EVTLIBELLE'));
        TOBEAct.PutValue ('EAC_UNITE', 'H');

        // $$$ JP 03/11/04
        TOBEAct.PutValue ('EAC_QTE', dQteHeure);
        // MD 23/11/06 - Clé calculée plus haut !
        // TOBEAct.PutValue ('EAC_NUMLIGNE', CalculNumCle ('EAC_NUMLIGNE', 'EACTIVITE', 'EAC_TYPEACTIVITE="REA" AND EAC_AFFAIRE="' + TOBAff.Detail [0].GetValue ('AFF_AFFAIRE') + '" AND EAC_RESSOURCE="' + TOBRess.Detail [0].GetValue ('ARS_RESSOURCE') + '" AND EAC_DATEACTIVITE="' + USDATETIME (Trunc (iDateDeb)) + '" AND EAC_TYPEARTICLE="PRE"'));

        // Insertion de la note de l'événement dans l'eactivité (si existe)
        vNote := TOBUneActivite.GetValue ('JEV_NOTEEVT');
        if (VarIsNull (vNote) = FALSE) and (string (vNote) <> '') then
           TOBEAct.PutValue ('EAC_DESCRIPTIF', vNote);

        // Mise à jour dans la base
        TOBEAct.InsertDB (nil);
        TOBEAct.Free;
        TOBEAct := nil;

        // Ok, les lignes ont été insérées, jour par jour
        Result := TRUE;
     finally
            TOBArt.Free;
            TOBAff.Free;
            TOBRess.Free;
            TOBEAct.Free;
     end;
end;

// $$$ JP 15/03/06 - il faut que les guids soient bien avec les ", exemple: "guid1","guid2","guid3"
function AgendaSendToGi (strGuidsEvt:string; bUpdateFait:boolean):integer;
var
   TOBActivite   :TOB;
   i             :integer; // $$$ JP 12/04/07 , iCompteur  :integer;
   strGuid       :string;
begin
     // Par défaut, le nb d'élément transférer est nul
     Result := 0;

     // Doit y avoir au moins une activité agenda à transférer, et que le transfert automatique vers la GI soit autorisé
     if strGuidsEvt = '' then
        exit;
     if GetParamSocSecur ('SO_AGEGENEREGI', True) = FALSE then
        exit;

     // Dans une transaction, car on doit transférer TOUT ou RIEN (si problème)
     TOBActivite := TOB.Create ('les activites', nil, -1);
     BeginTrans;
     try
     try
        // Pour calcul automatique de la clé e-Activité
        // $$$ JP 12/04/07 obsolète apparement iCompteur := -1;

        // Transfert en Gi des activité agenda spécifiées
        // $$$ JP 29/03/07: on ne prends pas en compte les activité déjà transférées en GI
        TOBActivite.LoadDetailFromSQL ('SELECT * FROM JUEVENEMENT WHERE JEV_GENEREGI<>"X" AND JEV_AFFAIRE<>"" AND JEV_GUIDEVT IN (' + strGuidsEvt + ')');
        for i := 0 to TOBActivite.Detail.Count-1 do
        begin
             // $$$ JP 30/03/07: il faut ne prendre en màj "GENEREGI" que ceux qui ont été bien transféré en GI
             if AgendaSendOneToGi (TOBActivite.Detail [i]) = TRUE then // $$$ JP 12/04/07 , iCompteur) = TRUE then
                 Result := Result + 1
             else
                 TOBActivite.Detail [i].PutValue ('JEV_GUIDEVT', '');
        end;

        // Passage à l'état fait de ces activité agenda
        if bUpdateFait = TRUE then
        begin
             // $$$ JP 30/03/07: mise à jour du FAIT séparément du GENEREGI, pour ne mettre à jour GENEREGI que pour ceux qui ont été réellement transféré en GI
             ExecuteSQL ('UPDATE JUEVENEMENT SET JEV_FAIT="X",JEV_DATEMODIF="' + UsDateTime (Now) + '",JEV_UTILISATEUR="' + V_PGI.User + '" WHERE JEV_FAIT<>"X" AND JEV_GUIDEVT IN (' + strGuidsEvt + ')');

             // $$$ JP 30/03/07: il faut ne prendre en màj "GENEREGI" que ceux qui ont été bien transféré en GI
             strGuidsEvt := '';
             for i := 0 to TOBActivite.Detail.Count-1 do
             begin
                  strGuid := Trim (TOBActivite.Detail [i].GetString ('JEV_GUIDEVT'));
                  if strGuid <> '' then
                  begin
                       if strGuidsEvt <> '' then
                          strGuidsEvt := strGuidsEvt + '","';
                       strGuidsEvt := strGuidsEvt + strGuid;
                  end;
             end;
             if strGuidsEvt <> '' then
                ExecuteSQL ('UPDATE JUEVENEMENT SET JEV_GENEREGI="X",JEV_DATEMODIF="' + UsDateTime (Now) + '",JEV_UTILISATEUR="' + V_PGI.User + '" WHERE JEV_GENEREGI<>"X" AND JEV_GUIDEVT IN ("' + strGuidsEvt + '")');
        end;

        // Valider les mises à jour dans la base de donnée
        CommitTrans;
        if Result > 0 then
           PgiInfo (IntToStr (Result) + ' activité(s) sur ' + IntToStr (TOBActivite.Detail.Count) + ' ont été générées dans l''e-activité');
     except
           Result := -1;
           PgiInfo ('Impossible de mettre à jour les activités sélectionnées');
           RollBack;
     end
     finally
            TOBActivite.Free;
     end;
end;

// $$$ JP 16/08/06: moulinette pour clé YX_CODE
procedure AgendaMajDroits_V7;
var
   TOBDroits   :TOB;
   TOBUnDroit  :TOB;
   i           :integer;
   strMaster   :string;
   strSlave    :string;
   strKey      :string;
   strPrevKey  :string;
begin
     TOBDroits := TOB.Create ('CHOIXEXT', nil, -1);
     try
        // Chargement de tous les droits agenda existants, sans doublon
        TOBDroits.LoadDetailDBFromSQL ('CHOIXEXT', 'SELECT YX_TYPE,YX_CODE,YX_LIBELLE,YX_ABREGE,YX_LIBRE FROM CHOIXEXT WHERE YX_TYPE="DAU" ORDER BY YX_LIBELLE,YX_ABREGE,YX_LIBRE');

        // Suppression de ces droits
        ExecuteSQL ('DELETE CHOIXEXT WHERE YX_TYPE="DAU"');

        // Mise à jour de la clé: user1 # user2, en gardant le droit déjà défini, et tant qu'à faire enlever les espaces en trop sur libelle et abregé
        strPrevKey := '';
        i          := 0;
        while i < TOBDroits.Detail.Count do
        begin
             TOBUnDroit := TOBDroits.Detail [i];
             strMaster  := Trim (TOBUnDroit.GetString ('YX_LIBELLE'));
             strSlave   := Trim (TOBUnDroit.GetString ('YX_ABREGE'));
             strKey     := strMaster + '#' + strSlave;

             // Si doublon de clé (à cause pb doublon v6.5 et v7.00), on l'ignore: revient à prendre le droit le plus faible, puisque trié par YX_LIBRE
             if strKey = strPrevKey then
                TOBUnDroit.Free
             else
             begin
                  TOBUnDroit.SetString ('YX_CODE',    strKey);
                  TOBUnDroit.SetString ('YX_LIBELLE', strMaster);
                  TOBUnDroit.SetString ('YX_ABREGE',  strSlave);

                  // Droit suivant
                  strPrevKey := strKey;
                  Inc (i);
             end;
        end;

        // Mise à jour en base
        TOBDroits.SetAllModifie (TRUE);
        TOBDroits.InsertOrUpdateDB;
     finally
            TOBDroits.Free;
     end;
end;

// CHARGEMENT DES DROITS AGENDA DE L'UTILISATEUR CONNECTE
procedure AgendaLoadDroits (TOBDroits:TOB; strGroupes:string);
var
   strRequete   :string;
begin
     // Chargement des droits de l'utilisateur connecté sur les autres utilisateurs (pour les groupes de travail définis)
     strRequete := 'SELECT YX_CODE,YX_LIBELLE,YX_ABREGE,US_LIBELLE,YX_LIBRE FROM CHOIXEXT ';
     strRequete := strRequete + 'JOIN UTILISAT ON YX_ABREGE=US_UTILISATEUR ';
     strRequete := strRequete + 'WHERE YX_TYPE="DAU" AND YX_LIBELLE="' + V_PGI.User + '" AND YX_LIBRE<>"0" ';
     if strGroupes <> '' then
        strRequete := strRequete +'AND EXISTS (SELECT 1 FROM USERCONF WHERE YX_ABREGE=UCO_USER AND UCO_GROUPECONF IN (' + strGroupes + ')) ';
     strRequete := strRequete + 'ORDER BY YX_CODE';
     TOBDroits.LoadDetailFromSQL (strRequete);
end;

function AgendaMajAutorise (bExterne, bAbsence:boolean; strDroit:string):boolean;
begin
     Result := (strDroit = '4') OR ((strDroit = '3') AND (bAbsence = FALSE)) OR ((strDroit = '2') AND (bAbsence = TRUE));
end;

function AgendaMajType (Item:TOB):integer;
begin
     if Item.GetValue ('I_ETAT') = 'ACT_DELEXCEPTION' then
          Result := aaNothing
     else if (Item.GetString ('I_PERS') = 'X') and (Item.GetValue ('I_CODE') <> V_PGI.User) then
          Result := aaNothing
     else if AgendaMajAutorise (Item.GetValue ('I_EXTERNE') = 'X', Item.GetValue ('I_ABSENCE') = 'X', string (Item.GetValue ('I_DROIT'))) = FALSE then
          Result := aaNothing
     else if Item.GetValue ('I_FAIT') = 'X' then
          Result := aaDoneOnly
     else
          Result := aaAll;
end;

function AgendaGetUsers (strUser:string; TOBDroits:TOB; bExterne:boolean; bAbsence:boolean; strForceUser:string; bWithQuote:boolean):string;
var
   TOBUnDroit    :TOB;
begin
     // Par défaut, aucun utilisateur permit pour l'utilisateur spécifié
     Result := '';

     // Enumération des utilisateurs permit pour l'utilisateur spécifié
     TOBUnDroit := TOBDroits.FindFirst (['YX_LIBELLE'], [strUser], TRUE);
     while TOBUnDroit <> nil do
     begin
          if (AgendaMajAutorise (bExterne, bAbsence, TOBUnDroit.GetString ('YX_LIBRE')) = TRUE) or (strForceUser = TOBUnDroit.GetString ('YX_ABREGE')) then
             if bWithQUote = TRUE then
                 Result := Result + '"' + TOBUnDroit.GetString ('YX_ABREGE') + '",'
             else
                 Result := Result + TOBUnDroit.GetString ('YX_ABREGE') + ',';

          // Droit suivant
          TOBUnDroit := TOBDroits.FindNext (['YX_LIBELLE'], [strUser], TRUE);
     end;
     if Result <> '' then
        Result := Copy (Result, 1, Length (Result)-1);
end;

function AgendaGetUsersMaxRights (strUser:string; TOBDroits:TOB):string;
var
   TOBUnDroit    :TOB;
begin
     // Par défaut, aucun utilisateur sur l'agenda duquel l'utilisateur spécifié a tous les droits agenda
     Result := '';

     // Enumération des utilisateurs sur l'agenda desquels l'utilisateur spécifié a tous les droits
     TOBUnDroit := TOBDroits.FindFirst (['YX_LIBELLE'], [strUser], TRUE);
     while TOBUnDroit <> nil do
     begin
          if TOBUnDroit.GetString ('YX_LIBRE') = '4' then
             Result := Result + '"' + TOBUnDroit.GetString ('YX_ABREGE') + '",';

          // Droit suivant
          TOBUnDroit := TOBDroits.FindNext (['YX_LIBELLE'], [strUser], TRUE);
     end;
     if Result <> '' then
        Result := Copy (Result, 1, Length (Result)-1);
end;

function AgendaGetUsersMinRights (strUser:string; TOBDroits:TOB):string;
var
   TOBUnDroit    :TOB;
begin
     // Par défaut, aucun utilisateur sur l'agenda duquel l'utilisateur spécifié a au moins le droit de lecture
     Result := '';

     // Enumération des utilisateurs sur l'agenda desquels l'utilisateur spécifié a au moins le droit de lecture
     TOBUnDroit := TOBDroits.FindFirst (['YX_LIBELLE'], [strUser], TRUE);
     while TOBUnDroit <> nil do
     begin
          if TOBUnDroit.GetString ('YX_LIBRE') <> '0' then
             Result := Result + '"' + TOBUnDroit.GetString ('YX_ABREGE') + '",';

          // Droit suivant
          TOBUnDroit := TOBDroits.FindNext (['YX_LIBELLE'], [strUser], TRUE);
     end;
     if Result <> '' then
        Result := Copy (Result, 1, Length (Result)-1);
end;

function AgendaSQLMonoUser (strUser:string; dtMin:TDateTime; dtMax:TDateTime):string;
begin
     Result := 'SELECT JEV_GUIDEVT,JEV_USER1,YX_LIBELLE,YX_ABREGE,YX_LIBRE,JEV_CODEEVT,JEV_DATE,JEV_DATEFIN,';
     Result := Result + 'JEV_EVTLIBELLE,JEV_LIEU,JEV_NODOSSIER,DOS_LIBELLE,JEV_GUIDPER,JEV_AFFAIRE,';
     Result := Result + 'JEV_FAIT,JEV_EXTERNE,JEV_ABSENCE,JEV_PERS,JEV_OCCURENCEEVT,JUEVTRECURRENCE.*,';
     Result := Result + 'JEX_GUIDEVT,JEX_RECURNUM,JEX_EXCEPTNUM,JEX_EXCEPTTYPE,JEX_GUIDEVTEXCEPT ';

     Result := Result + 'FROM JUEVENEMENT ';

     Result := Result + 'LEFT JOIN DOSSIER ON JEV_NODOSSIER=DOS_NODOSSIER ';
     Result := Result + 'LEFT JOIN JUEVTRECURRENCE ON JEV_GUIDEVT=JEC_GUIDEVT AND JEC_RECURNUM=1';
     Result := Result + 'LEFT JOIN JURECEXCEPTION ON ((JEX_GUIDEVT=JEC_GUIDEVT AND JEX_RECURNUM=JEC_RECURNUM) OR JEX_GUIDEVTEXCEPT=JEV_GUIDEVT) ';
     Result := Result + 'RIGHT JOIN CHOIXEXT ON (YX_TYPE="DAU" AND YX_LIBELLE="' + V_PGI.User + '" AND YX_ABREGE=JEV_USER1 AND YX_LIBRE<>"0") ';

     Result := Result + 'WHERE JEV_FAMEVT="ACT" AND JEV_DOMAINEACT<>"JUR" ';
     Result := Result + 'AND JEV_USER1="' + strUser + '" ';

     Result := Result + 'AND (JEV_DATEFIN>="' + USTIME (dtMin) + '" OR JEV_OCCURENCEEVT<>"SIN") ';
     Result := Result + 'AND JEV_DATE<"' + USTIME (dtMax) + '"';
end;

function AgendaSQLMultiUsers (strUsers:string; strGroupesConf:string; dtMin:TDateTime; dtMax:TDateTime):string;
begin
     // On filtre un peu les users et les groupes de conf
     strUsers       := Trim (UpperCase (strUsers));
     strGroupesConf := Trim (UpperCase (strGroupesConf));

     Result := 'SELECT DISTINCT JEV_GUIDEVT,JEV_USER1,YX_LIBRE,JEV_CODEEVT,JEV_DATE,JEV_DATEFIN,';
     Result := Result + 'JEV_EVTLIBELLE,JEV_LIEU,JEV_NODOSSIER,DOS_LIBELLE,JEV_GUIDPER,JEV_AFFAIRE,';
     Result := Result + 'JEV_FAIT,JEV_EXTERNE,JEV_ABSENCE,JEV_PERS,US_LIBELLE,JEV_OCCURENCEEVT,JUEVTRECURRENCE.*,';
     Result := Result + 'JEX_GUIDEVT,JEX_RECURNUM,JEX_EXCEPTNUM,JEX_EXCEPTTYPE,JEX_GUIDEVTEXCEPT ';

     Result := Result + 'FROM JUEVENEMENT ';

     Result := Result + 'LEFT JOIN DOSSIER ON JEV_NODOSSIER=DOS_NODOSSIER ';
     Result := Result + 'LEFT JOIN UTILISAT ON JEV_USER1=US_UTILISATEUR ';
     Result := Result + 'LEFT JOIN JUEVTRECURRENCE ON JEV_GUIDEVT=JEC_GUIDEVT ';
     Result := Result + 'LEFT JOIN JURECEXCEPTION ON ((JEX_GUIDEVT=JEC_GUIDEVT AND JEX_RECURNUM=JEC_RECURNUM) OR JEX_GUIDEVTEXCEPT=JEV_GUIDEVT) ';
     Result := Result + 'RIGHT JOIN CHOIXEXT ON (YX_TYPE="DAU" AND YX_LIBELLE="' + V_PGI.User + '" AND YX_ABREGE=JEV_USER1 AND YX_LIBRE<>"0") ';
     Result := Result + 'LEFT JOIN USERCONF ON JEV_USER1=UCO_USER ';

     Result := Result + 'WHERE JEV_FAMEVT="ACT" AND JEV_DOMAINEACT<>"JUR" ';
     if (strUsers <> '') and (strUsers <> '<<TOUS>>') then
        Result := Result + 'AND JEV_USER1 IN ("' + StringReplace (strUsers, ';', '","', [rfReplaceAll]) + '") ';
     if (strGroupesConf <> '') and (strGroupesConf <> '<<TOUS>>') then
        Result := Result + 'AND UCO_GROUPECONF IN ("' + StringReplace (strGroupesConf, ';', '","', [rfReplaceAll]) + '") ';

     Result := Result + 'AND (JEV_DATEFIN>="' + USTIME (dtMin) + '" OR JEV_OCCURENCEEVT<>"SIN") ';
     Result := Result + 'AND JEV_DATE<"' + USTIME (dtMax) + '"';
end;

function AgendaSQLMonoUserYPL (strUser:string; dtMin:TDateTime; dtMax:TDateTime):string;
begin
     Result := 'SELECT YRS_USER AS JEV_USER1,YX_LIBELLE,YX_ABREGE,YX_LIBRE,YPL_DATEDEBUT AS JEV_DATE,YPL_DATEFIN AS JEV_DATEFIN,';
     Result := Result + 'YPL_LIBELLE AS JEV_EVTLIBELLE,';
     Result := Result + '"" AS JEV_GUIDEVT,"" AS JEV_CODEEVT,"-" AS JEV_EXTERNE,"-" AS JEV_ABSENCE,"" AS JEV_LIEU,"" AS JEV_GUIDPER,';
     Result := Result + '"" AS JEV_NODOSSIER,"" AS DOS_LIBELLE,"" AS JEV_AFFAIRE,"SIN" AS JEV_OCCURRENCEEVT,';
     Result := Result + 'NULL AS JEC_GUIDEVT,NULL AS JEC_RECURRENCETYPE,NULL AS JEC_START,NULL AS JEC_END,NULL AS JEC_RECURNUM,';
     Result := Result + 'NULL AS JEC_INTERVALLE,NULL AS JEC_NBOCCURRENCE,NULL AS JEC_WEEKDAYS,';
     Result := Result + 'NULL AS JEC_WEEKOFMONTH,NULL AS JEC_DAYOFMONTH,NULL AS JEC_MONTHOFYEAR,';
     Result := Result + 'YPL_PRIVE AS JEV_PERS, YPL_PREFIXE FROM YPLANNING ';

     Result := Result + 'LEFT JOIN YRESSOURCE ON YRS_GUID=YPL_GUIDYRS ';
     Result := Result + 'RIGHT JOIN CHOIXEXT ON (YX_TYPE="DAU" AND YX_LIBELLE="' + V_PGI.User + '" AND YX_ABREGE=YRS_USER AND YX_LIBRE<>"0") ';

     Result := Result + 'WHERE YPL_PREFIXE<>"JEV" ';
     Result := Result + 'AND YRS_USER="' + strUser + '"';

     Result := Result + 'AND YPL_DATEFIN>="' + USTIME (dtMin) + '" ';
     Result := Result + 'AND YPL_DATEDEBUT<"' + USTIME (dtMax) + '"';
end;

function AgendaSQLMultiUsersYPL (strUsers:string; strGroupesConf:string; dtMin:TDateTime; dtMax:TDateTime):string;
begin
     // On filtre un peu les users et les groupes de conf
     strUsers       := Trim (UpperCase (strUsers));
     strGroupesConf := Trim (UpperCase (strGroupesConf));

     Result := 'SELECT YRS_USER AS JEV_USER1,YX_LIBELLE,YX_ABREGE,YX_LIBRE,YPL_DATEDEBUT AS JEV_DATE,YPL_DATEFIN AS JEV_DATEFIN,';
     Result := Result + 'YPL_LIBELLE AS JEV_EVTLIBELLE,';
     Result := Result + '"" AS JEV_GUIDEVT,"" AS JEV_CODEEVT,"-" AS JEV_EXTERNE,"-" AS JEV_ABSENCE,"" AS JEV_LIEU,"" AS JEV_GUIDPER,';
     Result := Result + '"" AS JEV_NODOSSIER,"" AS DOS_LIBELLE,"" AS JEV_AFFAIRE,"SIN" AS JEV_OCCURRENCEEVT,';
     Result := Result + 'NULL AS JEC_GUIDEVT,NULL AS JEC_RECURRENCETYPE,NULL AS JEC_START,NULL AS JEC_END,NULL AS JEC_RECURNUM,';
     Result := Result + 'NULL AS JEC_INTERVALLE,NULL AS JEC_NBOCCURRENCE,NULL AS JEC_WEEKDAYS,';
     Result := Result + 'NULL AS JEC_WEEKOFMONTH,NULL AS JEC_DAYOFMONTH,NULL AS JEC_MONTHOFYEAR,';
     Result := Result + 'YPL_PRIVE AS JEV_PERS, YPL_PREFIXE FROM YPLANNING ';
     Result := Result + 'LEFT JOIN YRESSOURCE ON YRS_GUID=YPL_GUIDYRS ';
     Result := Result + 'RIGHT JOIN CHOIXEXT ON (YX_TYPE="DAU" AND YX_LIBELLE="' + V_PGI.User + '" AND YX_ABREGE=YRS_USER AND YX_LIBRE<>"0") ';
     Result := Result + 'LEFT JOIN USERCONF ON YRS_USER=UCO_USER ';
     Result := Result + 'WHERE YPL_PREFIXE<>"JEV" ';

     if (strUsers <> '') and (strUsers <> '<<TOUS>>') then
        Result := Result + 'AND YRS_USER IN ("' + StringReplace (strUsers, ';', '","', [rfReplaceAll]) + '") ';
     if (strGroupesConf <> '') and (strGroupesConf <> '<<TOUS>>') then
        Result := Result + 'AND UCO_GROUPECONF IN ("' + StringReplace (strGroupesConf, ';', '","', [rfReplaceAll]) + '") ';

     Result := Result + 'AND YPL_DATEFIN>="' + USTIME (dtMin) + '" ';
     Result := Result + 'AND YPL_DATEDEBUT<"' + USTIME (dtMax) + '"';
end;

// $$$ JP 29/05/07: création d'une exception à partir d'une série
function AgendaCreateException (Item:TOB; iExceptType:ActExcType; bComplete:boolean):string; //; iExceptOcc:integer):string;
var
   strGuidEvt              :string;
   strGuidEvtExcept        :string;
   TOBOcc, TOBExc  :TOB;
begin
     Result  := '';

     // $$$ JP 28/08/07: si exception avec lien sur activité d'origine, il faut prendre ce lien en priorité
     strGuidEvt := Item.GetValue ('I_GUIDEVTREC');
     if strGuidEvt = '' then
        strGuidEvt := Item.GetValue ('I_GUIDEVT');

     // Guid de l'exception qui va être créée
     strGuidEvtExcept := '';

     // Création de l'activité représentant l'exception
     if bComplete = TRUE then
     begin
          TOBOcc := TOB.Create ('JUEVENEMENT', nil, -1);
          try
             TOBOcc.LoadDetailDB ('JUEVENEMENT', '"' + strGuidEvt + '"', '', nil, FALSE);
             if TOBOcc.Detail.Count > 0 then
             begin
                  strGuidEvtExcept := AGLGetGuid;
                  with TOBOcc.Detail [0] do
                  begin
                       PutValue ('JEV_GUIDEVT',      strGuidEvtExcept);
                       PutValue ('JEV_OCCURENCEEVT', 'EXC');

                       // $$$ JP 21/08/07: avoir les dates d'exception: celle de l'item fourni
                       PutValue ('JEV_DATE',    Item.GetValue ('I_DATEDEBUT'));
                       PutValue ('JEV_DATEFIN', Item.GetValue ('I_DATEFIN'));

                       // Ajout dans la base
                       InsertDB (nil);
                  end;
             end;
          finally
             TOBOcc.Free;
          end;
     end;

     // Création de l'exception (lié à la récurrence définie)
     TOBExc  := TOB.Create ('JURECEXCEPTION', nil, -1);
     try
        with TOB.Create ('JURECEXCEPTION', TOBExc, -1) do
        begin
             InitValeurs;

             PutValue ('JEX_GUIDEVT',       strGuidEvt);
             PutValue ('JEX_RECURNUM',      1);
             PutValue ('JEX_EXCEPTNUM',     Item.GetValue ('I_OCCURNUM')); // $$$ JP 29/05/07: le n° d'occurence pour identifier l'exception dans la série
             PutValue ('JEX_EXCEPTTYPE',    iExceptType);
             PutValue ('JEX_GUIDEVTEXCEPT', strGuidEvtExcept);
             if InsertDb (nil) = TRUE then
                Result := strGuidEvtExcept;
        end;
     finally
            TOBExc.Free;
     end;
end;

// $$$ JP 21/08/07 (strGuidEvt = guid activité d'origine pour laquelle il faut supprimer TOUTES les exceptions)
procedure AgendaDeleteExceptions (strGuidEvt:string);
begin
     // Suppression de toutes les exceptions lié à cette activité (dans JURECEXCEPTION et JUEVENEMENT)
     ExecuteSQL ('DELETE JUEVENEMENT WHERE EXISTS (SELECT 1 FROM JURECEXCEPTION WHERE JEX_GUIDEVTEXCEPT=JEV_GUIDEVT AND JEX_GUIDEVT="' + strGuidEvt + '")');
     ExecuteSQL ('DELETE JURECEXCEPTION WHERE JEX_GUIDEVT="' + strGuidEvt + '"');
end;

{$IFNDEF VER150}
function DayOf (const AValue: TDateTime): Word;
var
  LYear, LMonth: Word;
begin
     DecodeDate(AValue, LYear, LMonth, Result);
end;

function MonthOf (const AValue:TDateTime):Word;
var
  LYear, LDay: Word;
begin
     DecodeDate (AValue, LYear, Result, LDay);
end;

function DaysInAMonth (const AYear, AMonth:Word):Word;
begin
     Result := MonthDays [(AMonth = 2) and IsLeapYear(AYear), AMonth];
end;

function DaysInMonth (const AValue:TDateTime):Word;
var
  LYear, LMonth, LDay: Word;
begin
  DecodeDate(AValue, LYear, LMonth, LDay);
  Result := DaysInAMonth(LYear, LMonth);
end;
{$ENDIF}

// Conserve la partie "heure" (partie fractionnaire)
function AgendaGetDateOfDay (dtStart:TDateTime; iDayOfMonth:integer):TDateTime;
var
   wYear, wMonth, wDay  :Word;
begin
     DecodeDate (dtStart, wYear, wMonth, wDay);

     Result := dtStart - wDay + iDayOfMonth;
     if wDay > iDayOfMonth then
        Result := Result + DaysInAMonth (wYear, wMonth);
end;

// Ne conserve PAS la partie "heure" (partie fractionnaire)
function AgendaGetDateOfInstance (dtStart:TDateTime; strValidWeekDays:string; iInstance:integer):TDateTime;
var
   i, iWeekDay   :integer;
   strMonthDays  :string;
begin
     // Par défaut, date non trouvée
     Result := iDate1900;

     // On se positionne au début du mois
     dtStart := DebutDeMois (dtStart);

     // Jours éligibles dans le mois
     iWeekDay := DayOfWeek (dtStart);
     if iWeekDay > 2 then
          strMonthDays := Copy (strValidWeekDays, iWeekDay-1, 9-iWeekDay)
     else if iWeekDay < 2 then
          strMonthDays := strValidWeekDays [7];
     for i := 1 to 5 do
         strMonthDays := strMonthDays + strValidWeekDays;
     strMonthDays := Copy (strMonthDays, 1, DaysInMonth (dtStart));

     // On se positionne sur l'instance du jour voulu
     for i := 1 to Length (strMonthDays) do
     begin
          if strMonthDays [i] <> '-' then
          begin
               // Au pire le dernier jour valide trouvé
               Result := dtStart + i - 1;
               if iInstance < 5 then
               begin
                    Dec (iInstance);
                    if iInstance = 0 then
                       exit;
               end
          end;
     end;
end;

function AgendaGetInstanceOfDate (dtStart:TDateTime):integer;
var
   dtMonth       :TDateTime;
   i, iDayOfWeek :integer;
begin
     // Par défaut, date non trouvée
     Result := 0;

     // On se positionne au début du mois
     dtMonth := DebutDeMois (dtStart);

     // On compte le nb d'instance du jour de semaine de la date spécifié dans le mois, jusqu'à la date spécifiée
     iDayOfWeek := DayOfWeek (dtStart);
     for i := 1 to DayOf (dtStart) do
         if DayOfWeek (dtMonth+i-1) = iDayOfWeek then
            Inc (Result);
end;

// Si iOccurence spécifié (>0), dtOccurence recevra la date début de l'occurence n° iOccurence
function AgendaGetOccurence (iRecType:ActFrequence; var dtStart:TDateTime; var dtOccurence:TDateTime; var iOccurence:integer; iInterval:integer = 1; iInstance:integer = 0; strDaysOfWeek:string = ''; {iDayOfWeek:integer = 0;} iDayOfMonth:integer = 0; iMonthOfYear:integer = 0):boolean;
var
   iDayOfWeek       :integer;
   iNewDayOfWeek    :integer;
   i, iNbDays       :integer;
   wYearA, wYearB   :word;
   wMonthA, wMonthB :word;
   wDay             :word;
   dTime            :extended;
   dtDate           :TDateTime;
begin
     Result := FALSE;

     // $$$ JP 05/06/07: si pas de type de répétition, on ne peut rien calculer
     if iRecType = ACF_NONE then
        exit;

     // L'interval doit toujours être non nul, sauf en mode Annuel (non utilisé), ou bien hebdomadaire jours ouvrables
     if iInterval = 0 then
        if (iRecType <> ACF_YEARLY) and ((iRecType <> ACF_WEEKLY) or (strDaysOfWeek <> 'XXXXX--')) then
            exit
        else
            iInterval := 1;

     // $$$ JP 05/06/07: calcul de la première occurence désormais possible (et toujours ok: = dtStart)
     if iOccurence = 1 then
     begin
          dtOccurence := dtStart;
          Result      := TRUE;
          exit;
     end;

     // L'heure dans la date
     dTime := Frac (dtStart);

     // Calcul de l'occurence en fonction du type de périodicité
     case iRecType of
          ACF_DAILY:
          begin
               // Date de fin par rapport au nb d'occurence
               if iOccurence > 0 then
               begin
                    dtOccurence := dtStart + ((iOccurence-1) * iInterval);
                    Result      := dtOccurence > dtStart;
               end
               else
               begin
                    // Nb d'occurence minimum pour approcher la date de fin (sans la dépasser)
                    iOccurence := Trunc (dtOccurence - dtStart);
                    iOccurence := 1 + (iOccurence div iInterval);
                    Result     := iOccurence > 0;
               end;
          end;

          ACF_WEEKLY:
          begin
               iDayOfWeek := DayOfWeek (dtStart);
               if iDayOfWeek = 1 then
                   iDayOfWeek := 7
               else
                   Dec (iDayOfWeek);
               dtStart := dtStart - iDayOfWeek;

               iNewDayOfWeek := 0;
               for i := iDayOfWeek to 7 do
               begin
                    if strDaysOfWeek [i] <> '-' then
                    begin
                         iNewDayOfWeek := i;
                         break;
                    end;
               end;
               if iNewDayOfWeek = 0 then
               begin
                    for i := 1 to iDayOfWeek - 1 do
                    begin
                         if strDaysOfWeek [i] <> '-' then
                         begin
                              iNewDayOfWeek := i + 7;
                              break;
                         end;
                    end;
               end;

               if iNewDayOfWeek > 0 then
               begin
                    dtStart := dtStart + iNewDayOfWeek;
                    iNbDays := 0;
                    for i := 1 to 7 do
                        if strDaysOfWeek [i] <> '-' then
                           Inc (iNbDays);
                    if iOccurence > 0 then
                    begin
                         dtOccurence := dtStart + 7*((iOccurence-1) div iNbDays)*iInterval;
                         iNbDays := (iOccurence-1) mod iNbDays;
                         while iNbDays > 0 do
                         begin
                              dtOccurence := dtOccurence + 1;
                              Inc (iNewDayOfWeek);
                              if iNewDayOfWeek > 7 then
                              begin
                                   iNewDayOfWeek := 1;
                                   dtOccurence := dtOccurence + 7*(iInterval-1);
                              end;

                              if strDaysOfWeek [iNewDayOfWeek] <> '-' then
                                 Dec (iNbDays);
                         end;

                         Result := dtOccurence > dtStart;
                    end
                    else
                    begin
                         // Nb d'occurence minimum pour approcher la date de fin (sans la dépasser)
                         iOccurence := 0;
                         iDayOfWeek := DayOfWeek (dtStart);
                         if iDayOfWeek = 1 then
                             iDayOfWeek := 7
                         else
                             Dec (iDayOfWeek);
                         dtDate := dtStart;
                         while dtDate <= dtOccurence do
                         begin
                              if strDaysOfWeek [iDayOfWeek] <> '-' then
                                 Inc (iOccurence);

                              dtDate := dtDate + 1;
                              Inc (iDayOfWeek);
                              if iDayOfWeek > 7 then
                              begin
                                   iDayOfWeek := 1;
                                   dtDate     := dtDate + 7*(iInterval-1);
                              end;
                         end;

                         Result := iOccurence > 0;
                    end
               end
               else
                   Result := FALSE;
          end;

          ACF_MONTHLY:
          begin
               if iInstance = 0 then
               begin
                    // Date début correspondant au n° de jour du mois spécifié
                    dtStart := AgendaGetDateOfDay (dtStart, iDayOfMonth);

                    // Date fin théorique de la série
                    if iOccurence > 0 then
                    begin
                         dtOccurence := AgendaGetDateOfDay (IncMonth (dtStart, (iOccurence-1)*iInterval), iDayOfMonth);
                         Result := dtOccurence > dtStart;
                    end
                    else
                    begin
                         DecodeDate (dtStart,     wYearA, wMonthA, wDay);
                         DecodeDate (dtOccurence, wYearB, wMonthB, wDay);
                         iOccurence := 1 + 12*(wYearB-wYearA);
                         if wMonthB >= wMonthA then
                             iOccurence := iOccurence + (wMonthB - wMonthA)
                         else
                             iOccurence := iOccurence - (wMonthA - wMonthB);
                         iOccurence := iOccurence div iInterval;

                         Result := iOccurence > 0;
                    end;
               end
               else
               begin
                    // Date début correspondant au n° de jours de semaine spécifié (4ème jeudi, premier lundi, dernier jour ouvrable, ...)
                    dtStart := AgendaGetDateOfInstance (dtStart, strDaysOfWeek, iInstance);

                    // Date fin théorique de la série
                    if iOccurence > 0 then
                    begin
                         dtOccurence := AgendaGetDateOfInstance (IncMonth (dtStart, (iOccurence-1)*iInterval), strDaysOfWeek, iInstance);

                         // on remet l'heure dans les dates
                         dtStart     := dtStart     + dTime;
                         dtOccurence := dtOccurence + dTime;

                         Result := dtOccurence > dtStart;
                    end
                    else
                    begin
                         DecodeDate (dtStart,     wYearA, wMonthA, wDay);
                         DecodeDate (dtOccurence, wYearB, wMonthB, wDay);
                         iOccurence := 1 + 12*(wYearB-wYearA);
                         if wMonthB >= wMonthA then
                             iOccurence := iOccurence + (wMonthB - wMonthA)
                         else
                             iOccurence := iOccurence - (wMonthA - wMonthB);
                         iOccurence := iOccurence div iInterval;

                         Result := iOccurence > 0;
                    end;
               end;
          end;

          ACF_YEARLY:
          begin
               DecodeDate (dtStart, wYearA, wMonthA, wDay);
               if iInstance = 0 then
               begin
                    dtStart := EncodeDate (wYearA, iMonthOfYear, 1);
                    if iDayOfMonth < DaysInMonth (dtStart) then
                        dtStart := dtStart + iDayOfMonth - 1
                    else
                        dtStart := dtStart + DaysInMonth (dtStart) - 1;
                    if iOccurence > 0 then
                    begin
                         dtOccurence := EncodeDate (wYearA + iOccurence - 1, iMonthOfYear, 1);
                         if iDayOfMonth < DaysInMonth (dtOccurence) then
                             dtOccurence := dtOccurence + iDayOfMonth - 1
                         else
                             dtOccurence := dtOccurence + DaysInMonth (dtOccurence) - 1;

                         // on remet l'heure dans les dates
                         dtStart     := dtStart + dTime;
                         dtOccurence := dtOccurence   + dTime;

                         Result := dtOccurence > dtStart;
                    end
                    else
                    begin
                         DecodeDate (dtOccurence, wYearB, wMonthB, wDay);
                         iOccurence := wYearB-wYearA+1;

                         Result := iOccurence > 0;
                    end;
               end
               else
               begin
                    dtStart := AgendaGetDateOfInstance (EncodeDate (wYearA, iMonthOfYear, 1), strDaysOfWeek, iInstance);
                    if iOccurence > 0 then
                    begin
                         dtOccurence := AgendaGetDateOfInstance (EncodeDate (wYearA+iOccurence-1, iMonthOfYear, 1), strDaysOfWeek, iInstance);

                         // on remet l'heure dans les dates
                         dtStart     := dtStart + dTime;
                         dtOccurence := dtOccurence   + dTime;

                         Result := dtOccurence > dtStart;
                    end
                    else
                    begin
                         DecodeDate (dtOccurence, wYearB, wMonthB, wDay);
                         iOccurence := wYearB-wYearA+1;

                         Result := iOccurence > 0;
                    end;
               end;
          end;
     end;
end;

// $$$ JP 29/05/07: connaitre la portée souhaitée pour la modification (sur la série ou sur l'occurence en cours)
function AgendaSelectPortee (strTypeOcc:string; strTitle:string; bExceptEnabled:boolean=FALSE):ActPortee;
var
   FChoixPortee :TFAgendaChoixRecurrence;
begin
     // Par défaut, modification sur la série
     if strTypeOcc <> 'SIN' then
     begin
          Result := ACP_SERIE;

          // $$$ JP 31/08/07: ne pas activer pur l'instant, le faire dès que les alertes et la synchro outlook prendra en compte les exceptions
          exit;
          
          // Choix de la portee des modifications, uniquement si activité possède une série
          FChoixPortee := TFAgendaChoixRecurrence.Create (nil);
          try
             FChoixPortee.ModifCaption  := strTitle;
             FChoixPortee.ExceptEnabled := bExceptEnabled;
             if FChoixPortee.ShowModal = mrOk then
                 Result := FChoixPortee.ModifPortee
             else
                 Result := ACP_NONE;
          finally
                 FChoixPortee.Free;
          end;
     end
     else
         Result := ACP_OCCURENCE;
end;



//-----------------------------------------------
//           Objet de synchro avec Outlook
//-----------------------------------------------
constructor TOutlookSync.Create (iTypeSync:integer; iConflictSolver:integer; iSupprSolver:integer; bSupprFull:boolean);
begin
     m_iTypeSync       := iTypeSync;
     m_iConflictSolver := iConflictSolver;
     m_iSupprSolver    := iSupprSolver;
     m_bSupprFull      := bSupprFull;

     // Identifiant de la base Cegid prévue pour la synchro
     m_strGuidBase := dpGetBaseGuid;
end;

// $$$ JP 30/10/06
function TOutlookSync.WeekDaysToWeekMask (strDaysOfWeek:string):integer;
var
   iDayMask, i   :integer;
begin
     // Par défaut, aucun jour
     Result := 0;

     // Construction du masque
     if Length (strDaysOfWeek) = 7 then
         strDaysOfWeek := strDaysOfWeek [7] + Copy (strDaysOfWeek, 1, 6)
     else
         strDaysOfWeek := '-' + strDaysOfWeek + StringOfChar ('-', 6-Length (strDaysOfWeek));
     iDayMask := 1;
     for i := 1 to Length (strDaysOfWeek) do
     begin
          if strDaysOfWeek [i] = 'X' then
             Result := Result or iDayMask;
          iDayMask := iDayMask shl 1;
     end;
end;

// $$$ JP 30/10/06
function TOutlookSync.WeekMaskToWeekDays (iDaysOfWeek:integer):string;
var
   iDaysMask   :integer;
begin
     Result    := '';
     iDaysMask := 1;
     while iDaysMask < 128 do
     begin
          if (iDaysOfWeek and iDaysMask) = iDaysMask then
              Result := Result + 'X'
          else
              Result := Result + '-';
          iDaysMask := iDaysMask shl 1;
     end;
     Result := Result + Result [1];
     system.Delete (Result, 1, 1);
end;

// $$$ JP 27/10/06: comparaison des récurrences (car elles doivent être mises à jour que si elles changent réellement)
//                  entre autre conséquence de la modification d'une récurrence (dans outlook surtout): exceptions supprimées
function TOutlookSync.IsRecEqual (ActItem:TOB; ApptItem:variant):boolean;
var
   strActRec      :string;
   bApptRec       :boolean;
   vRecItem       :variant;
   strActRecType  :string;
   dtStart        :TDateTime;
   dtEnd          :TDateTime;
   iActTimeStart  :integer;
   iActTimeEnd    :integer;
   iApptTimeStart :integer;
   iApptTimeEnd   :integer;
   iNbOcc         :integer;
begin
     // On compare surtout le type de récurrence, et dans le même type les propriétés principales
     strActRec := ActItem.GetValue ('JEV_OCCURENCEEVT');
     bApptRec  := ApptItem.IsRecurring;
     if (strActRec = 'REC') and ( bApptRec = TRUE) then
     begin
          vRecItem := ApptItem.GetRecurrencePattern;
          if IDispatch (vRecItem) <> nil then
          begin
               // Comparaison de la plage de périodicité
               dtStart        := ActItem.GetValue ('JEC_START');
               dtEnd          := ActItem.GetValue ('JEC_END');
               iActTimeStart  := Round (Frac (ActItem.GetValue ('JEV_DATE')) * 86400);
               iActTimeEnd    := Round (Frac (ActItem.GetValue ('JEV_DATEFIN')) * 86400);
               iApptTimeStart := Round (Frac (vRecItem.StartTime) * 86400);
               iApptTimeEnd   := Round (Frac (vRecItem.EndTime) * 86400);

               iNbOcc       := ActItem.GetValue ('JEC_NBOCCURRENCE');
               if vRecItem.PatternStartDate <> dtStart then
               begin
                    Result := FALSE;
                    exit;
               end
               else if (iApptTimeStart <> iActTimeStart) or (iApptTimeEnd <> iActTimeEnd) then
               begin
                    Result := FALSE;
                    exit;
               end
               else if (vRecItem.NoEndDate = TRUE) and ((iNbOcc > 0) or (dtEnd > iDate1900)) then
               begin
                    Result := FALSE;
                    exit;
               end
               // $$$ JP 22/01/07: s'il y a une fin dans Outlook et pas dans l'agenda Cegid: différence
               else if (vRecItem.NoEndDate = FALSE) and ((iNbOcc = 0) or (dtEnd <= iDate1900)) then
               begin
                    Result := FALSE;
                    exit;
               end
               else if (dtEnd > iDate1900) and (vRecItem.PatternEndDate <> dtEnd) then
               begin
                    Result := FALSE; // $$$ JP 22/01/07 TRUE;
                    exit;
               end
               // $$$ JP 22/01/07: OccurRences et non pas Occurences!!
               else if (iNbOcc > 0) and (vRecItem.Occurrences <> iNbOcc) then
               begin
                    Result := FALSE; // $$$ JP 22/01/07 TRUE;
                    exit;
               end;

               // Comparaison des propriétés de récurrence selon le type de récurrence
               strActRecType := ActItem.GetValue ('JEC_RECURRENCETYPE');
               case vRecItem.RecurrenceType of
                    olRecursDaily:
                    begin
                         if strActRecType <> 'QUO' then
                             Result := FALSE
                         else
                             Result := vRecItem.Interval = ActItem.GetValue ('JEC_INTERVALLE');
                    end;

                    olRecursWeekly:
                    begin
                         if strActRecType <> 'HEB' then
                             Result := FALSE
                         else
                             Result := (vRecItem.Interval      = ActItem.GetValue ('JEC_INTERVALLE')) and
                                       (vRecItem.DayOfWeekMask = WeekDaysToWeekMask (ActItem.GetValue ('JEC_WEEKDAYS')));
                    end;

                    olRecursMonthly:
                    begin
                         if strActRecType <> 'MEN' then
                             Result := FALSE
                         else
                             Result := (vRecItem.Interval   = ActItem.GetValue ('JEC_INTERVALLE')) and
                                       (vRecItem.DayOfMonth = ActItem.GetValue ('JEC_DAYOFMONTH'));
                    end;

                    olRecursMonthNth:
                    begin
                         if strActRecType <> 'MEX' then
                             Result := FALSE
                         else
                             Result := (vRecItem.Interval      = ActItem.GetValue ('JEC_INTERVALLE')) and
                                       (vRecItem.Instance      = ActItem.GetValue ('JEC_WEEKOFMONTH')) and
                                       (vRecItem.DayOfWeekMask = WeekDaysToWeekMask (ActItem.GetValue ('JEC_WEEKDAYS')));
                    end;

                    olRecursYearly:
                    begin
                         if strActRecType <> 'ANN' then
                             Result := FALSE
                         else
                             Result := (vRecItem.MonthOfYear = ActItem.GetValue ('JEC_MONTHOFYEAR')) and
                                       (vRecItem.DayOfMonth  = ActItem.GetValue ('JEC_DAYOFMONTH'));
                    end;

                    olRecursYearNth:
                    begin
                         if strActRecType <> 'ANX' then
                             Result := FALSE
                         else
                             Result := (vRecItem.Instance      = ActItem.GetValue ('JEC_WEEKOFMONTH')) and
                                       (vRecItem.MonthOfYear   = ActItem.GetValue ('JEC_MONTHOFYEAR')) and
                                       (vRecItem.DayOfWeekMask = WeekDaysToWeekMask (ActItem.GetValue ('JEC_WEEKDAYS')));
                    end;
               end;

               vRecItem := unAssigned;
          end
          else
              Result := FALSE;
     end
     else
         Result := (strActRec <> 'REC') and (bApptRec = FALSE);
end;

procedure TOutlookSync.ActApptLink (ActItem:TOB; ApptItem:variant);
var
   vUserProp   :variant;
   strGuidEvt  :string;
begin
     // Un élément de plus synchronisé ou importé
     Inc (m_iNbSync);

     // Infos de synchro dans le rdv (seulement si synchronisation, pas si import)
     if (IDispatch (ApptItem) <> nil) and (m_iTypeSync <> ostImport) then
     begin
          try
             // Identifiant de la base Pgi dans le rdv outlook (on le stocke pour chaque élément, au cas où plus tard on pourrait synchroniser plusieurs bases Cegid avec le même calendrier Outlook)
             vUserProp := ApptItem.UserProperties.Find ('CEGIDBASEGUID');
             if IDispatch (vUserProp) = nil then
                vUserProp := ApptItem.UserProperties.Add ('CEGIDBASEGUID', olText, FALSE); // TRUE); FQ 11807
             if IDispatch (vUserProp) <> nil then
                vUserProp.Value := m_strGuidBase;

             // Identifiant de l'activité lié dans le rdv outlook
             if ActItem <> nil then
             begin
                  vUserProp := ApptItem.UserProperties.Find ('CEGIDEVTGUID', olText);
                  if IDispatch (vUserProp) = nil then
                     vUserProp := ApptItem.UserProperties.Add ('CEGIDEVTGUID', olText, FALSE);
                  if IDispatch (vUserProp) <> nil then
                  begin
                       strGuidEvt := ActItem.GetValue ('JEV_GUIDEVT');
                       vUserProp.Value := strGuidEvt;
                  end;
             end;
          finally
                 vUserProp := unAssigned;
          end;

          // On enregistre le rdv outlook
          ApptItem.Save;
     end;

     // Infos de synchro dans l'activité Pgi
     if ActItem <> nil then
     begin
          if IDispatch (ApptItem) <> nil then
          begin
               ActItem.PutValue  ('JEV_DUREE', ApptItem.LastModificationTime);
               if m_iTypeSync <> ostImport then
               begin
                    ActItem.PutValue  ('JEV_FOREIGNID',   ApptItem.EntryId);
                    ActItem.PutValue  ('JEV_FOREIGNAPP',  'OUT');
               end;
          end;
          ActItem.AddChampSupValeur ('SYNCHRO', 'SYNC');
     end;
end;

// Création d'un enreg' activité à partir d'un rdv outlook   $$$todo: comment récupérer catégorie de calendrier? (pour savoir si personnel, congés, ...)
function TOutlookSync.ApptToAct (ApptItem:variant; ActItem:TOB):TOB;
var
   strEvtGuid     :string;
   vRecItem       :variant;
begin
   try
     // Si pas d'activité à mettre à jour, c'est une création
     if ActItem = nil then
     begin
          ActItem := TOB.Create ('les activités', m_ActItems, -1);
          strEvtGuid := AglGetGuid;

          with ActItem do
          begin
               AddChampSupValeur ('JEV_GUIDEVT',  strEvtGuid);
               AddChampSup ('JEV_CODEEVT',        FALSE);
               AddChampSupValeur ('JEV_OCCURENCEEVT', 'SIN'); // $$$ JP 29/03/07: par défaut, pas de récurrence
               AddChampSup ('JEV_DATE',           FALSE);
               AddChampSup ('JEV_DATEFIN',        FALSE);
               AddChampSup ('JEV_EVTLIBELLE',     FALSE);
               AddChampSup ('JEV_LIEU',           FALSE);
               AddChampSup ('JEV_AFFAIRE',        FALSE);
               AddChampSup ('JEV_GUIDPER',        FALSE);
               AddChampSup ('JEV_NODOSSIER',      FALSE);
               AddChampSup ('JEV_EXTERNE',        FALSE);
               AddChampSup ('JEV_ABSENCE',        FALSE);
               AddChampSup ('JEV_PERS',           FALSE);
               AddChampSup ('JEV_ALERTE',         FALSE);
               AddChampSup ('JEV_ALERTEDATE',     FALSE);
               AddChampSup ('JEV_DUREE',          FALSE);
               AddChampSup ('JEV_FOREIGNID',      FALSE);
               AddChampSup ('JEV_FOREIGNAPP',     FALSE);

               // $$$ JP22/01/07: pas de récurrence par défaut
               AddChampSupValeur ('JEC_GUIDEVT',  ''); //strEvtGuid);
               AddChampSupValeur ('JEC_RECURNUM', 0); //1); // Pour l'instant qu'une seul récurrence par activité
               AddChampSup ('JEC_INTERVALLE',     FALSE);
               AddChampSup ('JEC_START',          FALSE);
               AddChampSup ('JEC_END',            FALSE);
               AddChampSup ('JEC_NBOCCURRENCE',   FALSE);
               AddChampSup ('JEC_RECURRENCETYPE', FALSE);
               AddChampSup ('JEC_WEEKDAYS',       FALSE);
               AddChampSup ('JEC_DAYOFMONTH',     FALSE);
               AddChampSup ('JEC_WEEKOFMONTH',    FALSE);
               AddChampSup ('JEC_MONTHOFYEAR',    FALSE);
          end;
     end
     else
        strEvtGuid := ActItem.GetValue ('JEV_GUIDEVT');

     // Nouvelle TOB (on récupérera les identifiant de l'activité spécifiée si elle existe)
     with ActItem do
     begin
          PutValue ('JEV_EVTLIBELLE', ApptItem.Subject);
          PutValue ('JEV_LIEU',       ApptItem.Location);
          if ApptItem.ReminderSet = TRUE then
          begin
                PutValue ('JEV_ALERTE',     'X');
                PutValue ('JEV_ALERTEDATE', ApptItem.Start - (ApptItem.ReminderMinutesBeforeStart/1440));
          end
          else
          begin
                PutValue ('JEV_ALERTE',     '-');
                PutValue ('JEV_ALERTEDATE', iDate1900);
          end;

          // Sensitivity <=> JEV_PERS
          if ApptItem.Sensitivity = olPrivate then
              PutValue ('JEV_PERS', 'X')
          else
              PutValue ('JEV_PERS', '-');

          // "Congé" dans Categories <=> JEV_ABSENCE
          if Pos ('Congé (Jour de)', ApptItem.Categories) > 0 then
          begin
               PutValue ('JEV_ABSENCE', 'X');
               PutValue ('JEV_EXTERNE', '-');
               PutValue ('JEV_CODEEVT', 'OUTABS');
          end
          else
          begin
               PutValue ('JEV_ABSENCE', '-');

               // Busy <=> JEV_EXTERNE et JEV_CODEEVT
               if ApptItem.BusyStatus = olOutOfOffice then
               begin
                    PutValue ('JEV_EXTERNE', 'X');
                    PutValue ('JEV_CODEEVT', 'OUTEXT');
               end
               else
               begin
                    PutValue ('JEV_EXTERNE', '-');
                    PutValue ('JEV_CODEEVT', 'OUTINT');
               end;
          end;

          // Création de la récurrence si elle a changée
          if IsRecEqual (ActItem, ApptItem) = FALSE then //if ApptItem.IsRecurring = TRUE then
          begin
               // Les dates
               PutValue ('JEV_DATE',    ApptItem.Start);
               PutValue ('JEV_DATEFIN', ApptItem.End);
               if ApptItem.AllDayEvent = TRUE then
                  PutValue ('JEV_DATEFIN', GetValue ('JEV_DATEFIN') - 1/1440);

               // Si récurrente ou pas, on aliement ou non les propriétés de récurrence de l'activité
               if ApptItem.IsRecurring = FALSE then
                    // Activité pas (ou plus) récurrente. Lors de l'enreg' de la TOB, l'enreg de récurrence sera supprimé
                    PutValue ('JEV_OCCURENCEEVT', 'SIN')
               else
               begin
                    // Lecture des propriétés de récurrence du rendez-vous outlook
                    vRecItem := ApptItem.GetRecurrencePattern;
                    if IDispatch (vRecItem) <> nil then
                    begin
                         // $$$ JP 22/01/07: il faut bien spécifier le guid activité pour pouvoir le supprimer de la base
                         PutValue ('JEC_GUIDEVT',  strEvtGuid);
                         PutValue ('JEC_RECURNUM', 1);

                         PutValue ('JEV_OCCURENCEEVT', 'REC');
                         PutValue ('JEC_INTERVALLE',   vRecItem.Interval);
                         PutValue ('JEC_START',        vRecItem.PatternStartDate);
                         if vRecItem.NoEndDate = FALSE then
                             PutValue ('JEC_END', vRecItem.PatternEndDate)
                         else
                             PutValue ('JEC_END', iDate1900);
                         PutValue ('JEC_NBOCCURRENCE', 0);

                         // Type de récurrence
                         case vRecItem.RecurrenceType of
                              0: // Quotidienne
                                 PutValue ('JEC_RECURRENCETYPE', 'QUO');

                              1: // Hebdomadaire
                                 PutValue ('JEC_RECURRENCETYPE', 'HEB');

                              2: // Mensuel
                                 PutValue ('JEC_RECURRENCETYPE', 'MEN');

                              3: // Mensuel par instance de jour
                                 PutValue ('JEC_RECURRENCETYPE', 'MEX');

                              5: // Annuel
                                 PutValue ('JEC_RECURRENCETYPE', 'ANN');

                              6: // Annuel par instance de jour
                                 PutValue ('JEC_RECURRENCETYPE', 'ANX');
                         end;

                         // Jours de semaine éligibles
                         PutValue ('JEC_WEEKDAYS', WeekMaskToWeekDays (vRecItem.DayOfWeekMask));

                         // Les autres propriétés
                         PutValue ('JEC_DAYOFMONTH',  vRecItem.DayOfMonth);
                         PutValue ('JEC_WEEKOFMONTH', vRecItem.Instance);
                         PutValue ('JEC_MONTHOFYEAR', vRecItem.MonthOfYear);
                    end
                    else
                         PutValue ('JEV_OCCURENCEEVT', 'SIN');
               end;

               vRecItem := unAssigned;
          end
          else
          begin
               // Ne pas oublier les dates (à faire en dernier, sinon IsRecEqual impossible)
               PutValue ('JEV_DATE',    ApptItem.Start);
               PutValue ('JEV_DATEFIN', ApptItem.End);
               if ApptItem.AllDayEvent = TRUE then
                  PutValue ('JEV_DATEFIN', GetValue ('JEV_DATEFIN') - 1/1440)

               // $$$ JP 22/01/07: pourquoi forcer à rdv non périodique, puisque les recurrences sont égales (donc peuvent exister)
               //PutValue ('JEV_OCCURENCEEVT', 'SIN');
          end;
     end;

     // Fait le lien entre l'activité et le rdv outlook (id de l'un dans l'autre, et vice-versa + date synchro=date modif' rdv outlook)
     ActApptLink (ActItem, ApptItem);

     Result := ActItem;
   except
         on E:Exception do
         begin
              PgiInfo ('Erreur de copie du rendez-vous "' + DateTimeToStr (ApptItem.Start) + ' - ' + ApptItem.Subject + '" dans l''activité associée'#10' Erreur: ' + E.Message);
              Result := nil;

              // Il faut marquer l'activité comme traitée (non synchronisée), sinon considérée en phase 2 comme à supprimer
              if (ActItem <> nil) and (ActItem.FieldExists ('SYNCHRO') = FALSE) then
                 ActItem.AddChampSupValeur ('SYNCHRO', 'NOSYNC');
         end;
   end;
end;

// Création d'un rdv outlook à partir d'un enreg' activité
function TOutlookSync.ActToAppt (ActItem:TOB; ApptItem:variant):variant;
var
   vRecItem      :variant;
   strRecType    :string;
   strValue      :string;
   iDaysOfWeek   :integer;
   i             :integer;
begin
   try
     // Si pas de rendez-vous outlook à mettre à jour, c'est une création
     if IDispatch (ApptItem) = nil then
        ApptItem := m_ApptItems.Add (olAppointmentItem);

     if IDispatch (ApptItem) <> nil then
     begin
          // Libellé (entête) et lieu (la note difficile: message de sécurité dans outlook + récupération uniquement du texte (pas d'image, ...))
          strValue          := ActItem.GetValue ('JEV_EVTLIBELLE');
          ApptItem.Subject  := strValue;
          strValue          := ActItem.GetValue ('JEV_LIEU');
          ApptItem.Location := strValue;

          // Alerte (rappel Outlook)
          if ActItem.GetValue ('JEV_ALERTE') = 'X' then
          begin
               ApptItem.ReminderSet                := TRUE;
               ApptItem.ReminderMinutesBeforeStart := (ActItem.GetValue ('JEV_DATE') - ActItem.GetValue ('JEV_ALERTEDATE'))*1440;
          end
          else
               ApptItem.ReminderSet                := FALSE;

          // Sensitivity <=> JEV_PERS
          if ActItem.GetValue ('JEV_PERS') = 'X' then
              ApptItem.Sensitivity := olPrivate
          else                     
              // On change la sensibilité si pas olConfidential avant (sinon on fait rien: le rdv doit garder son statut confidentiel, non géré en pgi)
              if ApptItem.Sensitivity <> olConfidential then
                 ApptITem.Sensitivity := olNormal;

          // "Congé" dans Categories <=> JEV_ABSENCE
          strValue := ApptItem.Categories;
          i        := Pos ('Congé (Jour de)', strValue);
          if ActItem.GetValue ('JEV_ABSENCE') = 'X' then
          begin
               if i = 0 then
                  ApptItem.Categories := strValue + ';Congé (Jour de)';
          end
          else if i > 0 then
               ApptItem.Categories := Copy (strValue, 1, i-1) + Copy (strValue, i+15, Length (strValue));

          // Busy <=> JEV_EXTERNE
          if ActItem.GetValue ('JEV_EXTERNE') = 'X' then
              ApptItem.BusyStatus := olOutOfOffice
          else
              // On change le statut occupé si olOutOfOffice avant (sinon on fait rien: le rdv doit garder son statut occupé ou tentative, non géré en pgi)
              if ApptItem.BusyStatus = olOutOfOffice then
                 ApptItem.BusyStatus := olFree;

          // Maj de la récurrence si nécessaire
          if IsRecEqual (ActItem, ApptItem) = FALSE then
          begin
               // On supprime l'ancienne récurrence (s'il le faut, on la re-créer de toute pièces)
               ApptItem.ClearRecurrencePattern;

               // Date de départ du rdv (la date début de la récurrence s'en déduise)
               ApptItem.Start       := ActItem.GetValue ('JEV_DATE');
               ApptItem.End         := ActItem.GetValue ('JEV_DATEFIN');
               ApptItem.AllDayEvent := (Frac (ActItem.GetValue ('JEV_DATE')) = 0.0) and (1440*Frac (ActItem.GetValue ('JEV_DATEFIN')) > 1438);

               // Si pas de récurrence pgi, on supprime la récurrence outlook
               if ActItem.GetValue ('JEV_OCCURENCEEVT') = 'REC' then
               begin
                    // On créer la récurrence outlook si c'était pas déjà le cas
                    vRecItem := ApptItem.GetRecurrencePattern;

                    // Type de récurrence (hebdomadaire, ...)
                    strRecType := ActItem.GetValue ('JEC_RECURRENCETYPE');
                    if strRecType = 'QUO' then
                    begin
                         vRecItem.RecurrenceType := olRecursDaily;
                         vRecItem.Interval       := ActItem.GetValue ('JEC_INTERVALLE');
                    end
                    else
                    begin
                         // Détermination des jours de semaine éligibles
                         iDaysOfWeek := WeekDaysToWeekMask (ActItem.GetValue ('JEC_WEEKDAYS'));

                         // Selon type de récurrence, on alimente le rdv outlook
                         if strRecType = 'HEB' then
                         begin
                              // $$$ JP 27/10/06: on ne change rien si déjà défini en "tous les jour de travail" (impossible de spécifier à outlook cette récurrence)
                              i := ActItem.GetValue ('JEC_INTERVALLE');
                              vRecItem.RecurrenceType := olRecursWeekly;
                              if i = 0 then
                                 i := 1; // Hélas, outlook renvoie 0 mais ne l'accepte pas en écriture!
                              vRecItem.Interval       := i;
                              vRecItem.DayOfWeekMask  := iDaysOfWeek;
                         end
                         else if strRecType = 'MEN' then
                         begin
                              vRecItem.RecurrenceType := olRecursMonthly;
                              vRecItem.DayOfMonth     := ActItem.GetValue ('JEC_DAYOFMONTH');
                              vRecItem.Interval       := ActItem.GetValue ('JEC_INTERVALLE');
                         end
                         else if strRecType = 'MEX' then
                         begin
                              vRecItem.RecurrenceType := olRecursMonthNth;
                              vRecItem.Instance       := ActItem.GetValue ('JEC_WEEKOFMONTH');
                              vRecItem.Interval       := ActItem.GetValue ('JEC_INTERVALLE');
                              vRecItem.DayOfWeekMask  := iDaysOfWeek;
                         end
                         else if strRecType = 'ANN' then
                         begin
                              vRecItem.RecurrenceType := olRecursYearly;
                              vRecItem.DayOfMonth     := ActItem.GetValue ('JEC_DAYOFMONTH');
                              vRecItem.MonthOfYear    := ActItem.GetValue ('JEC_MONTHOFYEAR');
                         end
                         else if strRecType = 'ANX' then
                         begin
                              vRecItem.RecurrenceType := olRecursYearNth;
                              vRecItem.Instance       := ActItem.GetValue ('JEC_WEEKOFMONTH');
                              vRecItem.MonthOfYear    := ActItem.GetValue ('JEC_MONTHOFYEAR');
                              vRecItem.DayOfWeekMask  := iDaysOfWeek;
                         end;
                    end;

                    // Fréquence, nb d'occurence et/ou date de fin de récurrence
                    if ActItem.GetValue ('JEC_NBOCCURRENCE') > 0 then
                         vRecItem.Occurrences := ActItem.GetValue ('JEC_NBOCCURRENCE')
                    else if ActItem.GetValue ('JEC_END') > iDate1900 then
                         vRecItem.PatternEndDate := ActItem.GetValue ('JEC_END');
               end;
          end
          else
          begin
               // On màj les dates seulement si le rdv non récurrent (l'activité ne l'est pas non plus, sinon IsRecEqual renverrait FALSE)
               if ApptItem.IsRecurring = FALSE then
               begin
                    ApptItem.Start       := ActItem.GetValue ('JEV_DATE');
                    ApptItem.End         := ActItem.GetValue ('JEV_DATEFIN');
                    ApptItem.AllDayEvent := (Frac (ActItem.GetValue ('JEV_DATE')) = 0.0) and (1440*Frac (ActItem.GetValue ('JEV_DATEFIN')) > 1438);
               end;
          end;

          // Lien entre l'activité et le rdv outlook
          ActApptLink (ActItem, ApptItem);

          vRecItem := unAssigned;
     end;

     Result := ApptItem;
   except
         on E:Exception do
         begin
              PgiInfo ('Erreur de copie activité "' + DateTimeToStr (ActItem.GetValue ('JEV_DATE')) + ' - ' + ActItem.GetValue ('JEV_EVTLIBELLE') + '" dans le rendez-vous associé'#10' Erreur: ' + E.Message);
              Result := unAssigned;

              // Il faut marquer l'activité comme traitée (non synchronisée), sinon considérée en phase 2 comme à supprimer
              if ActItem.FieldExists ('SYNCHRO') = FALSE then
                 ActItem.AddChampSupValeur ('SYNCHRO', 'NOSYNC');
         end;
   end;
end;

// Synchro rdv/activités à partir de l'énumération des rendez-vous outlook
procedure TOutlookSync.SyncAppt;
var
   vApptItem        :variant;
   vApptUserProp    :variant;
   strGuidInRdv     :string;
   ActItem          :TOB;
   iItemCopy        :integer;
   dtApptModif      :TDateTime;
   dtActModif       :TDateTime;
   dtLastSynchro    :TDateTime;
   bApptModified    :boolean;
   bActModified     :boolean;
begin
     // On énumère tout les appointmentitem
     vApptItem  := m_ApptItems.GetFirst;
     while IDispatch (vApptItem) <> nil do
     begin
          MoveCur (FALSE);

          // Que les éléments de type rendez-vous
          if vApptItem.Class = olAppointment then
          begin
               // Guid de l'activité déjà synchronisée
               vApptUserProp := vApptItem.UserProperties.Find ('CEGIDEVTGUID');
               if (IDispatch (vApptUserProp) <> nil) and (vApptUserProp.Type = olText) then
                   strGuidInRdv := vApptUserProp.Value
               else
                   strGuidInRdv := '';

               // En mode outlook vers Cegid, ou bidirectionnel, on créer la TOB si changement ou si pas encore connue dans base Cegid
               ActItem := nil;
               if strGuidInRdv <> '' then
               begin
                    ActItem := m_ActItems.FindFirst (['JEV_GUIDEVT'], [strGuidInRdv], TRUE);

                    // $$$ JP 11/12/06: si activité déjà synchronisé, le rdv est (normalement) un rdv dupliqué dans Outlook: on enlève le lien et on fait comme si c'était un rdv outlook non encore synchronisé avec une activité
                    if (ActItem <> nil) and (ActItem.FieldExists ('SYNCHRO') = TRUE) then
                    begin
                         strGuidInRdv := '';
                         ActItem      := nil;
                    end;
               end;

               // Détermination du sens de la copie (pgi vers outlook, outlook vers pgi, dans les deux sens, ou bien rien), dans les éléments existants ou dans des nouveaux éléments si nécessaire
               iItemCopy := ostNone;
               if ActItem = nil then
               begin
                    // Si pas d'activité liée à un rdv pourtant synchronisé, l'activité à donc été supprimée
                    if strGuidInRdv <> '' then
                    begin
                         case m_iSupprSolver of
                              odrKeep:
                                 // Recopie si l'hôte de destination est autorisé à recevoir un élément de ce genre
                                 if (m_iTypeSync <> ostFromCegid) or (m_bSupprFull = TRUE) then
                                    iItemCopy := ostFromOutlook;

                              odrDelete:
                                 // On supprime si l'hôte de suppression est autorisé à se voir enlever un élément de ce genre
                                 if (m_iTypeSync <> ostFromOutlook) or (m_bSupprFull = TRUE) then
                                    iItemCopy := ostDeleteInOutlook;
                         end;
                    end
                    else if m_iTypeSync <> ostFromCegid then
                         // Simple copie (rdv non encore lié à une activité Pgi)
                         iItemCopy := ostFromOutlook;
               end
               else
               begin
                    // $$$ JP 12/04/07: en cwas, il faudra faire autrement, car horloge client pas = à horloge serveur:
                    // pb surtout pour rdv outlook client. Créer un champ de référence JEV_SYNCLOCALDATE
                    // (et aussi JEV_SYNCSERVERDATE pour remplacer JEV_DUREE non prévu au départ pour ça)

                    // Activité et rdv liés: on regarde si l'une et/ou l'autre modifiée depuis dernière synchro (date de synchro référence=JEV_DUREE)
                    // (on prend une marge de latence d'une seconde, car date = virgule flottante)
                    // Il faut aussi ne pas resynchroniser des rdv/activités (généré lors de la création d'un rdv outlook dans une itération précédente)
                    // $$$ JP obligatoirement non synchronisé si on passe ici if ActItem.FieldExists ('SYNCHRO') = FALSE then
                    //begin
                    dtLastSynchro := ActItem.GetValue ('JEV_DUREE');
                    dtApptModif   := vApptItem.LastModificationTime;
                    bApptModified := Abs (dtApptModif  - dtLastSynchro) > dSecond;
                    dtActModif    := ActItem.GetValue ('JEV_DATEMODIF');
                    bActModified  := Abs (dtActModif   - dtLastSynchro) > dSecond;

                    // Selon le mode de synchro
                    case m_iTypeSync of
                         // Dans le sens unique Outlook vers Pgi, on copie le rdv dans l'activité si ce rdv est modifié depuis dernière synchro
                         ostFromOutlook:
                                 if bApptModified = TRUE then
                                    iItemCopy := ostFromOutlook;

                         // Dans le sens unique Pgi vers Outlook, on copie l'activité dans le rdv (sans se poser de question)
                         ostFromCegid:
                                 if bActModified = TRUE then
                                    iItemCopy := ostFromCegid;

                         // En synchro, on copie selon le mode de résolution de conflit (si conflit), sinon le rdv dans l'activité
                         ostFull:
                                 if bApptModified = TRUE then
                                 begin
                                      if bActModified = TRUE then
                                      begin
                                           // Conflit de modification (l'un ou l'autre a raison, ou bien le dernier modifié, ou bien encore création de doublon)
                                           case m_iConflictSolver of
                                                ocrCegid:
                                                   iItemCopy := ostFromCegid;

                                                ocrOutlook:
                                                   iItemCopy := ostFromOutlook;

                                                ocrFull:
                                                   iItemCopy := ostFull;

                                                ocrMostRecent:
                                                   if dtApptModif - dtActModif > dSecond then
                                                        iItemCopy := ostFromOutlook
                                                   else if dtActModif - dtApptModif > dSecond then
                                                        iItemCopy := ostFromCegid
                                                   else
                                                        iItemCopy := ostFull;
                                           end;
                                      end
                                      else
                                          // Uniquement le rdv modifié: on le copie dans l'activité
                                          iItemCopy := ostFromOutlook;
                                 end
                                 else if bActModified = TRUE then
                                      // Uniquement l'activité modifiée: on la copie dans le rdv
                                      iItemCopy := ostFromCegid;
                    end;
                    //end
                    //else
                    //    // Déjà synchronisé: ne rien faire du tout, même pas marquage pour synchro
                    //    iItemCopy := ostAlready;
               end;

               // Copie des éléments dans un sens ou dans les deux (full)
               case iItemCopy of
                    ostNone:
                       if ActItem <> nil then
                          ActItem.AddChampSupValeur ('SYNCHRO', 'NOSYNC');

                    ostFromCegid:
                       ActToAppt (ActItem, vApptItem);

                    ostFromOutlook:
                       ApptToAct (vApptItem, ActItem);

                    ostFull:
                    begin
                         ActToAppt (ActItem,   unAssigned);
                         ApptToAct (vApptItem, nil);
                    end;

                    ostDeleteInOutlook:
                    begin
                         vApptItem.Delete;
                         Inc (m_iNbSync);
                    end;
               end;
          end;

          // Element outlook suivant (normalement un rdv, mais peut être une tâche)
          vApptItem := m_ApptItems.GetNext;
     end;

     vApptItem := unAssigned;
end;

procedure TOutlookSync.CopyAndSaveAct;
var
   i            :integer;
   strEntryId   :string;
   strSync      :string;
   ActItem      :TOB;
   TOBEvt       :TOB;
   TOBRec       :TOB;
   TOBEvtDel    :TOB;
   TOBRecDel    :TOB;
   strSQL       :string;
begin
     // Les tob réelles pour mise à jour de la base une fois le traitement fini
     TOBEvt    := TOB.Create ('JUEVENEMENT', nil, -1);
     TOBRec    := TOB.Create ('JUEVTRECURRENCE', nil, -1);
     if m_iTypeSync <> ostImport then
     begin
          TOBEvtDel := TOB.Create ('JUEVENEMENT', nil, -1);
          TOBRecDel := TOB.Create ('JUEVTRECURRENCE', nil, -1);
     end;

     // On ne synchronise que les éléments non encore synchronisé dans la phase 1 (les nouvelles activités, ou celles déjà synchronisées, mais supprimées dans outlook et pas dans PGI)
     for i := 0 to m_ActItems.Detail.Count - 1 do
     begin
          MoveCur (FALSE);

          // L'activité à synchroniser
          ActItem := m_ActItems.Detail [i];

          // Synchro de cette activité si pas déjà fait
          if ActItem.FieldExists ('SYNCHRO') = FALSE then
          begin
               strEntryId := ActItem.GetValue ('JEV_FOREIGNID');
               if strEntryId <> '' then
               begin
                    // S'il y a un identifiant outlook et que ce rdv outlook n'a pas été traité dans phase 1, c'est qu'il n'existe plus dans outlook: on le supprime ou le garde selon ce qui est spécifié
                    case m_iSupprSolver of
                         odrKeep:
                            // Recopie si l'hôte de destination est autorisé à recevoir un élément de ce genre
                            if (m_iTypeSync <> ostFromOutlook) or (m_bSupprFull = TRUE) then
                                 ActToAppt (ActItem, unAssigned)
                            else
                                 ActItem.AddChampSupValeur ('SYNCHRO', 'NOSYNC'); // à priori inutile (car getvalue d'un champ inexistant renvoie '')

                         odrDelete:
                            // On marque pour suppression si l'hôte de suppression est autorisé à se voir enlever un élément de ce genre
                            if (m_iTypeSync <> ostFromCegid) or (m_bSupprFull = TRUE) then
                            begin
                                 ActItem.AddChampSupValeur ('SYNCHRO', 'DELETED');
                                 Inc (m_iNbSync);
                            end
                            else
                                ActItem.AddChampSupValeur ('SYNCHRO', 'NOSYNC'); // à priori inutile (car getvalue d'un champ inexistant renvoie '')

                         odrNone:
                            ActItem.AddChampSupValeur ('SYNCHRO', 'NOSYNC'); // à priori inutile (car getvalue d'un champ inexistant renvoie '')
                    end;
               end
               // Sinon, on copie l'activité dans un nouveau rdv outlook (si outlook "récepteur" d'élément)
               else if m_iTypeSync <> ostFromOutlook then
                    ActToAppt (ActItem, unAssigned)
               else
                    ActItem.AddChampSupValeur ('SYNCHRO', 'NOSYNC'); // à priori inutile
          end;

          // Si activité synchronisée (dans la phase 1 ou 2), on la copie dans la TOB des éléments à enregistrer
          strSync := ActItem.GetValue ('SYNCHRO');
          if strSync = 'SYNC' then
          begin
               with TOB.Create ('JUEVENEMENT', TOBEvt, -1) do
               begin
                    InitValeurs;
                    PutValue ('JEV_GUIDEVT',      ActItem.GetValue ('JEV_GUIDEVT'));
                    PutValue ('JEV_USER1',        V_PGI.User);
                    PutValue ('JEV_GUIDPER',      ActItem.GetValue ('JEV_GUIDPER'));
                    PutValue ('JEV_NODOSSIER',    ActItem.GetValue ('JEV_NODOSSIER'));
                    PutValue ('JEV_AFFAIRE',      ActItem.GetValue ('JEV_AFFAIRE'));
                    PutValue ('JEV_FAMEVT',       'ACT');
                    PutValue ('JEV_CODEEVT',      ActItem.GetValue ('JEV_CODEEVT'));
                    PutValue ('JEV_EXTERNE',      ActItem.GetValue ('JEV_EXTERNE'));
                    PutValue ('JEV_ABSENCE',      ActItem.GetValue ('JEV_ABSENCE'));
                    PutValue ('JEV_PERS',         ActItem.GetValue ('JEV_PERS'));
                    PutValue ('JEV_DATE',         ActItem.GetValue ('JEV_DATE'));
                    PutValue ('JEV_DATEFIN',      ActItem.GetValue ('JEV_DATEFIN'));
                    PutValue ('JEV_EVTLIBELLE',   ActItem.GetValue ('JEV_EVTLIBELLE'));
                    PutValue ('JEV_LIEU',         ActItem.GetValue ('JEV_LIEU'));
                    PutValue ('JEV_ALERTE',       ActItem.GetValue ('JEV_ALERTE'));
                    PutValue ('JEV_ALERTEDATE',   ActItem.GetValue ('JEV_ALERTEDATE'));
                    PutValue ('JEV_OCCURENCEEVT', ActItem.GetValue ('JEV_OCCURENCEEVT'));
                    if m_iTypeSync <> ostImport then
                    begin
                         PutValue ('JEV_FOREIGNID',    ActItem.GetValue ('JEV_FOREIGNID'));
                         PutValue ('JEV_FOREIGNAPP',   ActItem.GetValue ('JEV_FOREIGNAPP'));
                         PutValue ('JEV_DUREE',        ActItem.GetValue ('JEV_DUREE'));
                    end;

                    // La date de modification doit être exactement la même que celle du rdv outlook, donc on la spécifie nous-même
                    // $$$ JP 12/04/07: pb en eagl (peut être aussi 2/3): date modification souhaitée écrasée lors de la màj de la tob!
                    // On fait autrement: maj des datemodif par UPDATE (après l'insertorupdatedb) sur tous les éléments de la tob
                    //SetDateModif (ActItem.GetValue ('JEV_DUREE'));

                    // $$$ JP 11/12/06: il faut considérer que tous les champs sont modifiés (puisqu'on ne s'est pas fondé sur l'enreg éventuellement existant)
                    SetAllModifie (TRUE);
                    // FQ 12002 : Sauf le champ JEV_NOTEEVT qui doit être conservé
                    SetModifieField('JEV_NOTEEVT', False);
               end;

               // Gestion de la récurrence de l'activité
               if ActItem.GetValue ('JEV_OCCURENCEEVT') = 'REC' then
               begin
                    with TOB.Create ('JUEVTRECURRENCE', TOBRec, -1) do
                    begin
                         InitValeurs;
                         PutValue ('JEC_GUIDEVT',         ActItem.GetValue ('JEC_GUIDEVT'));
                         PutValue ('JEC_RECURNUM',        ActItem.GetValue ('JEC_RECURNUM')); // Pour l'instant, une seule récurrence possible
                         PutValue ('JEC_INTERVALLE',      ActItem.GetValue ('JEC_INTERVALLE'));
                         PutValue ('JEC_START',           ActItem.GetValue ('JEC_START'));
                         PutValue ('JEC_END',             ActItem.GetValue ('JEC_END'));
                         PutValue ('JEC_NBOCCURRENCE',    ActItem.GetValue ('JEC_NBOCCURRENCE'));
                         PutValue ('JEC_RECURRENCETYPE',  ActItem.GetValue ('JEC_RECURRENCETYPE'));
                         PutValue ('JEC_WEEKDAYS',        ActItem.GetValue ('JEC_WEEKDAYS'));
                         PutValue ('JEC_DAYOFMONTH',      ActItem.GetValue ('JEC_DAYOFMONTH'));
                         PutValue ('JEC_WEEKOFMONTH',     ActItem.GetValue ('JEC_WEEKOFMONTH'));
                         PutValue ('JEC_MONTHOFYEAR',     ActItem.GetValue ('JEC_MONTHOFYEAR'));

                         // $$$ JP 11/12/06: il faut considérer que tous les champs sont modifiés (puisqu'on ne s'est pas fondé sur l'enreg éventuellement existant)
                         SetAllModifie (TRUE);
                    end;
               end
               else if (m_iTypeSync <> ostImport) and (ActItem.GetValue ('JEC_GUIDEVT') <> '') then
               begin
                    // $$$ JP 30/10/06: si pas de récurrence, il faut supprimer l'éventuelle récurrence existante
                    with TOB.Create ('JUEVTRECURRENCE', TOBRecDel, -1) do
                    begin
                         InitValeurs;
                         PutValue ('JEC_GUIDEVT',  ActItem.GetValue ('JEC_GUIDEVT'));
                         PutValue ('JEC_RECURNUM', ActItem.GetValue ('JEC_RECURNUM'));
                    end;
               end;
          end
          else if (m_iTypeSync <> ostImport) and (strSync = 'DELETED') then
          begin
               with TOB.Create ('JUEVENEMENT', TOBEvtDel, -1) do
               begin
                    InitValeurs;
                    PutValue ('JEV_GUIDEVT', ActItem.GetValue ('JEV_GUIDEVT'));
               end;

               if ActItem.GetValue ('JEV_OCCURENCEEVT') = 'REC' then
               begin
                    with TOB.Create ('JUEVTRECURRENCE', TOBRecDel, -1) do
                    begin
                         InitValeurs;
                         PutValue ('JEC_GUIDEVT',  ActItem.GetValue ('JEC_GUIDEVT'));
                         PutValue ('JEC_RECURNUM', ActItem.GetValue ('JEC_RECURNUM')); // Pour l'instant, une seule récurrence possible
                    end;
               end;
          end;
     end;

     // On enregistre les activités (et récurrence d'activité)
     BeginTrans;
     try
        // Les activités supprimées
        if m_iTypeSync <> ostImport then
        begin
             for i := 0 to TOBEvtDel.Detail.Count - 1 do
                 TOBEvtDel.Detail [i].DeleteDB;
             for i := 0 to TOBRecDel.Detail.Count - 1 do
                 TOBRecDel.Detail [i].DeleteDB;
        end;

        // Les activités mises à jour ou ajoutées
        // Les activités mises à jour ou ajoutées
        if TOBEvt.Detail.Count > 0 then
        begin
             // $$$ JP 12/04/07: pas le choix, pour fixer sa propre datemodif (en eagl), faut passer par executesql
             if TOBEvt.InsertOrUpdateDB = TRUE then
             begin
                  strSQL := 'UPDATE JUEVENEMENT SET JEV_DATEMODIF=JEV_DUREE WHERE JEV_GUIDEVT IN ("';
                  for i := 0 to TOBEvt.Detail.Count - 1 do
                  begin
                       if i > 0 then
                          strSQL := strSQL + '","';
                       strSQL := strSQL + TOBEvt.Detail [i].GetValue ('JEV_GUIDEVT');
                  end;
                  strSQL := strSQL + '")';
                  TOBEvt.ClearDetail; // On sait jamais
                  ExecuteSQL (strSQL);
             end;
             // $$$ JP 12/04/07 fin
        end;

        // Les récurrences
        if TOBRec.Detail.Count > 0 then
           TOBRec.InsertOrUpdateDB;

        // Transaction terminée
        CommitTrans;
     except
           RollBack;
     end;

     if m_iTypeSync <> ostImport then
     begin
          TOBRecDel.Free;
          TOBEvtDel.Free;
     end;
     TOBRec.Free;
     TOBEvt.Free;
end;

procedure TOutlookSync.DesyncAppt;
var
   vApptItem        :variant;
   vApptUserProp    :variant;
   bDesync          :boolean;
begin
     try
        vApptItem  := m_ApptItems.GetFirst;
        while IDispatch (vApptItem) <> nil do
        begin
             MoveCur (FALSE);

             // Que les éléments de type rendez-vous
             if vApptItem.Class = olAppointment then
             begin
                  // Pour savoir si une désynchro a été effectuée (bien que normalement si CEGIDBASEGUID est présent, CEGIDEVTGUID aussi)
                  bDesync := FALSE;

                  // On enlève le Guid de la base de l'activité déjà synchronisée, et le guid de l'activité
                  vApptUserProp := vApptItem.UserProperties.Find ('CEGIDBASEGUID');
                  if IDispatch (vApptUserProp) <> nil then
                  begin
                       vApptUserProp.Delete;
                       bDesync := TRUE;
                  end;

                  // Guid de l'activité déjà synchronisée
                  vApptUserProp := vApptItem.UserProperties.Find ('CEGIDEVTGUID');
                  if IDispatch (vApptUserProp) <> nil then
                  begin
                       vApptUserProp.Delete;
                       bDesync := TRUE;
                  end;

                  // Si modif' faite, on enregistre le rdv (et on màj le compteur)
                  if bDesync = TRUE then
                  begin
                       vApptItem.Save;
                       Inc (m_iNbSync);
                  end;
             end;

             // Element outlook suivant (normalement un rdv, mais peut être une tâche)
             vApptItem := m_ApptItems.GetNext;
        end;
     finally
            vApptUserProp := unAssigned;
            vApptItem     := unAssigned;
     end;
end;

procedure TOutlookSync.DesyncAct;
begin
     // On enlève toute référence de synchro vers outlook dans la table des activités
     m_iNbSync := m_iNbSync + ExecuteSQL ('UPDATE JUEVENEMENT SET JEV_FOREIGNID="",JEV_FOREIGNAPP="" WHERE (JEV_FOREIGNID<>"") OR (JEV_FOREIGNAPP="OUT")');

     // Il faut enlever la référence vers le calendrier Outlook
     ExecuteSQL ('DELETE CHOIXEXT WHERE YX_TYPE="SYA" AND YX_TYPE="OUT"');
end;

procedure TOutlookSync.CreateOutlookEvtCode;
var
   TOBCodeEvt  :TOB;
   i, iPos     :integer;
   strCodes    :string;
   strCode     :string;
begin
     strCodes := 'OUTEXT;OUTINT;OUTABS;';
     TOBCodeEvt := TOB.Create ('code evt', nil, -1);
     TOBCodeEvt.LoadDetailFromSQL ('SELECT JTE_CODEEVT FROM JUTYPEEVT WHERE JTE_CODEEVT IN ("OUTEXT","OUTINT","OUTABS")');
     for i := 0 to TOBCodeEvt.Detail.Count - 1 do
     begin
          strCode := TOBCodeEvt.Detail [i].GetString ('JTE_CODEEVT');
          iPos := Pos (strCode, strCodes);
          if iPos > 0 then
             Delete (strCodes, iPos, 7);
     end;
     TOBCodeEvt.Free;

     strCode := Trim (ReadTokenST (strCodes));
     while strCode <> '' do
     begin
          with TOB.Create ('JUTYPEEVT', nil, -1) do
          begin
               InitValeurs;
               PutValue ('JTE_CODEEVT',      strCode);
               PutValue ('JTE_FAMEVT',       'ACT');
               PutValue ('JTE_EVTLIBELLE',   'Rendez-vous de Microsoft Outlook');
               PutValue ('JTE_EVTLIBABREGE', 'Rdv Ms-Outlook');
               PutValue ('JTE_VERSION',      1);
               if strCode = 'OUTEXT' then
                   PutValue ('JTE_EXTERNE',  'X')
               else
                   PutValue ('JTE_EXTERNE',  '-');
               if strCode = 'OUTABS' then
                   PutValue ('JTE_ABSENCE',  'X')
               else
                   PutValue ('JTE_ABSENCE',  '-');
               PutValue ('JTE_PREDEFINI',    'CEG');
               InsertDB (nil);
               Free;
          end;

          strCode := Trim (ReadTokenST (strCodes));
     end;
end;

function TOutlookSync.DoSync:integer; // renvoie le nb de rdv synchronisés, ou -1 si erreur
var
   strReq         :string;
   strValue       :string;
   vOutlook       :variant;
   vMapi          :variant;
   vCalFolder     :variant;
   vSyncItem      :variant;
   vUserProp      :variant;
begin
     // Par défaut, aucune synchronisation effectuée
     m_iNbSync := 0;

     // Si mode de synchro sur none, on fait rien
     if (m_iTypeSync = ostNone) or (m_iTypeSync = ostImport) then
        exit;

     // En mode dé-synchronisation, pas la peine de vérifier guid base en cours, ni type d'événement
     if m_iTypeSync <> ostDesynchro then
     begin
          // Identifiant de la base Cegid contenant les activités
          if m_strGuidBase = '' then
          begin
               PgiInfo ('Synchronisation impossible avec Outlook: l''identifiant de la base commune n''est pas initialisé');
               exit;
          end;

          // Avant tout, il faut créer un type d'activité spécifique pour import Outlook (comme pour synchro pda)
          CreateOutlookEvtCode;
     end;

     SourisSablier;
     try
        // Chargement de toutes les activités déjà synchro, et les nouvelles si elles sont dans les critères définis (date début, recurrence...)
        m_ActItems := TOB.Create ('les activités', nil, -1);
        strReq := 'SELECT * FROM JUEVENEMENT LEFT JOIN JUEVTRECURRENCE ON (JEV_GUIDEVT=JEC_GUIDEVT AND JEC_RECURNUM=1) WHERE JEV_FAMEVT="ACT" ';
        strReq := strReq + 'AND JEV_DOMAINEACT<>"JUR" AND JEV_USER1="' + V_PGI.User + '" ';
        strReq := strReq + 'AND (JEV_FOREIGNID="" OR JEV_FOREIGNAPP="OUT") ';
        strReq := strReq + 'ORDER BY JEV_FOREIGNID,JEV_DATE';
        m_ActItems.LoadDetailFromSQL (strReq);

        // Chargement des rendez-vous Outlook
        vOutlook := CreateOleObject ('Outlook.Application');
        if IDispatch (vOutlook) <> nil then
        begin
             // Ok, on se positionne sur le folder du calendrier par défaut et on récupère les rendez-vous synchronisable
             vMapi := vOutlook.GetNamespace ('MAPI');
             if IDispatch (vMapi) <> nil then
             begin
                  vCalFolder := vMapi.GetDefaultFolder (olFolderCalendar);
                  if IDispatch (vCalFolder) <> nil then
                  begin
                       // $$$ JP 03/11/06: on doit connaitre l'identifiant de stockage du calendrier outlook
                       m_strGuidOutlook := VOutlook.ProductCode; //vCalFolder.StoreId;

                       // LEs rdv outlook à synchroniser (peut contenir des éléments autre que des rdv, seront filtrer lors de la synchro)
                       m_ApptItems := vCalFolder.Items; // $$$ JP pour l'instant pas de restriction... Restrict ('[ResponseStatus] >= 0 And [ResponseStatus] <= 3');
                  end;
             end;
        end;

        // Quelques vérifications
        if ( (IDispatch (m_ApptItems) = nil) or (m_ApptItems.Count = 0)) and (m_ActItems.Detail.Count = 0) then
           exit;

        // Vérification de la cohérence de synchro (si déjà fait), sauf en mode désynchro
        if m_iTypeSync <> ostDesynchro then
        begin
             // $$$ JP 03/11/06: Pgi doit avoir CE calendrier Outlook comme référence de synchro (si déjà synchronisé avec Outlook)
             strValue := '';
             with TOB.Create ('le calendrier outlook', nil, -1) do
             begin
                  LoadDetailFromSQL ('SELECT YX_LIBELLE FROM CHOIXEXT WHERE YX_TYPE="SYA" AND YX_CODE="OUT"');
                  if Detail.Count > 0 then
                      strValue := Detail [0].GetValue ('YX_LIBELLE')
                  else
                      strValue := '';
                  Free;
             end;
             if (strValue <> '') and (strValue <> m_strGuidOutlook) then
             begin
                  // Si pas CE calendrier Outlook, on peut le remplacer, mais avec double confirmation
                  if PgiAsk ('L''agenda Expert est déjà synchronisé avec un autre calendrier Outlook.'#10' Désirez-vous tout de même faire la synchronisation avec ce calendrier Outlook?') <> mrYes then
                     exit;
                  if PgiAsk ('Deuxième confirmation:'#10#10' Ce calendrier Outlook deviendra le calendrier de référence pour la synchronisation'#10' Confirmez-vous la synchronisation avec ce calendrier Outlook?') <> mrYes then
                     exit;

                  // Ok, on prend cette référence de calendrier outlook comme référence de synchro
                  // $$$ JP 12/04/07: non, il faut justement conserver ce nouvel identifiant Outlook pour la synchro qui s'annonce
                  // m_strGuidOutlook := strValue;
             end;

             // $$$ JP 03/11/06: Outlook doit avoir CET agenda Expert comme référence de synchro (si déjà synchronisé avec Cegid Expert)
             strValue  := '';
             try
                vSyncItem := m_ApptItems.Find ('[CEGIDBASEGUID] <> "" and [CEGIDBASEGUID] <> "' + m_strGuidBase + '"'); // $$$ JP 12/04/07: chercher s'il y en a au moins un différent //Find ('[CEGIDBASEGUID] <> ""');

                if IDispatch (vSyncItem) <> nil then
                begin
                     vUserProp := vSyncItem.UserProperties.Find ('CEGIDBASEGUID');
                     if IDispatch (vUserProp) <> nil then
                     begin
                          strValue := vUserProp.Value;
                          vUserProp := unAssigned;
                     end;
                     vSyncItem := unAssigned;
                end;
             except
                   strValue := '';
             end;
             if (strValue <> '') and (strValue <> m_strGuidBase) then
             begin
                  if PgiAsk ('Certains rendez-vous Outlook sont déjà synchronisés avec un autre agenda Expert.'#10' Désirez-vous tout de même faire une synchronisation avec ce calendrier Outlook?') <> mrYes then
                     exit;
                  if PgiAsk ('Deuxième confirmation:'#10#10' Certains rendez-vous Outlook sont déjà synchronisés avec un autre agenda Expert.'#10' Confirmez-vous la synchronisation avec ce calendrier Outlook?') <> mrYes then
                     exit;

                  // $$$ JP 12/04/07: non, il faut justement conserver ce nouvel identifiant Cegid pour la synchro qui s'annonce
                  //m_strGuidBase := strValue;
             end;

             // Nombre d'éléments à synchronisé: étape
             InitMove (m_ApptItems.Count + m_ActItems.Detail.Count, '');

             // 1ère passe: synchro à partir des rendez-vous d'Outlook (en marquant les activités qui seront synchronisées pour ne pas le refaire en phase 2)
             SyncAppt;

             // 2ème passe: recopie des activités non synchronisées (car nouvelles dans outlook), et mise à jour base pgi des activités synchronisées
             CopyAndSaveAct;

             // On enregistre l'id du calendrier Outlook utilisé pour la synchro, si y'a bien eu une synchro
             if m_iNbSync > 0 then
             begin
                  with TOB.Create ('CHOIXEXT', nil, -1) do
                  begin
                       InitValeurs;
                       PutValue ('YX_TYPE',    'SYA');
                       PutValue ('YX_CODE',    'OUT');
                       PutValue ('YX_LIBELLE', m_strGuidOutlook);

                       InsertOrUpdateDB;
                       Free;
                  end;
             end;
        end
        else
        begin
             // 1ère passe: on dé-synchronise les rdv outlook
             DesyncAppt;

             // 2ème passe: on dé-synchronise les activité Expert
             DesyncAct;
        end;
     finally
            FiniMove;
            SourisNormale;
            Result       := m_iNbSync;
            m_ApptItems  := unAssigned;
            vCalFolder   := unAssigned;
            vMapi        := unAssigned;
            vOutlook     := unAssigned;
            m_ActItems.Free;
     end;
end;

function TOutlookSync.DoImport:integer;
var
   i, iIndex  :integer;
   vOutlook   :variant;
   vExplorer  :variant;
   vApptItem  :variant;
   bRecurring :boolean;
   strEntry   :string;
   RecList    :TStringList;
begin
     // Par défaut, pas de rendez-vous importés
     Result := 0;

     // Si mode de synchro sur none, on fait rien
     if m_iTypeSync <> ostImport then
        exit;

     SourisSablier;
     try
        // Avant tout, il faut créer un type d'activité spécifique pour import Outlook (comme pour synchro pda)
        CreateOutlookEvtCode;

        // On construit la TOB destinée à recevoir les nouveaux éléments
        m_ActItems := TOB.Create ('les activités', nil, -1);

        // Enumération rdv sélectionnés pour copie dans l'agenda Expert
        vOutlook := CreateOleObject ('Outlook.Application');
        if IDispatch (vOutlook) <> nil then
        begin
             vExplorer := vOutlook.ActiveExplorer;
             if IDispatch (vExplorer) <> nil then
             begin
                  m_ApptItems := vExplorer.Selection;
                  if IDispatch (m_ApptItems) <> nil then
                  begin
                       if m_ApptItems.Count < 1 then
                          exit;

                       // Liste des rdv répétables déjà traité (pour éviter d'importer x fois le même)
                       RecList            := TStringList.Create;
                       RecList.Sorted     := TRUE;
                       RecList.Duplicates := dupIgnore;

                       // Copie des rdv dans une nouvelle activité
                       try
                          // Nombre d'éléments à synchronisé: étape
                          InitMove (m_ApptItems.Count, '');

                          for i := 1 to m_ApptItems.Count do
                          begin
                               // Element pas forcément un rdv
                               vApptItem := m_ApptItems.Item (i);
                               if IDispatch (vApptItem) <> nil then
                               begin
                                    // Alimentation de la TOB des activités à enregistrer
                                    if (vApptItem.Class = olAppointment) then
                                    begin
                                         bRecurring := vApptItem.IsRecurring;
                                         strEntry   := vApptItem.EntryId;
                                         if (bRecurring = FALSE) or (RecList.Find (strEntry, iIndex) = FALSE) then
                                            if (ApptToAct (vApptItem, nil) <> nil) and (bRecurring = TRUE) then
                                               RecList.Add (strEntry);
                                    end;
                               end;
                          end;

                          // Enregistrement des activités générées
                          CopyAndSaveAct;
                       finally
                              RecList.Free;
                       end;
                  end;
             end;
        end;
     finally
            FiniMove;
            SourisNormale;
            Result       := m_iNbSync;
            m_ApptItems  := unAssigned;
            vExplorer    := unAssigned;
            vOutlook     := unAssigned;
            m_ActItems.Free;
     end;
end;


end.
