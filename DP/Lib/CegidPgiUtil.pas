{***********UNITE*************************************************
Auteur  ...... : Marc Desgoutte
Créé le ...... : 17/12/2002
Modifié le ... : 25/08/2005
Description .. : Contient les utilitaires Bureau PGI utilisables dans le
Suite ........ : process serveur et en mode EAGLCLIENT
Suite ........ : Les fonctions sont faites pour le mode Web Access, mais
Suite ........ : pas encore complètes en mode 2 tiers (voir galoutil
Suite ........ : pour le contenu 2 tiers)
Suite ........ : POUR L'INSTANT, NE PAS INCLURE EN 2/3
Mots clefs ... : S5;OUTILS;
*****************************************************************}
unit CegidPgiUtil;

{$IFNDEF BUREAU}
  #### ne pas utiliser hors projets Bureau ####
{$ENDIF}

{$IFNDEF EAGLCLIENT}
{$IFNDEF EAGLSERVER}
   $$$ JP 27/12/05 - pour l'instant, ne pas compiler ce source en 2/3
{$ENDIF}
{$ENDIF}

interface

uses
   Utob, Classes, HenT1, MajTable, hctrls, stdctrls, uHttp, FileCtrl,
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF DBXPRESS}
{$ENDIF EAGLClIENT}
{$IFDEF VER150}
   Variants,
{$ENDIF}
   Paramsoc, uWA, uWAIni, uWAIniBase, uWAInfoIni;
Type
   cWACegidPgi = class(cWA)
   private

   public
      constructor Create; override;

      // $$$ JP 16/12/05 - Création dossier avec fiche personne et options de création
      function WACreateDossier (TOBAnn:TOB; TOBOpt:TOB):integer; //CreeBaseDossier(NoDisq: Integer; NoDossier, LibDossier: String) : Boolean;

      // $$$ JP 15/12/05 - avoir la liste des "disques" de données
      // $$$ JP 04/09/06 - HTStringList
      function GetDisks      (var Disks:HTStringList):boolean;

      // $$$ JP 27/12/05 - Autoclose
      function SetAutoClose  (strBase:string; bAutoClose:boolean):boolean;

      function WAPartagesOk  (TOBOpt:TOB=nil):boolean; //(NoDisq : Integer): Boolean;

      // $$$ JP 27/12/05 - Existe Base   //SL
      //function ExisteBase    (strBase:string):boolean;

      // $$$ JP 28/12/05 - Suppression base
      function DeleteBase    (strBase:string):boolean;

      // $$$ JP 19/05/06 - FQ 10616: collation identique ou non entre  deux bases?
      function CheckCollationBase (strBase:string; strBaseRef:string; strCollRef:string):boolean;

      // MD 16/02/07 - Fonctions de transport Web Access
      function WATransportCreerFichierCommun        (TOBOpt: TOB): Boolean;
      function WATransportCreerFichierDossier       (TOBOpt: TOB): Boolean;
      function WATransportCreerFichierStd           (TOBOpt: TOB): Boolean;
      function WATransportDeleteFile                (TOBOpt: TOB): Boolean;
      function WATransportRecupListeFichiers        (TOBOpt: TOB): TOB;
      function WATransportRecupListeDossiersDistants(TOBOpt: TOB): TOB;
      function WATransportExecuteRemoteQuery        (TOBOpt: TOB): Boolean;
      function WATransportRecupFichiersDistants     (TOBOpt: TOB): Boolean;
      function WATransportIntegreFichiers           (TOBOpt: TOB): Boolean;
      function WATransportSupprContenuRep           (TOBOpt: TOB): Boolean;
      {+GHA - 12/2007}
      function WATransportSupprimeDossierDistant    (TOBOpt: TOB): Boolean;
      {-GHA - 12/2007}

      //SL 11/07/07 - Fonctions de transfert Web Access
      function WATransfertCreerFichierDossier   (TOBOpt: TOB): Boolean;

      // function WATransportCreerPathPortable   (NomPortable:String): Boolean;

{$IFDEF EAGLSERVER}
      class function Action (UneAction: string; RequestTOB:TOB; var ResponseTOB:TOB):boolean;
{$ENDIF EAGLSERVER}
   end ;



//////////////// IMPLEMENTATION //////////////////
implementation

uses
   Windows, SysUtils, Forms, controls,
   galDossierCste, galSystem,
{$IFDEF EAGLSERVER}
   uWABDSQLServer,
   eISAPI, eSession,
   uLanceProcess,
   LicUtil, adodb, tntclasses, DB,
{$ENDIF}
   UtilTranspDossier,
   hmsgbox, HDebug , uWABD, hQry;

const
     jpnbKeywordsMSSQL = 28;
     jpkeyWordsMSSQL: array[1..jpnbKeywordsMSSQL] of string = (
    // SUBSTRING ok
    'TRIM(%)', 'LTRIM(RTRIM(%1))',
    // LTRIM ok
    // RTRIM ok
    // STR ok
    // LEFT ok
    // RIGHT ok
    // LEN ok
    'CAST(% AS DATE)', 'CAST(%1 AS DATETIME)',
    //'YEAR(%)',            'DATEPART(dy,%1)',
    'MONTH(%)', 'DATEPART(month,%1)',
    'CHR(', 'CHAR(',
    'WEEK(%)', 'DATEPART(wk,%1)',
    'DAY(%)', 'DATEPART(day,%1)',
    'HOUR(%)', 'DATEPART(hh,%1)',
    'YDAY(%)', 'DATEPART(dy,%1)',
    'WDAY(%)', 'DATEPART(dw,%1)',
    'NOW', 'GETDATE( )',
    'IIF(%,%,%)', 'CASE WHEN %1 THEN %2 ELSE %3 END',
    'SI(%,%,%)', 'CASE WHEN %1 THEN %2 ELSE %3 END',
    'PGIGUID', '(CAST(NEWID( ) AS VARCHAR(36)))',
    // ALR le 29/11/04
    //'TRUNC(%)', 'CONVERT(CHAR(10), %1, 103)'
    // ALR le 30/01/07 pbm de format de date
    'TRUNC(%)', 'cast(CONVERT(CHAR(10), %1, 101) as datetime)'
    );


{$IFDEF EAGLSERVER}
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD / MD
Créé le ...... : 23/01/2004
Modifié le ... : 23/01/2004
Description .. : Fonction aiguillage
Suite ........ :
Mots clefs ... :
*****************************************************************}
class function cWACegidPgi.Action (UneAction:string; RequestTOB:TOB; var ResponseTOB:TOB):boolean;
var
   st                       :string;
   T1                       :TOB;
   tt                       :HTStringList; // $$$ JP 04/09/06
   WACegidPgi               :cWACegidPgi;
   iResCreate               :integer;
   TOBOptions, TOBAnnuaire  :TOB;
   //NoDisq : Integer;
   //NoDossier, LibDossier : String;
BEGIN
     // Par défaut, résultat ok
     Result := TRUE;

     // 3 façons de gérer le retour d'une action (voir uWAIniBase : fonction Action)
     T1 := nil;
     tt := nil;
     st := '';

     // Instance d'un plugin pour lancer l'action demandée
     WACegidPgi := cWACegidPgi.Create;
     try
        UneAction := UpperCase (Trim (UneAction));
        ForceTrace('---> cWACegidPgi : start action ' + UneAction);

        //--- Récupération des options
        // TOB des options, toujours présente à priori pour toutes les fonctions
        // Le nom de la TOB est important car RequestTOB peut contenir plein de tob encapsulés
        TOBOptions := RecupereUneTob (RequestTOB, 'OPTIONS PLUGIN BUREAU');
        // *** Attention : le mécanisme de RecupereUneTob nous fournit les paramètres
        // *** en tob fille de TOBOptions, c'est à dire dans TOBOptions.Detail[0]
        if UneAction='CREATEDOSSIER' then
           TOBAnnuaire := RecupereUneTOB (RequestTOB, 'ANNUAIRE');

        try
          //--- Aiguillage sur la fonction demandée
           if UneAction = 'TESTREPONSE' then
                ForceTrace ('---> cWACegidPgi : Reponse OK !')
           else if UneAction = 'PARTAGESOK' then
                Result := WACegidPgi.WAPartagesOk (TOBOptions)
           else if UneAction = 'CREATEDOSSIER' then
           begin
                iResCreate := WACegidPgi.WACreateDossier (TOBAnnuaire, TOBOptions);
                ResponseTOB.AddChampSupValeur ('OPB_RESULT', iResCreate, CSTInteger);
                if iResCreate <> TDOSS_CREATIONOK then
                   Result := FALSE;
           end
           else if UneAction = 'GETDISKS' then
                Result := WACegidPgi.GetDisks (tt)
           else if UneAction = 'SETAUTOCLOSE' then
                Result := WACegidPgi.SetAutoClose (TOBOptions.Detail [0].GetString ('OPB_BASE'), TOBOptions.Detail [0].GetInteger ('OPB_AUTOCLOSE') <> 0)
           {else if UneAction = 'EXISTEBASE' then
           begin
                Result := WACegidPgi.ExisteBase (TOBOptions.Detail [0].GetString ('OPB_BASE'));
                ResponseTOB.AddChampSupValeur ('OPB_RESULT', Result, CSTBoolean);
           end }
           else if UneAction = 'DELETEBASE' then
                Result := WACegidPgi.DeleteBase (TOBOptions.Detail [0].GetString ('OPB_BASE'))
           else if UneAction = 'CHECKCOLLATIONBASE' then
           begin
                Result := WACegidPgi.CheckCollationBase (TOBOptions.Detail [0].GetString ('OPB_BASE'), TOBOptions.Detail [0].GetString ('OPB_BASEREF'), TOBOptions.Detail [0].GetString ('OPB_COLLREF'));
                ResponseTOB.AddChampSupValeur ('OPB_RESULT', Result, CSTBoolean);
           end
//           else if UneAction = 'CREERPATHPORTABLE' then
//                Result := WACegidPgi.WATransportCreerPathPortable( TOBOptions.Detail [0].GetString ('OPB_NOMPORTABLE'))
           else if UneAction = 'TRANSPORTCREERFICHIERCOMMUN' then
           begin
                Result := WACegidPgi.WATransportCreerFichierCommun (TOBOptions);
                ResponseTOB.AddChampSupValeur ('OPB_RESULT', Result, CSTBoolean);
           end
           else if UneAction = 'TRANSPORTCREERFICHIERDOSSIER' then
           begin
                Result := WACegidPgi.WATransportCreerFichierDossier (TOBOptions);
                ResponseTOB.AddChampSupValeur ('OPB_RESULT', Result, CSTBoolean);
           end
           else if UneAction = 'TRANSPORTCREERFICHIERSTD' then
           begin
                Result := WACegidPgi.WATransportCreerFichierStd (TOBOptions);
                ResponseTOB.AddChampSupValeur ('OPB_RESULT', Result, CSTBoolean);
           end
           else if UneAction = 'TRANSPORTDELETEFILE' then
           begin
                Result := WACegidPgi.WATransportDeleteFile (TOBOptions);
           end
           else if UneAction = 'TRANSPORTRECUPLISTEFICHIERS' then
           begin
                T1 := WACegidPgi.WATransportRecupListeFichiers (TOBOptions);
           end
           else if UneAction = 'TRANSPORTRECUPLISTEDOSSIERSDISTANTS' then
           begin
                T1 := WACegidPgi.WATransportRecupListeDossiersDistants (TOBOptions);
           end
           else if UneAction = 'TRANSPORTEXECUTEREMOTEQUERY' then
           begin
                Result := WACegidPgi.WATransportExecuteRemoteQuery (TOBOptions);
                ResponseTOB.AddChampSupValeur ('OPB_RESULT', Result, CSTBoolean);
           end
           else if UneAction = 'TRANSPORTRECUPFICHIERSDISTANTS' then
           begin
                Result := WACegidPgi.WATransportRecupFichiersDistants (TOBOptions);
                ResponseTOB.AddChampSupValeur ('OPB_RESULT', Result, CSTBoolean);
           end
           else if UneAction = 'TRANSPORTINTEGREFICHIERS' then
           begin
                Result := WACegidPgi.WATransportIntegreFichiers (TOBOptions);
                ResponseTOB.AddChampSupValeur ('OPB_RESULT', Result, CSTBoolean);
           end
           else if UneAction = 'TRANSPORTSUPPRCONTENUREP' then
           begin
                Result := WACegidPgi.WATransportSupprContenuRep (TOBOptions);
           end
           else if UneAction = 'TRANSFERTCREERFICHIERDOSSIER' then
           begin
                Result := WACegidPgi.WATransfertCreerFichierDossier (TOBOptions);
                ResponseTOB.AddChampSupValeur ('OPB_RESULT', Result, CSTBoolean);
           end
           {+GHA - 12/2007}
           else if UneAction = 'TRANSPORTSUPPRIMEDOSSIERDISTANT' then
           begin
                Result := WACegidPgi.WATransportSupprimeDossierDistant(TOBOptions);
                ResponseTOB.AddChampSupValeur ('OPB_RESULT', Result, CSTBoolean);
           end
           {-GHA - 12/2007}
           // $$$ JP 11/09/06: plus de création base modèle, car déjà présent sous forme d'archive .bak

           else
           begin
                Result := FALSE;
                WACegidPgi.ErrorMessage := 'Action inconnue';
           end;
        except
              on E: Exception do WACegidPgi.ErrorMessage := E.Message;
        end;

        // Astuce : pour récupérer les infos de la TOB Server du coté client en cas d'erreur,
        // on pourrait cloner WACegidPgi : ResponseTOB.Dupliquer(WACegidPgi,true,true) ;
        //  if ResponseTOB=nil then ResponseTOB:=WACegidPgi.clone() else ResponseTOB.Dupliquer(WACegidPgi,true,true) ;

        // Formate la réponse, notamment ajout ErrorMessage dans un chp 'ERROR', si nécessaire
        if WACegidPgi.ErrorMessage<>'' then ddWriteln('####### ErrorMessage = '+WACegidPgi.ErrorMessage);
        WACegidPgi.SetRetour (ResponseTOB, Result, T1, tt, st) ;
     finally
            // Les TOB détachées par RecupereUneTob n'ont plus de parent
            // donc c'est à nous de les purger
            FreeAndNil (TOBOptions);
            if UneAction='CREATEDOSSIER' then FreeAndNil(TOBAnnuaire);
            FreeAndNil (tt); //.free ;
            FreeAndNil (T1); //.free ;
            FreeAndNil (WACegidPgi); //.free ;
     end;
end;
{$ENDIF EAGLSERVER}

constructor cWACegidPgi.create;
begin
     inherited;

     // la classe du plug-in qui répond n'est pas DBSS !!
     dbss_dll := 'sCegidPgi' ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marc Desgoutte
Créé le ...... : 25/08/2005
Modifié le ... :   /  /
Description .. : Crèe la base dossier S5 dans le partage précisé
Suite ........ : (voir tableau V_PGI.Partages ou PartagesLoc)
Mots clefs ... :
*****************************************************************}
function cWACegidPgi.WACreateDossier (TOBAnn:TOB; TOBOpt:TOB):integer; //CreeBaseDossier(NoDisq: Integer; NoDossier, LibDossier: String) : Boolean;
// Les vérifs "mono-utilisateur" via Bloqueur('CreationDossier'...) sont faites en amont
// Le partage correspondant à NoDisq est vérifié au préalable par PartagesOk(...)
// La base modèle est vérifiée au préalable par BaseModeleOk
// #### CODIFICATION DES BASES A REVOIR EN HEBERGEMENT MUTUALISE (plusieurs cabinets)
var
   TOBAnnuaire, TOBOptions :TOB;
{$IFDEF EAGLCLIENT}
   ResultTob    :TOB;
{$ELSE}
   StrMsg       :String;
{$ENDIF}
     procedure CleanExit;
     begin
       FreeAndNil(TOBOptions);
       FreeAndNil(TOBAnnuaire);
       Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
     end;
begin
     Result := TDOSS_CANTCREEDOSSIER;
     if (TOBAnn = nil) or (TOBOpt = nil) then exit;

     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
     TOBAnnuaire := TOB.Create ('ANNUAIRE', nil, -1);
     TOBAnnuaire.Dupliquer(TOBAnn, True, True);
{$IFDEF EAGLCLIENT}
     // TOB des options à envoyer (en tob filles de Self)
     AjouteUneTob (Self, TOBAnnuaire);
     AjouteUneTob (Self, TOBOptions);

     // Demande au serveur
     try
        ResultTOB := Request ('sCegidPgi.CreateDossier', '', Self, '', ''); // $$$ 19/12/05 - ne pas utiliser Param pour l'action //cWACegidPgi', 'CreateDossier', Self, '', '') ;
        GetRetour (ResultTob);
        if ResultTob.FieldExists('OPB_RESULT') then
          Result := ResultTob.GetInteger ('OPB_RESULT')
        else
        begin
          Result := TDOSS_UNKNOWNERROR;
          if ErrorMessage='' then ErrorMessage := 'Erreur inconnue pendant la creation de dossier.';
        end;
        FreeAndNil(ResultTOB);
     finally
     end;
{$ELSE}
  Result := CreateDossier(StrMsg,TOBAnnuaire,TOBOptions);
  ErrorMessage := StrMsg;
{$ENDIF}
  CleanExit;
end;

// $$$ JP 15/12/05 - avoir la liste des "disques" de données
function cWACegidPgi.GetDisks (var Disks:HTStringList):boolean;
{$IFDEF EAGLCLIENT}
var
   ResultTob    :TOB;
   st           :string;
{$ELSE}
{$IFDEF EAGLSERVER}
var
   LaSession             :TISession;
   i                     :integer;
   strDisk               :string;
{$ENDIF}
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
     try
        ResultTOB := Request ('sCegidPgi.GetDisks', '', Self, '', ''); // $$$ 19/12/05 - ne pas utiliser Param pour l'action //cWACegidPgi', 'GetDisks', Self, '', '');
        Result := GetRetour (ResultTOB, nil, Disks, st);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
{$IFDEF EAGLSERVER}
     // Recensement des disques non vide (même ceux non accessibles, ce qui devrait être rare)
     LaSession   := LookupCurrentSession;
     Disks       := HTStringList.Create;
     ddwriteln ('   -> sCegidPgi.GetDisks / session=' + LaSession.SessionId);
     for i := 1 to High (LaSession.Partages) do
     begin
          strDisk := Trim (LaSession.Partages [i]);
          if (strDisk <> '') then
             if DirectoryExists (strDisk) = TRUE then
                 Disks.Add ('Disque ' + IntToStr (i))
             else
                 Disks.Add ('Disque ' + IntToStr (i) + ' (non accessible)');
     end;

     Result := TRUE;
{$ENDIF}
{$ENDIF}
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;

// $$$ JP 27/12/05 - Autoclose
function cWACegidPgi.SetAutoClose (strBase:string; bAutoClose:boolean):boolean;
{$IFDEF EAGLCLIENT}
var
   ResultTob   :TOB;
   TOBOptions  :TOB;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
     try
        TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
        with TOBOptions do
        begin
             AddChampSupValeur ('OPB_BASE',      strBase,    CSTString);
             AddChampSupValeur ('OPB_AUTOCLOSE', bAutoClose, CSTBoolean);
        end;
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);

        ResultTOB := Request ('sCegidPgi.SetAutoClose', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
{$IFDEF EAGLSERVER}
     with cWABDSQLSERVER.Create do
     begin
          Base         := strBase;
          Result       := AutocloseBase (bAutoClose);
          Free;
     end;
{$ENDIF}
{$ENDIF}
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;

function cWACegidPgi.WAPartagesOk (TOBOpt:TOB):boolean; //(NoDisq : Integer): Boolean;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   ResultTOB   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
     procedure CleanExit;
     begin
       FreeAndNil(TOBOptions);
       Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
     end;
begin
     Result := FALSE;
     if TOBOpt=Nil then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);

{$IFDEF EAGLCLIENT}
     // TOB des options à envoyer (en tob fille de Self)
     AjouteUneTob (Self, TOBOptions);

     // Demande au serveur
     ResultTOB := Request ('sCegidPgi.PartagesOk', '', Self, '', ''); // $$$ 19/12/05 - ne pas utiliser Param pour l'action //cWACegidPgi', 'PartagesOk', Self, '', '');
     Result    := GetRetour (ResultTOB) ;

     ResultTOB.free ;
{$ELSE}
     Result := PartagesOk(StrMsg, TobOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     CleanExit;
end;

// $$$ JP 27/12/05 - Existe Base
{function cWACegidPgi.ExisteBase (strBase:string):boolean;
{$IFDEF EAGLCLIENT}
{var
   ResultTob   :TOB;
   TOBOptions  :TOB;
{$ENDIF}
{begin
{$IFDEF EAGLCLIENT}
{     try
        TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
        with TOBOptions do
        begin
             AddChampSupValeur ('OPB_BASE', strBase, CSTString);
        end;
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);

        ResultTOB := Request ('sCegidPgi.ExisteBase', '', Self, '', '');
        Result    := ResultTOB.GetInteger ('OPB_RESULT') <> 0;
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
{{$IFDEF EAGLSERVER}
{     DDWriteLn ('   sCegidPgi.ExisteBase? ' + strBase);
     with cWABDSQLSERVER.Create do
     begin
          Base         := strBase;
          Result       := ExisteBase;
          Free;
     end;
{$ENDIF}
{{$ENDIF}
{    Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end; }

// $$$ JP 28/12/05 - Suppression base
function cWACegidPgi.DeleteBase (strBase:string):boolean;
{$IFDEF EAGLCLIENT}
var
   ResultTob   :TOB;
   TOBOptions  :TOB;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
     try
        TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
        with TOBOptions do
        begin
             AddChampSupValeur ('OPB_BASE', strBase, CSTString);
        end;
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);

        ResultTOB := Request ('sCegidPgi.DeleteBase', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
{$IFDEF EAGLSERVER}
     with cWABDSQLSERVER.Create do
     begin
          Base         := strBase;
          Result       := SupprimerBase;
          Free;
     end;
{$ENDIF}
{$ENDIF}
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;

// $$$ JP 28/12/05 - Vérification de la collation base
function cWACegidPgi.CheckCollationBase (strBase:string; strBaseRef:string; strCollRef:string):boolean;
{$IFDEF EAGLCLIENT}
var
   ResultTob   :TOB;
   TOBOptions  :TOB;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
     try
        TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
        with TOBOptions do
        begin
             AddChampSupValeur ('OPB_BASE',    strBase,    CSTString);
             AddChampSupValeur ('OPB_BASEREF', strBaseRef, CSTString);
             AddChampSupValeur ('OPB_COLLREF', strCollRef, CSTString);
        end;
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);
        ResultTOB := Request ('sCegidPgi.CheckCollationBase', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
     // Vérification sur la collation spécifiée, ou bien par rapport à la collation de la base de référence
     if strCollRef <> '' then
         Result := GetCollationBase (strBase) = strCollRef
     else
         Result := GetCollationBase (strBase) = GetCollationBase (strBaseRef);
{$ENDIF}
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;

{
function cWACegidPgi.WATransportCreerPathPortable(TOBopt:TOB): Boolean;
var
  StrMsg: String;
begin
  Result := TransportCreerPathPortable(StrMsg,Nomportable);
  ErrorMessage := StrMsg;
end;
}

function cWACegidPgi.WATransportCreerFichierCommun (TOBOpt:TOB): boolean;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   ResultTob   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
begin
     Result := False;
     if (TOBOpt = nil) then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
{$IFDEF EAGLCLIENT}
     try
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);
        ResultTOB := Request ('sCegidPgi.TransportCreerFichierCommun', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
     Result := TransportCreerFichierCommun(StrMsg, TobOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     FreeAndNil(TOBOptions);
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;


function cWACegidPgi.WATransportCreerFichierDossier (TOBOpt:TOB): boolean;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   ResultTob   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
begin
     Result := False;
     if (TOBOpt = nil) then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
{$IFDEF EAGLCLIENT}
     try
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);
        ResultTOB := Request ('sCegidPgi.TransportCreerFichierDossier', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
     Result := TransportCreerFichierDossier(StrMsg, TobOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     FreeAndNil(TOBOptions);
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;


function cWACegidPgi.WATransportCreerFichierStd (TOBOpt:TOB): boolean;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   ResultTob   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
begin
     Result := False;
     if (TOBOpt = nil) then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
{$IFDEF EAGLCLIENT}
     try
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);
        ResultTOB := Request ('sCegidPgi.TransportCreerFichierStd', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
     Result := TransportCreerFichierStd(StrMsg, TobOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     FreeAndNil(TOBOptions);
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;

function cWACegidPgi.WATransportDeleteFile (TOBOpt: TOB): Boolean;
var
  TOBOptions:       TOB;
  StrMsg:           String;
begin
  {$IFDEF EAGLSERVER}
  ddWriteLn('Lancement WATransportDeleteFile');
  Result := False;
  if (TOBOpt = nil) then
  begin
    ddWriteLn('TOBOpt nil');
    exit;
  end;
  TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
  TOBOptions.Dupliquer(TOBOpt, True, True);
  ddWriteLn('Lancement TransportDeleteFile');
  Result := TransportDeleteFile(StrMsg, TOBOptions);
  ErrorMessage := StrMsg;
  ddWriteLn('ErrorMessage :'+ErrorMessage);
  FreeAndNil(TOBOptions);
  Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
  {$ENDIF}
end;


function cWACegidPgi.WATransportRecupListeFichiers (TOBOpt:TOB): TOB;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   Retour      :TOB;
   ResultTob   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
begin
     Result := Nil;
     if (TOBOpt = nil) then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
{$IFDEF EAGLCLIENT}
     try
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);

        ResultTOB := Request ('sCegidPgi.TransportRecupListeFichiers', '', Self, '', '');
        Retour := TOB.Create('LISTE FICHIERS', Nil, -1); // obligé pour GetRetour (xx, T1) : T1 doit exister
        GetRetour (ResultTOB, Retour);
        Result := TOB.Create('LISTE FICHIERS', Nil, -1);
        Result.Dupliquer(Retour, True, True);
        FreeAndNil(ResultTOB);
        FreeAndNil(Retour);
     finally
     end;
{$ELSE}
     Result := TransportRecupListeFichiers(StrMsg, TOBOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     FreeAndNil(TOBOptions);
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;

function cWACegidPgi.WATransportRecupListeDossiersDistants (TOBOpt:TOB): TOB;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   Retour      :TOB;
   ResultTob   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
begin
     Result := Nil;
     if (TOBOpt = nil) then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
{$IFDEF EAGLCLIENT}
     try
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);

        ResultTOB := Request ('sCegidPgi.TransportRecupListeDossiersDistants', '', Self, '', '');
        Retour := TOB.Create('LISTE DOSSIERS', Nil, -1); // obligé pour GetRetour (xx, T1) : T1 doit exister
        GetRetour (ResultTOB, Retour);
        Result := TOB.Create('LISTE DOSSIERS', Nil, -1);
        Result.Dupliquer(Retour, True, True);
        FreeAndNil(ResultTOB);
        FreeAndNil(Retour);
     finally
     end;
{$ELSE}
     Result := TransportRecupListeDossiersDistants(StrMsg, TOBOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     FreeAndNil(TOBOptions);
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;

function cWACegidPgi.WATransportExecuteRemoteQuery (TOBOpt: TOB): Boolean;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   ResultTob   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
begin
     Result := False;
     if (TOBOpt = nil) then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
{$IFDEF EAGLCLIENT}
     try
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);

        ResultTOB := Request ('sCegidPgi.TransportExecuteRemoteQuery', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
     Result := TransportExecuteRemoteQuery(StrMsg, TOBOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     FreeAndNil(TOBOptions);
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;

function  cWACegidPgi.WATransportRecupFichiersDistants(TOBOpt: TOB): Boolean;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   ResultTob   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
begin
     Result := False;
     if (TOBOpt = nil) then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
{$IFDEF EAGLCLIENT}
     try
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);

        ResultTOB := Request ('sCegidPgi.TransportRecupFichiersDistants', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
     Result := TransportRecupFichiersDistants(StrMsg, TOBOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     FreeAndNil(TOBOptions);
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;

function cWACegidPgi.WATransportIntegreFichiers (TOBOpt: TOB): Boolean;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   ResultTob   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
begin
     Result := False;
     if (TOBOpt = nil) then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
{$IFDEF EAGLCLIENT}
     try
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);
        ResultTOB := Request ('sCegidPgi.TransportIntegreFichiers', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
     Result := TransportIntegreFichiers(StrMsg, TobOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     FreeAndNil(TOBOptions);
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;

function cWACegidPgi.WATransportSupprContenuRep (TOBOpt: TOB): Boolean;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   ResultTob   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
begin
     Result := False;
     if (TOBOpt = nil) then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
{$IFDEF EAGLCLIENT}
     try
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);
        ResultTOB := Request ('sCegidPgi.TransportSupprContenuRep', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
     Result := TransportSupprContenuRep(StrMsg, TobOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     FreeAndNil(TOBOptions);
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;

function cWACegidPgi.WATransfertCreerFichierDossier (TOBOpt:TOB): boolean;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   ResultTob   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
begin
     Result := False;
     if (TOBOpt = nil) then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
{$IFDEF EAGLCLIENT}
     try
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);
        ResultTOB := Request ('sCegidPgi.TransfertCreerFichierDossier', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
     Result := TransfertCreerFichierDossier(StrMsg, TobOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     FreeAndNil(TOBOptions);
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marc Desgoutte
Créé le ...... : 25/08/2005
Modifié le ... :   /  /
Description .. : Lecture d'un paramsoc dans la base commune
Mots clefs ... : si elle existe (sinon dans la base en cours)
*****************************************************************}
{function  cWACegidPgi.GetParamSocFromDP(Nom: String): Variant;
var Q : TQuery;
    St, Data : String;
    zz : Char;
begin
  Result := #0;
  Q := OpenSQL('SELECT * FROM ##DP##.PARAMSOC WHERE SOC_NOM="'+Nom+'"', True);
  if Not Q.Eof then
     BEGIN
     St := Q.FindField('SOC_DESIGN').AsString ;
     St := ReadTokenSt(St) ; if St<>'' then zz:=St[1] else zz:=' ' ;
     Data := Q.FindField('SOC_DATA').AsString ;
     Result := ParamSocDataToVariant(Nom,Data,zz) ;
     END;
  Ferme(Q);
end;}

{+GHA - 12/2007}
function cWACegidPgi.WATransportSupprimeDossierDistant(TOBOpt: TOB): Boolean;
var
   TOBOptions  :TOB;
{$IFDEF EAGLCLIENT}
   ResultTob   :TOB;
{$ELSE}
   StrMsg      :String;
{$ENDIF}
begin
     Result := False;
     if (TOBOpt = nil) then exit;
     TOBOptions := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
     TOBOptions.Dupliquer(TOBOpt, True, True);
{$IFDEF EAGLCLIENT}
     try
        // TOB des options à envoyer (en tob fille de Self)
        AjouteUneTob (Self, TOBOptions);
        ResultTOB := Request ('sCegidPgi.TransportSupprimeDossierDistant', '', Self, '', '');
        Result    := GetRetour (ResultTOB);
        FreeAndNil (ResultTOB);
     finally
     end;
{$ELSE}
     Result := TransportSupprimeDossierDistant(StrMsg, TobOptions);
     ErrorMessage := StrMsg;
{$ENDIF}
     FreeAndNil(TOBOptions);
     Self.ClearDetail; // sinon perte de param lors des AjouteUneTob successifs, car VH_DP.ePlugIn est un objet persistant
end;
{-GHA - 12/2007}
end.

