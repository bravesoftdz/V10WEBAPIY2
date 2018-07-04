unit UtilDossier;

interface

uses HCtrls, UTob, Forms, Controls, Windows,
{$IFDEF VER150}
     Variants,
{$ENDIF}
{$IFDEF EAGLCLIENT}
     uHttp,
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     HDTLinks, SysUtils, AGLUtilOle, HMsgBox, {AGLUtilXLS,} UtilXLS,
{$IFDEF BUREAU}
     galOutil, CBPPath,
{$ENDIF}
     HEnt1, PGIAppli,
{$IFDEF MAJPCL}
     UtilPgi,
{$ENDIF}
     HStatus,
     Paramsoc;

function  GetNomDossier(nodoss: String) : String;
function  NouvelleCle(NomChamp, NomTable : String) : Integer;
function  LanceAppliMaquette( sNomDoc_p, sAppli_p : string; bIsNew_p : boolean = FALSE ) : boolean;

{$IFDEF BUREAU}
function  ConnecteDossier000STD: Boolean;
function  CreeDossier000STD: Boolean;

{$IFNDEF EAGLCLIENT}
procedure LancePGIMajVer(VersionDossier : Integer=0);
procedure UploadBackupDBMODELE;
{$IFDEF MAJPCL}
function  OptimiseDossier(nodoss : String) : Boolean;
procedure OptimiserDossierPCL (BaseDossier : String);
procedure OptimiserTable (BaseDossier : String; NomTable : String);
procedure DroperTable (BaseDossier : String; NomTable : String; bPhysique : boolean );
function  DonnerPrefixe (BaseDossier : String; NomTable : String): String;
{$ENDIF MAJPCL}
{$ENDIF EAGLCLIENT}

function  IsTableExiste (BaseDossier : String; NomTable : String): Boolean;
{$ENDIF BUREAU}

{$IFDEF EAGLCLIENT}
function CheckInstallAppli(Lappli: TPGIAppli): Boolean;
{$ENDIF}

const
  cstAppliWord_g : string = '.DOC';
  cstAppliExcel_g : string = '.XLS';
  cstAppliNotePad_g : string = '.TXT';
  cstAppliAutre_g : string = '.XXX';


//////////// IMPLEMENTATION /////////////
implementation

{$IFDEF BUREAU}
uses
     galSystem, UDossierSelect, EntDp, galPatience;
{$ENDIF}

function GetNomDossier(nodoss: String) : String;
var
   Q: TQuery;
begin
     Result := '';
     if nodoss='' then exit;

     // $$$ JP 26/04/06 - pb codeper, remplacé désormais par guidper
     try
        Q := OpenSQL('SELECT ANN_NOM1 FROM ANNUAIRE,DOSSIER WHERE ANN_GUIDPER=DOS_GUIDPER AND DOS_NODOSSIER="'+ nodoss + '"', True); // $$$ JP 26/04/06 - ANN_CODEPER=DOS_CODEPER'
        if not Q.Eof then
           Result := Q.FindField('ANN_NOM1').AsString;
        Ferme (Q);
     except
           Result := '???';
     end;
end;


function NouvelleCle(NomChamp, NomTable : String) : Integer;
var Q : TQuery;
begin
  Result := 1;
  Q := OpenSQL('SELECT MAX('+NomChamp+') FROM '+NomTable,TRUE) ;
  if (Not Q.Eof) and Not VarIsNull(Q.Fields[0].AsVariant) then
    Result := Q.Fields[0].AsInteger + 1;
  Ferme(Q) ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function LanceAppliMaquette( sNomDoc_p, sAppli_p : string; bIsNew_p : boolean = FALSE ) : boolean;
var
   MonExcel : OleVariant;
begin
   result := false;
   if not FileExists(sNomDoc_p) then
   begin
      PGIInfo('Le document n''existe pas :'#10#13 + sNomDoc_p, 'Lancement d''application');
      exit;
   end;
   sAppli_p := UpperCase(sAppli_p);
   if (sAppli_p = cstAppliWord_g) or (sAppli_p = cstAppliNotePad_g) then
   begin
      OpenDoc( sNomDoc_p, wdWindowStateMaximize, wdPrintView);
      result := true;
   end
   else if sAppli_p = cstAppliExcel_g then
   begin
//      OpenDocXLS( sNomDoc_p, wdWindowStateMaximize, wdPrintView);
      OpenExcel (FALSE, MonExcel, bIsNew_p);
      if VarIsEmpty (MonExcel) = TRUE then
         PgiInfo ('Connexion à Microsoft Excel impossible', 'Lancement d''application')
      else
      begin
         try
            MonExcel.Visible := TRUE;
            MonExcel.WorkBooks.Open (sNomDoc_p, FALSE, FALSE);
            result := true;
         finally
            MonExcel := unAssigned;
         end;
      end;
   end
   else
      PGIInfo('Application non supportée'#10#13 + sAppli_p, 'Lancement d''application');
end;

{$IFDEF BUREAU}
function ConnecteDossier000STD: Boolean;
var Q : TQuery;
    VStd : Integer;
begin
  Result := False;

  // Existence de la base des standards
  if not DBExists ('DB000STD') then
     if Not CreeDossier000STD then
       exit; // SORTIE

  // vérifie la version
  VStd := 0;
  Q := OpenSQL('SELECT SO_VERSIONBASE FROM DB000STD.dbo.SOCIETE', True);
  if Not Q.Eof then VStd := Q.Fields[0].Value;
  Ferme(Q);

  // connecte le dossier
  VH_Doss.NoDossier := '000STD';
  if VH_Doss.NoDossier<>'000STD' then exit;

  // connexion ok, vérif version, et maj si besoin
  Result := True;

{$IFDEF EAGLCLIENT}
  // #### tant qu'on ne fait pas de majstruct en Web Access
  if V_PGI.NumVersionSoc<>VStd then
    PGIInfo('Attention : la base des standards n''est pas à jour (v'+IntToStr(VStd)+' au lieu de '+IntToStr(V_PGI.NumVersionSoc)+')#13#10'
           +' Contactez un administrateur pour la mettre à jour.');
{$ELSE}
  // MD 20/02/04 - ne plus détruire, mais mettre à jour la base
  if V_PGI.NumVersionSoc<>VStd then LancePGIMajVer;
{$ENDIF EAGLCLIENT}

  // FQ 11405 - Traitement spécifique des menus déversés
  // #### attention longueur du champ MN_ACCESGRP
  ExecuteSQL('UPDATE DB000STD.dbo.MENU SET MN_ACCESGRP="0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000", '
                                        + 'MN_VERSIONDEV="-" WHERE MN_VERSIONDEV="X"');
end;


function CreeDossier000STD: Boolean;
var
    Guidper, NoDossier : String;
    TOBAnnuaire, TOBOptions : TOB;
    F : TFPatience ;
    Q : TQuery;
{$IFDEF EAGLCLIENT}
{$ELSE}
    dbtool : TCreeDoss;
{$ENDIF}
    iTryBloque : Integer;
begin
    Result := False;

{$IFDEF EAGLCLIENT}
    // Si pas d'interface sur le plugin cwas, impossible de faire quoi que ce soit
    if VH_DP.ePlugin = nil then
      exit;
{$ENDIF}

    // On tente de rentrer dans le traitement bloquant "CreationDossier"
    iTryBloque := 0;
    while EstBloque ('CreationDossier', FALSE) = TRUE do
    begin
         if iTryBloque > 9 then
         begin
              if PgiAsk ('Un dossier est en cours de création'#10' Désirez-vous attendre la fin de cette création pour enchainer la votre ?') <> mrYes then
              begin
                   Result := False;
                   exit;
              end
              else
                  iTryBloque := 0;
         end
         else
         begin
              Inc (iTryBloque);
              Delay (3000);
              Application.ProcessMessages;
         end;
    end;
    // Bloque la création de dossier pour les autres
    Bloqueur ('CreationDossier', TRUE);

    // Chargement de la fiche personne liée au futur dossier
    F := FenetrePatience('Création de la base des standards DB000STD');
    // MD 10/09/07 - au lieu de systématiquement recréer la fiche annuaire "STANDARDS",
    // on essaye de récupérer le guidper si il en existait un...
    Q := OpenSQL('SELECT ANN_GUIDPER FROM DOSSIER, ANNUAIRE WHERE ANN_GUIDPER=DOS_GUIDPER AND DOS_NODOSSIER="000STD"', True);
    if Not Q.Eof then
    begin
      GuidPer := Q.FindField('ANN_GUIDPER').AsString;
      Ferme(Q);
    end
    else
    begin
      Ferme(Q);
      Guidper := ForcePersonneExiste('000STD', 'STANDARDS');
    end;
    TOBAnnuaire := TOB.Create('ANNUAIRE',Nil,-1) ;
    TOBAnnuaire.SelectDB('"'+Guidper+'"',Nil,TRUE) ;
    NoDossier := '000STD';

{$IFDEF EAGLCLIENT}
    // Mise en place des options de création dossier
    TOBOptions  := TOB.Create ('OPTIONS PLUGIN BUREAU', nil, -1);
    with TOBOptions do
    begin
         AddChampSupValeur ('OPB_NODISK',            1,             CSTInteger);
         AddChampSupValeur ('OPB_NODOSSIER',         '000STD',      CSTString);
         AddChampSupValeur ('OPB_PWD',               '',            CSTString);
         AddChampSupValeur ('OPB_GROUPE',            '',            CSTString);
         AddChampSupValeur ('OPB_BASEPROD',          True,          CSTBoolean);
         AddChampSupValeur ('OPB_REPRISEBASE',       False,         CSTBoolean);
         AddChampSupValeur ('OPB_ONEXISTINGDOSSIER', True,          CSTBoolean);
         AddChampSupValeur ('OPB_INITETABL',         False,         CSTBoolean);
    end;
    try
       // Invocation du serveur cwas pour création dossier
       Result := (VH_DP.ePlugin.WACreateDossier (TOBAnnuaire, TOBOptions) = 0);
    finally
           // Terminares, on ne bloque plus le traitement "CreationDossier"
           Bloqueur ('CreationDossier', FALSE);
    end;
    TOBOptions.Free;
{$ELSE}
    try
       dbtool := TCreeDoss.Create;
       // FQ 11789 => OnExistingDossier à True car on ne supprime jamais l'enregistrement 000STD dans la table DOSSIER
       Result := (dbtool.CreeDossier(TOBAnnuaire, 1, NoDossier, F.lAide, F.lCreation, Nil, False, True, '', False, Nil, '', False, True)=0);
    finally
           // Terminares, on ne bloque plus le traitement "CreationDossier"
           Bloqueur ('CreationDossier', FALSE);
           dbtool.Free;
    end;
{$ENDIF}
    // purges
    if F<>Nil then begin F.Close; F.Free; end;
    TOBAnnuaire.Free;
    if Not Result then
       PGIInfo('*** Impossible de créer la base des standards (DB000STD)', TitreHalley);
end;

{$IFNDEF EAGLCLIENT}
{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 02/09/2005
Modifié le ... :   /  /
Description .. : Lance la mise à jour de structure sur le dossier en cours
Suite ........ : (dossier identifié par l'objet VH_Doss de UDossierSelect)
Suite ........ :
Suite ........ : Traite l'aiguillage entre mise à jour complète et mise à jour
Suite ........ : allégée "DossierPCL".
Mots clefs ... :
*****************************************************************}
procedure LancePGIMajVer(VersionDossier : Integer=0);
var
    LigneCmd, ScriptSql, ScriptSqlLoc, ScriptToUpload, DossierPcl, ChemLog, ChemLogLoc : String;
    appli : TPGIAppli;
    F : TFPatience;
    // Q     : TQuery;
    // oldverrou, olduser : String;
begin
  if (V_Applis=Nil) or (V_Applis.Applis.Count=0) then exit;

  // On peut la lancer sur un dossier transporté (dos_verrou='PAR')
  // $$$ JP 12/07/06: ne plus faire ceci, car PgiMajVer s'en occupe
  {oldverrou := 'ENL';
  olduser := V_PGI.User;
  Q := OpenSQL('SELECT DOS_VERROU, DOS_UTILISATEUR FROM DOSSIER WHERE DOS_NODOSSIER="'+VH_Doss.NoDossier+'"', True);
  if Not Q.Eof then
    begin
    oldverrou := Q.FindField('DOS_VERROU').AsString;
    olduser := Q.FindField('DOS_UTILISATEUR').AsString;
    end;
  Ferme(Q);
  // Précise qu'on est en cours de mise à jour, et par qui
  ExecuteSQL('UPDATE DOSSIER SET DOS_VERROU="MAJ", DOS_UTILISATEUR="'+V_PGI.User+'"'
   +' WHERE DOS_NODOSSIER="'+VH_Doss.NoDossier+'"');
  }

  DossierPcl := RetourneDossierPcl;
  // ScriptSql := TCbpPath.GetCegidDistriStd+'\BUREAU\MajVer_SQL_750_'+IntToStr(V_PGI.NumVersionSoc)+'.sql';
  // 15/06/07 - Désormais : script côté serveur
  ScriptSql := V_PGI.DatPath+'\SOCREF\MajVer_SQL_750_'+IntToStr(V_PGI.NumVersionSoc)+'.sql';
  // avec son chemin local vu du serveur
  ScriptSqlLoc := GetParamSocSecur('SO_ENVPATHDATLOC', V_PGI.DatPath)+'\SOCREF\MajVer_SQL_750_'+IntToStr(V_PGI.NumVersionSoc)+'.sql';

  // #### Moment arbitraire pour uploader le backup de la base modèle pour les créations de dossier en Web Access
  UploadBackupDBMODELE;

  // Le script ne pouvait pas être déployé sur le serveur dans PGI01\DAT\SOCREF
  // donc on fait un upload depuis le répertoire d'install de la socref (local)
  if Not FileExists(ScriptSQL) then
  begin
       ScriptToUpload := TCbpPath.GetCegidData+'\MajVer_SQL_750_'+IntToStr(V_PGI.NumVersionSoc)+'.sql';
       if FileExists(ScriptToUpload) then
          try
             ForceDirectories(V_PGI.DatPath+'\SOCREF');
             CopyFile(PChar(ScriptToUpload), PChar(ScriptSQL), False);
          except
             on E:Exception do if V_PGI.SAV then PGIInfo(E.Message, 'Téléchargement du script');
          end;
  end;

  // La maj par Script Transac SQL ne s'applique que pour des dossiers PCL 750
  if (VH_DP.VersionMsSql>'SQL 7') and (DossierPcl<>'') and FileExists(ScriptSQL) and (VersionDossier=750) then
  begin
      // ChemLog := TCbpPath.GetCegidUserTempPath+'BUREAU';
      // ForceDirectories(ChemLog);
      // ChemLog := ChemLog + '\MajVer_SQL_750_'+IntToStr(V_PGI.NumVersionSoc)+'_Dossier_'+VH_Doss.NoDossier+'.log';
      ChemLog := V_PGI.DatPath + '\SOCREF\MajVer_SQL_750_'+IntToStr(V_PGI.NumVersionSoc)+'_Dossier_'+VH_Doss.NoDossier+'.log';
      ChemLogLoc := GetParamSocSecur('SO_ENVPATHDATLOC', V_PGI.DatPath) + '\SOCREF\MajVer_SQL_750_'+IntToStr(V_PGI.NumVersionSoc)+'_Dossier_'+VH_Doss.NoDossier+'.log';
      DeleteFile(ChemLog);
      LigneCmd := // 'cmd /c echo Ne pas fermer cette fenêtre : mise à jour en cours...'
                  //  +' & '+ // => concaténation en ligne de commande
                     'osql -S'+VH_DP.LeServeurSql
                     +' -d'+VH_Doss.DbSocName
                     +' -U'+VH_Doss.Soc.User
                     +' -P'+VH_Doss.Soc.Password
                     +' -i"'+ScriptSqlLoc+'"'
                     +' -o"'+ChemLogLoc+'"'
                     +' -n';
      // Lancement de la mise à jour
      // FileExecAndWait(LigneCmd);
      // 15/06/07 - Exécution directe sur le serveur
      try
         F := FenetrePatience('Mise à jour en cours, veuillez patientez...');
         F.lCreation.Caption := 'Exécution du script MajVer_SQL_750_'+IntToStr(V_PGI.NumVersionSoc)+'.sql';
         F.lAide.Caption     := 'sur le dossier '+VH_Doss.NoDossier+' ('+VH_Doss.LibDossier+')';
         F.Invalidate;
         Application.ProcessMessages;
         ExecCmdShell(LigneCmd);
      Finally
             if F<>Nil then begin F.Close; F.Free; end;
      end;
      if V_PGI.SAV then FileExecOrTop('notepad.exe '+ChemLog, False, True, False, '');
  end
  // Utilisation standard du PgiMajVer
  else
  begin
      appli := V_Applis.Lappli('PGIMAJVER.EXE');
      // on n'utilise pas ActiveAppli car on veut passer des param. à la ligne de commande
      // if appli<>Nil then V_Applis.ActiveAppli(appli);

      LigneCmd := ExtractFilePath(Application.ExeName) + 'PGIMAJVER.EXE';
      LigneCmd := LigneCmd + ParamLigneCommande (appli, VH_Doss.NoDossier, V_PGI.CurrentAlias, V_PGI.DbName, VH_Doss.GuidPer, VH_Doss.CodeSoc);
      LigneCmd := LigneCmd + ' /MAJSTRUCTAUTO=TRUE';
      // pour mise à jour en mode allégé "DOSSIERPCL"
      LigneCmd := LigneCmd + DossierPcl;

      LigneCmd := LigneCmd + ' /SHRINK=TRUE';

      // Lancement de la mise à jour
      FileExecAndWait(LigneCmd);
  end;

  // Rétablissement (important pour le transport en nomade)
  // $$$ JP 12/07/06: ne plus faire ceci, car PgiMajVer s'en occupe
  //ExecuteSQL('UPDATE DOSSIER SET DOS_VERROU="'+oldverrou+'", DOS_UTILISATEUR="'+olduser+'" WHERE DOS_NODOSSIER="'+VH_Doss.NoDossier+'"');

 {// MD 15/05/06 : le pgimajver fait lui même l'optimisation avant la maj
  // et à la fin de la maj, le MCD logique reste cohérent par rapport au physique
  // donc la proc ci-dessous est inutile :

  // Sur un dossier parti, il faut obligatoirement faire l'optimisation,
  // sinon il sera impossible de réintégrer
  if oldverrou='PAR' then // #### and "JeSuisEnLocal" ?
    OptimiseDossier(VH_Doss.NoDossier);
  }
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : MD
Créé le ...... : 26/06/2007
Modifié le ... :   /  /    
Description .. : Procédure temporaire
Suite ........ : pour copie fichier .bak de la base modèle
Suite ........ : (inclus dans le kit client socref)
Suite ........ : vers le répertoire des standards cabinet
Suite ........ : (sur le serveur sql)
Mots clefs ... :
*****************************************************************}
procedure  UploadBackupDBMODELE;
var
    FichierBak, FichierBakToUpload : String;
begin
  FichierBak := V_PGI.DatPath+'\SOCREF\S5_DBMODELE_'+IntToStr(V_PGI.NumVersionSoc)+'.bak';

  // Le script ne pouvait pas être déployé sur le serveur dans PGI01\DAT\SOCREF
  // donc on fait un upload depuis le répertoire d'install de la socref (local)
  if Not FileExists(FichierBak) then
  begin
       FichierBakToUpload := TCbpPath.GetCegidData+'\S5_DBMODELE_'+IntToStr(V_PGI.NumVersionSoc)+'.bak';
       if FileExists(FichierBakToUpload) then
          try
             ForceDirectories(V_PGI.DatPath+'\SOCREF');
             CopyFile(PChar(FichierBakToUpload), PChar(FichierBak), False);
          except
             on E:Exception do if V_PGI.SAV then PGIInfo(E.Message, 'Téléchargement du backup');
          end;
  end;



end;

{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 01/02/2006
Modifié le ... : 01/02/2006
Description .. : Optimise un dossier allégé "PCL"
Suite ........ : (=suppression physique des objets)
Suite ........ : Renvoit True si l'optimisation est ok
Suite ........ : (ou déjà faite)
Mots clefs ... :
*****************************************************************}
{$IFDEF MAJPCL}
function OptimiseDossier(nodoss : String) : Boolean;
begin
  Result := True;
  // La réparation plantera sur une structure allégée non optimisée
  if IsDossierMajPcl('DB'+nodoss) then
    begin
    // donc il faut d'abord optimiser la base (purge des objets)
    if Not IsDossierOptimise('DB'+nodoss) then
      begin
      ChgStatus('Optimisation de la structure du dossier '+nodoss+'...');

      //--- Optimisation du dossier en paramétre
      OptimiserDossierPCL ('DB'+nodoss);

      // Compression de la base
      ShrinkDatabaseMsSql('DB'+nodoss);
      DefStatus;

      Result := IsDossierOptimise('DB'+nodoss);
      end;
    end;
end;

//--------------------------------
//--- Nom : OptimiserDossierPCL
//--------------------------------
procedure OptimiserDossierPCL (BaseDossier : String);
var RSql: TQuery;
begin
 try
   SourisSablier;
   //--- Initialisation des variables
   InitPGIpourDossierPCL;

   //--- Suppression des Vues
   ExecuteSQL('DELETE FROM '+BaseDossier+'.DBO.DEVUES where DV_DOMAINE in ('+V_PGI.PCLDomainesExclus+')');

   //--- Suppresion des tables : Sauf table affaire
   RSql:=OpenSQL('SELECT DT_NOMTABLE FROM '+BaseDossier+'.DBO.DETABLES where DT_DOMAINE in ('+V_PGI.PCLDomainesExclus+') AND DT_NOMTABLE<>"AFFAIRE"',True,-1,'',True) ;
   while (not Rsql.eof) do
    begin
     OptimiserTable (BaseDossier,RSql.Fields[0].AsString);
     RSql.Next;
    end;
   Ferme(RSql);

   //--- Suppression des tables en dehors du domaine PCL et supprimées depuis la 663
   OptimiserTable (BaseDossier,'QGRPOFCAR');
   OptimiserTable (BaseDossier,'QPROFILUSERGRP');
   OptimiserTable (BaseDossier,'WDEPOTAPP');
   OptimiserTable (BaseDossier,'WORDREAUTO');

   //--- Suppression des menus
   executeSQl('Delete from '+BaseDossier+'.DBO.menu where mn_1=0 and mn_2 not in ('+V_PGI.PCLModulesInclus+')');
   executeSQl('Delete from '+BaseDossier+'.DBO.menu where mn_1<>0 and mn_1 not in ('+V_PGI.PCLModulesInclus+')');
   executeSQl('Delete from '+BaseDossier+'.DBO.paramsoc where not ('+V_PGI.PCLParamSocInclus+')');

   //--- On supprime la table AFFAIRE
   OptimiserTable (BaseDossier,'AFFAIRE');
 Finally
   SourisNormale;
 end;
end;

//-------------------------------
//--- Nom : DropTableOptimiser
//-------------------------------
procedure OptimiserTable (BaseDossier : String; NomTable : String);
begin
 if IsTableExiste(BaseDossier,NomTable) then
  DroperTable (BaseDossier,NomTable,True)
 else
  DroperTable (BaseDossier,NomTable,False);
end;

//------------------------
//--- Nom : DroperTable
//------------------------
procedure DroperTable (BaseDossier : String; NomTable : String; bPhysique : boolean );
var Prefixe  : String ;
    OkDelete : Boolean ;
begin
 OkDelete:=True;
 Prefixe:=DonnerPrefixe (BaseDossier,NomTable);

 Try
  if bPhysique then ExecuteSQL('DROP TABLE '+BaseDossier+'.DBO.'+NomTable) ;
 Except
  OkDelete:=False ;
 End ;

 if ((OkDelete) and (Prefixe<>'')) then
  begin
   ExecuteSQL('DELETE FROM '+BaseDossier+'.DBO.DETABLES WHERE DT_PREFIXE="'+Prefixe+'"') ;
   ExecuteSQL('DELETE FROM '+BaseDossier+'.DBO.DECHAMPS WHERE DH_PREFIXE="'+Prefixe+'"') ;
   ExecuteSQL('DELETE FROM '+BaseDossier+'.DBO.DESHARE WHERE DS_NOMTABLE="'+NomTable+'"');
  end;
end;

//--------------------------
//--- Nom : DonnerPrefixe
//--------------------------
function DonnerPrefixe (BaseDossier : String; NomTable : String): String;
var SSql : String;
    RSql : TQuery;
begin
 Result:='';
 SSql:='SELECT DT_PREFIXE FROM '+BaseDossier+'.DBO.DETABLES WHERE DT_NOMTABLE="'+NomTable+'"';
 RSql:=OpenSQL(SSql,True) ;
 if (not Rsql.eof) then Result:=RSql.Fields[0].AsString;
 Ferme(RSql);
end;
{$ENDIF MAJPCL}
{$ENDIF EAGLCLIENT}

//--------------------------
//--- Nom : IsTableExiste
//--------------------------
function IsTableExiste (BaseDossier : String; NomTable : String): Boolean;
begin
 Result:=ExisteSQL('SELECT NAME FROM '+BaseDossier+'.DBO.SYSOBJECTS WHERE NAME="'+NomTable+'" AND ID=OBJECT_ID("'+BaseDossier+'.DBO.'+nomtable+'")');
end;
{$ENDIF BUREAU}


{$IFDEF EAGLCLIENT}
function CheckInstallAppli(Lappli : TPGIAppli): Boolean;
var
    iResInstall, lgser: Integer;
    ErrMsg, PathMnu, ligne, cle, valeur: String;
    MnuFile: TextFile;
begin
  // if Lappli.Args<>'###A INSTALLER###' then exit;
  // => en fait non : on checkmodule systématiquement, car une appli doit
  // être mise à jour par le bureau AVANT d'être lancée, sinon elle fera
  // son auto-update, puis ne saura pas se connecter en dbxxx@section

  // C'était pour gérer plusieurs exe rattachés à la même appli (ex : comsx rattaché à ccs5)
  // mais inutile : il suffit de mettre un .inf par appli dans le kit
  // Result:=(AppServer.CheckModule(InfoAppli.CheckNom)=1);
  Result := (AppServer.CheckModule(Lappli.Nom)=1);
  if Not Result then
  begin
       AppServer.InstallModule(Lappli.Nom);
       // iResInstall:=AppServer.CheckModule(Lappli.CheckNom) ;
       iResInstall := AppServer.CheckModule(Lappli.Nom) ;
       Result := (iResInstall=1) or ( (iResInstall=0) and (FileExists(IncludeTrailingPathDelimiter(Lappli.Path)+Lappli.Nom))) ;
       if not Result then
       begin
            ErrMsg := 'Impossible d''installer l''application '+Lappli.Titre;
            if iResInstall=2 then
                ErrMsg := ErrMsg + #13#10 + ' Au moins un composant a une version mineure trop ancienne'
            else if iResInstall=0 then
                ErrMsg := ErrMsg + #13#10 + ' Au moins un composant est absent, ou a une version majeure trop ancienne';
            PGIInfo(ErrMsg);
            exit; // SORTIE
       end;
  end;

 // L'appli s'est installée, on relit le .mnu qu'elle a éventuellement maj !
 PathMnu := ExtractFilePath(Application.ExeName)+ChangeFileExt(Lappli.Nom, 'MNU');
 // #### le code ci-dessous est dupliqué de PGIAppli et serait mieux dans un MnuToAppli(...)
 if FileExists(PathMnu) then
  begin
   // valeurs par défaut
   Lappli.Titre := '';
   Lappli.Serie := 'S5';
   Lappli.Path := ExtractFilePath(Application.ExeName);
   Lappli.Tag := '0';
   Lappli.Tool := False;
   Lappli.Pgi := True;
   Lappli.DbCom := False;
   Lappli.Std := False;
   Lappli.Args := '';
   Lappli.Visible := True;
   Lappli.Plaq := False;
   Lappli.NomHal := '';
   Lappli.Eagl := False;
   Lappli.Nomade := True;
   lgser := 0;
   SetLength(Lappli.CodSerias, lgser);
   // ouverture
   AssignFile(MnuFile, PathMnu);
   Reset(MnuFile);
   // lecture
   Readln(MnuFile, ligne);
   while ligne<>'' do
    begin
     cle := ReadTokenPipe(ligne, '=');
     valeur := ReadTokenPipe(ligne, '=');
     // selon la cle
     if valeur<>'' then
      begin
       if cle='EXE' then
        Lappli.Nom := valeur
       else if cle='TITRE' then
        Lappli.Titre := valeur
       else if cle='SERIE' then
        Lappli.Serie := valeur
       else if cle='ORDRE' then
        Lappli.Ordre := valeur
       else if cle='PATH' then
        begin
         if Copy(valeur, Length(Valeur), 1)<>'\' then valeur := valeur + '\';
         if FileExists(valeur+Lappli.Nom) then Lappli.Path := valeur;
        end
       else if cle='TAG' then
        Lappli.Tag := valeur
       else if cle='TOOL' then
        Lappli.Tool := (valeur='TRUE')
       else if (cle='PGI') then
        Lappli.Pgi := (valeur<>'FALSE')
       else if (cle='NOPGIEXEC') then  // flag rajouté par ALR mais pour
        Lappli.Pgi := (valeur<>'TRUE') // l'instant même rôle que le flag PGI...
       else if cle='DBCOM' then
        Lappli.DbCom := (valeur='TRUE')
       else if cle='STD' then
        Lappli.Std := (valeur='TRUE')
       else if Copy(cle, 1, 3)='ARG' then
        Lappli.Args := valeur
       else if cle='VISIBLE' then
        Lappli.Visible := (valeur<>'FALSE')
       else if cle='PLAQ' then
        Lappli.Plaq := (valeur='TRUE')
       else if Copy(cle, 1, 6)='CODSER' then
        begin
         // exemple :
         // CODSER1=00140039;Comptabilité
         // CODSER2=00290039;Budget
         // ...
         lgser := lgser+1;
         SetLength(Lappli.CodSerias, lgser);
         Lappli.CodSerias[lgser-1] := valeur;
        end
       else if cle='NOMHAL' then
        Lappli.NomHal := valeur
       else if cle='EAGL' then
        Lappli.Eagl := (valeur='TRUE')
       else if cle='NOMADE' then
        Lappli.Nomade := (valeur<>'FALSE');
      end; // fin du if valeur<>''

     // ligne suivante
     Readln(MnuFile, ligne);
    end;
   CloseFile(MnuFile);
  end;

end;
{$ENDIF}

end.





