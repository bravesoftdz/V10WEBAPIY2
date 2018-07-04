{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. : Regroupement variables globales JURI
Mots clefs ... :
*****************************************************************}
unit EntJUR;
/////////////////////////////////////////////////////////////////
interface
/////////////////////////////////////////////////////////////////
uses
   {$ifdef eAGLClient}
   MenuOLX,
   {$else}
   MenuOLG,
   {$ENDIF EAGLCLIENT}
   Forms, Controls, UTOB, PGIAppli, UComOutils, ParamSoc, CBPPath,
   USynchroConst, USynchroOutils, HMsgBox, UJurOutilsAPP, UJurOutilsCFE;
/////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////
Type
   LaVariableJUR = Class

         sRepStd_c, sRepDoc_c, sRepDos_c, sLectRes_c,
         sRepLog_c, sRepMailing_c, sCurGuidPer_c : string;
         //
         bSeria_c, bSeriaGed_c, bSeriaMess_c, bSeriaDoc_c, bSeriaRF_c : boolean;
         bVersionRF_c, bValidRF_c : boolean;

         bSynchroOK_c, bCFEAutorise_c : boolean;
      	bBeatroceAdm_c  : boolean;
         bPilotage_c : boolean;

         sCmdLine_c : string;
         OBSysAppli_c, OBSysEve_c : TOB;

      Private

         function  SyncOptionOK(sOption_p : string; iValOpt_p : integer) : boolean;
         function  SyncLance(sTagApp_p, sMode_p, sEve_p, sRestriction_p : string) : boolean;
         procedure SyncOptionMaj(sTagApp_p, sModeSync_p, sEve_p : string;
                      iIndOpt_p : integer; cValOpt_p : char = #0);

      Public
         procedure InitParamsSynchro;
         procedure LibereParamsSynchro;
         procedure ChargeParamsSynchro;

         procedure BuildMenuSynchro;

         procedure SyncLanceImp;
         procedure SyncLanceExp;
         procedure SyncLanceAppli(sTagApp_p, sMode_p, sEve_p : string);

         function  IsAppSynchro(sTagApp_p : string) : boolean;
   end;
/////////////////////////////////////////////////////////////////
var
   VH_JUR : LaVariableJUR;
/////////////////////////////////////////////////////////////////
procedure InitLaVariableJUR;
procedure LibereLaVariableJUR;
procedure ChargeParamsJUR;

function  SetModeFonc(sModeFonc_p : String) : boolean;

//////////////////////// IMPLEMENTATION /////////////////////////
implementation
/////////////////////////////////////////////////////////////////
uses
   sysutils, hent1, hctrls, FileCtrl;
/////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 11/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure InitLaVariableJUR;
begin
   VH_JUR := LaVariableJUR.Create;
   VH_JUR.bBeatroceAdm_c := JURGetEnvVar('BEATROCE', 'BEATROCE');
   VH_JUR.bPilotage_c := JURGetEnvVar('PILOTAGE', 'TRUE');
   VH_JUR.bValidRF_c := JURGetEnvVar('VALID_RF', 'VALID_RF');

   VH_JUR.InitParamsSynchro;

   if V_Applis = Nil then
      V_Applis := TPGILesApplis.Create;      // libéré dans le finalization de PGIAppli;         //????
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 12/03/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure ChargeParamsJUR;
var
   sRacine_l, sRep_l : string;
   IPos_l : integer;
begin
   {$ifdef EAGLSERVER}
   {$else}
//   VH_JUR.sRepStd_c := ChangeStdDatPath('$STD');
//   VH_JUR.sRepDat_c := ChangeStdDatPath('$DAT');
   {$endif}

   VH_JUR.sRepDoc_c := TCBPPath.GetCegidDistriDoc;// + '\' + V_PGI.CodeProduit;
   VH_JUR.sRepStd_c := TCBPPath.GetCegidDistriStd + '\' + V_PGI.CodeProduit;

   VH_JUR.sRepLog_c := TCBPPath.GetCegidDataDistri + '\LOG\' + V_PGI.CodeProduit;
   VH_JUR.sRepMailing_c := GetParamSocSecur('SO_ENVDISK1', '') + '\DAT\JURI\MAILING';

   InitRep(VH_JUR.sRepLog_c);
   InitRep(VH_JUR.sRepMailing_c);

//   GetParamsoc('SO_ENVPATHDAT');  //V_PGI.DatPath;
//   VH_JUR.sRepDatLoc_c := GetParamsoc('SO_ENVPATHDATLOC');

//   VH_JUR.ChargeParamsSynchro;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure LibereLaVariableJUR;
Begin
   if Assigned(VH_JUR) then
   begin
      VH_JUR.LibereParamsSynchro;
      VH_JUR.Free;
      VH_JUR := Nil;
   end;
End;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 06/08/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function  SetModeFonc(sModeFonc_p : String) : boolean;
begin
   result := true;
   try
      ExecuteSQL('UPDATE SOCIETE SET SO_MODEFONC = ' + sModeFonc_p);
   except
      result := false;
    // PGIInfo('Le mode de fonctionnement par défaut n''est pas renseigné dans cette société.');
   end;
end;

/////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/06/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LaVariableJUR.InitParamsSynchro;
begin
   OBSysAppli_c := TOB.Create('JUSYSAPPLI', nil, -1);
   OBSysEve_c := TOB.Create('JUSYSEVE', nil, -1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/06/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure LaVariableJUR.LibereParamsSynchro;
begin
   if OBSysAppli_c <> nil then
      FreeAndNil(OBSysAppli_c);
   if OBSysEve_c <> nil then
      FreeAndNil(OBSysEve_c);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/06/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure LaVariableJUR.ChargeParamsSynchro;
var
   iInd_l : integer;
   sTag_l, sRepSync_l : string;
begin
   bSynchroOK_c := false;
//   BOBDeletePreSocref('SYNC', V_PGI.NumVersionBase);
//   BOBMajParams('SYNC', 'Paramètres Synchronisation', false);

   // Applis à synchroniser
   OBSysAppli_c.LoadDetailDBFromSQL('JUSYSAPPLI',
                                    'SELECT * FROM JUSYSAPPLI '+
                                    'WHERE JSA_TAG NOT LIKE "X%"');

   OBSysEve_c.LoadDetailDBFromSQL('JUSYSEVE','SELECT * FROM JUSYSEVE');

   for iInd_l := 0 to OBSysAppli_c.Detail.Count - 1 do
   begin
      sTag_l := OBSysAppli_c.Detail[iInd_l].GetString('JSA_TAG');
      // au moins une appli autorisée??
      if SyncAppliOK(sTag_l) then
         bSynchroOK_c := true;
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/06/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure LaVariableJUR.BuildMenuSynchro;
var
   iAdm_l, iInd_l : integer;
   sTag_l, sLib_l : string;
begin
   if not bSynchroOK_c then Exit;

   iAdm_l := FMenuG.OutLook.Groups.Count;
   FMenuG.OutLook.AddGroup('Synchronisation', '', 261000);
   for iInd_l := 0 to OBSysAppli_c.Detail.Count - 1 do
   begin
      sTag_l := OBSysAppli_c.Detail[iInd_l].GetString('JSA_TAG');
      if not SyncAppliOK(sTag_l) then
         Continue;
      sLib_l := OBSysAppli_c.Detail[iInd_l].GetString('JSA_LIBAPPLI');
      sLib_l := StringReplace(sLib_l, 'Applications ', '', []);
      FMenuG.OutLook.Groups.Items[iAdm_l].AddItem(sLib_l, '', 261000 + StrToInt(sTag_l), 39);
//      Inc(iAdm_l);
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/06/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
function  LaVariableJUR.IsAppSynchro(sTagApp_p : string) : boolean;
begin
   result := bSynchroOK_c and
             SyncAppliOK(sTagApp_p) and
             (OBSysAppli_c.FindFirst(['JSA_TAG'], [sTagApp_p], true) <> nil);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/06/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function LaVariableJUR.SyncOptionOK(sOption_p : string; iValOpt_p : integer) : boolean;
begin
   result := (sOption_p[iValOpt_p] = '0');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 07/06/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure LaVariableJUR.SyncLanceImp;//Auto(sMode_p : string);
var
   OBEve_l : TOB;
   sTagApp_l, sOption_l, sLibApp_l, sEve_l : string;
   iInd_l : integer;
   bChecked_l : boolean;
begin
   if not bSynchroOK_c then exit;

   for iInd_l := 0 to OBSysAppli_c.Detail.count - 1 do
   begin
      sTagApp_l := OBSysAppli_c.Detail[iInd_l].GetString('JSA_TAG');
      if not SyncAppliOK(sTagApp_l) then continue;
      
      OBEve_l := OBSysEve_c.FindFirst(['JSE_TAG', 'JSE_MODE'], [sTagApp_l, csModeImp_g], true);
      sOption_l := OBEve_l.GetString('JSE_OPTION');
      if not SyncOptionOK(sOption_l, ciOptUnic_g) and
         not SyncOptionOK(sOption_l, ciOptDone_g) then
         Continue;

      sLibApp_l := OBSysAppli_c.Detail[iInd_l].GetString('JSA_LIBAPPLI');
      sLibApp_l := StringReplace(sLibApp_l, 'Applications ', '', []);

      if (PGIAsk('Voulez-vous synchroniser l''annuaire avec les données "' + sLibApp_l + '"?',
                 'Synchronisation PGI') = mrNo) then
      begin
         if (PGIAsk('Ne plus afficher ce message?',
                    'Synchronisation PGI') = mrYes) then
         SyncOptionMaj(sTagApp_l, csModeImp_g, sEve_l, ciOptUnic_g, '1');
         Continue;
      end;

      sEve_l := OBEve_l.GetString('JSE_EVENEMENT');
      if not SyncLance(sTagApp_l, csModeImp_g, sEve_l, '') then break;

      bChecked_l := ExisteSQL('SELECT * FROM JUSYSCLE ' +
                              'WHERE JSC_TAG = "' +sTagApp_l + '" ' +
                              '  AND JSC_MODE = "' +csModeImp_g + '" ' +
                              '  AND JSC_CODEI IS NOT NULL ' +
                              '  AND JSC_CODEX IS NOT NULL');

      if bChecked_l then
         SyncOptionMaj(sTagApp_l, csModeImp_g, sEve_l, ciOptDone_g, '1');
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 13/06/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure LaVariableJUR.SyncLanceExp;
var
   OBEve_l : TOB;
   sRequete_l, sEve_l, sOption_l : string;
   dtDate_l{, dtDatePre_l} : TDateTime;
begin
   if not bSynchroOK_c then exit;

   // Quadratus : Export des évènements
   sEve_l := 'EVENEMENTS';
   if not SyncAppliOK(csTagQuadra_g) then exit;

   OBEve_l := OBSysEve_c.FindFirst(['JSE_TAG', 'JSE_MODE', 'JSE_EVENEMENT'], [csTagQuadra_g, csModeExp_g, 'EVENEMENTS'], true);
   sOption_l := OBEve_l.GetString('JSE_OPTION');
   if not SyncOptionOK(sOption_l, ciOptDone_g) then Exit;

   dtDate_l := OBEve_l.GetDateTime('JSE_DATE');
//   dtDatePre_l := OBEve_l.GetDateTime('JSE_DATEPRE');

   sRequete_l := 'SELECT JEV_CODEEVT, JEV_DATEMODIF ' +
                 'FROM JUEVENEMENT ' +
                 'WHERE JEV_DOMAINEACT = "JUR" ';

//   if dtDate_l = dtDatePre_l then
//      sRequete_l := sRequete_l + ''

   if dtDate_l <> iDate1900 then
      sRequete_l := sRequete_l +
                 '  AND (JEV_DATECREATION >= "' + USDATETIME(dtDate_l) + '"' +
                 '   OR JEV_DATEMODIF >= "' + USDATETIME(dtDate_l) + '")';

   if not ExisteSQL(sRequete_l) then exit;

   if (PGIAsk('Voulez-vous exporter les évènements vers "Quadratus"?',
              'Synchronisation PGI') = mrNo) then
   begin
      if (PGIAsk('Ne plus afficher ce message?',
                 'Synchronisation PGI') = mrYes) then
         SyncOptionMaj(csTagQuadra_g, csModeExp_g, sEve_l, ciOptDone_g, '1');
      exit;
   end;

   if not SyncLance(csTagQuadra_g, csModeExp_g, sEve_l, '"' + USDATETIME(dtDate_l) + '"') then exit;


{   if ExisteSQL(sRequete_l) then
   begin
      OBEve_l.PutValue('JSE_DATEPRE', OBEve_l.GetDateTime('JSE_DATE'));
      OBEve_l.PutValue('JSE_DATE', Date + Time);
      OBEve_l.UpdateDB;
   end;}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/06/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure LaVariableJUR.SyncLanceAppli(sTagApp_p, sMode_p, sEve_p : string);
begin
   if not bSynchroOK_c then exit;
   
   if SyncAppliOK(sTagApp_p) then
      SyncLance(sTagApp_p, sMode_p, sEve_p, '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 08/06/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function LaVariableJUR.SyncLance(sTagApp_p, sMode_p, sEve_p, sRestriction_p : string) : boolean;
var
   sCmdParam_l : string;
begin
   // Application Synchro installée??
   if sCmdLine_c = '' then
      sCmdLine_c := APPCmdLine(TCBPPath.GetCegidDistriApp, csExeSync_g, '', '');

   // paramètres passés à Synchro
   sCmdParam_l := ' /TAG=' + sTagApp_p +
                  ' /MODE=' + sMode_p +
                  ' /RESTRICTION=' + sRestriction_p +
                  ' /EVENEMENT=' + sEve_p;

   result := APPLance(sCmdLine_c, sCmdParam_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 09/06/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure LaVariableJUR.SyncOptionMaj(sTagApp_p, sModeSync_p, sEve_p : string;
                      iIndOpt_p : integer; cValOpt_p : char = #0);
var
   OBEve_l : TOB;
   sValOpt_l : string;
begin
   if sEve_p = '' then
      OBEve_l := OBSysEve_c.FindFirst(['JSE_TAG', 'JSE_MODE'], [sTagApp_p, sModeSync_p], true)
   else
      OBEve_l := OBSysEve_c.FindFirst(['JSE_TAG', 'JSE_MODE', 'JSE_EVENEMENT'], [sTagApp_p, sModeSync_p, sEve_p], true);

   while OBEve_l <> nil do
   begin
      sValOpt_l := OBEve_l.GetString('JSE_OPTION');
      sValOpt_l[iIndOpt_p] := cValOpt_p;
      OBEve_l.PutValue('JSE_OPTION', sValOpt_l);
      if sEve_p = '' then
         OBEve_l := OBSysEve_c.FindNext(['JSE_TAG', 'JSE_MODE'], [sTagApp_p, sModeSync_p], true)
      else
         OBEve_l := OBSysEve_c.FindNext(['JSE_TAG', 'JSE_MODE', 'JSE_EVENEMENT'], [sTagApp_p, sModeSync_p, sEve_p], true);
   end;
   OBSysEve_c.UpdateDB;
end;
/////////////////////////////////////////////////////////////////
end.

