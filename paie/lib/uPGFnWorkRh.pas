{***********UNITE*************************************************
Auteur  ...... : SB
Créé le ...... : 13/07/2010
Modifié le ... : 13/07/2010
Description .. : Fonctions workflow RH
Mots clefs ... : PROCESSUS;SERVICES*
*****************************************************************}
{
03/08/2010 Ajout des processus de la formation et du budget, création de function générique pour les traitements similaires
PT1 04/11/2010 Suppression seulement des éléments existants en cas de re-maj de base via cptx
PT2 20/12/2010 Suivi de dde : Affectation du role Adm si pas de référent trouvé
PT3 20/12/2010 Création mvt par un référent : affectation de l'intervenant sur ligne d'ordre 0
}
unit uPGFnWorkRh;

interface
uses   Classes,Sysutils, uTob,UtilOutils;

   procedure PgInitTabletteTypeDesti ;  { appel CPTX }

   procedure PGGenereProcessusAbs(WithSupp : Boolean; StWhere, StCondition : String);  { appel CPTX }
   procedure PGGenereProcessusFor(WithSupp : Boolean; StWhere : String);  { appel CPTX }
   procedure PGGenereProcessusInscFor(WithSupp : Boolean; StWhere : String);  { appel CPTX }
   Function  PGRecupProcess(TypProcess : String) : TOB ;
   procedure PGCreeProcess(TypProcess : String;  LcTob_Process,LcTFille : Tob);
   Function  PgCreatLignSuiviDde (Tob_SuivDDe,LcTFille:Tob; TypProcess,StRole,StMod,Ref,TbMaitre,ChpEtat,Suf : string; cpt : integer) : Tob;  { PT2 }
   function  PgRechReferentRH(Sal : string; LcTP : Tob) : String;
   function  PgRechReferentRHAdmin(StRole,StMod : string) : String;    { PT2 }

   procedure PgSupprSuiviDemande (StTyp,StWhere : string);
   function  IfExistSuiviDemande (StTyp,StWhere,StEtat : string) : Boolean ;
   function  PgMajSuiviDemande(TypProc,Guid,Mode,Sal,ValidPar,FicPrec,StRef,TBMaitre : String) : Boolean ;  { PT10 Ajout FicPrec }
   function  PgGetGuidAbsence(TypMvt,Sal: string; Ordre : Integer) : String;  { PT10 }


   procedure PGRecupAffectRoleRh;                                        { appel CPTX }
   procedure PgCreateAffectRoleRH(Creat : Boolean; Serv,Role,Module,Ref,AncRef :String);


   procedure PgMajTabletteTypeDesti(Mode, Niv : string ) ;
   procedure PGGenereTypDestModule ( Niv,SMod,actif : string ) ;
   procedure PGSupprTypDestModule ( Niv,SMod: string ) ;
   function PgGetPrefModuleRH(Smod : String) : string; { PT11 }



type
  PGParamRoleRh = record
    stChamp     : string;
    stTypeAffect: String;
    stRoleRH: String;
    stModuleRh: String;
  end;


implementation

uses
    HCtrls,
    Hent1,
    CBPDates
{$IFNDEF EAGLCLIENT}
     //,FE_Main
     ,uDbxDataSet
{$ELSE}
     //,MainEAGL
{$ENDIF}
;
Var
TabElement : array[1..9] of PGParamRoleRh =
  (
    {1} (stChamp: 'PSE_RESPSERVICE'; stTypeAffect: 'SAL';  stRoleRH: 'RES';  stModuleRh: 'GAL'; ),
    {2} (stChamp: 'PSE_RESPONSABS';  stTypeAffect: 'SAL';  stRoleRH: 'RES';  stModuleRh: 'ABS'; ),
    {3} (stChamp: 'PSE_ASSISTABS';   stTypeAffect: 'SAL';  stRoleRH: 'ASS';  stModuleRh: 'ABS'; ),
    {4} (stChamp: 'PSE_RESPONSNDF';  stTypeAffect: 'SAL';  stRoleRH: 'RES';  stModuleRh: 'NDF'; ),
    {5} (stChamp: 'PSE_ASSISTNDF';   stTypeAffect: 'SAL';  stRoleRH: 'ASS';  stModuleRh: 'NDF'; ),
    {6} (stChamp: 'PSE_RESPONSVAR';  stTypeAffect: 'SAL';  stRoleRH: 'RES';  stModuleRh: 'VAR'; ),
    {7} (stChamp: 'PSE_ASSISTVAR';   stTypeAffect: 'SAL';  stRoleRH: 'ASS';  stModuleRh: 'VAR'; ),
    {8} (stChamp: 'PSE_RESPONSAUG';  stTypeAffect: 'SAL';  stRoleRH: 'RES';  stModuleRh: 'AUG'; ),
    {9} (stChamp: 'PSE_ASSISTAUG';   stTypeAffect: 'SAL';  stRoleRH: 'ASS';  stModuleRh: 'AUG'; )
//   {10} (stChamp: 'PSE_ADMINISTABS'; stTypeAffect: 'SAL';  stRoleRH: 'ADM';  stModuleRh: 'ABS'; )

   );
TabService : array[1..18] of PGParamRoleRh =
  (
    {1} (stChamp: 'PGS_RESPSERVICE';    stTypeAffect: 'SER';  stRoleRH: 'RES';  stModuleRh: 'GAL'; ),
    {2} (stChamp: 'PGS_SECRETAIRE';     stTypeAffect: 'SER';  stRoleRH: 'ASS';  stModuleRh: 'GAL'; ),
    {3} (stChamp: 'PGS_ADJOINTSERVICE'; stTypeAffect: 'SER';  stRoleRH: 'ADJ';  stModuleRh: 'GAL'; ),
    {4} (stChamp: 'PGS_RESPONSABS';     stTypeAffect: 'SER';  stRoleRH: 'RES';  stModuleRh: 'ABS'; ),
    {5} (stChamp: 'PGS_SECRETAIREABS';  stTypeAffect: 'SER';  stRoleRH: 'ASS';  stModuleRh: 'ABS'; ),
    {6} (stChamp: 'PGS_ADJOINTABS';     stTypeAffect: 'SER';  stRoleRH: 'ADJ';  stModuleRh: 'ABS'; ),
    {7} (stChamp: 'PGS_RESPONSNDF';     stTypeAffect: 'SER';  stRoleRH: 'RES';  stModuleRh: 'NDF'; ),
    {8} (stChamp: 'PGS_SECRETAIRENDF';  stTypeAffect: 'SER';  stRoleRH: 'ASS';  stModuleRh: 'NDF'; ),
    {9} (stChamp: 'PGS_ADJOINTNDF';     stTypeAffect: 'SER';  stRoleRH: 'ADJ';  stModuleRh: 'NDF'; ),
   {10} (stChamp: 'PGS_RESPONSVAR';     stTypeAffect: 'SER';  stRoleRH: 'RES';  stModuleRh: 'VAR'; ),
   {11} (stChamp: 'PGS_SECRETAIREVAR';  stTypeAffect: 'SER';  stRoleRH: 'ASS';  stModuleRh: 'VAR'; ),
   {12} (stChamp: 'PGS_ADJOINTVAR';     stTypeAffect: 'SER';  stRoleRH: 'ADJ';  stModuleRh: 'VAR'; ),
   {13} (stChamp: 'PGS_RESPONSFOR';     stTypeAffect: 'SER';  stRoleRH: 'RES';  stModuleRh: 'FOR'; ),
   {14} (stChamp: 'PGS_SECRETAIREFOR';  stTypeAffect: 'SER';  stRoleRH: 'ASS';  stModuleRh: 'FOR'; ),
   {15} (stChamp: 'PGS_ADJOINTFOR';     stTypeAffect: 'SER';  stRoleRH: 'ADJ';  stModuleRh: 'FOR'; ),
   {16} (stChamp: 'PGS_RESPONSAUG';     stTypeAffect: 'SER';  stRoleRH: 'RES';  stModuleRh: 'AUG'; ),
   {17} (stChamp: 'PGS_SECRETAIREAUG';  stTypeAffect: 'SER';  stRoleRH: 'ASS';  stModuleRh: 'AUG'; ),
   {18} (stChamp: 'PGS_ADJOINTAUG';     stTypeAffect: 'SER';  stRoleRH: 'ADJ';  stModuleRh: 'AUG'; )
   );

TabManquant : array[1..9] of PGParamRoleRh =
  (
    {1} (stChamp: 'PGS_RESPONSFOR';      stTypeAffect: 'SAL';  stRoleRH: 'RES';  stModuleRh: 'FOR'; ),
    {2} (stChamp: 'PGS_SECRETAIRE';      stTypeAffect: 'SAL';  stRoleRH: 'ASS';  stModuleRh: 'GAL'; ),
    {3} (stChamp: 'PGS_SECRETAIREFOR';   stTypeAffect: 'SAL';  stRoleRH: 'ASS';  stModuleRh: 'FOR'; ),
    {4} (stChamp: 'PGS_ADJOINTSERVICE';  stTypeAffect: 'SAL';  stRoleRH: 'ADJ';  stModuleRh: 'GAL'; ),
    {5} (stChamp: 'PGS_ADJOINTABS';      stTypeAffect: 'SAL';  stRoleRH: 'ADJ';  stModuleRh: 'ABS'; ),
    {6} (stChamp: 'PGS_ADJOINTNDF';      stTypeAffect: 'SAL';  stRoleRH: 'ADJ';  stModuleRh: 'NDF'; ),
    {7} (stChamp: 'PGS_ADJOINTVAR';      stTypeAffect: 'SAL';  stRoleRH: 'ADJ';  stModuleRh: 'VAR'; ),
    {8} (stChamp: 'PGS_ADJOINTFOR';      stTypeAffect: 'SAL';  stRoleRH: 'ADJ';  stModuleRh: 'FOR'; ),
    {9} (stChamp: 'PGS_ADJOINTAUG';      stTypeAffect: 'SAL';  stRoleRH: 'ADJ';  stModuleRh: 'AUG'; )
   );


procedure PGRecupAffectRoleRh;
Var
Q : TQuery;
Tob_PgAffectRoleRh,Tob_Recup : TOB;
TRH,TDS : Tob;
i,j : integer;
St : String;
begin

  ExecuteSql('DELETE FROM PGAFFECTROLERH WHERE PFH_MODULERH IN ("ABS","AUG","FOR","GAL","NDF","VAR")'); { PT1 }
  Tob_PgAffectRoleRh :=  TOB.Create('PGAFFECTROLERH',nil,-1);

  { Récupération de la table deportsal }
  For i := 1 to 9 do
    Begin
      Tob_Recup:=TOB.Create('Les_affects',nil,-1);
      { Récupération de l'information }
      St :='SELECT DISTINCT '+TabElement[i].stChamp+',PSE_SALARIE,PSE_CODESERVICE FROM DEPORTSAL WHERE '+TabElement[i].stChamp+'<>""  ORDER BY '+TabElement[i].stChamp;
      Q := OpenSql(St,True);
      if Not Q.eof then
        Tob_Recup.LoadDetailDB('Les_affects','','',Q,False);
      Ferme(Q);
      For j := 0 to  Tob_Recup.detail.count-1 do
        Begin
        TDS:= Tob_Recup.detail[j];
        if (Trim(TDS.GetValue(TabElement[i].stChamp))) <> '' then
           Begin
           { Création de la tob fille d'affectation du rôle }
           TRH := TOB.Create('PGAFFECTROLERH' ,Tob_PgAffectRoleRh,-1);
           TRH.PutValue('PFH_TYPEAFFECTROLE'  ,TabElement[i].stTypeAffect );
           TRH.PutValue('PFH_ROLERH'          ,TabElement[i].stRoleRH);
           TRH.PutValue('PFH_REFERENTRH'      ,TDS.GetValue(TabElement[i].stChamp));
           TRH.PutValue('PFH_MODULERH'        ,TabElement[i].stModuleRh);
           TRH.PutValue('PFH_CODESERVICE'     ,TDS.GetValue('PSE_CODESERVICE'));
           TRH.PutValue('PFH_SALARIE'         ,TDS.GetValue('PSE_SALARIE'));
           TRH.InsertDB(nil);
           End;
        End;
      Tob_Recup.Free;
    End;


   { Récupération des administrateurs des absences }
      St :='SELECT DISTINCT PSE_ADMINISTABS,PSE_SALARIE FROM DEPORTSAL WHERE PSE_ADMINISTABS="X"';
      Q := OpenSql(St,True);
      if Not Q.eof then
         Begin
         Tob_Recup:=TOB.Create('Les_affects',nil,-1);
         Tob_Recup.LoadDetailDB('Les_affects','','',Q,False);
         Ferme(Q);
         For j := 0 to  Tob_Recup.detail.count-1 do
            Begin
            TDS:= Tob_Recup.detail[j];
            { Création de la tob fille d'affectation du rôle }
            TRH := TOB.Create('PGAFFECTROLERH' ,Tob_PgAffectRoleRh,-1);
            TRH.PutValue('PFH_TYPEAFFECTROLE'  ,'SAL');
            TRH.PutValue('PFH_ROLERH'          ,'ADM');
            TRH.PutValue('PFH_MODULERH'        ,'ABS');
            TRH.PutValue('PFH_CODESERVICE'     ,'***');
            TRH.PutValue('PFH_SALARIE'         ,TDS.GetValue('PSE_SALARIE'));
            TRH.PutValue('PFH_REFERENTRH'      ,TDS.GetValue('PSE_SALARIE'));    
            TRH.InsertDB(nil);
            End;
         Tob_Recup.Free;
         End
       else
         Ferme(Q);

    { Récupération des affectations service }
    St :='SELECT * FROM SERVICES';
    Q := OpenSql(St,True);
    if Not Q.eof then
      Begin
      Tob_Recup:=TOB.Create('Les_affects',nil,-1);
      Tob_Recup.LoadDetailDB('Les_affects','','',Q,False);
      Ferme(Q);
      For i := 1 to 18 do
        Begin
        { Récupération de l'information }
        For j := 0 to  Tob_Recup.detail.count-1 do
          Begin
          TDS:= Tob_Recup.detail[j];
          St := TDS.GetValue(TabService[i].stChamp);
          if (Trim(St) <> '') then
             Begin
             { Création de la tob fille d'affectation du rôle }
             TRH := TOB.Create('PGAFFECTROLERH' ,Tob_PgAffectRoleRh,-1);
             TRH.PutValue('PFH_TYPEAFFECTROLE'  ,TabService[i].stTypeAffect );
             TRH.PutValue('PFH_ROLERH'          ,TabService[i].stRoleRH);
             TRH.PutValue('PFH_REFERENTRH'      ,TDS.GetValue(TabService[i].stChamp));
             TRH.PutValue('PFH_MODULERH'        ,TabService[i].stModuleRh);
             TRH.PutValue('PFH_CODESERVICE'     ,TDS.GetValue('PGS_CODESERVICE'));
             TRH.PutValue('PFH_SALARIE'         ,'***');
             TRH.InsertDB(nil);
             End;
          End;
        End;
      Tob_Recup.Free;
      End
     else
       Ferme(Q);




    { Création des affectations salariés manquantes }
      { Récupération de l'information }
      {PGS_RESPONSFOR,PGS_SECRETAIRE,PGS_SECRETAIREFOR,
      PGS_ADJOINTSERVICE,PGS_ADJOINTABS,PGS_ADJOINTNDF,PGS_ADJOINTVAR,PGS_ADJOINTFOR,PGS_ADJOINTAUG }
    St :='SELECT PGS_RESPONSFOR,PGS_SECRETAIRE,PGS_SECRETAIREFOR,PGS_ADJOINTSERVICE,PGS_ADJOINTABS,'+
         'PGS_ADJOINTNDF,PGS_ADJOINTVAR,PGS_ADJOINTFOR,PGS_ADJOINTAUG,'+
         'PGS_CODESERVICE,PSE_CODESERVICE,PSE_SALARIE '+
         'FROM SERVICES,DEPORTSAL WHERE PSE_CODESERVICE=PGS_CODESERVICE';
      Q := OpenSql(St,True);
      if Not Q.eof then
        Begin
        Tob_Recup:=TOB.Create('Les_affects',nil,-1);
        Tob_Recup.LoadDetailDB('Les_affects','','',Q,False);
        Ferme(Q);
        For j := 0 to  Tob_Recup.detail.count-1 do
          Begin
             TDS:= Tob_Recup.detail[j];
             For i := 1 to 9 do
               Begin
               if (Trim(TDS.GetValue(TabManquant[i].stChamp))) <> '' then
                  Begin
                  { Création de la tob fille d'affectation du rôle }
                  TRH := TOB.Create('PGAFFECTROLERH' ,Tob_PgAffectRoleRh,-1);
                  TRH.PutValue('PFH_TYPEAFFECTROLE'  ,TabManquant[i].stTypeAffect );
                  TRH.PutValue('PFH_ROLERH'          ,TabManquant[i].stRoleRH);
                  TRH.PutValue('PFH_REFERENTRH'      ,TDS.GetValue(TabManquant[i].stChamp));
                  TRH.PutValue('PFH_MODULERH'        ,TabManquant[i].stModuleRh);
                  TRH.PutValue('PFH_CODESERVICE'     ,TDS.GetValue('PSE_CODESERVICE'));
                  TRH.PutValue('PFH_SALARIE'         ,TDS.GetValue('PSE_SALARIE'));
                  TRH.InsertDB(nil);
                  End;
               End;
          End;
        Tob_Recup.Free;
        End
      else
        Ferme(Q);


    Tob_PgAffectRoleRh.free;

end;

procedure PGGenereProcessusAbs(WithSupp : Boolean; StWhere, StCondition : String);
Var
Tob_Abs,Tob_Process,TAbs, TSuividde, TSuiviDdeD, Tcreat, TCreatD : Tob;
St : String;
Q : TQuery;
i : integer;
Begin
TSuividde:= nil;	//PT5
//PT7
if (StCondition<>'') then
   StCondition:= 'WHERE '+StCondition;
//FIN PT7
if WithSupp then
//PT5
   begin
   St:= 'SELECT * FROM PGSUIVIDEMANDE '+
        'WHERE PSN_GUIDTABLE IN (SELECT PCN_GUID FROM ABSENCESALARIE '+
                                 'WHERE (PCN_TYPEMVT="ABS" OR (PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI")) '+
                                 'AND PCN_EXPORTOK="-" AND PCN_CODETAPE="..." '+
                                 'AND PCN_GUID NOT IN (SELECT DISTINCT PSN_GUIDTABLE FROM PGSUIVIDEMANDE '+
                                                      {'WHERE '+ PT7}StCondition+') '+StWhere+')';
   Q:= OpenSql (St, True);
   If not Q.eof then
      Begin
      TSuividde := Tob.create('Les_suivis_dde',nil,-1);
      TSuividde.LoadDetailDB('Les_suivis_dde','','',Q,False);
      Ferme(Q);
      end
   else
      Ferme(Q);
//FIN PT5

   ExecuteSql('DELETE FROM PGSUIVIDEMANDE '+
              'WHERE PSN_GUIDTABLE IN (SELECT PCN_GUID FROM ABSENCESALARIE '+
                                    'WHERE (PCN_TYPEMVT="ABS" OR (PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI")) '+
                                    'AND PCN_EXPORTOK="-" AND PCN_CODETAPE="..." '+
                                    'AND PCN_GUID NOT IN (SELECT DISTINCT PSN_GUIDTABLE FROM PGSUIVIDEMANDE '+
{PT5
                                                 'WHERE PSN_TYPPROCESSDDE="ABS" AND PSN_ETATSUIVIDDE<>"ATT"))');
}
                                                 {'WHERE '+ PT7}StCondition+') '+StWhere+')');
   end;
//FIN PT5

St := 'SELECT * FROM ABSENCESALARIE '+
       'WHERE (PCN_TYPEMVT="ABS" OR (PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI")) '+
       'AND PCN_EXPORTOK="-" AND PCN_CODETAPE="..." '+
       'AND PCN_GUID NOT IN (SELECT DISTINCT PSN_GUIDTABLE FROM PGSUIVIDEMANDE '+
{PT5
                            'WHERE PSN_TYPPROCESSDDE="ABS" AND PSN_ETATSUIVIDDE<>"ATT") '+StWhere+
}
                            {'WHERE '+ PT7}StCondition+') '+StWhere+
//FIN PT5
       'ORDER BY PCN_SALARIE';
Q := OpenSql(St,True);
If not Q.eof then
  Begin

  Tob_Abs := Tob.create('Les_absences_Att',nil,-1);
  Tob_Abs.LoadDetailDB('Les_absences_Att','','',Q,False);
  Ferme(Q);

  Tob_Process  :=  PGRecupProcess('ABS');
  if Tob_Process <> nil then
    For i:=0 to Tob_Abs.Detail.count-1 do
      Begin
      TAbs := Tob_Abs.Detail[i];
      PGCreeProcess('ABS',Tob_Process,TAbs);
      End;
  if Tob_Abs <> nil then Tob_Abs.free;
  if Tob_Process <> nil then Tob_Process.free;
  End
else
  Ferme(Q);

//PT5
if (TSuividde<>nil) then
   begin
   St:= 'SELECT * FROM PGSUIVIDEMANDE '+
        'WHERE PSN_GUIDTABLE IN (SELECT PCN_GUID FROM ABSENCESALARIE '+
                                 'WHERE (PCN_TYPEMVT="ABS" OR (PCN_TYPEMVT="CPA" AND PCN_TYPECONGE="PRI")) '+
                                 'AND PCN_EXPORTOK="-" AND PCN_CODETAPE="..." '+
                                 'AND PCN_GUID NOT IN (SELECT DISTINCT PSN_GUIDTABLE FROM PGSUIVIDEMANDE '+
                                                      {'WHERE '+ PT7}StCondition+') '+StWhere+')';
   Q:= OpenSql (St, True);
   If not Q.eof then
      Begin
      Tcreat := Tob.create('PGSUIVIDEMANDE',nil,-1);
      Tcreat.LoadDetailDB('PGSUIVIDEMANDE','','',Q,False);
      Ferme(Q);
      end
   else
      Ferme(Q);
//   PSN_TABLEMAITRE,PSN_GUIDTABLE,PSN_TYPPROCESSDDE,PSN_ORDRESUIVI
   For i:= 0 to TCreat.Detail.Count-1 do
       begin
       TCreatD:= TCreat.Detail[i];
       TSuiviDdeD:= TSuiviDde.FindFirst (['PSN_TABLEMAITRE','PSN_GUIDTABLE','PSN_TYPPROCESSDDE','PSN_ORDRESUIVI'],
                                         [TCreatD.GetValue ('PSN_TABLEMAITRE'),
                                         TCreatD.GetValue ('PSN_GUIDTABLE'),
                                         TCreatD.GetValue ('PSN_TYPPROCESSDDE'),
                                         TCreatD.GetValue ('PSN_ORDRESUIVI')], False);
       if (TSuiviDdeD<>nil) then
          begin
          TCreatD.PutValue('PSN_ETATSUIVIDDE', TSuiviDdeD.GetValue ('PSN_ETATSUIVIDDE'));

          TCreatD.UpdateDB;
          end;
       end;
   FreeAndNil (TSuividde);
   FreeAndNil (TCreat);
   end;
//FIN PT5
End;


procedure PGGenereProcessusFor(WithSupp : Boolean; StWhere : String);  { appel CPTX }
Var
Tob_For,Tob_Process,TFor : Tob;
St : String;
Q : TQuery;
i : integer;
Begin
if WithSupp then
   ExecuteSql('DELETE FROM PGSUIVIDEMANDE '+
              'WHERE PSN_GUIDTABLE IN '+
                   '(SELECT PFO_GUID FROM FORMATIONS '+
                   'WHERE PFO_ETATINSCFOR = "ATT"  '+
                   'AND PFO_GUID NOT IN (SELECT DISTINCT PSN_GUIDTABLE FROM PGSUIVIDEMANDE '+
                   'WHERE PSN_TYPPROCESSDDE="FOR" AND PSN_ETATSUIVIDDE<>"ATT"))');

St := 'SELECT * FROM FORMATIONS WHERE PFO_ETATINSCFOR = "ATT" '+
       'AND PFO_GUID NOT IN (SELECT DISTINCT PSN_GUIDTABLE FROM PGSUIVIDEMANDE '+
                            'WHERE PSN_TYPPROCESSDDE="FOR" AND PSN_ETATSUIVIDDE<>"ATT") '+StWhere+
       'ORDER BY PFO_SALARIE';

Q := OpenSql(St,True);
If not Q.eof then
  Begin
  Tob_For := Tob.create('formations_Att',nil,-1);
  Tob_For.LoadDetailDB('formations_Att','','',Q,False);
  Ferme(Q);

  Tob_Process  :=  PGRecupProcess('FOR');
  if Tob_Process <> nil then
    For i:=0 to Tob_For.Detail.count-1 do
      Begin
      TFor := Tob_For.Detail[i];
      PGCreeProcess('FOR',Tob_Process,TFor);
      End;
  if Tob_For <> nil then Tob_For.free;
  if Tob_Process <> nil then Tob_Process.free;
  End
else
  Ferme(Q);
End;


procedure PGGenereProcessusInscFor(WithSupp : Boolean; StWhere : String);  { appel CPTX }
Var
Tob_For,Tob_Process,TFor : Tob;
St : String;
Q : TQuery;
i : integer;
Begin
if WithSupp then
   ExecuteSql('DELETE FROM PGSUIVIDEMANDE '+
              'WHERE PSN_GUIDTABLE IN '+
                   '(SELECT PFI_GUID FROM INSCFORMATION '+
                   'WHERE PFI_ETATINSCFOR = "ATT"  '+
                   'AND PFI_GUID NOT IN (SELECT DISTINCT PSN_GUIDTABLE FROM PGSUIVIDEMANDE '+
                   'WHERE PSN_TYPPROCESSDDE="BUD" AND PSN_ETATSUIVIDDE<>"ATT"))');

St := 'SELECT * FROM INSCFORMATION WHERE PFI_ETATINSCFOR = "ATT" '+
       'AND PFI_GUID NOT IN (SELECT DISTINCT PSN_GUIDTABLE FROM PGSUIVIDEMANDE '+
       'WHERE PSN_TYPPROCESSDDE="BUD" AND PSN_ETATSUIVIDDE<>"ATT") '+StWhere+
       'ORDER BY PFI_SALARIE';

Q := OpenSql(St,True);
If not Q.eof then
  Begin
  Tob_For := Tob.create('Inscforms_Att',nil,-1);
  Tob_For.LoadDetailDB('Inscforms_Att','','',Q,False);
  Ferme(Q);

  Tob_Process  :=  PGRecupProcess('BUD');
  if Tob_Process <> nil then
    For i:=0 to Tob_For.Detail.count-1 do
      Begin
      TFor := Tob_For.Detail[i];
      PGCreeProcess('BUD',Tob_Process,TFor);
      End;
  if Tob_For <> nil then Tob_For.free;
  if Tob_Process <> nil then Tob_Process.free;
  End
else
  Ferme(Q);
End;    


Function  PGRecupProcess(TypProcess : String) : TOB ;
var
St : String;
Q : TQuery;
Begin
  Result := nil;
 { Recupération du paramétrage processus Absence type predefini Dossier }
  St := 'SELECT * FROM PGPROCESSUSDDE WHERE PPG_TYPPROCESSDDE="'+TypProcess+'" AND ##PPG_PREDEFINI## PPG_PREDEFINI="DOS"'+
        'ORDER BY PPG_ORDREPROCESS' ;
  Q := OpenSql(St,True);
  if not Q.eof then
    Begin
    Result := Tob.create('Le_process',nil,-1);
    Result.LoadDetailDB('Le_process','','',Q,False);
    End;
  Ferme(Q);
  { Recupération du paramétrage processus Absence type predefini Standard }
   if Result = nil then
      Begin
      St := 'SELECT * FROM PGPROCESSUSDDE WHERE PPG_TYPPROCESSDDE="'+TypProcess+'" AND ##PPG_PREDEFINI## PPG_PREDEFINI="STD"'+
        'ORDER BY PPG_ORDREPROCESS' ;
      Q := OpenSql(St,True);
      if not Q.eof then
         Begin
         Result := Tob.create('Le_process',nil,-1);
         Result.LoadDetailDB('Le_process','','',Q,False);
         End;
      Ferme(Q);
      End;
  { Recupération du paramétrage processus Absence type predefini Cegid }
   if Result = nil then
      Begin
      St := 'SELECT * FROM PGPROCESSUSDDE WHERE PPG_TYPPROCESSDDE="'+TypProcess+'" AND ##PPG_PREDEFINI## PPG_PREDEFINI="CEG"'+
        'ORDER BY PPG_ORDREPROCESS' ;
      Q := OpenSql(St,True);
      if not Q.eof then
         Begin
         Result := Tob.create('Le_process',nil,-1);
         Result.LoadDetailDB('Le_process','','',Q,False);
         End;
      Ferme(Q);
      End;
End;


procedure PGCreeProcess(TypProcess : String; LcTob_Process,LcTFille : Tob);  
Var
Tob_SuivDDe,TSDCol,TSD,TP : Tob;
i,cpt : integer;
St,ChpEtat,TbMaitre,Suf,sMod : string;
Begin
//PSN_TABLEMAITRE,PSN_GUIDTABLE,PSN_TYPPROCESSDDE,PSN_ORDRESUIVI
cpt := 0;

if TypProcess = 'ABS' then
  Begin
  TbMaitre := 'ABSENCESALARIE';
  ChpEtat := 'PCN_VALIDRESP';
  Suf := 'PCN';
  sMod := 'ABS';
  End
else
if TypProcess = 'FOR' then
  Begin
  TbMaitre := 'FORMATIONS';
  ChpEtat := 'PFO_ETATINSCFOR';
  Suf := 'PFO';
  sMod := 'FOR';
  End
else
if TypProcess = 'BUD' then
  Begin
  TbMaitre := 'INSCFORMATION';
  ChpEtat := 'PFI_ETATINSCFOR';
  Suf := 'PFI';
  sMod := 'FOR';
  End
  else
 exit;

Tob_SuivDDe := Tob.create('PGSUIVIDEMANDE',nil,-1);
{Affectation du demandeur }
TSDCol := PgCreatLignSuiviDde(Tob_SuivDDe,LcTFille,TypProcess,'SAL',sMod,LcTFille.GetValue(Suf+'_SALARIE'),TbMaitre,ChpEtat,Suf,cpt);  { PT2 Création d'une sous fonction }
{Affectation des référents }
For i := 0 to LcTob_Process.detail.count-1 do
  Begin
  TP := LcTob_Process.detail[i];
  { Dans le cas où l'on recherche l'administrateur }
  if TP.GetValue('PPG_ROLERH')='ADM' then
    Begin
    St:=PgRechReferentRHAdmin(TP.GetValue('PPG_ROLERH'),TP.GetValue('PPG_MODULERH'));  { PT2 Modif }
    end
  { Dans les autres cas }
  else
    Begin
    St:=PgRechReferentRH(LcTFille.GetValue(Suf+'_SALARIE'),TP);
    if St = '' then continue;
    End;
    Inc(cpt);
    TSD := PgCreatLignSuiviDde(Tob_SuivDDe,LcTFille,TypProcess,TP.GetValue('PPG_ROLERH'),TP.GetValue('PPG_MODULERH'),St,TbMaitre,ChpEtat,Suf,cpt);  { PT2 Création d'une sous fonction }
    if Assigned(TSD) then  { PT4 }
      Begin
      TSD.InsertDB(nil);
      TSD.free;
      end;
  End;

{ PT2 Si pas de référent alors affectation du référent administrateur }
if cpt = 0 then
  begin
    St:=PgRechReferentRHAdmin('ADM',sMod);
    cpt := 1;
    TSD := PgCreatLignSuiviDde(Tob_SuivDDe,LcTFille,TypProcess,'ADM',sMod,St,TbMaitre,ChpEtat,Suf,cpt);
    if Assigned(TSD) then  { PT4 }
      Begin
      TSD.InsertDB(nil);
      TSD.free;
      End;
  end;

{ Seulement s'il existe un référent rattaché }  
if (cpt > 0) and (Assigned(TSDCol)) then  { PT4 }
  Begin
  TSDCol.InsertDB(nil);
  TSDCol.free;
  End;
End;

{ PT2 Création d'une sous fonction }
Function PgCreatLignSuiviDde (Tob_SuivDDe,LcTFille:Tob; TypProcess,StRole,StMod,Ref,TbMaitre,ChpEtat,Suf : string; cpt : integer) : Tob;
var
  st : string;
begin
  st := LcTFille.GetValue(Suf+'_GUID');  { PT4 }
  if st = '' then begin Result :=nil; Exit; end;
  Result := Tob.create('PGSUIVIDEMANDE',Tob_SuivDDe,-1);
  Result.PutValue('PSN_TYPPROCESSDDE',TypProcess);
  Result.PutValue('PSN_REFERENTRH',Ref);
  Result.PutValue('PSN_INTERVENANTRH','');
  {$IFDEF EABSENCES}
  if (Cpt = 0) and (V_Pgi.UserSalarie<>Ref) then  { PT3 }
     Result.PutValue('PSN_INTERVENANTRH',V_Pgi.UserSalarie);  { PT3 }
  {$ENDIF}
  Result.PutValue('PSN_GUIDTABLE',LcTFille.GetValue(Suf+'_GUID'));
  Result.PutValue('PSN_TABLEMAITRE',TbMaitre);
  Result.PutValue('PSN_ORDRESUIVI',cpt);
  Result.PutValue('PSN_ROLERH',StRole);
  Result.PutValue('PSN_MODULERH',StMod);
  if LcTFille.GetValue(ChpEtat)='ATT' then
      Begin
      if StRole='SAL' then
        Result.PutValue('PSN_DATEINTERV',LcTFille.GetValue(Suf+'_DATECREATION'))
      else
        Result.PutValue('PSN_DATEINTERV',idate1900);
      Result.PutValue('PSN_ETATSUIVIDDE',LcTFille.GetValue(ChpEtat));
      Result.PutValue('PSN_STATUTSUIVIDDE','OUV');
      End
  Else
      Begin
      if StRole='SAL' then
        Result.PutValue('PSN_DATEINTERV',LcTFille.GetValue(Suf+'_DATECREATION'))
      else
        Result.PutValue('PSN_DATEINTERV',LcTFille.GetValue(Suf+'_DATEMODIF'));
      Result.PutValue('PSN_ETATSUIVIDDE',LcTFille.GetValue(ChpEtat));
      Result.PutValue('PSN_STATUTSUIVIDDE','CLO');
      End;
end;




function PgRechReferentRH(Sal : string; LcTP : Tob) : String;
Var
St,StNiv : string;
Q : TQuery;
Niv : integer;
Begin
Result := '';
Niv := LcTP.GetValue('PPG_NIVEAUH');
StNiv := IntToStr(Niv);
if LcTP.GetValue('PPG_TYPAFFPROCESS') = 'REF' then
  Begin
  Niv := Niv - 1;
  StNiv := IntToStr(Niv);
  if Niv=0 then { Responsable directe }
    St := 'SELECT DISTINCT PFH_REFERENTRH FROM PGAFFECTROLERH '+
        'WHERE PFH_TYPEAFFECTROLE="SAL" AND PFH_ROLERH="'+LcTP.GetValue('PPG_ROLERH')+'" '+
        'AND PFH_MODULERH="'+LcTP.GetValue('PPG_MODULERH')+'" '+
        'AND PFH_SALARIE="'+Sal+'"'
  else        { Responsable niveau supérieur }
    St := 'SELECT DISTINCT PFH_REFERENTRH FROM PGAFFECTROLERH '+
          'WHERE PFH_TYPEAFFECTROLE="SER" AND PFH_ROLERH="'+LcTP.GetValue('PPG_ROLERH')+'" '+
          'AND PFH_MODULERH="'+LcTP.GetValue('PPG_MODULERH')+'" '+
          'AND PFH_CODESERVICE=(SELECT PSO_SERVICESUP FROM SERVICEORDRE '+
                               'WHERE PSO_CODESERVICE=(SELECT PSE_CODESERVICE FROM DEPORTSAL '+
                                                      'WHERE PSE_SALARIE="'+Sal+'") '+
                               'AND PSO_NIVEAUSUP='+StNiv+')';
  End
else
  if LcTP.GetValue('PPG_TYPAFFPROCESS') = 'NIV' then
  Begin
  St := 'SELECT DISTINCT PFH_REFERENTRH FROM PGAFFECTROLERH '+
        'WHERE PFH_TYPEAFFECTROLE="SER" '+
        'AND PFH_ROLERH="'+LcTP.GetValue('PPG_ROLERH')+'" '+
        'AND PFH_MODULERH="'+LcTP.GetValue('PPG_MODULERH')+'" '+
        'AND PFH_CODESERVICE IN (SELECT PSO_CODESERVICE FROM SERVICEORDRE WHERE PSO_NIVEAUSERVICE='+StNiv+' '+
        'AND PSO_CODESERVICE IN (SELECT PSO_SERVICESUP FROM SERVICEORDRE WHERE PSO_CODESERVICE =(SELECT PSE_CODESERVICE FROM DEPORTSAL '+
                                                                                               'WHERE PSE_SALARIE="'+Sal+'")))';
  End
else
  if LcTP.GetValue('PPG_TYPAFFPROCESS') = 'SER' then
  Begin
  St := 'SELECT DISTINCT PFH_REFERENTRH FROM PGAFFECTROLERH '+
        'WHERE PFH_TYPEAFFECTROLE="SER" AND PFH_ROLERH="'+LcTP.GetValue('PPG_ROLERH')+'" '+
        'AND PFH_MODULERH="'+LcTP.GetValue('PPG_MODULERH')+'" '+
        'AND PFH_CODESERVICE="'+LcTP.GetValue('PPG_CODESERVICE')+'" ';
  End;
  if St <> '' then
    Begin
    Q := OpenSql(St,True);
    If not Q.eof then
      Result := Q.FindField('PFH_REFERENTRH').AsString;
    Ferme(Q);
    End;
End;


function  PgRechReferentRHAdmin(StRole,StMod : string) : String;  { PT2 Modif }
Var
St : String;
Q : TQuery;
Begin
Result := '';
St := 'SELECT DISTINCT PFH_REFERENTRH FROM PGAFFECTROLERH '+
      'WHERE PFH_ROLERH="'+StRole+'" '+		{ PT2 Modif }
      'AND PFH_MODULERH="'+StMod+'"';		{ PT2 Modif }
Q := OpenSql(St,True);
If not Q.eof then
  Begin
  if Q.RecordCount = 1 then  { Si pls adm exist alors on ne précise pas de matricule }
     Result := Q.FindField('PFH_REFERENTRH').AsString;
  End;
Ferme(Q);
End;


procedure PgSupprSuiviDemande (StTyp,StWhere : string);
Begin
if (StTyp='') or (StWhere = '') then exit;

if StTyp =  'BUD' then
   StWhere := 'WHERE PSN_GUIDTABLE IN '+
              '(SELECT PFI_GUID FROM INSCFORMATION '+
              'WHERE '+StWhere+')'
else
  if StTyp =  'FOR' then
     StWhere := 'WHERE PSN_GUIDTABLE IN '+
              '(SELECT PFO_GUID FROM FORMATIONS '+
              'WHERE '+StWhere+')';

ExecuteSql('DELETE FROM PGSUIVIDEMANDE '+StWhere);

End;

function IfExistSuiviDemande (StTyp,StWhere,StEtat : string) : Boolean ;
Begin
result := False;
if StTyp =  'BUD' then
   StWhere := 'WHERE PSN_GUIDTABLE IN '+
              '(SELECT PFI_GUID FROM INSCFORMATION '+
              'WHERE '+StWhere+') '+ StEtat
else
  if StTyp =  'FOR' then
     StWhere := 'WHERE PSN_GUIDTABLE IN '+
              '(SELECT PFO_GUID FROM FORMATIONS '+
              'WHERE '+StWhere+') '+ StEtat;

if ExisteSql('SELECT * FROM PGSUIVIDEMANDE '+StWhere) then
   result := True;

End;

function  PgMajSuiviDemande(TypProc,Guid,Mode,Sal,ValidPar,FicPrec,StRef,TBMaitre : String) : Boolean ;
Var
St : string;
Tot : integer;
Q   : TQuery;
Begin
result := False;
{ DEB PT10 }
if FicPrec = 'ADM' then StRef := ''
else
  if FicPrec <> '' then
  StRef := 'AND ((PSN_REFERENTRH="'+ValidPar+'") '+
           'OR (PSN_REFERENTRH IN (SELECT DISTINCT PFH_REFERENTRH FROM PGAFFECTROLERH, SERVICEORDRE '+
           'WHERE PFH_CODESERVICE=PSO_CODESERVICE AND PFH_ROLERH="RES" AND PFH_MODULERH="ABS" '+
           'AND PSO_SERVICESUP IN (SELECT PARH.PFH_CODESERVICE FROM PGAFFECTROLERH PARH '+
                                   'WHERE PARH.PFH_REFERENTRH="'+ValidPar+'" '+
                                   'AND PARH.PFH_ROLERH="RES" AND PARH.PFH_MODULERH="ABS" '+
                                   'AND PARH.PFH_TYPEAFFECTROLE="SER"))))';
{ FIN PT10 }
St := 'UPDATE PGSUIVIDEMANDE SET PSN_ETATSUIVIDDE="'+Mode+'",PSN_INTERVENANTRH="'+ValidPar+'", '+
      'PSN_DATEINTERV="'+USDateTime(Now)+'" '+
      'WHERE PSN_GUIDTABLE = "'+Guid+'" '+
      'AND PSN_TABLEMAITRE="'+TBMaitre+'" '+
      //'AND PSN_ETATSUIVIDDE="ATT" '+
      'AND PSN_ORDRESUIVI<>0 '+
      'AND PSN_TYPPROCESSDDE="'+TypProc+'" '+StRef;
Tot := ExecuteSql (st);
if (Mode<>'ATT') AND (Tot > 0) then { PT10 Ajout cond. Mode }
   Begin
   { ON VERIFIE S'IL Y A ENCORE UN SUIVI A TRAITER }
   St := 'SELECT PSN_ORDRESUIVI FROM PGSUIVIDEMANDE '+
      'WHERE PSN_GUIDTABLE = "'+Guid+'" '+
      'AND PSN_TABLEMAITRE="'+TBMaitre+'" '+
      'AND PSN_TYPPROCESSDDE="'+TypProc+'" AND PSN_ETATSUIVIDDE="ATT" '+
      'AND PSN_REFERENTRH<>"'+Sal+'"';
   Q := OpenSql(st,true);
   if not Q.eof then Tot := Q.RecordCount
   else Tot := 0;
   Ferme(Q);
   { Dans le cas où plus rien n'est à valider  }
   if (Tot = 0) then
     Begin
            { On clôture le processus }
     St := 'UPDATE PGSUIVIDEMANDE SET PSN_STATUTSUIVIDDE="CLO" '+
           'WHERE PSN_GUIDTABLE = "'+Guid+'" '+
           'AND PSN_TABLEMAITRE="'+TBMaitre+'" '+
           'AND PSN_TYPPROCESSDDE="'+TypProc+'"';
     ExecuteSql (st);
                { On met à jour le suivi de demande collaborateur }
     St := 'UPDATE PGSUIVIDEMANDE SET PSN_STATUTSUIVIDDE="CLO",PSN_ETATSUIVIDDE="'+Mode+'" '+
           'WHERE PSN_GUIDTABLE = "'+Guid+'" '+
           'AND PSN_TABLEMAITRE="'+TBMaitre+'" '+
           'AND PSN_ORDRESUIVI=0 '+
           'AND PSN_TYPPROCESSDDE="'+TypProc+'"';
     ExecuteSql (st);
                { On renvoie true pour la mise à jour des indicateurs de la table maitre  }
     result := true;
     End;
{ DEB PT10 Traitement inverse dans le cas d'une mise en attente de la demande }
   end
 else
   if (Mode='ATT') AND (Tot > 0) then
     Begin
     { On déclôture le processus }
     St := 'UPDATE PGSUIVIDEMANDE SET PSN_STATUTSUIVIDDE="OUV" '+
           'WHERE PSN_GUIDTABLE = "'+Guid+'" '+
           'AND PSN_TABLEMAITRE="'+TBMaitre+'" '+
           'AND PSN_TYPPROCESSDDE="'+TypProc+'"';
     ExecuteSql (st);
     { On met à jour le suivi de demande collaborateur }
     St := 'UPDATE PGSUIVIDEMANDE SET PSN_STATUTSUIVIDDE="OUV",PSN_ETATSUIVIDDE="'+Mode+'" '+
           'WHERE PSN_GUIDTABLE = "'+Guid+'" '+
           'AND PSN_TABLEMAITRE="'+TBMaitre+'" '+
           'AND PSN_ORDRESUIVI=0 '+
           'AND PSN_TYPPROCESSDDE="'+TypProc+'"';
     ExecuteSql (st);     
     { On renvoie true pour la mise à jour des indicateurs de la table maitre  }
     result := true;
     end;
{ FIN PT10 }
End;
{ DEB PT10 Création d'une fn qui renvoie le guid de l'absence }
function  PgGetGuidAbsence(TypMvt,Sal: string; Ordre : Integer) : String;
var st : string;
Q : TQuery;
Begin
st := 'SELECT PCN_GUID FROM ABSENCESALARIE '+
      'WHERE PCN_TYPEMVT="'+TypMvt+'" AND PCN_SALARIE="'+Sal+'" '+
      'AND PCN_ORDRE='+IntToStr(Ordre);
Q := OpenSql(st,true);
   if not Q.eof then Result := Q.Findfield('PCN_GUID').AsString
   else Result := '';
Ferme(Q);
end;
{ FIN PT10 }

procedure PgCreateAffectRoleRH(Creat : Boolean; Serv,Role,Module,Ref,AncRef :String);
Var
st : String;
Tob_Sal,TS,Tob_PgAffectRoleRh,TRH : Tob;
Q : Tquery;
i : integer;
Begin
if Creat then
  Begin
  { Création des affectations salariés }
  St := 'SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_CODESERVICE="'+Serv+'"';
  Q := OpenSql(St,True);
  if not Q.eof then
    Begin
    Tob_Sal :=  Tob.create('Les_salaries',nil,-1);
    Tob_Sal.LoadDetailDb('Les_salaries','','',Q,False);
    Ferme(Q);
    Tob_PgAffectRoleRh :=  TOB.Create('PGAFFECTROLERH',nil,-1);
    For i:=0 to Tob_Sal.detail.count-1 do
      Begin
      TS :=  Tob_Sal.Detail[i];
      TRH := TOB.Create('PGAFFECTROLERH' ,Tob_PgAffectRoleRh,-1);
      TRH.PutValue('PFH_TYPEAFFECTROLE'  ,'SAL');
      TRH.PutValue('PFH_ROLERH'          ,Role);
      TRH.PutValue('PFH_REFERENTRH'      ,Ref);
      TRH.PutValue('PFH_MODULERH'        ,Module);
      TRH.PutValue('PFH_CODESERVICE'     ,Serv);
      TRH.PutValue('PFH_SALARIE'         ,TS.GetValue('PSE_SALARIE'));
      TRH.InsertDB(nil);
      End;
    Tob_Sal.free;
    Tob_PgAffectRoleRh.free;
    End
  else
    Ferme(Q);
  End
else
    Begin
    ExecuteSql('UPDATE PGAFFECTROLERH SET PFH_REFERENTRH="'+Ref+'" '+
               'WHERE PFH_TYPEAFFECTROLE="SAL" '+
               'AND PFH_ROLERH="'+Role+'" '+
               'AND PFH_MODULERH="'+Module+'" '+
               'AND PFH_REFERENTRH="'+AncRef+'" '+
               'AND PFH_CODESERVICE="'+Serv+'"');
    End;



End;

procedure PgInitTabletteTypeDesti;
var
Q : TQuery;
Begin
{ DEB PT11 Reconstitution des éléments du module AUG écrasant les élements du module ABS } 
ExecuteSql('DELETE FROM COMMUN WHERE CO_TYPE="PG1" '+
           'AND CO_CODE<>"001" AND CO_CODE<>"002" AND CO_CODE<>"003" '+
           'AND CO_CODE<>"004" AND CO_CODE<>"005" AND CO_LIBELLE NOT LIKE "%Administrateur%"');
{ FIN PT11 } 
Q := OpenSql('SELECT * FROM HIERARCHIE',True);
While not Q.eof do
  Begin
  PgMajTabletteTypeDesti('CREATION',Q.FindField('PHO_NIVEAUH').AsString);
  Q.next;
  end;
Ferme(Q); 
end;



procedure PgMajTabletteTypeDesti(Mode, Niv : string ) ;
Var
Q : TQuery;
Begin
{ Parcours des modules }
Q := OpenSql('SELECT * FROM COMMUN WHERE CO_TYPE="POH"',True);
if Not Q.eof then
      Begin
      Q.First;
      While not Q.eof do
        Begin
        if Mode = 'CREATION' then
          PGGenereTypDestModule(niv,Q.FindField('CO_CODE').AsString,Q.FindField('CO_LIBRE').AsString)
        else
          PGSupprTypDestModule(niv,Q.FindField('CO_CODE').AsString);
        Q.Next;
        End;
      Avertirtable('PGALERTGPRDESTI');  
      End;
Ferme(Q);
End;

procedure PGGenereTypDestModule ( Niv,SMod,actif : string ) ;
Var
TTypDesti,T1 : Tob;
Q : TQuery;
Begin
//if niv = '1' then exit; { PT11 On regénère le niveau 1 }
TTypDesti := TOB.Create('COMMUN',nil,-1);
Q := OpenSql('SELECT * FROM COMMUN WHERE CO_TYPE="PRH" AND CO_CODE<>"ADM" ' ,True); { exclut les administrateurs }
if Not Q.eof then
      Begin
      Q.First;
      While not Q.eof do
        Begin
        { Premier caractère correspond au rôle }
        { Second caractère correspond au module }
        { Troisième caractère correspond au niveau }
        T1 := TOB.Create('COMMUN',TTypDesti,-1);
        T1.PutValue('CO_TYPE','PG1');
        T1.PutValue('CO_CODE',Q.FindField('CO_ABREGE').AsString + Copy(PgGetPrefModuleRH(sMod),1,1) + Niv);  { PT11 }
        T1.PutValue('CO_LIBELLE',Q.FindField('CO_LIBELLE').AsString+' '+RechDom('PGMODULERH',sMod,False)+' N+'+Niv);
        T1.PutValue('CO_ABREGE',Actif);
        T1.PutValue('CO_LIBRE',sMod);
        T1.InsertOrUpdateDB;
        Q.Next;
        End;
      End;
TTypDesti.free;      
Ferme(Q);
End;


procedure PGSupprTypDestModule ( Niv,SMod : string ) ;
Var
Q : TQuery;
St : string;
Begin
if niv = '1' then exit;
Q := OpenSql('SELECT * FROM COMMUN WHERE CO_TYPE="PRH"',True);
if Not Q.eof then
      Begin
      Q.First;
      While not Q.eof do
        Begin
        { Premier caractère correspond au rôle }
        { Second caractère correspond au module }
        { Troisième caractère correspond au niveau }
        St :=Q.FindField('CO_ABREGE').AsString + Copy(PgGetPrefModuleRH(sMod),1,1) + Niv; { PT11 }
        ExecuteSql('DELETE FROM COMMUN WHERE CO_TYPE="PG1" AND CO_CODE="'+St+'"');
        Q.Next;
        End;
      End;
Ferme(Q);
end;
{ DEB PT11 Utilisation d'un autre code pour le module AUG } 
function PgGetPrefModuleRH(Smod : String) : string;
Begin
result := smod;
if result = 'AUG' then result := 'MAU';
End;
{ FIN PT11 } 

end.



