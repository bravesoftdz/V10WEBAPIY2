unit MajHalleyUtil;

interface

Uses  DB,UtilOutils;


procedure MoveCommunToChoixcode (CoTyp,CCTyp : string);
procedure MoveChoixcodeToChoixExt (CcTyp,YxTyp : string);
procedure AddParamGC ( TSoc,TRef : TDataSet ) ;
Procedure AddNewNaturesGC ;
procedure InsertChoixCode (stType,stCode,stLibelle,stAbrege,stlibre : string);
procedure InsertChoixExt (stType,stCode,stLibelle,stAbrege,stlibre : string);
Function  TableToVersion ( sNomTable : String ) : Integer ;
Procedure TestDropTable ( NomTable : String ) ;
Procedure MAJDropTable ( NomTable : String ) ;
Function  IncoherTable ( NomTable,PrefOK : String ) : boolean ;
Procedure MajDropMauvaisesVues ;
Procedure DropVuePourrie ( NomVue : String ) ;
Function  IsMonoOuCommune : boolean;
Procedure UpdateChoixCodeLibre(sType, sCode, sLibre : String);
Procedure MAJListeProduits;
Procedure ForceConfidentialiteMenu;
Procedure SupprimeEtat( TypeEtat,NatureEtat,CodeEtat : string; VerifSiModele : boolean = False );
procedure MajSouche;
procedure PrepareBasePCL;
{$IFDEF MAJPCL}
procedure OptimiseDossierPCL;
procedure UpdateDomaine (stNomTable, stDomaine : string);
{$ENDIF}

function DupliqueFiltre (pstNomFiltre, pstNomFiltreDest : string) : integer;
function DupliqueFiltreRef (pstNomFiltre, pstNomFiltreDest : string) : integer;
procedure ForceEmptyParamSoc(nom: string; value: string);
procedure AGLNettoieFiltre( vStNomFiltre : string; vStOldZones : string; vStNewZones : String; vBoCtxStd : Boolean = False);

//js1 100706 log spécifique à PCL
procedure LogPCL(St,Fichier:string);
function FicLogPCL : string;
//js1 100706 shrink PCL
function OkShrinkPCL : boolean;
function GetDb0name : string;
function isVerrou : boolean;
function GetUser : string;

//js1 24102007 maj du journal d'evenement ==> a terme a déplacer de utilgc vers utilpgi 
function MAJJnalEvent(TypeEvt, Etat, Libelle, BlocNote : string) : integer;
function ExecuteSQLNoPCL(const Sql: WideString): Integer;

implementation

Uses HEnt1,HCtrls,MajTable,StdCtrls,ComCtrls,Forms,
    {$IFNDEF EAGLCLIENT}
      {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
    {$ENDIF}
     MajStruc,HQry,SysUtils,Classes,
     DBCtrls,ParamSoc, Dialogs,UTOB,Ent1,hMsgBox,utilpgi,hdebug ;

{$IFDEF PASUTILISE}
function SurBaseCommune: Boolean;
// True si on est en MULTI-dossier et sur la base commune
begin
  Result := ((V_PGI_Env<>Nil) and (V_PGI_Env.ModeFonc='MULTI') and (V_PGI_Env.SocCommune=V_PGI.TempAlias));
end;

function EnMulti : Boolean;
// True si on est en MULTI-dossier et sur la base commune
begin
  Result := ((V_PGI_Env<>Nil) and (V_PGI_Env.ModeFonc='MULTI') );
end;
{$ENDIF}

procedure ForceEmptyParamSoc(nom: string; value: string);
var
  q: tquery;
  s: string;
begin

{ Pour forcer une valeur de PARAMSOC alors que ce PARAMSOC n'existe pas encore.
 Ce PARAMSOC doit imérativement être créé dans le même traitement
 Exemple : ajout d'un paramsoc pour gerer la mise à jour du DICO
          mais on n'a pas la main entre la mise à jour des PARAMSOC et celle du DICO
          donc valeur forcée dans MajAvant pour anticuper sur le traitement
}

  Q := OpenSQL('SELECT SOC_NOM, SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "' + nom + '"', FALSE);
  if not Q.EOF then
  begin
    s := q.fields[1].asString;
    if (s = '') then
    begin
      q.edit();
      q.fields[1].asString := value;
      q.post();
    end;
  end
  else
  begin
    q.insert();
    q.fields[0].asString := nom;
    q.fields[1].asString := value;
    q.post();
  end;

  ferme(Q);
end;

procedure MoveCommunToChoixcode (CoTyp,CCTyp : string);
Var Q,QCC : TQuery ;
BEGIN
Q:=OpenSQL('SELECT * FROM COMMUN WHERE CO_TYPE="'+CoTyp+'"',True) ;
QCC:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+CCTyp+'"',False) ;
if QCC.eof then
begin
  While Not Q.EOF do
     BEGIN
     QCC.Insert ; InitNew(QCC) ;
     QCC.FindField('CC_TYPE').AsString:=CCTyp ;
     QCC.FindField('CC_CODE').AsString:=Q.FindField('CO_CODE').AsString ;
     QCC.FindField('CC_LIBELLE').AsString:=Q.FindField('CO_LIBELLE').AsString ;
     QCC.FindField('CC_ABREGE').AsString:=Q.FindField('CO_ABREGE').AsString ;
     QCC.FindField('CC_LIBRE').AsString:=Q.FindField('CO_LIBRE').AsString ;
     QCC.Post ;
     Q.Next ;
     END ;
end;
Ferme(Q) ;
Ferme(QCC) ;
// if not EnMulti then executeSQL('DELETE FROM COMMUN WHERE CO_TYPE="'+CoTyp+'"') ;  // pas supprimer si dossier commun sinon pas de copie sur les dossiers
END ;

procedure MoveChoixcodeToChoixExt (CcTyp,YxTyp : string);
Var QCC,QYX : TQuery ;
BEGIN
QCC:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+CcTyp+'"',True) ;
QYX:=OpenSQL('SELECT * FROM CHOIXEXT WHERE YX_TYPE="'+YxTyp+'"',False) ;
if QYX.eof then
begin
  While Not QCC.EOF do
     BEGIN
     QYX.Insert ; InitNew(QYX) ;
     QYX.FindField('YX_TYPE').AsString:=YxTyp ;
     QYX.FindField('YX_CODE').AsString:=QCC.FindField('CC_CODE').AsString ;
     QYX.FindField('YX_LIBELLE').AsString:=QCC.FindField('CC_LIBELLE').AsString ;
     QYX.FindField('YX_ABREGE').AsString:=QCC.FindField('CC_ABREGE').AsString ;
     QYX.FindField('YX_LIBRE').AsString:=QCC.FindField('CC_LIBRE').AsString ;
     QYX.Post ;
     QCC.Next ;
     END ;
end;
Ferme(QCC) ;
Ferme(QYX) ;
executeSQL('DELETE  FROM CHOIXCOD WHERE CC_TYPE="'+CcTyp+'"') ;
END ;


procedure InsertChoixCode (stType,stCode,stLibelle,stAbrege,stlibre : string);
var Q : TQuery;
begin
Q:=OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+stType+'" AND CC_CODE="'+stCode+'"',FALSE) ;
If Q.EOF Then
  BEGIN
  Q.Insert ;
  Q.FindField('CC_TYPE').AsString:=stType ;
  Q.FindField('CC_CODE').AsString:=stCode ;
  Q.FindField('CC_LIBELLE').AsString:=stLibelle ;
  Q.FindField('CC_ABREGE').AsString:=stAbrege ;
  Q.FindField('CC_LIBRE').AsString:=stLibre ;
  Q.Post ;
  END ;
Ferme(Q) ;
end;

procedure InsertChoixExt (stType,stCode,stLibelle,stAbrege,stlibre : string);
var Q : TQuery;
begin
Q:=OpenSQL('SELECT * FROM CHOIXEXT WHERE YX_TYPE="'+stType+'" AND YX_CODE="'+stCode+'"',FALSE) ;
If Q.EOF Then
  BEGIN
  Q.Insert ;
  Q.FindField('YX_TYPE').AsString:=stType ;
  Q.FindField('YX_CODE').AsString:=stCode ;
  Q.FindField('YX_LIBELLE').AsString:=stLibelle ;
  Q.FindField('YX_ABREGE').AsString:=stAbrege ;
  Q.FindField('YX_LIBRE').AsString:=stLibre ;
  Q.Post ;
  END ;
Ferme(Q) ;
end;

procedure AddParamGC ( TSoc,TRef : TDataSet ) ;
var i : Integer ;
    TT : TField ;
begin
TSoc.Insert ;
For i:=0 to TSoc.FieldCount-1 do
    BEGIN
    TT:=TRef.FindField(TSoc.Fields[i].Fieldname) ;
    if TT<>NIL then TSoc.Fields[i].Value:=TT.Value ;
    END ;
TSoc.Post ;
end ;

Procedure AddNewNaturesGC ;
Var TSoc,TRef : THTable ;
    NatRef,SoucheRef,TypeSoucheRef : String ;
    ListeNat,ListeSouche : TStrings ;
BEGIN
TSoc:=THTable.Create(Application) ; TSoc.DatabaseName:=DBSoc.DatabaseName ;
ListeNat:=TStringList.Create ; ListeSouche:=TStringList.Create ;
{Copie PARPIECE}
TSoc.Tablename:='PARPIECE' ; TSoc.Indexname:='GPP_CLE1' ; TSoc.Open ;
TRef:=OpenTableRef('PARPIECE','GPP_CLE1') ;
While Not TRef.EOF do
   BEGIN
   if TRef.FindField('GPP_NIVEAUPARAM').AsString='EDI' then
      BEGIN
      NatRef:=TRef.FindField('GPP_NATUREPIECEG').AsString ; ListeNat.Add(NatRef) ;
      SoucheRef:=TRef.FindField('GPP_SOUCHE').AsString ;
      if Not TSoc.Findkey([NatRef]) then BEGIN AddParamGC(TSoc,TRef) ; ListeSouche.Add(SoucheRef) ; END ;
      END ;
   TRef.Next ;
   END ;
Ferme(TRef) ; TSoc.Close ;
{Copie PARPIECEGRP}
TSoc.Tablename:='PARPIECEGRP' ; TSoc.Indexname:='GPR_CLE1' ; TSoc.Open ;
TRef:=OpenTableRef('PARPIECEGRP','GPR_CLE1') ;
While Not TRef.EOF do
   BEGIN
   NatRef:=TRef.FindField('GPR_NATUREPIECEG').AsString ;
   if ListeNat.IndexOf(NatRef)>=0 then
      BEGIN
      if Not TSoc.Findkey([NatRef,TRef.FindField('GPR_NATPIECEGRP').AsString]) then AddParamGC(TSoc,TRef) ;
      END ;
   TRef.Next ;
   END ;
Ferme(TRef) ; TSoc.Close ;
{Copie SOUCHE}
TSoc.Tablename:='SOUCHE' ; TSoc.Indexname:='SH_CLE1' ; TSoc.Open ;
TRef:=OpenTableRef('SOUCHE','SH_CLE1') ;
While Not TRef.EOF do
   BEGIN
   TypeSoucheRef:=TRef.FindField('SH_TYPE').AsString ;
   SoucheRef:=TRef.FindField('SH_SOUCHE').AsString ;
   if ((TypeSoucheRef='GES') and (ListeSouche.IndexOf(SoucheRef)>=0)) then
      BEGIN
      if Not TSoc.Findkey([TypeSoucheRef,SoucheRef]) then AddParamGC(TSoc,TRef) ;
      END ;
   TRef.Next ;
   END ;
Ferme(TRef) ; TSoc.Close ;
{Libérations}
TSoc.Free ;
ListeNat.Clear ; ListeNat.Free ;
ListeSouche.Clear ; ListeSouche.Free ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : JCF
Créé le ...... : 02/08/2001
Modifié le ... :   /  /
Description .. : Retourne le numéro de version d'un table
Mots clefs ... : TABLE;VERSION
*****************************************************************}
Function TableToVersion ( sNomTable : String ) : Integer ;
Var iTable   : Integer ;
    sPrefixe : String ;
BEGIN
sPrefixe := TableToPrefixe(sNomTable) ;
iTable := PrefixeToNum(sPrefixe);
Result := V_PGI.DETables[iTable].NumVersion ;
END ;

Procedure MAJDropTable ( NomTable : String ) ;
Var Pref : String ;
    OkDel : Boolean ;
BEGIN
Pref:=TableToPrefixe(NomTable) ; OkDel:=true ;
 Try
  ExecuteSQL('DROP TABLE '+NomTable) ;
 Except
  OkDel:=False ;
 End ;
if ((OkDel) and (Pref<>'')) then
   BEGIN
   ExecuteSQL('DELETE FROM DETABLES WHERE DT_PREFIXE="'+Pref+'"') ;
   ExecuteSQL('DELETE FROM DECHAMPS WHERE DH_PREFIXE="'+Pref+'"') ;
   END ;
END ;

Procedure MAJDropTableOptim ( NomTable : String; bPhysique : boolean ) ;
Var Pref : String ;
    OkDel : Boolean ;
BEGIN
Pref:=TableToPrefixe(NomTable) ; OkDel:=true ;
 Try
  if bPhysique then ExecuteSQL('DROP TABLE '+NomTable) ;
 Except
  OkDel:=False ;
 End ;
if ((OkDel) and (Pref<>'')) then
   BEGIN
   ExecuteSQL('DELETE FROM DETABLES WHERE DT_PREFIXE="'+Pref+'"') ;
   ExecuteSQL('DELETE FROM DECHAMPS WHERE DH_PREFIXE="'+Pref+'"') ;
   ExecuteSQL('DELETE FROM DESHARE WHERE DS_NOMTABLE="'+NomTable+'"');
   END ;
END ;

Procedure TestDropTable ( NomTable : String ) ;
BEGIN
if TableExiste(NomTable) then MAJDroptable(NomTable) ;
END ;

Procedure TestDropTableOptim ( NomTable : String ) ;
BEGIN
if TableExiste(NomTable) then MAJDroptableOptim(NomTable, True)
else MAJDroptableOptim(NomTable, False);
END ;

Function IncoherTable ( NomTable,PrefOK : String ) : boolean ;
Var PrefLog,PrefPhy,NomChamp : String ;
    TT      : HTStrings ;
    QQ      : TQuery ;
BEGIN
Result:=False ;
if Not TableExiste(NomTable) then Exit ;
PrefLog:=TableToPrefixe(NomTable) ; if PrefLog='' then Exit ;
if PrefLog<>prefOK then BEGIN Result:=True ; Exit ; END ;
TT:=HTStringList.Create ;
QQ:=PrepareSQL('SELECT * FROM '+NomTable,TRUE) ; QQ.GetFieldnames(TT) ;
Ferme(QQ) ;
if TT.Count>0 then
   BEGIN
   NomChamp:=TT[0] ; PrefPhy:=ExtractPrefixe(NomChamp) ;
   if PrefPhy<>PrefLog then Result:=True ;
   END ;
TT.Clear ; TT.Free ;
END ;

Procedure MajDropMauvaisesVues ;
Var LPHY,LLOG : HTStrings ;
    StDataBaseName : String ;
    QQ             : TQuery ;
    i,Ind          : integer ;
    NomT,NomVue    : String ;
BEGIN
LPHY:=HTStringList.Create ;
LLOG:=HTStringList.Create ;
StDatabaseName:=DBSOC.DataBaseName ;
Session.GetTableNames(StDatabaseName,'*.*',False,False,LPHY) ;
QQ:=OpenSQL('SELECT DT_NOMTABLE FROM DETABLES',True) ;
While Not QQ.EOF do
   BEGIN
   LLOG.Add(uppercase(QQ.FIELDS[0].AsString)) ;
   QQ.Next ;
   END ;
Ferme(QQ) ;
for i:=0 to LPHY.Count-1 do LPHY[i]:=uppercase(VireDBO(LPHY[i])) ;
for i:=0 to LLOG.Count-1 do
    BEGIN
    NomT:=LLOG[i] ;
    ind:=LPHY.IndexOf(NomT) ;
    if Ind>0 then LPHY.Delete(ind) ;
    END ;
for i:=0 to LPHY.Count-1 do
    BEGIN
    NomVue:=LPHY[i] ;
    if Copy(NomVue,1,3)='SYS' then Continue ;
     Try
      DBDeleteView(DBSOC,V_PGI.Driver,NomVue) ;
     Except
     End ;
    END ;
LPHY.Free ; LLOG.Free ;
END ;

Procedure DropVuePourrie ( NomVue : String ) ;
BEGIN
 Try
  DBDeleteView(DBSOC,V_PGI.Driver,NomVue) ;
 Except
 End ;
END ;

Function IsMonoOuCommune : boolean;
begin
  // test base commune existe, mode mono ou multi
  // Result := ((V_PGI.InBaseCommune) or (V_PGI.DefaultSectionName='')); // je suis sur la base commune en direct ou non
  // suite à bug qui fait que  V_PGI.InBaseCommune retourne toujours True PCS 28/09/2004
  Result := ((V_PGI.DbName=V_PGI.DefaultSectionDbName) or (V_PGI.DefaultSectionName='')); // je suis sur la base commune en direct ou non

end;

Procedure UpdateChoixCodeLibre(sType, sCode, sLibre : String);
begin
  if ExisteSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+sType+'" AND CC_CODE="'+sCode+'"')then
    ExecuteSQL('UPDATE CHOIXCOD SET CC_LIBRE="'+ sLibre + '" WHERE CC_CODE="'+sCode+'" AND CC_TYPE="'+sType+'"');
end;

Procedure ForceConfidentialiteMenu;
// Force confidentialité à masqher les nouveaux points d'entrés.
var stAccesGrp : string;
    SavEnableDeShare:boolean; //js1 pour c ayel 260606 on active le deshare
Begin
// OK en PCL à partir de 657 (V6.00) PCS 31/08/2004 If V_PGI.ModePCL='1' then Exit;
  SavEnableDeShare := V_PGI.enableDEShare;
  V_PGI.enableDEShare := True;

  if TableCount('USERGRP','',True) < 3 then exit;
  stAccesGrp:=StringOfChar('-',100);
  ExecuteSQL('UPDATE MENU SET MN_ACCESGRP="'+stAccesGrp+'" WHERE MN_VERSIONDEV="X" AND MN_1<>0');

  V_PGI.enableDEShare := SavEnableDeShare;
End;

procedure MAJListeProduits;
var
  stTable,stIndex : string;
  TSoc, TRef: THTable;
begin
  stTable:='YLISTEPRODUITS';
  stIndex:='YLP_CLE1';
  // suppression des elemenents existants
  ExecuteSQL('DELETE FROM '+stTable);
  TSoc := THTable.Create(Application);
  TSoc.DatabaseName := DBSoc.DatabaseName;
  TSoc.Tablename := stTable;
  TSoc.Indexname := stIndex;
  TSoc.Open;
  TRef := OpenTableRef(stTable, stIndex);
  while not TRef.EOF do
  begin
    AddParamGC(TSoc, TRef);
    TRef.Next;
  end;
  Ferme(TRef);
  TSoc.Close;
  TSoc.Free;
end;


Procedure SupprimeEtat( TypeEtat,NatureEtat,CodeEtat : string; VerifSiModele : boolean = False  );
Begin
  if VerifSiModele and ( not ExisteSQL('select MO_CODE FROM MODELES WHERE MO_TYPE="'+TypeEtat+'" AND MO_NATURE="'+NatureEtat+'" AND MO_CODE="'+CodeEtat+'" AND MO_MODELE="X"') )
     then exit;
  ExecuteSQL('DELETE FROM MODELES WHERE MO_TYPE="'+TypeEtat+'" AND MO_NATURE="'+NatureEtat+'" AND MO_CODE="'+CodeEtat+'"');
  ExecuteSQL('DELETE FROM MODEDATA WHERE MD_CLE LIKE "'+TypeEtat+NatureEtat+CodeEtat+'%"');
End;

procedure MajSouche;
var TSoc, TRef: THTable;
  TypeSoucheRef,SoucheRef : string;
begin
  TSoc := THTable.Create(Application);
  TSoc.DatabaseName := DBSoc.DatabaseName;
  TSoc.Tablename := 'SOUCHE';
  TSoc.Indexname := 'SH_CLE1';
  TSoc.Open;
  TRef := OpenTableREF('SOUCHE', 'SH_CLE1');
  while not TRef.EOF do
  begin
    TypeSoucheRef:=TRef.FindField('SH_TYPE').AsString ;
    SoucheRef:=TRef.FindField('SH_SOUCHE').AsString ;
    If not TSoc.Findkey([TypeSoucheRef,SoucheRef]) then
      AddParamGc(TSoc,TRef);
    TRef.Next;
  end;
  Ferme(TRef);
  TSoc.Close;
  TSoc.Free;
end;

procedure PrepareBasePCL;
var QQ: TQuery;
    stMenu,stCommun,stDomaineVue : string;
Begin

  // vues
  stDomaineVue:='("A","0","B","E","G","H","O","Q","R","U","W","X")';
  QQ:=OpenSQL('SELECT DV_NOMVUE FROM DEVUES where DV_DOMAINE in '+stDomaineVue,True) ;
  While Not QQ.EOF do
     BEGIN
     DropVuePourrie(QQ.FIELDS[0].AsString) ;
     QQ.Next ;
     END ;
  Ferme(QQ) ;
  ExecuteSQL('DELETE FROM DEVUES where DV_DOMAINE in '+stDomaineVue);
  // autres vues
  DropVuePourrie('YYAGENDA_MUL') ;
  ExecuteSQL('DELETE FROM DEVUES where DV_NOMVUE = "YYAGENDA_MUL"');


  //Tables
  QQ:=OpenSQL('SELECT DT_NOMTABLE FROM DETABLES where DT_DOMAINE in ("A","0","B","E","G","H","O","Q","R","U","W","X")',True) ;
  While Not QQ.EOF do
     BEGIN
     TestDropTable(QQ.FIELDS[0].AsString) ;
     QQ.Next ;
     END ;
  Ferme(QQ) ;

  //Menus
  stMenu:='(30,31,32,33,34,36,60,65,260,266,267,279,280,281)';
  executeSQl('delete from menu where mn_1=0 and mn_2 in '+stMenu );
  executeSQl('delete from menu where mn_1 in '+stMenu);

  //liste
  executeSQl('delete from Liste where LI_LISTE like "GC%"' );
  executeSQl('delete from Liste where LI_LISTE like "BT%"' );
  executeSQl('delete from Liste where LI_LISTE like "EB%"' );
  executeSQl('delete from Liste where LI_LISTE like "MO%"' );
  executeSQl('delete from Liste where LI_LISTE like "FO%"' );
  executeSQl('delete from Liste where LI_LISTE like "HR%"' );
  executeSQl('delete from Liste where LI_LISTE like "II%"' );
  executeSQl('delete from Liste where LI_LISTE like "QU%"' );
  executeSQl('delete from Liste where LI_LISTE like "RT%"' );
  executeSQl('delete from Liste where LI_LISTE like "UM%"' );
  executeSQl('delete from Liste where LI_LISTE like "WP%"' );
  executeSQl('delete from Liste where LI_LISTE like "X%"' );

  //Etats
  executeSQl('truncate table modeles ' );
  executeSQl('truncate table modedata ' );
  //traduction
  executeSQl('truncate table TRADDICO ' );

  //Commun
  stCommun:='("0","B","E","G","H","O","Q","R","U","W","X")';
  executeSQl('delete from commun where co_type in  (select do_type from decombos  where do_domaine in '+stCommun+')');

  PGIInfo('Traitement terminé');
end;

{$IFDEF MAJPCL}

procedure UpdateDomaine (stNomTable, stDomaine : string);
begin
  ExecuteSQL ('UPDATE DETABLES SET DT_DOMAINE="'+stDomaine+'" WHERE DT_NOMTABLE="'+stNomTable+'"');
end;

procedure OptimiseDossierPCL;
var QQ: TQuery;
    stMenu,stCommun,stDomaineVue : string;
    SaveEnableDEShare : boolean;
Begin
  InitPGIpourDossierPCL;

  // force à travailler sans DESHARE
  SaveEnableDEShare := v_pgi.enableDEShare;
  v_pgi.enableDEShare := False;

  // vues
  stDomaineVue:='('+V_PGI.PCLDomainesExclus+')';
(*
On ne fait pas la destruction physique car pour l'instant DBDDeleteAllView
  QQ:=OpenSQL('SELECT DV_NOMVUE FROM DEVUES where DV_DOMAINE in '+stDomaineVue,True) ;
  While Not QQ.EOF do
     BEGIN
     DropVuePourrie(QQ.FIELDS[0].AsString) ;
     QQ.Next ;
     END ;
  Ferme(QQ) ;
  *)
  ExecuteSQL('DELETE FROM DEVUES where DV_DOMAINE in '+stDomaineVue);

  // MD le 19/01/07
  // YYAGENDA_MUL = left join sur table affaire
  ExecuteSQL('DELETE FROM DEVUES where DV_NOMVUE="YYAGENDA_MUL"');


  //Tables : suppression sauf table AFFAIRE
  QQ:=OpenSQL('SELECT DT_NOMTABLE FROM DETABLES where DT_DOMAINE in ('+V_PGI.PCLDomainesExclus+') AND DT_NOMTABLE<>"AFFAIRE"',True) ;
  While Not QQ.EOF do
     BEGIN
     TestDropTableOptim(QQ.FIELDS[0].AsString) ;
     QQ.Next ;
     END ;
  Ferme(QQ) ;

  // Suppression des tables en dehors du domaine PCL et supprimées depuis la 663
  TestDropTableOptim('QGRPOFCAR');
  TestDropTableOptim('QPROFILUSERGRP');
  TestDropTableOptim('WDEPOTAPP');
  TestDropTableOptim('WORDREAUTO');


  //Menus

  stMenu:='('+V_PGI.PCLModulesInclus+')';
  executeSQl('delete from menu where mn_1=0 and mn_2 not in '+stMenu );
  executeSQl('delete from menu where mn_1<>0 and mn_1 not in '+stMenu);

  //liste
  // si on les supprime elles vont être crées à la prochaine mise à jour
 { executeSQl('delete from Liste where LI_LISTE like "GC%"' );
  executeSQl('delete from Liste where LI_LISTE like "BT%"' );
  executeSQl('delete from Liste where LI_LISTE like "EB%"' );
  executeSQl('delete from Liste where LI_LISTE like "MO%"' );
  executeSQl('delete from Liste where LI_LISTE like "FO%"' );
  executeSQl('delete from Liste where LI_LISTE like "HR%"' );
  executeSQl('delete from Liste where LI_LISTE like "II%"' );
  executeSQl('delete from Liste where LI_LISTE like "QU%"' );
  executeSQl('delete from Liste where LI_LISTE like "RT%"' );
  executeSQl('delete from Liste where LI_LISTE like "UM%"' );
  executeSQl('delete from Liste where LI_LISTE like "WP%"' );
  executeSQl('delete from Liste where LI_LISTE like "X%"' ); }

  executeSQl('delete from paramsoc where NOT ('+V_PGI.PCLParamSocInclus+')');

  // On supprime la table AFFAIRE 
  TestDropTableOptim('AFFAIRE') ;
  //
  v_pgi.enableDEShare := SaveEnableDEShare;


end;

{$ENDIF}


 

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 07/12/2005
Modifié le ... : 19/12/2005
Description .. : Cette fonction permet de dupliquer un filtre.
Suite ........ : Retour : nombre d'enregistrements traités (-1 si filtre de destination existe déjà)
Suite ........ : CA - 19/12/2005 : Test existence filtre de départ sinon pb INSERT sous ORACLE
Mots clefs ... :
****************************************************************}

function DupliqueFiltre (pstNomFiltre, pstNomFiltreDest : string) : integer;
var
 stSQLInsert, stSQLSelect : string;
begin
  if not ExisteSQL ('SELECT FI_TABLE FROM FILTRES WHERE FI_TABLE="'+pstNomFiltreDest+'"') then
  begin
    { On teste l'existence du filtre initial pour éviter problème INSERT NULL sous ORACLE }
    if ExisteSQL ('SELECT FI_TABLE FROM FILTRES WHERE FI_TABLE="'+pstNomFiltre+'"') then
    begin
      stSQLInsert := 'INSERT INTO FILTRES (FI_TABLE,FI_LIBELLE,FI_CRITERES,FI_CREATEUR,FI_DATECREATION,FI_DATEMODIF,FI_UTILISATEUR) ';
      stSQLSelect := 'SELECT "'+pstNomFiltreDest+'", FI_LIBELLE,FI_CRITERES,'
                        + 'FI_CREATEUR,FI_DATECREATION,FI_DATEMODIF,FI_UTILISATEUR ';
      stSQLSelect := stSQLSelect + 'FROM FILTRES WHERE FI_TABLE="'+pstNomFiltre+'"';
      Result := ExecuteSQL ( stSQLInsert + stSQLSelect );
    end else Result := -2;
  end else Result := -1;
end;

function DupliqueFiltreRef (pstNomFiltre, pstNomFiltreDest : string) : integer;
var
  stSQLInsert, stSQLSelect : string;
begin
  if not ExisteSQL ('SELECT FIR_TABLE FROM FILTRESREF WHERE FIR_TABLE="'+pstNomFiltreDest+'"') then
  begin
    if ExisteSQL ('SELECT FIR_TABLE FROM FILTRESREF WHERE FIR_TABLE="'+pstNomFiltre+'"') then
    begin
      stSQLInsert := 'INSERT INTO FILTRESREF (FIR_NUMPLAN,FIR_TABLE,FIR_LIBELLE,FIR_CRITERES,FIR_PREDEFINI) ';
      stSQLSelect := 'SELECT FIR_NUMPLAN,"'+pstNomFiltreDest+'", FIR_LIBELLE,FIR_CRITERES,FIR_PREDEFINI ';
      stSQLSelect := stSQLSelect + 'FROM FILTRESREF WHERE FIR_TABLE="'+pstNomFiltre+'"';
      Result := ExecuteSQL ( stSQLInsert + stSQLSelect );
    end else Result := -2;
  end else Result := -1;
end;   
////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************

Auteur  ...... : Gilles COSTE

Créé le ...... : 13/06/2006

Modifié le ... :   /  /

Description .. :

Mots clefs ... :

*****************************************************************}

procedure AGLNettoieFiltre( vStNomFiltre : string; vStOldZones : string; vStNewZones : String; vBoCtxStd : Boolean = False);
var lQuery : TQuery;
    lTobEnreg : Tob;
    i : integer;
    lStTemp : string;
    lStNewZones : string;
    lStOldZones : string;
    lStNewZone  : string;
    lStOldZone  : string;
    lParams : TStrings;
begin
  try
    try
      lParams := TStringList.Create ;

      if vBoCtxStd then
        lQuery := OpenSQL('SELECT * FROM FILTRESREF WHERE FIR_TABLE = "' + vStNomFiltre + '" ORDER BY FIR_TABLE', True)
      else
        lQuery := OpenSQL('SELECT * FROM FILTRES WHERE FI_TABLE = "' + vStNomFiltre + '" ORDER BY FI_TABLE', True);

      if not lQuery.Eof then
      begin
        lTobEnreg := TOB.Create('', nil, -1);

        if vBoCtxStd then
          lTobEnreg.LoadDetailDB('FILTRESREF', '', '', lQuery, False)
        else
          lTobEnreg.LoadDetailDB('FILTRES', '', '', lQuery, False);


        for i := 0 to lTobEnreg.Detail.Count - 1 do
        begin
          lStNewZones := vStNewZones;
          lStOldZones := vStOldZones;

          if vBoCtxStd then
            lParams.Text := lTobEnreg.Detail[i].GetValue('FIR_CRITERES')
          else
            lParams.Text := lTobEnreg.Detail[i].GetValue('FI_CRITERES');

          if (lParams.Count > 0) then
          begin
            if (Copy(lParams[0], 1, 5) = '<?xml') then
            begin // Contenu XML
              lStTemp := lParams.Text;
              lStNewZones := vStNewZOnes;
              lStOldZones := vStOldZOnes;

              repeat
                lStOldZone := ReadTokenSt(lStOldZones);
                lStNewZone := ReadTokenSt(lStNewZones);

                if (lStNewZone <> '') and (lStOldZone <> '') then
                begin
                  if Pos('<N>' + lStOldZone + '</N>', lStTemp) > 1 then
                   lStTemp := FindEtReplace( lStTemp, '<N>' + lStOldZone + '</N>', '<N>' + lStNewZone + '</N>' , TRUE);
                end;
              until (lStNewZone = '');
            end;

            if vBoCtxStd then
              lTObEnreg.Detail[i].PutValue('FIR_CRITERES', lStTemp)
            else
              lTObEnreg.Detail[i].PutValue('FI_CRITERES', lStTemp);

            lTobEnreg.InsertOrUpdateDB;
          end;
        end;
      end; // End Query EOF
    except
    end;
  finally
    FreeAndNil(lTobEnreg);
    FreeAndNil(lParams);
    Ferme(lQuery);
  end;
end;
//js1 100706 prise en compte d'un log spécfique pcl
function FicLogPCL : string;
var
  i: integer;
  s: string;
  StFicLogPCL : string;
begin
  Result:='';
  for i := 0 to ParamCount do
  begin
    s := UpperCase(ParamStr(i));
    if Copy(s, 1, 8) = '/LOGFILE' then
    begin
      StFicLogPCL := s;
      ReadTokenPipe(stFicLogPCL, '=');
      Result := stFicLogPCL;
      Break;
    end;
  end;
end;

//js1 100706 log specifique à pcl
procedure LogPCL(St,Fichier: string);
begin
  if (Fichier <> '') then
  try
    with TLogFile.create(Fichier) do
    begin
      Write(PChar(st));
      Free;
    end;
  finally
  end;
end;

//js1 100706 prise en compte d'un parametre PCL pour le shrink
function OkShrinkPCL : boolean;
var
  i: integer;
  s: string;
  StShrinkPCL : string;
begin
  Result:=false;
  for i := 0 to ParamCount do
  begin
    s := UpperCase(ParamStr(i));
    if Copy(s, 1, 7) = '/SHRINK' then
    begin
      StshrinkPCL := s;
      ReadTokenPipe(stShrinkPCL, '=');
      Result := ( stShrinkPCL='TRUE' );
      Break;
    end;
  end;
end;

//js1 110706 nom de la db0
function GetDb0name : string;
var
  i: integer;
  s: string;
  Stnom : string;
begin
  Result:='';
  for i := 0 to ParamCount do
  begin
    s := UpperCase(ParamStr(i));
    if Copy(s, 1, 8) = '/DOSSIER' then
    begin
      Stnom := s;
      ReadTokenPipe(Stnom, '@');
      Result := Stnom;
      Break;
    end;
  end;
end;

//js1 120706 teste verrou
function isVerrou : boolean;
var stSavVerrou,sSql:string;
    savEnableDeshare:boolean;
    QVerrou:Tquery;
begin
  Result:=false;
  savEnableDeshare:=V_PGI.EnableDeshare;
  V_PGI.EnableDeshare:=True;

  sSQL := 'SELECT DOS_VERROU FROM DOSSIER WHERE DOS_NODOSSIER="'+V_PGI.NoDossier+'"';
  QVerrou := OpenSQL(sSQL, True);
  if not QVerrou.Eof then stSavVerrou := QVerrou.FindField('DOS_VERROU').AsString;
  Ferme(QVerrou);

  if stSavVerrou = 'MAJ' then
    begin
    LogPCL('BASE VERROUILLEE ' + DateTimeToStr(Now),FicLogPCL);
    PGIBOX('Le dossier '+V_PGI.NODOSSIER+' est déjà en cours de mise à jour ou doit être '+
    'débloqué (outil d''assistance)');
    V_PGI.EnableDeshare:=savEnableDeshare;
    Result:=true;
  end
  else V_PGI.EnableDeshare:=savEnableDeshare;

end;

//js1 120706 nom de la db0
function GetUser : string;
var
  i: integer;
  s: string;
  Stnom ,sSql: string;
  QUser:Tquery;
begin
  Result:='';
  for i := 0 to ParamCount do
  begin
    s := UpperCase(ParamStr(i));
    if Copy(s, 1, 5) = '/USER' then
    begin
      Stnom := s;
      ReadTokenPipe(Stnom, '=');
      Break;
    end;
  end;
  sSQL := 'SELECT US_UTILISATEUR FROM UTILISAT WHERE US_ABREGE="'+Stnom+'"';

  QUser := OpenSQL(sSQL, True);
  if not QUser.Eof then Stnom := QUser.FindField('US_UTILISATEUR').AsString;
  Result := Stnom;
end;

{***********A.G.L.***********************************************
Auteur  ...... : JS
Créé le ...... : 22/04/2004
Modifié le ... :   /  /
Description .. : Mise à jour de la table journal d'événement
Mots clefs ... : JOURNAL;EVENEMENT;
*****************************************************************}
function MAJJnalEvent(TypeEvt, Etat, Libelle, BlocNote : string) : integer;
var TobJournal : TOB;
    TSqlJournal : TQuery;
begin
  Result := 0;
  TobJournal := TOB.Create('JNALEVENT', Nil, -1) ;
  try
    TobJournal.PutValue('GEV_TYPEEVENT',TypeEvt);
    TobJournal.PutValue('GEV_ETATEVENT',Etat);
    TobJournal.PutValue('GEV_LIBELLE',TraduireMemoire(Libelle));
    TobJournal.PutValue('GEV_DATEEVENT',Now);
    TobJournal.PutValue('GEV_UTILISATEUR',V_PGI.User);
    TOBJournal.PutValue('GEV_BLOCNOTE', TraduireMemoire(BlocNote));
    TSqlJournal := OpenSQL ('SELECT MAX(GEV_NUMEVENT) FROM JNALEVENT', True);
    try
      if Not TSqlJournal.EOF then Result := TSqlJournal.Fields[0].AsInteger;
    finally
      Ferme (TSqlJournal);
    end;
    Inc (Result);
    TOBJournal.PutValue ('GEV_NUMEVENT', Result);
    TOBJournal.InsertDB (Nil) ;
  finally
    TobJournal.Free;
  end;
end;

function ExecuteSQLNoPCL(const Sql: WideString): Integer;
begin
  if not IsDossierPCL then
    Result := ExecuteSQLContOnExcept(Sql)
  else
    Result := 0
end;

end.
