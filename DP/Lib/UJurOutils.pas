{***********UNITE*************************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
unit UJurOutils;
/////////////////////////////////////////////////////////////////
interface
/////////////////////////////////////////////////////////////////
uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
   {$IFDEF EAGLCLIENT}
   MaineAgl,
   {$ELSE}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
   FE_Main,
   {$ENDIF}
   {$IFDEF CJS5}
   UFJURSynthese,
   {$ENDIF CJS5}   
   PGIExec, SysUtils, HEnt1, Db, HRichEdt, HCtrls, windows, hmsgbox,
   dpJurOutilsGen, Dialogs, ExtCtrls, 
   FileCtrl, ParamSoc, CBPPath, UComOutils, DPJUROutilsDossiers,
   UTOB, controls, Classes, dpjuroutils, ugedfiles;

/////////////////////////////////////////////////////////////////

type
  TListeLibre = class
    Texte1 : string;
    Texte2 : string;
    Texte3 : string;
    Date1 : string;
    Date2 : string;
    Date3 : string;
    Coche1 : string;
    Coche2 : string;
    Coche3 : string;
    Montant1 : string;
    Montant2 : string;
    Montant3 : string;
  protected
    constructor Create;
  public
    procedure Detruit;
  end;

/////////////////////////////////////////////////////////////////
{$IFDEF CJS5}
procedure ChargeDossierJuridique;
{$ENDIF CJS5}

/////////////////////////////////////////////////////////////////
procedure InitEnteteEnregTable(var ChDateCreat : TDateTime;
                                var ChUserCreat : String;
                                var ChDateModif : TDateTime;
                                var ChUserModif : String);
                                {var ChAppModif : String);}

procedure LectureBlob(Q : TQuery; Champ : string ;ControlChamp : TControl; var TT : HTStrings);
procedure EcritureBlob(Q : TQuery; Champ : string ;ControlChamp : TControl; TT : HTStrings);

function GetFormeDos(ValChamp : string) : string;
function GetTypeDos(ValChamp : string ) : string;
function GetLibDos(ValChamp : string) : string;
function GetTypeDosBibAppli(ValChamp : string) : string;
function SelectionDossier : string;
function SelectionOperation(CodeDos : string ) : string;

function CalculDateEve(Limite,Reel : TDateTime; NbJour : integer) : TDateTime;
procedure CopieDocPerso(sCodeAct_p, sModule_p : string);
function Decale(Source : string; Decalage : integer) : string;
function DecodeAlRue1(laRue : string) : string;
procedure RecupeCleDossier(CodeDosJur : string; var sGuidPerDos_p,TypeDossier,NoOrdre : string);
function  DropLaTable(sLaTable_p : String; nLaVersion_p : integer) : boolean; overload;
function  DropLaTable(sLaTable_p : String) : boolean; overload;

function PosComposant( nNoComp_p, nMaxComp_p : integer;
                        var nTop_p, nLeft_p : integer;
                        nTopPlus_p, nLeftPlus_p : integer;
                        bHautVersBas_p : boolean = true ) : integer;
function PosCompHaB( nNoComp_p, nMaxComp_p : integer;
                        var nTop_p, nLeft_p : integer;
                        nTopPlus_p, nLeftPlus_p : integer ) : integer;
function PosCompGaD( nNoComp_p, nMaxComp_p : integer;
                        var nTop_p, nLeft_p : integer;
                        nTopPlus_p, nLeftPlus_p : integer ) : integer;

//procedure TOBToFile(sFile_p : string; OBTable_p : TOB; bEntete_p : boolean);
//function FileToTOB(sFile_p : string) : TOB;
function FileToTOBWithoutFormat(sFileParam_p, sTable_p : string) : TOB;
function FileToTOB(sFileParam_p, sTable_p : string;
                    var aosChamps_p, aosLibelle_p : TArrayOfString;
                    var sEntete_p : string; var iVersion_p : integer) : TOB;

procedure TOBToFile(sFileParam_p, sTable_p : string; OBParam_p : TOB;
                    aosChamps_p, aosLibelle_p : tarrayofstring;
                    sEntete_p : string; iVersion_p : integer);
procedure LectureEntete(sLigne_p : string; var sFormat_p, sEntete_p : string; var iVersion_p : integer);

function FieldToVariant(sChamp_p , sValeur_p : string) : variant;
function FieldToQuotedStr(sChamp_p : string; vValeur_p : variant) : string;
procedure JURSetFlagAppli(sExeName_p : string; bFlag_p : boolean);

// *** Nom des documents pour la fusion ***
function GetDocToLoad(sMaquette_p : string; bPerso_p : boolean;
                      var sDoc_p, sCrit1_p, sCrit2_p, sPredef_p : string) : boolean;

function GetNomDocToSave(sNoDossier_p, sNomDoc_p : string; bModeDev_p, bWithNoOrdre_p : boolean) : string;

function GetNomDocFromGED(sDocGuid_p, sCodeDos_p : string; var sDocPath_p : string) : boolean;

// *** Import / Export bibles ***
function  DocExiste(aosCrit_p : array of string; sPredef_p, sDocMaq_p : string) : boolean;
function  DocExport(aosCrit_p : array of string; sPredef_p, sDocMaq_p : string;
 						 bWithMessage_p : boolean) : string;
procedure DocImport(aosCrit_p : array of string; sPredef_p, sDocMaq_p, sFromFile_p : string);
procedure DocDelete(aosCrit_p : array of string; sPredef_p, sDocMaq_p : string);
function  DocDuplique(aosCrit_p : array of string; sPredef_p, sDocMaq_p, sNewPredef_p : string) : boolean;

function  ChoixMaquette(sDefautExt_p, sFiltre_p : string) : string;

// *** Import / Export GED ***
function GetRepToSave : string;
function AffecteNoOrdre(sNomDoc_p : string) : string;

/////////////////////////////////////////////////////////////////
implementation
/////////////////////////////////////////////////////////////////
uses
    GalOutil, UYFileSTD;

/////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... : 01/09/2003
Description .. :
Mots clefs ... :
*****************************************************************}
constructor TListeLibre.Create;
begin
  inherited Create;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TListeLibre.Detruit;
begin
  Free;
end;
/////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure InitEnteteEnregTable(var ChDateCreat : TDateTime;
                                var ChUserCreat : String;
                                var ChDateModif : TDateTime;
                                var ChUserModif : String);
                                {var ChAppModif : String);}
begin
  if ChUSerCreat='' then
  begin
    ChDateCreat := Now;
    ChUSerCreat := V_PGI.UserName;
  end;
  ChDateModif := Now;
  ChUserModif := V_PGI.UserName;
//  ChAppModif := 'Juridique';
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure LectureBlob(Q : TQuery;Champ : string ;ControlChamp : TControl; var TT : HTStrings);
begin
    // $$$ JP 05/09/06 TT.Assign(TMemoField(Q.FindField(Champ)));
    TT.Assign (THMemoField (Q.FindField (Champ)));
    StringsToRich (THRichEdit (ControlChamp), TT);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure EcritureBlob(Q : TQuery;Champ : string ;ControlChamp: TControl; TT : HTStrings);
begin
  RichToStrings(THRichEdit(ControlChamp),TT) ;
  TMemoField(Q.FindField(Champ)).Assign(TT) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function SelectionDossier : string;
begin
  result := AGLLanceFiche('JUR','JURIDIQUE_SEL','','','');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function SelectionOperation (CodeDos : string ) : string;
begin
  result := AGLLanceFiche('JUR','DOSOPER_SEL','JOP_CODEDOS='+CodeDos,'','');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetLibDos(ValChamp : string) : string;
var
  QQ : TQuery;
begin
  QQ:=OpenSQL('SELECT JUR_DOSLIBELLE from JURIDIQUE where JUR_CODEDOS="'+ValChamp+'"',TRUE) ;
  result:='';
  if Not QQ.EOF then
    result:=QQ.FindField('JUR_DOSLIBELLE').AsString;
  Ferme(QQ) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetFormeDos(ValChamp : string) : string;
var
  QQ : TQuery;
begin
  //NCX 11/12/00 Pour permettre de connaître la forme telle que définie dans JFJ_BIBLE
  // QQ:=OpenSQL('SELECT JUR_FORME,JUR_CODEDOS from JURIDIQUE where JUR_CODEDOS="'+ValChamp+'"',TRUE) ;
  QQ:=OpenSQL('SELECT JUR_FORME,JUR_CODEDOS,JFJ_BIBLE,JFJ_FORME from JURIDIQUE,JUFORMEJUR where JUR_CODEDOS="'+ValChamp+'" AND JFJ_FORME=JUR_FORME',TRUE) ;
  result:='';
  if Not QQ.EOF then
    result:=QQ.FindField('JFJ_BIBLE').AsString;
  Ferme(QQ) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetTypeDos(ValChamp : string) : string;
var
  QQ : TQuery;
begin
  QQ:=OpenSQL('SELECT JUR_TYPEDOS from JURIDIQUE where JUR_CODEDOS="'+ValChamp+'"',TRUE) ;
  result:='';
  if Not QQ.EOF then
    result:=QQ.FindField('JUR_TYPEDOS').AsString;
  Ferme(QQ) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetTypeDosBibAppli(ValChamp : string) : string;
var
  QQ : TQuery;
begin
  QQ:=OpenSQL('SELECT CO_ABREGE from COMMUN where CO_TYPE="JTD" and CO_CODE="'+GetTypeDos(ValChamp)+'"',TRUE) ;
  result:='';
  if Not QQ.EOF then
    result:=QQ.FindField('CO_ABREGE').AsString;
  Ferme(QQ) ;
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function CalculDateEve(Limite,Reel : TDateTime; NbJour : integer) : TDateTime;
var
  DateRef : TDateTime;
begin
  DateRef := iDate1900;
  if Reel<>iDate1900 then DateRef := Reel
  else if Limite<>iDate1900 then DateRef := Limite;
  result := iDate1900;
  if DateRef<>iDate1900 then result:=PlusDate(DateRef,NbJour,'J');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure CopieDocPerso(sCodeAct_p, sModule_p : string);
var
   sCEGMaquette_l, sCABMaquette_l, sRepPerso_l, sDocResult_l, sCrit1_l, sCrit2_l : string;
begin
   if GetBible(sModule_p, sCodeAct_p, '', '',
               false, sCEGMaquette_l, sDocResult_l, sCrit1_l, sCrit2_l) <> 0 then
      Exit;

   if sDocResult_l <> '' then
   begin
   sRepPerso_l := GetParamSocSecur('SO_JUREPPERSO', '') + '\' + sCrit1_l + '\' + sCrit2_l;
      sCABMaquette_l := sRepPerso_l + '\' + ExtractFileName(sCEGMaquette_l);
   end
   else
   begin
   sRepPerso_l := ExtractFileDir(sCEGMaquette_l);
      sCABMaquette_l := sCEGMaquette_l;
   end;

     if not DirectoryExists(sRepPerso_l) then
      ForceDirectories(sRepPerso_l);
     if CopyFile(PChar(sCEGMaquette_l), PChar(sCABMaquette_l), false) then
      PGIInfo('La maquette a été copiée dans ' + sRepPerso_l, 'Maquette personnalisée')
   else
      PGIError('ERREUR : La maquette n''a pu être copiée dans ' + sRepPerso_l, 'Maquette personnalisée');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function Decale(Source : string; Decalage : integer) : string;
begin
  result := Copy(Source,Decalage+1,Length(Source)-Decalage+1);
end;


//{
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function DecodeAlRue1(laRue : string) : string;
var
  mot,sep,tmp,alrue1 : string;
begin
  tmp := laRue; sep:=''; alrue1:='';
  while tmp<>'' do
  begin
    mot := ReadTokenSt(tmp);
    alrue1:=alrue1+sep+mot;
    if mot<>'' then sep:=' ' else sep:='';
  end;
  result := alrue1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure RecupeCleDossier(CodeDosJur : string; var sGuidPerDos_p,TypeDossier,NoOrdre : string);
var
  QQ : TQuery;
begin
//NCX 16/02/01 On garde que "STE"
  QQ:=OpenSQL('SELECT ANL_GUIDPERDOS,ANL_TYPEDOS,ANL_NOORDRE from '+
              'ANNULIEN where ANL_CODEDOS="'+CodeDosJur+'" AND ANL_TYPEDOS="STE"' ,TRUE) ;
  if Not QQ.EOF then
  begin
    sGuidPerDos_p:=QQ.FindField('ANL_GUIDPERDOS').AsString;
    TypeDossier:=QQ.FindField('ANL_TYPEDOS').AsString;
    NoOrdre:=QQ.FindField('ANL_NOORDRE').AsString;
  end;
  Ferme(QQ) ;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 01/09/2003
Modifié le ... :   /  /
Description .. : affichage des dossiers juridiques Apres connection
Mots clefs ... :
*****************************************************************}
{$IFDEF CJS5}
procedure ChargeDossierJuridique;
var
  sRequete_l, sGuidPer_l, sAction_l : string;
begin
   If PGIApp <> Nil Then
   begin
      sGuidPer_l := PGIApp.GetParam('General', 'Codeper', '');
      if (sGuidPer_l = '0') or (sGuidPer_l = '') then
         exit;
   end
   else
      exit;

   // on regarde si le dossier existe deja
   sRequete_l := 'select JUR_CODEDOS, JUR_GUIDPERDOS, JUR_TYPEDOS, JUR_NOORDRE ' +
                 'from JURIDIQUE ' +
                 'where JUR_GUIDPERDOS = "' + sGuidPer_l + '"' +
                 '  and JUR_TYPEDOS = "STE" ' +
                 '  and JUR_NOORDRE = 1 ';

   if not ExisteSQL(sRequete_l) then
      DPJURCreationDossierJuridique(sGuidPer_l, 'STE', 1);
   sAction_l := 'ACTION=MODIFICATION';
{   if ExisteSQL( sRequete_l ) then
     sAction_l := 'ACTION=MODIFICATION'
   else
     sAction_l := 'ACTION=CREATION';}

   LanceFicheSynthese(sAction_l, sGuidPer_l, 'STE', 'JUR', 1);
end;
{$ENDIF CJS5}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 21/10/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure JURSetFlagAppli(sExeName_p : string; bFlag_p : boolean);
var
   sGuidPer_l, sNodossier_l : string;
   OBDossier_l : TOB;
begin
   if not V_PGI.RunFromLanceur then
      Exit;

   If PGIApp <> Nil Then
   begin
      sGuidPer_l := PGIApp.GetParam('General', 'CodePer', '');
      if (sGuidPer_l = '0') or (sGuidPer_l = '') then
         sGuidPer_l := PGIApp.GetParam('General', 'GuidPer', '');      
   end
   else
   begin
      pgierror('Pas de PGIApp');
      exit;
   end;

   if (sGuidPer_l = '0') or (sGuidPer_l = '') then
   begin
      //pgierror('Pas de Guidper "' + sGuidPer_l + '"');
      exit;
   end;

   OBDossier_l := TOB.Create('DOSSIER', nil, -1);
   OBDossier_l.LoadDetailDBFromSQL('DOSSIER',
               'select DOS_NODOSSIER from DOSSIER where DOS_GUIDPER = "' + sGuidPer_l + '"');
   if OBDossier_l.Detail.Count > 0 then
      sNodossier_l := OBDossier_l.Detail[0].GetValue('DOS_NODOSSIER');
   OBDossier_l.Free;

   if sNodossier_l = '' then
   begin
      pgierror('Pas de Nodossier "' + sNodossier_l + '"');
      exit;
   end;

   SetFlagAppliDossier(sExeName_p, sNodossier_l, bFlag_p);
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 10/10/02
Fonction ..... : PoseComposants
Description .. : Calcule la position d'un champ dans une fiche à partir
                 de données de base
Paramètres ... : Ordre du composant
                 Nombre de composants par ligne ou par colonne
                 Position haute
                 Position gauche
                 Déplacement vers le bas
                 Déplacement vers la droite
                 Direction du déplacement
*****************************************************************}

function PosComposant( nNoComp_p, nMaxComp_p : integer;
                        var nTop_p, nLeft_p : integer;
                        nTopPlus_p, nLeftPlus_p : integer;
                        bHautVersBas_p : boolean = true ) : integer;
var
   nDimA_l, nDimAPlus_l, nDimB_l, nDimBPlus_l, nDimTmp_l : integer;
begin
   if ( nNoComp_p <> 0 ) and ( nMaxComp_p <> 0 ) then
   begin
      if bHautVersBas_p then
      begin
         nDimA_l := nTop_p;
         nDimAPlus_l := nTopPlus_p;
         nDimB_l := nLeft_p;
         nDimBPlus_l := nLeftPlus_p;
      end
      else
      begin
         nDimB_l := nTop_p;
         nDimBPlus_l := nTopPlus_p;
         nDimA_l := nLeft_p;
         nDimAPlus_l := nLeftPlus_p;
      end;

      if ( nNoComp_p mod nMaxComp_p ) <> 0 then
         nDimA_l := nDimA_l + nDimAPlus_l;

      if ( nNoComp_p mod nMaxComp_p ) = 0 then
      begin
         nDimB_l := nDimB_l + nDimBPlus_l;
         nDimTmp_l := nDimA_l - ( (nMaxComp_p - 1) * nDimAPlus_l );
         if nDimTmp_l > 0 then
            nDimA_l := nDimTmp_l;
      end;

      if bHautVersBas_p then
      begin
         nTop_p := nDimA_l;
         nLeft_p := nDimB_l;
      end
      else
      begin
         nTop_p := nDimB_l;
         nLeft_p := nDimA_l;
      end;

   end;
   result := nNoComp_p + 1;
end;
{*****************************************************************
Auteur ....... : BM
Date ......... : 10/10/02
Fonction ..... : PoseCompGad
Description .. : Calcule la position d'un champ dans une fiche à partir
                 de données de base de gauche à droite
Paramètres ... : Ordre du composant
                 Nombre de composants par ligne ou par colonne
                 Position haute
                 Position gauche
                 Déplacement vers le bas
                 Déplacement vers la droite
*****************************************************************}

function PosCompGaD( nNoComp_p, nMaxComp_p : integer;
                        var nTop_p, nLeft_p : integer;
                        nTopPlus_p, nLeftPlus_p : integer ) : integer;
begin
   result := PosComposant( nNoComp_p, nMaxComp_p, nTop_p, nLeft_p,
                            nTopPlus_p, nLeftPlus_p, false );
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 10/10/02
Fonction ..... : PoseComposants
Description .. : Calcule la position d'un champ dans une fiche à partir
                 de données de base de haut en bas
Paramètres ... : Ordre du composant
                 Nombre de composants par ligne ou par colonne
                 Position haute
                 Position gauche
                 Déplacement vers le bas
                 Déplacement vers la droite
*****************************************************************}

function PosCompHaB( nNoComp_p, nMaxComp_p : integer;
                        var nTop_p, nLeft_p : integer;
                        nTopPlus_p, nLeftPlus_p : integer ) : integer;
begin
   result := PosComposant( nNoComp_p, nMaxComp_p, nTop_p, nLeft_p,
                            nTopPlus_p, nLeftPlus_p, true );
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 08/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function DropLaTable(sLaTable_p : String; nLaVersion_p : integer) : boolean;
var
   nVersion_l : integer;
   bOkDel_l : boolean;
begin
   bOkDel_l := False ;
   nVersion_l := AglGetVersionOfTable(sLaTable_p);

   if (nVersion_l <> -1) and (nVersion_l < nLaVersion_p) then
      bOkDel_l := DropLaTable(sLaTable_p);
   result := bOkDel_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 08/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function DropLaTable(sLaTable_p : String) : boolean;
var
   sPref_l : String ;
   bOkDel_l : boolean;
begin
   sPref_l := TableToPrefixe(sLaTable_p) ;
   bOkDel_l := true;
   Try
      ExecuteSQL('DROP TABLE ' + sLaTable_p) ;
   Except
      bOkDel_l := False ;
   End ;
   if ((bOkDel_l) and (sPref_l <> '')) then
   BEGIN
      ExecuteSQL('DELETE FROM DETABLES WHERE DT_PREFIXE = "' + sPref_l + '"') ;
      ExecuteSQL('DELETE FROM DECHAMPS WHERE DH_PREFIXE = "' + sPref_l + '"') ;
   END ;
   result := bOkDel_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 15/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOBToFile(sFileParam_p, sTable_p : string; OBParam_p : TOB;
                    aosChamps_p, aosLibelle_p : tarrayofstring;
                    sEntete_p : string; iVersion_p : integer);
var
   iEnrInd_l, iChampInd_l : integer;
   tfTable_l : TextFile;
   sLigne_l : string;
begin
   AssignFile(tfTable_l, sFileParam_p);
   ReWrite(tfTable_l);
   //$|0|format|entete fichier|version|Date modififcation
   sLigne_l := '$|0|' + sTable_p + '|' + sEntete_p + '|' + IntToStr(iVersion_p) + '|' + DateTimeToStr(Date + Time);
   WriteLn(tfTable_l, sLigne_l);

   //$|2|format|entete champs
   sLigne_l := '$|2|' + sTable_p + '|';
   for iChampInd_l := 0 to Length(aosLibelle_p) - 1 do
      sLigne_l := sLigne_l + aosLibelle_p[iChampInd_l] + '|';
   WriteLn(tfTable_l, sLigne_l);

   //$|1|nomtable|nomchamp|
   sLigne_l := '$|1|' + sTable_p + '|';
   for iChampInd_l := 0 to Length(aosChamps_p) - 1 do
      sLigne_l := sLigne_l + aosChamps_p[iChampInd_l] + '|';
   WriteLn(tfTable_l, sLigne_l);


   for iEnrInd_l := 0 to OBParam_p.Detail.Count - 1 do
   begin
      sLigne_l := '=|1|' + sTable_p + '|';
      for iChampInd_l := 0 to Length(aosChamps_p) - 1 do
         sLigne_l := sLigne_l + VarToStr(OBParam_p.Detail[iEnrInd_l].GetValue(aosChamps_p[iChampInd_l])) + '|';
      WriteLn(tfTable_l, sLigne_l);
   end;
   CloseFile(tfTable_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 25/01/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure LectureEntete(sLigne_p : string; var sFormat_p, sEntete_p : string; var iVersion_p : integer);
var
   sPrefixe_l, sVers_l : string;
begin
   sPrefixe_l := Copy(sLigne_p, 1, 4);
   READTOKENPipe(sLigne_p, sPrefixe_l);
   sFormat_p := READTOKENPipe(sLigne_p, '|');
   sEntete_p := READTOKENPipe(sLigne_p, '|');
   sVers_l := READTOKENPipe(sLigne_p, '|');
   if sVers_l = '' then
      iVersion_p := 0
   else
      iVersion_p := StrToInt(sVers_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 15/09/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function FileToTOBWithoutFormat(sFileParam_p, sTable_p : string) : TOB;
var
   sEntete_l : string;
   aosChamps_l, aosLibelles_l : TArrayOfString;
   iVersion_l : integer;
begin
   result := FileToTOB(sFileParam_p, sTable_p, aosChamps_l, aosLibelles_l, sEntete_l, iVersion_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 18/01/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function FileToTOB(sFileParam_p, sTable_p : string;
                    var aosChamps_p, aosLibelle_p : TArrayOfString;
                    var sEntete_p : string; var iVersion_p : integer) : TOB;
var
   iInd_l : integer;
   tfTable_l : TextFile;
   sLigne_l, sPrefixe_l, sValeur_l, sVersion_l : string;
   OBParam_l : TOB;
begin
   result := nil;
   if not FileExists(sFileParam_p) then exit;

   OBParam_l := TOB.create(sTable_p, nil, -1);
   AssignFile(tfTable_l, sFileParam_p);
   Reset(tfTable_l);

   while not Eof(tfTable_l) do
   begin
      ReadLn(tfTable_l, sLigne_l);
      sPrefixe_l := Copy(sLigne_l, 1, 4);
      if sPrefixe_l = '$|0|' then  //$|0|format|entete fichier|version|Date modififcation
      begin
         LectureEntete(sLigne_l, sTable_p, sEntete_p, iVersion_p);
         Continue;
      end;

      sPrefixe_l := Copy(sLigne_l, 1, 4);
      READTOKENPipe(sLigne_l, sPrefixe_l);
      READTOKENPipe(sLigne_l, '|');

      if sPrefixe_l = '$|2|' then  //$|2|format|entete champs
      begin
         iInd_l := 0;
         while sLigne_l <> '' do
         begin
            Inc(iInd_l);
            SetLength(aosLibelle_p, iInd_l);
            aosLibelle_p[iInd_l-1] := READTOKENPipe(sLigne_l, '|');
         end;
         Continue;
      end;

      if sPrefixe_l = '$|1|' then  //$|1|nomtable|nomchamp|
      begin
         iInd_l := 0;
         while sLigne_l <> '' do
         begin
            Inc(iInd_l);
            SetLength(aosChamps_p, iInd_l);
            aosChamps_p[iInd_l-1] := READTOKENPipe(sLigne_l, '|');
         end;
         Continue;
      end;

      with TOB.Create(sTable_p, OBParam_l, -1) do  //=|1|nomtable|valeurchamp|
      begin
         iInd_l := 0;
         while sLigne_l <> '' do
         begin
            sValeur_l := READTOKENPipe(sLigne_l, '|');
            if TableToPrefixe(sTable_p) = '' then
               AddChampSup(aosChamps_p[iInd_l], true);
            PutValue(aosChamps_p[iInd_l], FieldToVariant(aosChamps_p[iInd_l], sValeur_l));
            Inc(iInd_l);
         end;
      end;
   end;
   CloseFile(tfTable_l);
   result := OBParam_l;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 16/09/2004
Modifié le ... :   /  /    
Description .. : Convertit une chaine en valeur selon le type du champ
Mots clefs ... : 
*****************************************************************}
function FieldToVariant(sChamp_p, sValeur_p : string) : variant;
var
   sType_l : string;
   vValeur_l : variant;
begin
   sType_l := UpperCase(ChampToType(sChamp_p));

   if (sType_l = 'DATE') then
      vValeur_l := StrToDate(sValeur_p)
   else if (sType_l = 'DOUBLE') then
      vValeur_l := StrToFloat(sValeur_p)
   else if (sType_l = 'INTEGER') then
      vValeur_l := StrToInt(sValeur_p)
   else
      vValeur_l := sValeur_p;

   result := vValeur_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 16/09/2004
Modifié le ... :   /  /    
Description .. : Rajoute "" si le champ est de type string
Mots clefs ... : 
*****************************************************************}
function FieldToQuotedStr(sChamp_p : string; vValeur_p : variant) : string;
var
   sType_l, sValeur_l : string;
begin
   sType_l := UpperCase(ChampToType(sChamp_p));

   if (sType_l = 'DATE') then
      sValeur_l := '"' + USDATETIME(vValeur_p) + '"'
   else if (sType_l = 'DOUBLE') then
      sValeur_l := FloatToStr(vValeur_p)
   else if (sType_l = 'INTEGER') then
      sValeur_l := IntToStr(vValeur_p)
   else
      sValeur_l := '"' + vValeur_p + '"';

   result := sValeur_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 22/07/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetDocToLoad(sMaquette_p : string; bPerso_p : boolean;
                      var sDoc_p, sCrit1_p, sCrit2_p, sPredef_p : string) : boolean;
var
	sBefore_l, sTmp_l : string;
begin
   result := false;
   if sMaquette_p = '' then exit;

   sMaquette_p := Uppercase(sMaquette_p);
   sBefore_l := sMaquette_p;
   sTmp_l := READTOKENPipe(sMaquette_p, '\');
   if (sTmp_l = '') or (sTmp_l = sBefore_l) then exit;
  	sCrit1_p := sTmp_l;

   sTmp_l := READTOKENPipe(sMaquette_p, '\');
   if (sTmp_l = '') or (sTmp_l = sBefore_l) then exit;
  	sCrit2_p := sTmp_l;

   sDoc_p := sMaquette_p;
   if sDoc_p = '' then exit;

   if (sPredef_p = '') then
   begin
      if bPerso_p then
         sPredef_p := 'STD'
      else
         sPredef_p := 'CEG';
   end;
   result := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 21/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetNomDocToSave(sNoDossier_p, sNomDoc_p : string; bModeDev_p, bWithNoOrdre_p : boolean) : string;
var
   sRepDest_l, sPathDoc_l : string;
begin
   result := '';
   
   sRepDest_l := GetParamSocSecur('SO_JUREPDOC', '');
   if sRepDest_l = '' then exit;

   if bModeDev_p then
   begin
      InitRep(sRepDest_l);
      sPathDoc_l := sRepDest_l + '\Resultat.rtf';
   end
   else
   begin
      if sNoDossier_p <> '' then
         sRepDest_l := sRepDest_l + '\' + sNoDossier_p;

      InitRep(sRepDest_l);
      sPathDoc_l := sRepDest_l + '\' + sNomDoc_p;
      if bWithNoOrdre_p then
         sPathDoc_l := AffecteNomDocMaquette(sPathDoc_l);
   end;

{   sNomRep_l := GetRepToSave;

   if bModeDev_p then
   begin
      sPathDoc_l := sNomRep_l + '\' + sNoDossier_p + '_Resultat.rtf';
   end
   else if bWithNoOrdre_p then
   begin
      sPathDoc_l := sNomRep_l + '\' + AffecteNoOrdre(sNoDossier_p + '_' + sNomDoc_p);
   end
   else
   begin
      sPathDoc_l := sNomRep_l + '\' + sNoDossier_p + '_' + sNomDoc_p;
   end;
 }
	result := UpperCase(sPathDoc_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 20/09/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetNomDocFromGED(sDocGuid_p, sCodeDos_p : string; var sDocPath_p : string) : boolean;
var
  	FileInfos_l : TGedFileInfos;
   sGEDName_l, sDocName_l : string;
begin
   result := false;

   if sDocGuid_p = '' then exit;

   V_GedFiles.GetGedFileInfos(sDocGuid_p, FileInfos_l);
   if (FileInfos_l.FileName = '') then exit;

   sGEDName_l := FileInfos_l.FileName;
   READTOKENPipe(sGEDName_l, '_');
   sDocName_l := READTOKENPipe(sGEDName_l, '_');
   READTOKENPipe(sGEDName_l, '.');
   sDocName_l := sDocName_l + '.' + sGEDName_l;

   sDocPath_p := GetNomDocToSave(sCodeDos_p, sDocName_l, false, true);
   V_GedFiles.Extract(sDocPath_p, sDocGuid_p);

   result := true;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/07/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function GetRepToSave : string;
var
   sNomRep_l : string;
begin
   sNomRep_l := TcbpPath.GetCegidUserTempPath;
   if V_PGI.RunFromLanceur then
	   sNomRep_l := sNomRep_l + 'PGI\DAT\' + V_PGI.DefaultSectionDbName
   else
	   sNomRep_l := sNomRep_l + 'PGI\DAT\' + V_PGI.DBName;

   InitRep(sNomRep_l);

	result := UpperCase(sNomRep_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 05/07/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function AffecteNoOrdre(sNomDoc_p : string) : string;
var
  iNoOrdre_l : integer;
  sRacine_l, sExt_l, sFileName_l, sNoOrdre_l : string;
  OBYFiles_l : TOB;
begin
  sRacine_l := UpperCase(Copy(sNomDoc_p, 0, Length(sNomDoc_p) - 4));
  sExt_l := UpperCase(Copy(sNomDoc_p, Length(sNomDoc_p) - 3, 4));

  OBYFiles_l := TOB.Create('YFILES', nil, -1);
  OBYFiles_l.LoadDetailDBFromSQL('YFILES',
  											'SELECT MAX(YFI_FILENAME) AS YFI_FILENAME ' +
                                 'FROM YFILES ' +
                                 'WHERE YFI_FILENAME LIKE "' + sRacine_l + '_%' + sExt_l + '"');

	iNoOrdre_l := 0;
	if OBYFiles_l.Detail.Count > 0 then
   begin
    	sFileName_l := OBYFiles_l.Detail[0].GetString('YFI_FILENAME');
      READTOKENPipe(sFileName_l, sRacine_l + '_');
      iNoOrdre_l := StrToIntDef(READTOKENPipe(sFileName_l, sExt_l), 0);
   end;

   Inc(iNoOrdre_l);
   sFileName_l := sRacine_l + '_' + IntToStr(iNoOrdre_l) + sExt_l;

  OBYFiles_l.Free;
  result := sFileName_l;
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 28/06/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function DocExport(aosCrit_p : array of string; sPredef_p, sDocMaq_p : string;
						bWithMessage_p : boolean) : string;
var
	sRepPath_l, sDocPath_l, sLangue_l : string;
	iErrCode_l : integer;
   iCritInd_l : integer;
   bCritOK_l : boolean;
   aosCrit_l : array [0..4] of string;
begin
  	result := '';
   bCritOK_l := true;

   for iCritInd_l := 0 to 4 do
   begin
      if iCritInd_l < Length(aosCrit_p) then
      begin
         bCritOK_l := (aosCrit_p[iCritInd_l] <> '');
         aosCrit_l[iCritInd_l] := aosCrit_p[iCritInd_l]
      end;
   end;

//   if not bCritOK_l then
//      exit;

	if (sPredef_p = '') or (sDocMaq_p = '') then
		Exit;

   sLangue_l := V_PGI.LangueParDefaut;
   if sLangue_l = '' then
      sLangue_l := 'FRA';

   sRepPath_l := TCBPPath.GetCegidUserTempPath + 'PGI\STD\' + V_PGI.CodeProduit;
   for iCritInd_l := 0 to length(aosCrit_p) - 1 do
      sRepPath_l := sRepPath_l + '\' + aosCrit_p[iCritInd_l];
   sRepPath_l := sRepPath_l + '\' + sLangue_l + '\' + sPredef_p;

   if not InitRep(sRepPath_l) then
   begin
      PGIError('Répertoire ' + sRepPath_l + ' innaccessible.#13#10 Droits d''accès insuffisants');
      Result := '';
      exit;
   end;

   sDocPath_l := sRepPath_l + '\' + sDocMaq_p;
   DeleteFile(PChar(sDocPath_l));

   iErrCode_l := AGL_YFILESTD_EXTRACT(sDocPath_l, V_PGI.CodeProduit, sDocMaq_p,
                                      aosCrit_l[0], aosCrit_l[1], aosCrit_l[2], aosCrit_l[3], aosCrit_l[4],
                                      false, sLangue_l, sPredef_p);

   if iErrCode_l > 0 then
   begin
      if bWithMessage_p then
         PGIInfo(AGL_YFILESTD_GET_ERR(iErrCode_l) + ' : ' + sDocPath_l);
      sDocPath_l := '';
   end;

	result := sDocPath_l;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 02/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure DocImport(aosCrit_p : array of string; sPredef_p, sDocMaq_p, sFromFile_p : string);
var
	sExt_l, sLangue_l : string;
	iErrCode_l : integer;
   aosCrit_l : array [0..4] of string;
   iCritInd_l : integer;
   bCritOK_l : boolean;
begin
   bCritOK_l := true;
   for iCritInd_l := 0 to 4 do
   begin
      if iCritInd_l < Length(aosCrit_p) then
         aosCrit_l[iCritInd_l] := aosCrit_p[iCritInd_l]
   end;

   sExt_l := UpperCase(ExtractFileExt(sDocMaq_p));
   sExt_l := Copy(sExt_l, 2, Length(sExt_l) - 1);

   sLangue_l := V_PGI.LangueParDefaut;
   if sLangue_l = '' then
      sLangue_l := 'FRA';

   iErrCode_l := AGL_YFILESTD_IMPORT(sFromFile_p, V_PGI.CodeProduit, sDocMaq_p, sExt_l,
                          aosCrit_l[0], aosCrit_l[1], aosCrit_l[2], aosCrit_l[3], aosCrit_l[4],
                          '-', '-', '-', '-', '-',
                          sLangue_l, sPredef_p,
                          'Maquette CJS5', '000000');
   if iErrCode_l > 0 then
   begin
      PGIInfo(AGL_YFILESTD_GET_ERR(iErrCode_l) + ' : ' + sFromFile_p);
   end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 29/01/2008
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function DocDuplique(aosCrit_p : array of string; sPredef_p, sDocMaq_p, sNewPredef_p : string) : boolean;
var
   sDocPath_l : string;
begin
  	sDocPath_l := DocExport(aosCrit_p, sPredef_p, sDocMaq_p, false);
   if sDocPath_l <> '' then
      DocImport(aosCrit_p, sNewPredef_p, sDocMaq_p, sDocPath_l);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 23/11/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function DocExiste(aosCrit_p : array of string; sPredef_p, sDocMaq_p : string) : boolean;
var
	sLangue_l : string;
   aosCrit_l : array [0..4] of string;
   iCritInd_l : integer;
   bCritOK_l : boolean;
begin
   for iCritInd_l := 0 to 4 do
   begin
      if iCritInd_l < Length(aosCrit_p) then
         aosCrit_l[iCritInd_l] := aosCrit_p[iCritInd_l]
   end;

   sLangue_l := V_PGI.LangueParDefaut;
   if sLangue_l = '' then
      sLangue_l := 'FRA';

   result := ExisteSQL('select * from yfilestd ' +
                       'where yfs_codeproduit = "' + V_PGI.CodeProduit +'" ' +
                       '  and yfs_nom = "' + sDocMaq_p + '" ' +
                       '  and yfs_langue = "' + sLangue_l + '" ' +
                       '  and yfs_predefini = "' + sPredef_p + '" ' +
                       '  and yfs_crit1 = "' + aosCrit_l[0] + '" ' +
                       '  and yfs_crit2 = "' + aosCrit_l[1] + '" ' +
                       '  and yfs_crit3 = "' + aosCrit_l[2] + '" ' +
                       '  and yfs_crit4 = "' + aosCrit_l[3] + '" ' +
                       '  and yfs_crit5 = "' + aosCrit_l[4] + '" ');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : BM
Créé le ...... : 28/06/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure DocDelete(aosCrit_p : array of string; sPredef_p, sDocMaq_p : string);
var
	sFileGuid_l, sLangue_l : string;
   OBYFileSTD_l, OBNFiles_l, OBNFileParts_l : TOB;
   iInd_l : integer;
   aosCrit_l : array [0..4] of string;
   iCritInd_l : integer;
   bCritOK_l : boolean;
begin
   bCritOK_l := true;
   for iCritInd_l := 0 to 4 do
   begin
      if iCritInd_l < Length(aosCrit_p) then
      begin
         bCritOK_l := (aosCrit_p[iCritInd_l] <> '');
         aosCrit_l[iCritInd_l] := aosCrit_p[iCritInd_l]
      end;
   end;

   if not bCritOK_l then
      exit;

	if (sPredef_p = '') or (sDocMaq_p = '') then
		Exit;

   sLangue_l := V_PGI.LangueParDefaut;
   if sLangue_l = '' then
      sLangue_l := 'FRA';

   OBYFileSTD_l := TOB.Create('YFILESTD', nil, -1);
   OBNFiles_l := TOB.Create('NFILES', nil, -1);
   OBNFileParts_l := TOB.Create('NFILEPARTS', nil, -1);

   OBYFileSTD_l.LoadDetailDBFromSQL('YFILESTD',
               'SELECT * FROM YFILESTD ' +
               'WHERE YFS_CODEPRODUIT = "' + V_PGI.CodeProduit + '" ' +
               '  AND YFS_NOM = "' + sDocMaq_p  + '" ' +
               '  AND YFS_LANGUE = "' + sLangue_l + '" ' +
               '  AND YFS_CRIT1 = "' + aosCrit_l[0] + '" ' +
               '  AND YFS_CRIT2 = "' + aosCrit_l[1] + '" ' +
               '  AND YFS_CRIT3 = "' + aosCrit_l[2] + '" ' +
               '  AND YFS_CRIT4 = "' + aosCrit_l[3] + '" ' +
               '  AND YFS_CRIT4 = "' + aosCrit_l[4] + '" ' +
               '  AND YFS_PREDEFINI = "' + sPredef_p + '"');

   if OBYFileSTD_l.Detail.Count > 0 then
   begin
      sFileGuid_l := OBYFileSTD_l.Detail[0].GetString('YFS_FILEGUID');
      OBNFiles_l.LoadDetailDBFromSQL('NFILES',
                  'SELECT * FROM NFILES WHERE NFI_FILEGUID = "' + sFileGuid_l + '"');
      OBNFileParts_l.LoadDetailDBFromSQL('NFILEPARTS',
                  'SELECT * FROM NFILEPARTS WHERE NFS_FILEID = "' + sFileGuid_l + '"');

      for iInd_l := 0 to OBYFileSTD_l.Detail.Count - 1 do
         OBYFileSTD_l.Detail[iInd_l].DeleteDB;
      for iInd_l := 0 to OBNFiles_l.Detail.Count - 1 do
         OBNFiles_l.Detail[iInd_l].DeleteDB;
      for iInd_l := 0 to OBNFileParts_l.Detail.Count - 1 do
         OBNFileParts_l.Detail[iInd_l].DeleteDB;
   end;

   OBYFileSTD_l.Free;
   OBNFiles_l.Free;
   OBNFileParts_l.Free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 25/07/2005
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :
*****************************************************************}
{function ImportMaquette(sPathFic_p, sMaquette_p, sPerso_p, sPredef_p : string;
                        dtDateCre_p, dtDateMod_p : TDatetime;
                        bModif_p : boolean) : integer;
var
     sExt_l, sDocBase_l, sCrit1_l, sCrit2_l, sCopie_l : string;
   iErrCode_l : integer;
begin
   if (sMaquette_p = '') then
   begin
      iErrCode_l := 0;
   end
   else if (sPerso_p = 'X') then
   begin
      iErrCode_l := -1;
   end
   else
   begin
      GetDocToLoad(sMaquette_p, false, sDocBase_l, sCrit1_l, sCrit2_l, sPredef_p);
      // Copy de sauvegarde
      sCopie_l := sPathFic_p;
      if bModif_p then
      begin
         sCopie_l := GetParamSocSecur('SO_JUREPBIBLE', '') + '\' + sCrit1_l + '\' + sCrit2_l + '\' + V_PGI.LangueParDefaut + '\' + sPredef_p;
         ForceDirectories(sCopie_l);
         sCopie_l := sCopie_l + '\' + sDocBase_l;
         CopyFile(PChar(sPathFic_p), PChar(sCopie_l), false);
      end;

      sExt_l := UpperCase(ExtractFileExt(sDocBase_l));
      sExt_l := Copy(sExt_l, 2, Length(sExt_l) - 1);
      iErrCode_l := AGL_YFILESTD_IMPORT(sCopie_l, V_PGI.CodeProduit, sDocBase_l, sExt_l,
                             sCrit1_l, sCrit2_l, '', '', '',
                             '-', '-', '-', '-', '-',
                             V_PGI.LangueParDefaut, sPredef_p,
                             'Maquette CJS5', '000000');

      SetDatesMaquette(V_PGI.CodeProduit, sDocBase_l, sCrit1_l, sCrit2_l,
                       sPredef_p, V_PGI.LangueParDefaut, dtDateCre_p, dtDateMod_p);
   end;
   if iErrCode_l > 0 then
      PGIInfo(AGL_YFILESTD_GET_ERR(iErrCode_l) + ' : ' + sDocBase_l + ' : ' + sCopie_l);
   result := iErrCode_l;
end;}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 25/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function ExtractIdMaquette(sDir_p, sMaquette_p : string) : string;
var
   sCrit1_l, sCrit2_l : string;
begin
   while sDir_p <> sMaquette_p do
   begin
      sCrit1_l := sCrit2_l;
      sCrit2_l := '';
      sCrit2_l := READTOKENPipe(sDir_p, '\');
   end;
   sDir_p := '';
   if sCrit1_l <> ''  then sDir_p := sCrit1_l + '\';
   if sCrit2_l <> ''  then sDir_p := sDir_p + sCrit2_l + '\';
   sMaquette_p := sDir_p + sMaquette_p;
   result := sMaquette_p;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 11/10/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{procedure GetDatesMaquette(sCodeProd_p, sNom_p, sCrit1_p, sCrit2_p, sPredef_p, sLangue_p : string;
                           var dtDateCre_p, dtDateMod_p : TDatetime) ;
var
   OBMaq_l : TOB;
begin
   dtDateCre_p := iDate1900;
   dtDateMod_p := iDate1900;

   OBMaq_l := TOB.Create('YFILESTD', nil, -1);
   OBMaq_l.LoadDetailDBFromSQL('YFILESTD',
                    'SELECT YFS_DATECREATION, YFS_DATEMODIF ' +
                    'FROM YFILESTD ' +
                    'WHERE YFS_CODEPRODUIT = "' + sCodeProd_p + '" ' +
                    '  AND YFS_NOM = "' + sNom_p + '" ' +
                    '  AND YFS_LANGUE = "' + sLangue_p + '" ' +
                    '  AND YFS_PREDEFINI = "' + sPredef_p + '" ' +
                    '  AND YFS_CRIT1 = "' + sCrit1_p + '" ' +
                    '  AND YFS_CRIT2 = "' + sCrit2_p + '"');
   if OBMaq_l.Detail.Count > 0 then
   begin
      dtDateCre_p := OBMaq_l.Detail[0].GetDateTime('YFS_DATECREATION');
      dtDateMod_p := OBMaq_l.Detail[0].GetDateTime('YFS_DATEMODIF');
   end;
   OBMaq_l.Free;
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 11/10/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{procedure SetDatesMaquette(sCodeProd_p, sNom_p, sCrit1_p, sCrit2_p, sPredef_p, sLangue_p : string;
                           dtDateCre_p, dtDateMod_p : TDatetime) ;
var
   sRequete_l : string;
begin
   if dtDateCre_p = iDate1900 then
      dtDateCre_p := Date;

   sRequete_l := 'UPDATE YFILESTD ' +
                 'SET YFS_DATECREATION = "' + USTIME(dtDateCre_p) + '", ' +
                 '    YFS_DATEMODIF = "' + USTIME(dtDateMod_p) + '" ' +
                 'WHERE YFS_CODEPRODUIT = "' + sCodeProd_p + '" ' +
                 '  AND YFS_NOM = "' + sNom_p + '" ' +
                 '  AND YFS_LANGUE = "' + sLangue_p + '" ' +
                 '  AND YFS_PREDEFINI = "' + sPredef_p + '" ' +
                 '  AND YFS_CRIT1 = "' + sCrit1_p + '" ' +
                 '  AND YFS_CRIT2 = "' + sCrit2_p + '"';
   ExecuteSQL(sRequete_l);
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 25/07/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
{procedure DeleteMaquette(sMaquette_p, sPerso_p : string; sPredef_p : string = '');
var
     sDocBase_l, sCrit1_l, sCrit2_l : string;
   OBYFileSTD_l, OBNFiles_l, OBNFileParts_l : TOB;
   iFILEID_l, iInd_l : integer;
begin

   GetDocToLoad(sMaquette_p, false, sDocBase_l, sCrit1_l, sCrit2_l, sPredef_p);

   OBYFileSTD_l := TOB.Create('YFileSTD', nil, -1);
   OBNFiles_l := TOB.Create('NFiles', nil, -1);
   OBNFileParts_l := TOB.Create('NFileParts', nil, -1);

   OBYFileSTD_l.LoadDetailDBFromSQL('YFileSTD',
               'SELECT * FROM YFILESTD ' +
               'WHERE YFS_CODEPRODUIT = "' + V_PGI.CodeProduit + '" ' +
               '  AND YFS_NOM = "' + sDocBase_l  + '" ' +
               '  AND YFS_LANGUE = "' + V_PGI.LangueParDefaut + '" ' +
               '  AND YFS_PREDEFINI = "' + sPredef_p + '" ' +
               '  AND YFS_CRIT1 = "' + sCrit1_l + '" ' +
               '  AND YFS_CRIT2 = "' + sCrit2_l + '" ' +
               '  AND YFS_PREDEFINI = "' + sPredef_p + '"');

   if OBYFileSTD_l.Detail.Count > 0 then
   begin
      iFILEID_l := OBYFileSTD_l.Detail[0].GetInteger('YFS_FILEID');
      OBNFiles_l.LoadDetailDBFromSQL('NFiles',
                  'SELECT * FROM NFILES WHERE NFI_FILEID = ' + IntToStr(iFileID_l));
      OBNFileParts_l.LoadDetailDBFromSQL('NFileParts',
                  'SELECT * FROM NFILEPARTS WHERE NFS_FILEID = ' + IntToStr(iFileID_l));

      for iInd_l := 0 to OBYFileSTD_l.Detail.Count - 1 do
         OBYFileSTD_l.Detail[iInd_l].DeleteDB;
      for iInd_l := 0 to OBNFiles_l.Detail.Count - 1 do
         OBNFiles_l.Detail[iInd_l].DeleteDB;
      for iInd_l := 0 to OBNFileParts_l.Detail.Count - 1 do
         OBNFileParts_l.Detail[iInd_l].DeleteDB;
   end;
   
   OBYFileSTD_l.Free;
   OBNFiles_l.Free;
   OBNFileParts_l.Free;
end;}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 06/07/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function ChoixMaquette(sDefautExt_p, sFiltre_p : string) : string;
var
  	OD : TOpenDialog;
  	Panel : TPanel;
begin
	result := '';
  	Panel := TPanel.Create(nil);
  	Panel.Visible := False;
  	Panel.ParentWindow := GetDesktopWindow;
  	OD := TOpenDialog.Create(Panel);

  	OD.DefaultExt := sDefautExt_p;
  	OD.Filter := sFiltre_p;

  	if (OD.Execute) and (OD.FileName <> '') then
      result := OD.FileName;

  	OD.Free;
  	Panel.Free;
end;
end.


