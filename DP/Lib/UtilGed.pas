unit UtilGed;

interface

uses
  HCtrls,
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF}
  HDTLinks, SysUtils, UTob, Registry, Windows, ShellAPI,
  Hent1, UGedFiles, HMsgBox, comctrls, ComObj, OleCtnrs,
  {$IFNDEF EAGLSERVER}
  CBPPath,
  {$ENDIF}
  TlHelp32;


Type

  TInfoDocGed = Record
                 CheminFichierGed : String; // Chemin du fichier GED
                 NomFichierGed    : String; // Nom du fichier GED avec extension
                 DocGuid          : String;
                 FileGuid         : String;
                 CodeGed          : String;
                 NoDossier        : String; // Numéro du dossier
                 Libelle          : String; // Libelle du document
                 Auteur           : String; // Auteur du document
                end;

  TGedDPNode = class(TObject)
  public
    SFileGUID  : String;
    SDocGUID   : String;
    CodeGed    : String;
    TypeGed    : String;   // DCS (ens. de docs), DOC, FIC, LIS, COM, MSG
    TagMenu    : Integer;
    IconId     : Integer;
    IconSel    : Integer;  // Id icône quand noeud déployé
    Caption    : String;
    Url        : string;   // Url (si YDO_DOCTYPE="URL")
    EwsId      : String;   // Ged - Id de la branche eWs associée

    // $$$ JP 11/04/06
    EwsRegle   : string;  // Ged - Règle de publication eWS associée

    // $$$ JP 30/09/04 - pour la publication html sur CD
    CDPubli    : String;  // Ged - code de publication sur CD (voir description YGD_CDPUBLI
    Auteur     : String;  // Ged - auteur du document (si document)
    Nature     : String;  // Ged - nature document (si document)
    Mois       : String;  // Ged - mois du document (si document)
    Annee      : String;  // Ged - année du document (si document)

    DateDeb    : TDateTime; // Ged - date début de validité (si document)
    DateFin    : TDateTime; // Ged - date fin de validité (si document)
    NoDossier  : String;    // Ged - No de dossier (si document)
    CodeUser   : String;    // Ged - Code utilisateur (si document)
    EwsDatePub : TDateTime; // Ged - Date publication eWS (si document)

    DocFileName       : String;    // Ged - Nom du fichier
    DocExt            : String;    // Ged - Extension
    DocEtat           : Integer;   // Ged - Etat du document (En cours de modif ou périmé)
    DocPrive          : Boolean;   // Ged - Document prive
    DocDroitGed       : String;    // Ged - Libelle droit GED
    DroitDocument     : String;    // Ged - Droit document
    DocCreateur       : String;    // Ged - Createur du document dans la GED
    DocDateCreation   : TDateTime; // Ged - Date de création du document dans la GED
    DocDateModif      : TDateTime; // Ged - Date d'extration

    function  CodeReg    (strTypeReg:string):string;
    function  DisplayReg (strTypeReg:string):string;
  end;

function  ExtensionToLibelleApplication (Extension: string): string;
function  GetCheminGed(sCodeGed: String; bDetaille : Boolean): String;
function  CodeGedToEwsId (CodeGed: String) : String;
function  CodeGedToEwsRegle (CodeGed: String) : String; // $$$ JP 11/04/06

function  InsertDocumentGed(TobDoc, TobDocGed: TOB; SFileGUID: string; var LastError: String): string;
function  GetFileNameGed(SFileGUID: string): String;
function  SupprimeDocument(SDocGUID : string): Boolean;
function  GetFileId(SModeleGUID : string ) : String;
function  ExtraitDocument(SDocGUID : string; var sTitre: String; parle:Boolean=TRUE): String;
procedure MarquageEwsAPublier(SDocGUID: string; bApublier: Boolean);
procedure SupprimeDocumentGed(SDocGUID: string);

{$IFNDEF EAGLSERVER}
function InitialiserInfoDocGed (NoDossier : String; DocGuid : String) : TInfoDocGed; OverLoad;
function InitialiserInfoDocGed (UnChemin : String; UnFichier : String; UnDossier : String; UnCodeGed : String; UnLibelle : String; UnAuteur : String) : TInfoDocGed; OverLoad;
{$ENDIF}
function ExtraireDocumentGed (DocGed : TInfoDocGed) : Boolean;
function IntegrerDocumentGed (DocGed : TInfoDocGed) : Boolean;
function ReintegrerDocumentGed (DocGed : TInfoDocGed) : Boolean;
procedure OuvrirDocumentGed (DocGed : TInFoDocGed);

function CreerDocumentDOC (UnChemin : String; UnFichier : String) : Boolean;
function CreerDocumentXLS (UnChemin : String; UnFichier : String) : Boolean;
function ArreterProcess(const ProcessName : string) : boolean;

function IsValideParametreGedOLE (UnUtilisateur : String; UnCodeGed : String; UnRepertoire : String) : Boolean;
function ExtraireDocumentsGedOLE (UnUtilisateur : String; UnDossier : String; UnCodeGed : String; UnRepertoire : String) : Boolean;
function ExtraireDocGed (UnDossier : String; UnCodeGed : String; UnRepertoire : String) : Boolean;
function TraiteAjoutFichierDansGed (Fichier, strDossier, CodeGed, Description, Auteur, Annee, Mois:string; majwinner:Boolean):variant; OverLoad;

//////////// IMPLEMENTATION ////////////
implementation

{$IFNDEF EAGLSERVER} // ajout me 18/11/2005
uses UtilDossier, YDocuments_Tom;
{$ENDIF}

function GetFileNameGed(SFileGUID: String): String;
var Q: TQuery;
begin
   Result := '';
   if SFileGUID = '' then
      exit;
   Q := OpenSQL('SELECT YFI_FILENAME FROM YFILES WHERE YFI_FILEGUID = "' + SFileGUID+'"', True);
   if Not Q.Eof then
      Result := Q.FindField('YFI_FILENAME').AsString;
   Ferme(Q);
end;

function GetFileId(SModeleGUID : String) : String;
var QRYDocFiles_l : TQuery;
    SFileGUID : String;
begin
   Result := '';
   if SModeleGUID ='' then exit;
   QRYDocFiles_l := OpenSQL('SELECT YDF_FILEGUID ' +
                            'FROM YDOCUMENTS, YDOCFILES ' +
                            'WHERE YDO_DOCGUID = YDF_DOCGUID ' +
                            '  AND YDO_DOCGUID = "' + SModeleGUID + '"' +
                            '  AND (YDF_FILEROLE IS NULL OR YDF_FILEROLE = "PAL") ', true);
   if not QRYDocFiles_l.Eof then
      SFileGUID := QRYDocFiles_l.FindField('YDF_FILEGUID').AsString;
   Ferme(QRYDocFiles_l);
   result := SFileGUID;
end;

procedure SupprimeDocumentGed(SDocGUID: string);
begin
  // pour l'instant pas besoin de tom sur DPDOCUMENT (niveau le plus haut)
  ExecuteSQL('DELETE DPDOCUMENT WHERE DPD_DOCGUID="' + SDocGUID + '"' );

  // suppression dans YDOCUMENTS, YDOCFILES, YFILES, YFILEPARTS :
  SupprimeDocument(SDocGUID);
end;

function SupprimeDocument(SDocGUID : string): Boolean;
var
  tDocs, tResult: Tob;
begin
  tDocs := Tob.Create('_YDOCUMENTS_', nil, -1);
  tResult := Tob.Create('_RESULT_', nil, -1);

  try
    tDocs.LoadDetailDBFromSQL('YDOCUMENTS', 'SELECT * FROM YDOCUMENTS WHERE (YDO_DOCGUID = "' + SDocGUID + '")');
{$IFDEF EAGLSERVER} // ajout me 18/11/2005
    tDocs.DeleteDB;
{$ELSE}
    tDocs.DeleteDBTom('YDOCUMENTS', False, tResult);
{$ENDIF}

    if tResult.Detail.Count > 0 then
    begin
      Result := False;
    end
    else
      Result := True;

  except
//    FLastErrorMsg := 'Erreur dans le processus';
    Result := False;
  end;

  tDocs.Free;
  tResult.Free;
end;


{***********A.G.L.***********************************************
Auteur  ...... : MD
Créé le ...... : 01/01/2007
Modifié le ... : 29/05/2007
Description .. : Fonction bas niveau d'insertion dans la GED,
Suite ........ : sans IHM.
Suite ........ : Retour le DocGuid obtenu.
Mots clefs ... : 
*****************************************************************}
function InsertDocumentGed(TobDoc, TobDocGed: TOB; SFileGUID: string; var LastError: String): string;
var SDocGUID : String;
    TobDocFiles : TOB;
begin
  Result := '';

  // Contrôles avant validation d'un nouvel enreg
  if TobDocGed.GetValue('DPD_NODOSSIER')='' then
    begin
    LastError := 'N° de dossier vide (DPD_NODOSSIER)';
    exit;
    end;
  // Contrôles généraux
  if TobDoc.GetValue('YDO_LIBELLEDOC')='' then
    begin
    LastError := 'Pas de libellé descriptif du document (YDO_LIBELLEDOC)';
    exit;
    end;
  if (TobDoc.GetValue('YDO_ANNEE')<>'') and (Not IsNumeric(TobDoc.GetValue('YDO_ANNEE'))) then
    begin
    LastError := 'L''année doit être numérique ou vide (YDO_ANNEE)';
    exit;
    end;
  if (TobDoc.GetValue('YDO_MOIS')<>'') and (Not IsNumeric(TobDoc.GetValue('YDO_MOIS'))) then
    begin
    LastError := 'Le mois doit être numérique ou vide (YDO_MOIS)';
    exit;
    end;

  // Valeurs non visibles dans l'arborescence DP
  if TobDocGed.GetValue('DPD_CODEGED')='' then
    TobDocGed.PutValue('DPD_CODEGED', '###');

  // Calcul de la clé primaire
  SDocGUID:=AGLGetGUID();
  TobDoc.PutValue('YDO_DOCGUID', SDocGUID);
  TobDocGed.PutValue('DPD_DOCGUID', SDocGUID);

  // MD 31/05/06 : Ancienne clé mise à -2
  // pour éviter pb de maj des clés étrangères (comme dans ECRCOMPL)
  // qui prennent les enreg ydo_docid à 0 pour des enreg ged valides
  // lors de la conversion des integer en guid...
  TobDoc.PutValue('YDO_DOCID', -2);
  TobDocGed.PutValue('DPD_DOCID', -2);

  // #### Calcul d'une référence documentaire à revoir
  TobDoc.PutValue ('YDO_DOCREF', TobDocGed.GetValue('DPD_NODOSSIER') + ' ' + TobDocGed.GetValue('DPD_CODEGED') + ' '+ SFileGUID + ' ' + SDocGUID);

  // $$$ JP 19/12/06: doctype à "DOC"
  TobDoc.PutValue ('YDO_DOCTYPE', 'DOC');

  TobDocFiles := Nil;
  Try
    // #### pour l'instant, insère un seul fichier pour le document
    // avec comme rôle "PRINCIPAL"
    if (SFileGUID<>'') then
      begin
      TobDocFiles := Tob.Create('YDOCFILES', Nil, -1);
      TobDocFiles.PutValue('YDF_DOCGUID', SDocGUID);
      TobDocFiles.PutValue('YDF_FILEGUID', SFileGUID);
      TobDocFiles.LoadDB;
      TobDocFiles.PutValue('YDF_FILEROLE', 'PAL');
      TobDocFiles.InsertOrUpdateDB;
      TobDocFiles.Free;
      end;

    TobDoc.InsertOrUpdateDB;
    TobDocGed.InsertOrUpdateDB;
    Result := SDocGUID;
  except
    if TobDocFiles<>Nil then TobDocFiles.Free;
    LastError := 'Erreur lors de la sauvegarde (InsertDocumentGed).';
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MD
Créé le ...... : 01/01/2005
Modifié le ... : 29/05/2007
Description .. : Extrait le document de la GED et retourne le chemin
Suite ........ : complet du fichier obtenu dans le rép. temporaire
Suite ........ : Attention, ce n'est pas une extraction pour modification
Suite ........ : mais pour lecture ou traitement direct
Suite ........ : (ne pas confondre avec ExtraireDocumentGed :
Suite ........ : extraction dans le répertoire
Suite ........ : TCbpPath.GetCegidUserRoamingAppData)
Mots clefs ... : 
*****************************************************************}
function  ExtraitDocument(SDocGUID : string; var sTitre: String; parle:Boolean=TRUE): String;
var TobDoc, TEnreg : TOB;
    sTempFileName : String;
begin
  Result := '';

  // Chargement tob de l'enreg
  TobDoc := TOB.Create('Les documents', Nil, -1);
  // INNER car il y a toujours un document, même s'il n'a pas de fichiers
  TobDoc.LoadDetailFromSQL('SELECT DPDOCUMENT.*, YDOCUMENTS.*, YDOCFILES.*, YFI_FILEGUID, YFI_FILENAME'
   + ' FROM DPDOCUMENT INNER JOIN YDOCUMENTS ON DPD_DOCGUID=YDO_DOCGUID'
                   + ' LEFT JOIN YDOCFILES ON YDO_DOCGUID=YDF_DOCGUID'
                   + ' LEFT JOIN YFILES ON YDF_FILEGUID=YFI_FILEGUID'
   + ' WHERE DPD_DOCGUID="'+SDocGUID+'" AND YDF_FILEROLE="PAL"');

  if TobDoc.Detail.Count=0 then
    begin
    if parle then // ajout me 23/11/2005
       PGIInfo('Document n° '+SDocGUID+' non trouvé.', 'Extraction d''un document');
    TobDoc.Free;
    exit;
    end;

  // #### 1 seul enreg, tant qu'on ne fait référence qu'à un fichier PRINCIPAL du document
  TEnreg := TobDoc.Detail[0];

  if TEnreg.GetString('YFI_FILEGUID')='' then
    begin
    PGIInfo('Pas de fichier associé au document n° '+TEnreg.GetValue('YDO_DOCGUID'),TEnreg.GetValue('YDO_LIBELLEDOC'));
    TobDoc.Free;
    exit; // SORTIE (pas de document)
    end;

  // FQ 11795 sTitre := 'Document - '+TEnreg.GetValue('YDO_LIBELLEDOC')+' - '+TEnreg.GetValue('YDO_ANNEE');
  sTitre := 'Document - '+TEnreg.GetValue('YDO_LIBELLEDOC'); 

  // Nom du fichier temporaire
  sTempFileName := V_GedFiles.TempPath + TEnreg.GetValue('YFI_FILENAME');
  // si la tentative échoue avec le nom d'origine
  if FileExists(sTempFileName) then
    // on récupère un nom temporaire quelconque, avec la bonne extension
    sTempFileName := V_GedFiles.TempFileName(ExtractFileExt(TEnreg.GetValue('YFI_FILENAME')));
  // Extraction

  if V_GedFiles.Extract(sTempFileName, TEnreg.GetString('YFI_FILEGUID')) then
    Result := sTempFileName;
  TobDoc.Free;
end;

procedure MarquageEwsAPublier(SDocGUID: String; bApublier: Boolean);
var marque: String;
begin
 if bApublier then marque := 'X' else marque := '-';
 ExecuteSQL('UPDATE DPDOCUMENT SET DPD_EWSAPUBLIER="'+marque+'" WHERE DPD_DOCGUID="'+SDocGUID+'"');
end;

//-----------------------------------------------------------------------

function GetCheminGed(sCodeGed: String; bDetaille : Boolean): String;
var Q : TQuery;
    Niveau : Integer;
    CodeGed3,CodeGed2,CodeGed1 : String;
    SQL, CheminGed : String;

// Retourne Armoire \ Classeur \ Intercalaire si bDetaille=True
// ou Code Armoire;Code Classeur;Code Intercalaire si bDetaille=False
begin
  Result := 'EN ATTENTE';
  if sCodeGed='' then exit;

  // Voir P.G. pour utiliser éventuelt fcts des tablettes hiérarchiques
  {SQL := 'SELECT * FROM YGEDDICO AS YG2 WHERE EXISTS('+GetTabletteSQL('YYGEDNIVEAU2',
   DataTypeLink.GetSql(TGN1.CodeGed, False)+' AND (YG2.YGD_CODEGED=YGEDDICO.YGD_CODEGED)',
   True )+')'
   +' ORDER BY YGD_TRIGED';}

  Niveau := 0;CheminGed:='';CodeGed3:=SCodeGed;CodeGed2:='';CodeGed1:='';
  Q := OpenSQL('SELECT YGD_NIVEAUGED FROM YGEDDICO WHERE YGD_CODEGED="'+sCodeGed+'"', True);
  if Not Q.Eof then Niveau := Q.Fields[0].AsInteger;
  Ferme(Q);

  if Niveau=3 then
   begin
    if bDetaille then
     begin
      SQL := GetTabletteSQL('YYGEDNIVEAU3', 'YGD_CODEGED="'+sCodeGed+'"');
      Q := OpenSQL(SQL, True);
      if Not Q.Eof then
        CheminGed := ' \ '+Q.FindField('YGD_LIBELLEGED').AsString;
      Ferme(Q);
     end
    else
     begin
      // recherche code ged père
      SQL := 'SELECT YGD_CODEGED, YGD_LIBELLEGED FROM YGEDDICO WHERE'
             + ' EXISTS (SELECT YDT_CODEHDTLINK FROM ##DP##.YDATATYPETREES'
             + ' WHERE (YDT_CODEHDTLINK = "YYGEDNIV2GEDNIV3") AND (YDT_SCODE = "'+SCodeGed+'")'
             + ' AND (##DP##.YDATATYPETREES.YDT_MCODE = YGEDDICO.YGD_CODEGED) )';
             // rq : autres syntaxes qui fonctionnent malgré la redirection
             // + ' AND (YDATATYPETREES.YDT_MCODE = YGEDDICO.YGD_CODEGED) )';
             // + ' AND (YDT_MCODE = YGD_CODEGED) )';

      Q := OpenSQL(SQL, True);
      if Not Q.Eof then
       CodeGed2:=Q.FindField('YGD_CODEGED').AsString;
      Ferme(Q);
      Niveau := 2;
     end;
   end;

  if Niveau=2 then
   begin
    if bDetaille then
     begin
      SQL := GetTabletteSQL('YYGEDNIVEAU2', 'YGD_CODEGED="'+sCodeGed+'"');
      Q := OpenSQL(SQL, True);
      if Not Q.Eof then
       CheminGed := ' \ '+Q.FindField('YGD_LIBELLEGED').AsString+CheminGed;
      Ferme(Q);
     end
    else
     begin
      if CodeGed2='' then CodeGed2:=SCodeGed;
      // recherche code ged père
      SQL := 'SELECT YGD_CODEGED, YGD_LIBELLEGED FROM YGEDDICO WHERE'
           + ' EXISTS (SELECT YDT_CODEHDTLINK FROM ##DP##.YDATATYPETREES'
           + ' WHERE (YDT_CODEHDTLINK = "YYGEDNIV1GEDNIV2") AND (YDT_SCODE = "'+CodeGed2+'")'
           + ' AND (##DP##.YDATATYPETREES.YDT_MCODE = YGEDDICO.YGD_CODEGED) )';
           // rq : autres syntaxes qui fonctionnent malgré la redirection
           // + ' AND (YDATATYPETREES.YDT_MCODE = YGEDDICO.YGD_CODEGED) )';
           // + ' AND (YDT_MCODE = YGD_CODEGED) )';

      Q := OpenSQL(SQL, True);
      if Not Q.Eof then
       CodeGed1:=Q.FindField('YGD_CODEGED').AsString;
      Ferme(Q);
      Niveau := 1;
     end;
   end;

  if Niveau=1 then
   begin
    if bDetaille then
     begin
      SQL := GetTabletteSQL('YYGEDNIVEAU1', 'YGD_CODEGED="'+SCodeGed+'"');
      Q := OpenSQL(SQL, True);
      if Not Q.Eof then
       CheminGed := Q.FindField('YGD_LIBELLEGED').AsString + CheminGed;
      Ferme(Q);
     end
    else
     begin
      if (CodeGed1='') then CodeGed1:=SCodeGed;
      CheminGed:=CodeGed1+';'+CodeGed2+';'+CodeGed3;
     end
   end;

 Result:=CheminGed;
end;

function  CodeGedToEwsId (CodeGed: String) : String;
// Extrait, si existe, le code Ews associé à la branche Ged
var Q: TQuery;
begin
  Result := '';
  if CodeGed='' then exit;
  Q := OpenSQL('SELECT YGD_EWSID FROM YGEDDICO WHERE YGD_CODEGED="'+CodeGed+'"', True);
  if Not Q.Eof then
    if Not VarIsNull(Q.Fields[0].Value) then Result := Q.Fields[0].AsString;
  Ferme(Q);
end;

// $$$ JP 11/04/06
function  CodeGedToEwsRegle (CodeGed: String):string;
// Extrait, si existe, la regle Ews associé à la branche Ged
var Q: TQuery;
begin
     Result := '';
     if CodeGed='' then exit;
     Q := OpenSQL('SELECT YGD_EWSREGLEPUB FROM YGEDDICO WHERE YGD_CODEGED="'+CodeGed+'"', True);
     if Not Q.Eof then
        if Not VarIsNull(Q.Fields[0].Value) then Result := Q.Fields[0].AsString;
     Ferme(Q);
end;

function ExtensionToLibelleApplication (Extension: string): string;
var SCleDefaut : String;
begin
 Result := 'Document Inconnu';
 with TRegistry.Create do
  try
   RootKey := HKEY_CLASSES_ROOT;
   if OpenKey(Extension, False) then
    SCleDefaut:=ReadString ('');
   if OpenKey('\'+SCleDefaut, False) then
    Result := ReadString('');
  finally
   Free;
  end;
end;

                            {************}
                            { TGedDPNode }
                            {************}

function TGedDPNode.CodeReg (strTypeReg:string):string;
begin
     Result := '';

     if strTypeReg = '+M&A' then
          Result := IntToStr (12*StrToInt (Annee) + StrToInt (Mois))
     else if strTypeReg = '+AUT' then
          Result := Auteur
     else if strTypeReg = '+NAT' then
          Result := Nature; //RechDom ('YYNATDOC', Nature, FALSE);
end;

function TGedDPNode.DisplayReg (strTypeReg:string):string;
begin
     Result := '';

     if strTypeReg = '+M&A' then
          Result := Mois + '/' + Annee
     else if strTypeReg = '+AUT' then
          Result := Auteur
     else if strTypeReg = '+NAT' then
          Result := RechDom ('YYNATDOC', Nature, FALSE);
end;

{$IFNDEF EAGLSERVER}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : DC
Créé le ...... : 01/01/2007
Modifié le ... :   /  /
Description .. : Initialise un objet TInfoDocGed
Suite ........ : à partir d'un document existant
Mots clefs ... :
*****************************************************************}
function InitialiserInfoDocGed (NoDossier : String; DocGuid : String) : TInfoDocGed; OverLoad;
var UneInfoDocGed : TInfoDocGed;
    SSql          : String;
    QSql          : TQuery;
begin
 if (DocGuid<>'') then
  begin
   SSql :=  ' SELECT DPDOCUMENT.*, YDOCUMENTS.*, YDOCFILES.*, YFILES.*'
          + ' FROM DPDOCUMENT, YDOCUMENTS, YDOCFILES, YFILES'
          + ' WHERE YDF_DOCGUID="'+DocGuid+'"'
          + ' AND DPD_DOCGUID=YDO_DOCGUID'
          + ' AND DPD_NODOSSIER="'+NoDossier+'"'
          + ' AND YDO_DOCGUID=YDF_DOCGUID'
          + ' AND YDF_FILEGUID=YFI_FILEGUID'
          + ' AND (YDF_FILEROLE IS NULL OR YDF_FILEROLE="PAL")'
          + ' ORDER BY YDO_DATEMODIF DESC';

   QSql:=OpenSql (SSql,True);
   if not (QSql.Eof) then
    begin
     UneInfoDocGed.CheminFichierGed:=TCbpPath.GetCegidUserRoamingAppData+'\GED\'+DocGuid;
     UneInfoDocGed.NomFichierGed:=QSql.FindField ('YFI_FILENAME').AsString;
     UneInfoDocGed.DocGuid:=DocGuid;
     UneInfoDocGed.FileGuid:=QSql.FindField ('YFI_FILEGUID').AsString;
     UneInfoDocGed.NoDossier:=NoDossier;
     UneInfoDocGed.CodeGed:=QSql.FindField ('DPD_CODEGED').AsString;
     UneInfoDocGed.Libelle:=QSql.FindField ('YDO_LIBELLEDOC').AsString;
     UneInfoDocGed.Auteur:=QSql.FindField ('YDO_AUTEUR').AsString;
    end;
   QSql.Free;
  end;
 Result:=UneInfoDocGed;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : DC
Créé le ...... : 01/01/2007
Modifié le ... :   /  /
Description .. : Initialise un objet TInfoDocGed
Suite ........ : à partir d'un nouveau document sur disque
Mots clefs ... :
*****************************************************************}
function InitialiserInfoDocGed (UnChemin : String; UnFichier : String; UnDossier : String; UnCodeGed : String; UnLibelle : String; UnAuteur : String) : TInfoDocGed; OverLoad;
var UneInfoDocGed : TInfoDocGed;
begin
 UneInfoDocGed.CheminFichierGed:=UnChemin;
 UneInfoDocGed.NomFichierGed:=UnFichier;
 UneInfoDocGed.DocGuid:='';
 UneInfoDocGed.FileGuid:=V_GedFiles.Import (UnChemin+UnFichier);
 UneInfoDocGed.NoDossier:=UnDossier;
 UneInfoDocGed.CodeGed:=UnCodeGed;
 UneInfoDocGed.Libelle:=UnLibelle;
 UneInfoDocGed.Auteur:=UnAuteur;
 Result:=UneInfoDocGed;
end;
{$ENDIF EAGLSERVER}


{***********A.G.L.Privé.*****************************************
Auteur  ...... : DC
Créé le ...... : 01/01/2007
Modifié le ... : 29/05/2007
Description .. : Extrait un document pour modification.
Suite ........ : Paramètres d'extraction dans DocGed :
Suite ........ :   .DocGuid = clé du document
Suite ........ :   .FileGuid = clé du fichier
Suite ........ :   .NomFichierGed = nom physique
Suite ........ :   .CheminFichierGed = répertoire d'extraction
Suite ........ :   (normalement TCbpPath.GetCegidUserRoamingAppData
Suite ........ :   +'\GED\'+DocGuid)
Suite ........ : Voir InitialiserInfosDocGed pour alimenter DocGed
Mots clefs ... : 
*****************************************************************}
function ExtraireDocumentGed (DocGed : TInfoDocGed) : Boolean;
begin
 Result := False;
 if not DirectoryExists (DocGed.CheminFichierGed) then
  if not ForceDirectories(DocGed.CheminFichierGed) then
   begin
    PGIInfo('Impossible de créer le répertoire : '+DocGed.CheminFichierGed);
    exit;
   end;

 if (V_GedFiles.Extract(DocGed.CheminFichierGed+'\'+DocGed.NomFichierGed, DocGed.FileGuid)) then
  begin
   ExecuteSql ('UPDATE YDOCUMENTS SET YDO_DOCSTATE="EXT", YDO_DATEMODIF="'+UsDateTime (Now)+'", YDO_UTILISATEUR="'+V_PGI.USER+'" WHERE YDO_DOCGUID="'+DocGed.DocGuid+'"');
   Result:=True;
  end
 else
  begin
   PgiInfo('Impossible d''extraire le document.', TitreHalley);
   Result:=False;
  end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : DC
Créé le ...... : 01/01/2007
Modifié le ... : 29/05/2007
Description .. : <Fonction non utilisée>
Suite ........ : Insertion d'un document dans la Ged à partir des
Suite ........ : caractéristiques fournies dans l'objet DocGed :
Suite ........ :   .Libelle   = libellé du document
Suite ........ :   .Auteur    = nom libre
Suite ........ :   .NoDossier = dossier rattaché
Suite ........ :   .CodeGed   = emplacement dans l'arborescence Ged
Suite ........ :   .FileGuid  = clé du fichier
Mots clefs ... :
*****************************************************************}
function IntegrerDocumentGed (DocGed : TInfoDocGed) : Boolean;
var TobDoc,TobDocGed :Tob;
    Erreur : String;
begin
 Result:=False;

 if DocGed.FileGUID <>'' then
  begin
   DocGed.DocGUID := '';
   TobDoc    := Tob.Create('YDOCUMENTS', nil, -1);
   TobDocGed := Tob.Create('DPDOCUMENT', nil, -1);

   try
    TobDoc.LoadDb;
    TobDoc.PutValue('YDO_LIBELLEDOC',DocGed.Libelle);
    TobDoc.PutValue('YDO_AUTEUR',DocGed.Auteur);

    TobDocGed.LoadDb;
    TobDocGed.PutValue('DPD_NODOSSIER',DocGed.NoDossier);
    TobDocGed.PutValue('DPD_CODEGED',DocGed.CodeGed);

    DocGed.DocGUID := InsertDocumentGed (TobDoc, TobDocGed, DocGed.FileGUID, Erreur);
   finally
    TobDoc.Free;
    TobDocGed.Free;
   end;

   if DocGed.DocGUID='' then
    begin
     PgiInfo (Erreur,TitreHalley);
     V_GedFiles.Erase (DocGed.FileGUID);
     Result := False;
    end
   else
    Result := True;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : DC
Créé le ...... : 01/01/2007
Modifié le ... : 29/05/2007
Description .. : <Fonction non utilisée>
Suite ........ : Réintègre en Ged un document extrait par la
Suite ........ : fonction ExtraireDocument (c'est à dire extrait
Suite ........ : pour modification dans le répertoire
Suite ........ : TCbpPath.GetCegidUserRoamingAppData)
Mots clefs ... :
*****************************************************************}
function ReintegrerDocumentGed (DocGed : TInfoDocGed) : Boolean;
var UnGuid : String;
begin
 Result:=False;
 if not FileExists (DocGed.CheminFichierGed+'\'+DocGed.NomFichierGed) then
  exit
 else
  begin
   UnGuid:=V_GedFiles.Import (DocGed.CheminFichierGed+'\'+DocGed.NomFichierGed);
   if (UnGuid<>'') then
    begin
     ExecuteSql ('UPDATE YDOCFILES SET YDF_FILEGUID="'+UnGuid+'" WHERE YDF_DOCGUID="'+DocGed.DocGuid+'"');
     if (DeleteFile (PChar (DocGed.CheminFichierGed+'\'+DocGed.NomFichierGed))) then
      begin
       V_GedFiles.Erase (DocGed.FileGuid);
       ExecuteSql ('UPDATE YDOCUMENTS SET YDO_DOCSTATE="", YDO_DATEMODIF="'+UsDateTime (Now)+'", YDO_UTILISATEUR="'+V_PGI.USER+'" WHERE YDO_DOCGUID="'+DocGed.DocGuid+'"');
       Result:=True;
      end
     else
      begin
       PgiInfo('Impossible de réintégrer, le document '+DocGed.NomFichierGed+' est peut-être en cours d''utilisation.', TitreHalley);
       Result:=False;
      end;
    end;
  end;
end;

//------------------------------
//--- Nom : OuvrirDocumentGed
//------------------------------
procedure OuvrirDocumentGed (DocGed : TInFoDocGed);
begin
 ShellExecute(0, PChar('open'), PChar (DocGed.CheminFichierGed+'\'+DocGed.NomFichierGed), nil, nil, SW_RESTORE);
end;

//-----------------------------
//--- Nom : CreerDocumentDoc
//-----------------------------
function CreerDocumentDOC (UnChemin : String; UnFichier : String) : Boolean;
var MsWord                : OleVariant;
    MsWordDocs, MsWordDoc : OleVariant;
begin
 Result:=False;
 try
  MsWord:=GetActiveOleObject('Word.Application');
 except
  on EOleSysError do
   try
    ArreterProcess ('winword.exe');
    MsWord:=CreateOleObject ('Word.Application');
   except
    MsWord:=UnAssigned;
   end;
 end;

 MsWordDocs:=MSWord.Documents;
 MsWordDoc:=MSWordDocs.add;
 MsWordDoc.SaveAs (UnChemin+UnFichier);
 MsWordDoc.Close (True);
 MsWord:=Unassigned;

 if FileExists (UnChemin+UnFichier) then Result:=True;
end;

//-----------------------------
//--- Nom : CreerDocumentXls
//-----------------------------
function CreerDocumentXLS (UnChemin : String; UnFichier : String) : Boolean;
var MsExcel                 : OleVariant;
    MsExcelDocs, MsExcelDoc : OleVariant;
begin
 Result:=False;
 try
  MsExcel:=GetActiveOleObject('Excel.Application');
 except
  on EOleSysError do
   try
    ArreterProcess ('Excel.exe');
    MsExcel:=CreateOleObject ('Excel.Application');
   except
    MsExcel:=UnAssigned;
   end;
 end;

 MsExcelDocs:=MSExcel.Workbooks;
 MsExcelDoc:=MSExcelDocs.add;
 MsExcelDoc.SaveAs (UnChemin+UnFichier);
 MsExcelDoc.Close (True);
 MsExcel:=Unassigned;

 if FileExists (UnChemin+UnFichier) then Result:=True;
end;

//---------------------------------
//--- Nom : ArreterProcess
//---------------------------------
function ArreterProcess(const ProcessName : string) : boolean;
var ProcessEntry32 : TProcessEntry32;
    HSnapShot : THandle;
    HProcess : THandle;
begin
 Result := False;

 HSnapShot := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
 if HSnapShot = 0 then exit;

 ProcessEntry32.dwSize := sizeof(ProcessEntry32);
 if Process32First(HSnapShot, ProcessEntry32) then
  repeat
   if CompareText(ProcessEntry32.szExeFile, ProcessName) = 0 then
    begin
     HProcess := OpenProcess(PROCESS_TERMINATE, False, ProcessEntry32.th32ProcessID);
     if HProcess <> 0 then
      begin
       Result := TerminateProcess(HProcess, 0);
       CloseHandle(HProcess);
      end;
     Break;
    end;
  until not Process32Next(HSnapShot, ProcessEntry32);

 CloseHandle(HSnapshot);
end;

//--------------------------------------
//--- Nom : IsValideParametreAppelOLE
//--------------------------------------
function IsValideParametreGedOLE (UnUtilisateur : String; UnCodeGed : String; UnRepertoire : String) : Boolean;
begin
 Result:=True;

 //--- Vérification de l'existence du code user
 if not ExisteSql ('SELECT 1 FROM UTILISAT WHERE US_UTILISATEUR="'+UnUtilisateur+'"') then
  begin
   Result:=False;
   Exit;
  end;

 //--- Vérification de l'existence de l'arborescence Ged indiquée
 if not ExisteSql ('SELECT 1 FROM YGEDDICO WHERE YGD_CODEGED="'+UnCodeGed+'"') then
  begin
   Result:=False;
   Exit;
  end;

 //--- Vérification du répertoire Path indiqué
 if not DirectoryExists (UnRepertoire) then
  if not ForceDirectories(UnRepertoire) then
   begin
    PGIInfo('Le répertoire '+UnRepertoire+' n''existe pas.', TitreHalley);
    Result:=False;
    exit;
   end;
end;

//------------------------------------
//--- Nom : ExtraireDocumentsGedOLE
//------------------------------------
function ExtraireDocumentsGedOLE (UnUtilisateur : String; UnDossier : String; UnCodeGed : String; UnRepertoire : String) : Boolean;
var IndiceNiv1,IndiceNiv2,IndiceNiv3          : Integer;
    UnCodeGedNiv1,UnCodeGedNiv2,UnCodeGedNiV3 : String;
    CodeGedNiv1,CodeGedNiV2,CodeGedNiv3       : String;
    TobNiv1, TobNiv2, TobNiv3                 : TOB;
    DataTypeLink                              : TYDataTypeLink;
    DataTypeLink2                             : TYDataTypeLink;
    SSql, UnCodeArboGed, CodeGed              : String;
    TmpIndex                                  : Integer;
    Resultat                                  : Boolean;
begin
 Resultat:=False;
 UnCodeArboGed:=copy (UnCodeGed,1,1);
 UnCodeGedNiv1:=copy (UnCodeGed,2,1);
 UnCodeGedNiv2:=copy (UnCodeGed,3,2);
 UnCodeGedNiv3:=copy (UnCodeGed,5,2);

 //--- Niveau1 : Armoire
 SSql:='SELECT * FROM YGEDDICO AS YG1 WHERE EXISTS('+GetTabletteSQL('YYGEDNIVEAU1', 'YG1.YGD_CODEGED=YGEDDICO.YGD_CODEGED', True )+')'+
       ' AND YGD_ARBOGED="'+UnCodeArboGed+'" ORDER BY YGD_TRIGED';
 TobNiv1 := Tob.Create('Ged Niveau 1', nil, -1);
 TobNiv1.LoadDetailFromSQL(SSql);

 DataTypeLink := V_DataTypeLinks.FirstMaster('YYGEDNIVEAU1', tmpIndex);
 DataTypeLink2 := V_DataTypeLinks.FirstMaster('YYGEDNIVEAU2', tmpIndex);

 for IndiceNiv1:=0 to TobNiv1.Detail.Count-1 do
  begin
   CodeGed:=UnCodeArboGed+UnCodeGedNiv1;
   CodeGedNiv1:=TobNiv1.Detail[IndiceNiv1].GetValue('YGD_CODEGED');

   if (CodeGed=CodeGedNiv1) then
    begin
     Resultat:=ExtraireDocGed (UnDossier,CodeGedNiv1,UnRepertoire);
     //--- Niveau 2 : Classeur
     SSql:='SELECT * FROM YGEDDICO AS YG2 WHERE EXISTS('+GetTabletteSQL('YYGEDNIVEAU2',DataTypeLink.GetSql(CodeGedNiv1, False)+
           ' AND (YG2.YGD_CODEGED=YGEDDICO.YGD_CODEGED)',True )+') AND YGD_ARBOGED="'+UnCodeArboGed+'" ORDER BY YGD_TRIGED';
     TobNiv2:=Tob.Create('Ged Niveau 2', nil, -1);
     TobNiv2.LoadDetailFromSQL(SSql);

     for IndiceNiv2:=0 to TobNiv2.Detail.Count-1 do
      begin
       CodeGed:=UnCodeArboGed+UnCodeGedNiv1+UnCodeGedNiv2;
       CodeGedNiv2:=TobNiv2.Detail[IndiceNiv2].GetValue('YGD_CODEGED');

       if (CodeGed=CodeGedNiv2) or (UnCodeGedNiv2='') then
        begin
         Resultat:=ExtraireDocGed (UnDossier,CodeGedNiv2,UnRepertoire);

         //--- Niveau 3 : Intercalaire
         SSql:='SELECT * FROM YGEDDICO AS YG3 WHERE EXISTS('+GetTabletteSQL('YYGEDNIVEAU3',DataTypeLink2.GetSql(CodeGedNiv2, False)+
               ' AND (YG3.YGD_CODEGED=YGEDDICO.YGD_CODEGED)',True )+')'+' AND YGD_ARBOGED="'+UnCodeArboGed+'"'+' ORDER BY YGD_TRIGED';
         TobNiv3 := Tob.Create('Ged Niveau 3', nil, -1);
         TobNiv3.LoadDetailFromSQL(SSql);

         for IndiceNiv3:=0 to TobNiv3.Detail.Count-1 do
          begin
           CodeGedNiv3:=TobNiv3.Detail[IndiceNiv3].GetValue('YGD_CODEGED');
           CodeGed:=UnCodeArboGed+UnCodeGedNiv1+UnCodeGedNiv2+UnCodeGedNiv3;
           if (CodeGed=UnCodeGedNiv3) or (UnCodeGedNiv3='') then
            Resultat:=ExtraireDocGed (UnDossier,CodeGedNiv3,UnRepertoire);
          end;
         TobNiv3.Free;
        end;
      end;
     TobNiv2.Free;
    end;
  end;
 TobNiv1.Free;

 Result:=Resultat;
end;

//---------------------------
//--- Nom : ExtraireDocGed
//---------------------------
function ExtraireDocGed (UnDossier : String; UnCodeGed : String; UnRepertoire : String) : Boolean;
var SArmoire, SClasseur, SIntercalaire : String;
    SSql,UnNiveau                      : String;
    UneTob                             : Tob;
    Indice                             : Integer;
begin
 Result:=True;

 //--- Constitution de l'arborescence
 SArmoire:=copy (UnCodeGed,1,2);
 SClasseur:=copy (UnCodeGed,3,2);
 if (SClasseur<>'') then
  begin
   SIntercalaire:=copy (UnCodeGed,5,2);
   if (SIntercalaire<>'') then
    UnNiveau:=SArmoire+'\'+SClasseur+'\'+SIntercalaire
   else
    UnNiveau:=SArmoire+'\'+SClasseur;
  end
 else
  UnNiveau:=Sarmoire;

 //--- Extraction des documents GED
 SSql:=' SELECT YFI_FILENAME,YFI_FILEGUID,YDO_DOCGUID FROM DPDOCUMENT, YDOCUMENTS, YDOCFILES, YFILES'+
       ' WHERE DPD_DOCGUID=YDO_DOCGUID AND YDO_DOCGUID=YDF_DOCGUID AND YDF_FILEGUID=YFI_FILEGUID'+
       ' AND DPD_NODOSSIER="'+UnDossier+'"'+
       ' AND DPD_CODEGED="'+UnCodeGed+'"'+
       ' AND (YDF_FILEROLE IS NULL OR YDF_FILEROLE="PAL") ORDER BY YDO_DATEMODIF DESC';

 UneTob:=TOB.Create('ListeDocument', Nil, -1);
 UneTob.LoadDetailFromSQL(SSql);

 if (UneTob.Detail.Count>0) then
  begin
   UnRepertoire:=UnRepertoire+'\'+UnDossier+'\'+UnNiveau;
   if not DirectoryExists (UnRepertoire) then ForceDirectories(UnRepertoire);
  end;

 for Indice:=0 To UneTob.Detail.Count-1 Do
  begin
   if not (V_GedFiles.Extract(UnRepertoire+'\'+UneTob.Detail[Indice].GetValue ('YFI_FILENAME'),UneTob.Detail[Indice].GetValue ('YFI_FILEGUID'))) then
    begin
     PgiInfo('Impossible d''extraire le document : '+UneTob.Detail[Indice].GetValue ('YFI_FILENAME'), TitreHalley);
     Result:=False;
     Continue;
    end;
  end;

 UneTob.Free;
end;


//----------------------------------------------------
//--- Nom   : TraiteAjoutFichierDansGed
//--- Objet :
//----------------------------------------------------
function TraiteAjoutFichierDansGed (Fichier, strDossier, CodeGed, Description, Auteur, Annee, Mois : String; majwinner:Boolean): Variant; OverLoad;
var
   SFileGUID, SDocGUID : String;
   TobDoc,TobDocGed :Tob;
   Erreur           :string;
begin
{     // Par défaut, fichier non intégré en GED Cegid
     Result := False;

     //PGR 09/2007 Traitement des éléments GED avec la migration WINNER
     //majwinner := false;

     if (GetParamSocSecur ('SO_MDLIENWINNER', False)) and (Trim(CodeTiers) = 'WINNER') then
     begin
       majwinner := true;
       CodeTiers := '';
     end;

     // PHC 23/01/06
     if CodeDossier = '' then
        strDossier := Trim (DonnerDossierFromGuid (DonnerGuidFromTiers (CodeTiers)))
     else
        strDossier := Trim (CodeDossier);

     // $$$ JP 06/10/05 - si code dossier inconnu par rapport au tiers fourni, on n'autorise pas l'insertion du document
     if strDossier = '' then exit;
}
     // Par défaut, fichier non intégré en GED Cegid
     Result := False;

     // Insertion du document
     SFileGUID:=V_GedFiles.Import (Fichier);
     if SFileGUID <>'' then
     begin
          SDocGUID  := '';
          TobDoc    := Tob.Create('YDOCUMENTS', nil, -1);
          TobDocGed := Tob.Create('DPDOCUMENT', nil, -1);
          try
              TobDoc.LoadDb;
              TobDoc.PutValue('YDO_LIBELLEDOC',Copy(Description,1,70) );
              TobDoc.PutValue('YDO_AUTEUR',    Copy(Auteur,1,35)      );
              TobDoc.PutValue('YDO_ANNEE',     Copy(Annee,1,4)        );
              TobDoc.PutValue('YDO_MOIS',      Copy(Mois,1,2)         );
              //PGR 09/2007 Traitement des éléments GED avec la migration WINNER
              if majwinner = true then
                TobDoc.PutValue('YDO_DATEFIN',   '31/12/2099'        );

              TobDocGed.LoadDb;
              TobDocGed.PutValue('DPD_NODOSSIER', strDossier);
              TobDocGed.PutValue('DPD_CODEGED', CodeGed);
              //PGR 01/2008 Traitement des éléments GED WINNER
              //PGR 09/2007 Traitement des éléments GED avec la migration WINNER
              //if majwinner = true then
              //  TobDocGed.PutValue('DPD_ARBOGED','D')
              TobDocGed.PutValue('DPD_ARBOGED',Copy(CodeGed,1,1));

              SDocGUID := InsertDocumentGed (TobDoc, TobDocGed, SFileGUID, Erreur);
              //PGR 01/2008 Traitement des éléments GED WINNER
              //ExecuteSQL('UPDATE DPDOCUMENT SET DPD_ARBOGED="A" WHERE DPD_CODEGED<>"###" AND DPD_ARBOGED=""');
              ExecuteSQL('UPDATE DPDOCUMENT SET DPD_ARBOGED=SUBSTRING(DPD_CODEGED,1,1) WHERE DPD_CODEGED<>"###" AND DPD_ARBOGED=""');
            finally
              TobDoc.Free;
              TobDocGed.Free;
          end;

          // Si fichier non référencé, on l'enlève de la Ged
          if SDocGUID='' then
          begin
              V_GedFiles.Erase (SFileGUID);
              if Erreur <> '' then PgiInfo (Erreur, TitreHalley);
          end
          else
              Result := True;
     end;

end;

end.




