{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/04/2003
Modifié le ... :  07 /02  /2006  PCS adaptation  GED GUID
Description .. : Source TOM de la TABLE : YFILES (YFILES)
Mots clefs ... : TOM;YFILES
*****************************************************************}
Unit YFiles_Tom ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob ;

Type
  TOM_YFILES = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    function ChangeFileID(OldFileGUID, NewFileGUID: String): Boolean;


    end ;

Implementation

uses UGedFiles;

const
	TexteMessage: array[1..6] of string =
      (
      {1}  'Erreur de suppression du FileParts',
      {2}  'Erreur de suppression dans une des tables en relation',
      {3}  'Utilisé dans la table YMAILFILES',
      {4}  'Utilisé dans la table YMSGFILES',
      {5}  'Utilisé dans la table YDOCFILES',
      {6}  'Utilisé dans la table NBOOKDFPF'
      // #### Ajouter ici les tables qui pointent sur YFILES avec une clé étrangère FILEGUID
      // #### manque les tables de paie (non communes) : PCURRICULUM, PSUIVITCANDIDAT, PSUIVICV
       );

procedure TOM_YFILES.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_YFILES.OnDeleteRecord ;
var
  FileGUID: String;


begin
  Inherited ;
  LastError := 0;

  FileGUID := GetField('YFI_FILEGUID');

  //Begin : Liste des tables utilisants des fichiers YFILES

  try
    if ExisteSQL('SELECT 1 FROM YMAILFILES WHERE (YMF_FILEGUID ="' + FileGUID + '")') then
      LastError := 3;
    if ExisteSQL('SELECT 1 FROM YMSGFILES WHERE (YMG_FILEGUID ="' + FileGUID + '")') then
      LastError := 4;
    if ExisteSQL('SELECT 1 FROM YDOCFILES WHERE (YDF_FILEGUID ="' + FileGUID + '")') then
      LastError := 5;
    if ExisteSQL('SELECT 1 FROM NBOOKDFPF WHERE (NBF_FILEGUID ="' + FileGUID + '")') then
      LastError := 6;
    // #### manque les tables de paie (non communes) : PCURRICULUM, PSUIVITCANDIDAT, PSUIVICV
  except
    LastError := 2;
  end;

  //End : Liste des tables utilisants des fichiers YFILES

  if  LastError = 0 then
  begin
    if not V_GedFiles.EraseFileParts(FileGUID) then
      LastError := 1;
  end;

  if LastError <> 0 then
  begin
    LastErrorMsg := texteMessage[LastError];
  end;
end ;

procedure TOM_YFILES.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_YFILES.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_YFILES.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_YFILES.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_YFILES.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_YFILES.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_YFILES.OnCancelRecord ;
begin
  Inherited ;
end ;

function TOM_YFILES.ChangeFileID(OldFileGUID, NewFileGUID: String): Boolean;
begin
  BeginTrans;

  try
    V_GedFiles.ChangeFileIDParts(OldFileGUID, NewFileGUID);

    ExecuteSql('UPDATE YFILES SET YFI_FILEGUID = "' + NewFileGUID + '" WHERE (YFI_FILEGUID ="' + OldFileGUID + '")');
    ExecuteSql('UPDATE YMAILFILES SET YMF_FILEGUID = "' + NewFileGUID + '" WHERE (YMF_FILEGUID ="' + OldFileGUID + '")');

    CommitTrans;
    Result := True;
  except
    RollBack;
    Result := False;
  end;
end;


Initialization
  registerclasses ( [ TOM_YFILES ] ) ;
end.

