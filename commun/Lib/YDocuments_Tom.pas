{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 23/07/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : YDOCUMENTS (YDOCUMENTS)
Mots clefs ... : TOM;YDOCUMENTS
*****************************************************************}
unit YDocuments_Tom;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbtables {BDE}, {$ELSE}uDbxDataSet, {$ENDIF}
  Fiche,
  FichList,
{$ELSE}
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
  UTob;

type
  TOM_YDOCUMENTS = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  end;

implementation

uses YFiles_Tom; // pour ne pas l'oublier

const
  TexteMessage: array[1..4] of string =
  (
      {1}'Erreur de suppression dans YDOCFILES',
      {2}'Utilisé dans la table DPDOCUMENT',
      {3}'Utilisé dans la table YMODELES',
      {4}'Utilisé dans la table RTDOCUMENT' {rq : non commune}
      // #### Ajouter ici les tables qui pointent sur YDOCUMENTS avec une clé étrangère DOCGUID
      // #### manque les tables de compta (non communes) : ECRCOMPL, IMMO
    );

procedure TOM_YDOCUMENTS.OnNewRecord;
begin
  inherited;
end;

procedure TOM_YDOCUMENTS.OnDeleteRecord;
var
  sDocGuid: string;
  tDocFiles, tFiles: Tob;
begin
  inherited;
  LastError := 0;

  sDocGuid := GetField('YDO_DOCGUID');

  //Begin : Liste des tables utilisants des enreg YDOCUMENTS

  if ExisteSQL('SELECT 1 FROM DPDOCUMENT WHERE DPD_DOCGUID ="' + sDocGuid + '"') then
    LastError := 2;

  if ExisteSQL('SELECT 1 FROM YMODELES WHERE YMO_DOCGUID ="' + sDocGuid + '"') then
    LastError := 3;

  if ExisteSQL('SELECT 1 FROM RTDOCUMENT WHERE RTD_DOCGUID ="' + sDocGuid + '"') then
    LastError := 4;

  //End : Liste des tables utilisants des enreg YDOCUMENTS

  if LastError = 0 then
  begin
    // Chargement pour suppression dans les tables dépendantes via tom
    tFiles := Tob.Create('_YFILES_', nil, -1);
    tFiles.LoadDetailDBFromSQL('YFILES', 'SELECT * FROM YFILES WHERE' +
      ' EXISTS(SELECT YDF_DOCGUID FROM YDOCFILES' +
      ' WHERE (YDF_DOCGUID = "' + sDocGuid + '")' +
      ' AND (YDF_FILEGUID = YFI_FILEGUID))');

    tDocFiles := Tob.Create('_YDOCFILES_', nil, -1);
    tDocFiles.LoadDetailDBFromSQL('YDOCFILES', 'SELECT * FROM YDOCFILES WHERE (YDF_DOCGUID = "' + sDocGuid + '")');

    // cas d'erreur de transaction
    if not tDocFiles.DeleteDB then
      LastError := 1
    else
      tFiles.DeleteDBTom('YFILES');

    tDocFiles.Free;
    tFiles.Free;
  end;

  if LastError <> 0 then
  begin
    LastErrorMsg := TexteMessage[LastError];
  end;
end;

procedure TOM_YDOCUMENTS.OnUpdateRecord;
begin
  inherited;
end;

procedure TOM_YDOCUMENTS.OnAfterUpdateRecord;
begin
  inherited;
end;

procedure TOM_YDOCUMENTS.OnLoadRecord;
begin
  inherited;
end;

procedure TOM_YDOCUMENTS.OnChangeField(F: TField);
begin
  inherited;
end;

procedure TOM_YDOCUMENTS.OnArgument(S: string);
begin
  inherited;
end;

procedure TOM_YDOCUMENTS.OnClose;
begin
  inherited;
end;

procedure TOM_YDOCUMENTS.OnCancelRecord;
begin
  inherited;
end;

initialization
  registerclasses([TOM_YDOCUMENTS]);
end.
