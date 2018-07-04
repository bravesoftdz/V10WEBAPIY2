{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : YMAILS (YMAILS)
Mots clefs ... : TOM;YMAILS
*****************************************************************}
Unit YMails_Tom ;

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
  TOM_YMAILS = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    end ;

Implementation

uses YFiles_Tom; // pour ne pas l'oublier

const
	TexteMessage: array[1..2] of string =
      (
      {1}  'Erreur de suppression du YMailFiles',
      {2}  'Utilisé dans la table YMESSAGES'
      );

procedure TOM_YMAILS.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_YMAILS.OnDeleteRecord ;
var
  sUser: string;
  iMailID: Integer;

  tMailFiles, tFiles: Tob;
begin
  Inherited ;
  LastError := 0;

  sUser := GetField('YMA_UTILISATEUR');
  iMailID := GetField('YMA_MAILID');

  //Begin : Liste des tables utilisants des fichiers YMAILS

  if ExisteSQL('SELECT 1 FROM YMESSAGES WHERE (YMS_USERMAIL="'+sUser+'")'
   +' AND (YMS_MAILID ='+IntToStr(iMailID)+')') then
    LastError := 2;

  //End : Liste des tables utilisants des fichiers YMAILS

  if  LastError = 0 then
  begin
    // Chargement pour invoquer la suppression via la tom yfiles
    tFiles := Tob.Create('_YFILES_', nil, -1);
    tFiles.LoadDetailDBFromSQL('YFILES', 'SELECT * FROM YFILES WHERE' +
                              ' EXISTS(SELECT YMF_UTILISATEUR, YMF_MAILID FROM YMAILFILES' +
                                        ' WHERE (YMF_UTILISATEUR = "' + sUser + '") AND (YMF_MAILID = ' + IntToStr(iMailID) + ')' +
                                        ' AND (YMF_FILEGUID = YFI_FILEGUID))');

    tMailFiles := Tob.Create('_YMAILFILES_', nil, -1);
    tMailFiles.LoadDetailDBFromSQL('YMAILFILES', 'SELECT * FROM YMAILFILES WHERE (YMF_UTILISATEUR = "' + sUser + '") AND (YMF_MAILID = ' + IntToStr(iMailID) + ')');

    // cas d'erreur de transaction
    if not tMailFiles.DeleteDB then
      LastError := 1
    else
      tFiles.DeleteDBTom('YFILES');

    tMailFiles.Free;
    tFiles.Free;
  end;

  if LastError <> 0 then
  begin
    LastErrorMsg := TexteMessage[LastError];
  end;
end ;

procedure TOM_YMAILS.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_YMAILS.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_YMAILS.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_YMAILS.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_YMAILS.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_YMAILS.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_YMAILS.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_YMAILS ] ) ; 
end.

