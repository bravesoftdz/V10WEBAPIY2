{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 23/07/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : YMESSAGES (YMESSAGES)
Mots clefs ... : TOM;YMESSAGES
*****************************************************************}
Unit YMessages_Tom ;

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
  TOM_YMESSAGES = Class (TOM)
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
      {1}  'Erreur de suppression dans YMSGFILES',
      {2}  'Erreur de suppression dans YMSGADDRESS'
      );

procedure TOM_YMESSAGES.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_YMESSAGES.OnDeleteRecord ;
var
  sMsgGuid : String;
  tMsgFiles, tMsgAddress, tFiles: Tob;
begin
  Inherited ;
  LastError := 0;
  sMsgGuid := GetField('YMS_MSGGUID');

  //Begin : Liste des tables utilisants des fichiers YMESSAGES

  {  if ExisteSQL('SELECT 1 FROM ... WHERE ..._MSGGUID ="'+sMsgGuId+'"') then
    LastError := 2; }

  //End : Liste des tables utilisants des fichiers YMESSAGES

  if  LastError = 0 then
  begin
    // Chargement pour invoquer la suppression via la tom yfiles
    tFiles := Tob.Create('_YFILES_', nil, -1);
    tFiles.LoadDetailDBFromSQL('YFILES', 'SELECT * FROM YFILES WHERE' +
                              ' EXISTS(SELECT YMG_MSGGUID FROM YMSGFILES' +
                                        ' WHERE (YMG_MSGGUID = "'+sMsgGuid+'")' +
                                        ' AND (YMG_FILEGUID = YFI_FILEGUID))');

    tMsgFiles := Tob.Create('_YMSGFILES_', nil, -1);
    tMsgFiles.LoadDetailDBFromSQL('YMSGFILES', 'SELECT * FROM YMSGFILES WHERE (YMG_MSGGUID = "' +sMsgGuid+ '")');

    tMsgAddress := Tob.Create('_YMSGADDRESS_', nil, -1);
    tMsgAddress.LoadDetailDBFromSQL('YMSGADDRESS', 'SELECT * FROM YMSGADDRESS WHERE (YMR_MSGGUID = "' +sMsgGuid+ '")');

    // cas d'erreur de transaction
    if not tMsgFiles.DeleteDB then
      LastError := 1
    else
      begin
      if not tMsgAddress.DeleteDB then
        LastError := 2
      else
        tFiles.DeleteDBTom('YFILES');
      end;

    tMsgAddress.Free;
    tMsgFiles.Free;
    tFiles.Free;
  end;

  if LastError <> 0 then
  begin
    LastErrorMsg := TexteMessage[LastError];
  end;
end ;

procedure TOM_YMESSAGES.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_YMESSAGES.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_YMESSAGES.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_YMESSAGES.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_YMESSAGES.OnArgument ( S: String ) ;
begin
  Inherited ;
end ;

procedure TOM_YMESSAGES.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_YMESSAGES.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_YMESSAGES ] ) ; 
end.
