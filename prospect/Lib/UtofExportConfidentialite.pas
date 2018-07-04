
{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 10/04/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : RTEXPORTCONFID ()
Mots clefs ... : TOF;RTEXPORTCONFID
*****************************************************************}
Unit UtofExportConfidentialite ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     FE_Main,db,
     {$IFNDEF DBXPRESS} {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
     Maineagl,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,Utob
     ,M3FP,Ed_tools
         ;
Type
  TOF_YYEXPORTCONFID = Class (TOF)
  end ;
  TOF_YYIMPORTCONFID = Class (TOF)
  end ;

Function YYLanceFiche_ExportConfidentialite(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
Function YYLanceFiche_ImportConfidentialite(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

implementation

Function YYLanceFiche_ExportConfidentialite(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
V_PGI.ZoomOLE := true;  //Affichage en mode modal
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
V_PGI.ZoomOLE := false;  //Affichage en mode modal
end;

Function YYLanceFiche_ImportConfidentialite(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
V_PGI.ZoomOLE := true;  //Affichage en mode modal
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
V_PGI.ZoomOLE := false;  //Affichage en mode modal
end;

function AGLYYEXPORTCONFID( parms: array of variant; nb: integer ) : variant;
var FileName : String ;
    Q : TQuery;
    TobMenu : Tob;
    F : TForm;
BEGIN
  F:=TForm(Longint(Parms[0])) ;
  Result:=0;
  FileName:=string(parms[1]) ;
  Q := OpenSQL('SELECT MN_TAG,MN_ACCESGRP,MN_1 FROM MENU WHERE MN_TAG<>0'{+' AND MN_ACCESGRP LIKE "%-%"'},True) ;
  if Not Q.EOF then
  begin
    TobMenu:=TOB.create ('Export Confidentialité',NIL,-1);
    TobMenu.LoadDetailDB('','','',Q, False,true) ;
    InitMoveProgressForm (F, 'Sauvegarde Menu','Sauvegarde Confidentialité', TobMenu.Detail.Count-1,false,false) ;
    TobMenu.SaveToFile (FileName, false, false, true);
    FiniMoveProgressForm;
    TobMenu.Free;
  end;
  Ferme(Q);
END;
function AGLYYIMPORTCONFID( parms: array of variant; nb: integer ) : variant;
var FileName : String ;
    TobMenu : Tob;
    i : integer;
    F : TForm;
BEGIN
  F:=TForm(Longint(Parms[0])) ;
  Result:=0;
  FileName:=string(parms[1]) ;
  TobMenu:=TOB.create ('Import Confidentialité',NIL,-1);
  if FileExists (FileName) then
     TobLoadFromFile( FileName, Nil, TobMenu ) ;
  if (TobMenu<>nil) then
  begin
    InitMoveProgressForm (F, 'Mise à jour Menu','Mise à jour Confidentialité', TobMenu.Detail.Count-1,false,false) ;
    for i:=0 to TobMenu.Detail.Count-1 Do
    begin
      ExecuteSQL('UPDATE MENU SET MN_ACCESGRP="'+TobMenu.Detail[i].GetValue('MN_ACCESGRP')+'" WHERE MN_TAG='+IntToStr(TobMenu.Detail[i].GetValue('MN_TAG'))+' AND MN_1='+IntToStr(TobMenu.Detail[i].GetValue('MN_1'))) ;
      MoveCurProgressForm ('');
    end;
    FiniMoveProgressForm;
  end;
  TobMenu.Free;
END;

Initialization
  registerclasses ( [ TOF_YYEXPORTCONFID ] ) ;
  registerclasses ( [ TOF_YYIMPORTCONFID ] ) ;  
  RegisterAglFunc( 'YYExportConfid', true,1,AGLYYEXPORTCONFID) ;
  RegisterAglFunc( 'YYImportConfid', true,1,AGLYYIMPORTCONFID) ;
end.
