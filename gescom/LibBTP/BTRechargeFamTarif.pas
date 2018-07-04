unit BTRechargeFamTarif;

interface

Uses StdCtrls, 
     Controls,
     Classes,
     grids,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_Main,
{$else}
		 Maineagl,
     eMul,
     uTob,
{$ENDIF}
		 AglInit,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOB,
     HTB97,
     UTOF ;

  	Procedure InitialisationTARIF;
    
implementation

{ BTRechargeFamTarif }

procedure InitialisationTARIF;
Var StSQL       : String;
    QQ          : TQUERY;
    Ind         : Integer;
    TobDetail   : TOB;
    TOBFAMTAR   : TOB;
    TOBSFAMTAR  : TOB;
    TOBCHOIXCOD : TOB;
begin

  if not ExisteSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="TAR"') then exit;
  //
  TOBFAMTAR := TOB.Create('FAMILLE TARIF', nil, -1);
  TOBSFAMTAR:= TOB.Create('SOUS-FAMILLE TARIF', nil, -1);

  //Lecture de la tablette Famille TARIF
  StSQL := 'SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="TAR"';
  QQ := OpenSQL(StSQL,true,-1,'',true);

  If Not QQ.Eof then
  begin
    TOBCHOIXCOD := TOB.Create('TABLETTE', nil, -1);
    TOBCHOIXCOD.LoadDetailDB('CHOIXCOD','','',QQ, False);
    For Ind := 0 to TOBCHOIXCOD.Detail.count - 1 do
    begin
      TobDetail := TOB.Create('BTFAMILLETARIF',TOBFAMTAR,-1);
      TOBDetail.PutValue('BFT_FAMILLETARIF', TobCHOIXCOD.Detail[Ind].GetString('CC_CODE'));
      TOBDetail.PutValue('BFT_LIBELLE'     , TobCHOIXCOD.Detail[Ind].GetString('CC_LIBELLE'));
    end;
    FreeAndNil(TOBCHOIXCOD);
  end;

  Ferme (QQ);

  TOBFAMTAR.SetAllModifie(True);
  TOBFAMTAR.InsertOrUpdateDB;

  //Lecture de la tablette Sous-Famille TARIF
  //Lecture de la tablette Famille TARIF
  StSQL := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="BFT"';
  QQ := OpenSQL(StSQL,true,-1,'',true);

  If Not QQ.Eof then
  begin
    TOBCHOIXCOD := TOB.Create('TABLETTE', nil, -1);
    TOBCHOIXCOD.LoadDetailDB('CHOIXCOD','','',QQ, False);
    For Ind := 0 to TOBCHOIXCOD.Detail.count - 1 do
    begin
      TobDetail := TOB.Create('BTSOUSFAMILLETARIF',TOBSFAMTAR,-1);
      TOBDetail.PutValue('BSF_FAMILLETARIF' , TOBCHOIXCOD.Detail[Ind].GetString('CC_LIBRE'));
      TOBDetail.PutValue('BSF_SOUSFAMTARART', TOBCHOIXCOD.Detail[Ind].GetString('CC_CODE'));
      TOBDetail.PutValue('BSF_LIBELLE'      , TOBCHOIXCOD.Detail[Ind].GetString('CC_LIBELLE'));
    end;
    FreeAndNil(TOBCHOIXCOD);
  end;

  Ferme (QQ);

  TOBSFAMTAR.SetAllModifie(True);
  TOBSFAMTAR.InsertOrUpdateDB;

  //Réinitialisation tablette Famille TARIF
  StSQL := 'DELETE CHOIXCOD WHERE CC_TYPE="TAR"';
  ExecuteSQL(StSQL);

  //Réinitialisation Tablette Sous-Famille TARIF
  StSQL := 'DELETE CHOIXCOD WHERE CC_TYPE="BFT"';
  ExecuteSQL(StSQL);

  //Libération des TOB
  FreeAndNil(TOBFAMTAR);
  FreeAndNil(TOBSFAMTAR);

end;

end.
