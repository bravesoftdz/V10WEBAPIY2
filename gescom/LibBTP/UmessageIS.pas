unit UmessageIS;

interface
uses  StdCtrls,Controls,Classes, forms,sysutils,ComCtrls,paramsoc,
      HCtrls,HEnt1,HMsgBox,UTOF, AffaireUtil,Grids,EntGC,
      utilarticle,FactComm, FactCalc,Facture,UtilPGI,UtilGC,
      SaisUtil,Ent1,Formule,dicobtp,tiersutil,
      factaffaire,facttiers,utilgrp,factcpta,
{$IFDEF EAGLCLIENT}

{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,
{$ENDIF}
      FactUtil,UTob, UtilTarif{,voirtob}, factpiece,facttob,factarticle,factadresse,uEntCommun,UtilTOBPiece;

function GenereMessageIS (CodeMessage: string; TheMessages,ThePiece,TheBases,TheTiers : TOB) : boolean;

implementation

function GenereMessageIS (CodeMessage: string; TheMessages,ThePiece,TheBases,TheTiers : TOB) : boolean;
var EmplacIO,EANSOCIETE,EANDEST,FileNameD,FileNameE : string;
    ffileD,FfileE : TextFile;
    II : integer;
    TheMessage : string;
    ThisGuid : string;
begin
  result := false;
  ThisGuid := AglGetGuid;
  EANSOCIETE := GetParamSocSecur ('SO_BTECHGEAN','');
  if EANSOCIETE = '' then
  begin
    PGIINfo ('ERREUR : Code EAN de votre société non défini dans les paramètres sociétés');
    V_PGI.ioerror := oeUnknown;
    exit;
  end;
  EANDEST := TheTiers.getString('T_EAN');
  if EANDEST = '' then
  begin
    PGIINfo ('ERREUR : Code EAN du client non défini');
    V_PGI.ioerror := oeUnknown;
    exit;
  end;
  EmplacIO := GetParamSocSecur('SO_BTECHGEMPLAC','');
  if EmplacIO = '' then
  begin
    PGIINfo ('ERREUR : Emplacement des échanges inter sociétés non défini');
    V_PGI.ioerror := oeUnknown;
    exit;
  end;
  if CodeMessage='001' then TheMessage := 'INV';
  // bon la tout est OK --> on essaye de générer un fichier
  if not DirectoryExists (EmplacIO) then CreateDir (EmplacIO);
  //
  FileNameD := IncludeTrailingBackslash(EmplacIO)+ThisGuid+'.LET';
  //
  AssignFile(ffileD, FileNameD) ;
  ReWrite(ffileD) ;
  // --- pas de détail pour l'instant ---
  // ------------------------------------
  Flush(ffileD);
  CloseFile(ffileD);
  // --------------------
  FileNameE := IncludeTrailingBackslash(EmplacIO)+ThisGuid+'.ENV';
  AssignFile(ffileE, FileNameE);
  rewrite(ffileE);
  if IoResult = 0 then
  begin
    if CodeMessage='001' then Writeln (ffileE,'MESSAGE=INVOICE');
    Writeln (ffileE,'EMETTEUR='+EANSOCIETE);
    Writeln (ffileE,'DESTINATAIRE='+EANDEST);
    Writeln (ffileE,'NUMPIECE='+ThePiece.GetString('GP_NUMERO'));
    Writeln (ffileE,'DATE='+DateTimetoStr(ThePiece.GetDateTime('GP_DATEPIECE')));
    Writeln (ffileE,'REFINTERNE='+ThePiece.getString('GP_REFEXTERNE'));
    Writeln (ffileE,'CHANTIER='+ThePiece.getString('GP_AFFAIRE'));
    //
    for II := 0 to TheBases.detail.count -1 do
    begin
      Writeln(ffileE,'BASE'+IntToStr(II+1)+'='+STRFPOINT (TheBases.detail[II].getDouble('GPB_BASEDEV')));
      Writeln(ffileE,'TAUX'+IntToStr(II+1)+'='+STRFPOINT (TheBases.detail[II].getDouble('GPB_TAUXTAXE')));
      Writeln(ffileE,'VALEUR'+IntToStr(II+1)+'='+STRFPOINT (TheBases.detail[II].getDouble('GPB_VALEURDEV')));
    end;
    if TheBases.detail.count < 5 then
    begin
      For II := TheBases.detail.count to 4 do
      begin
        Writeln(ffileE,'BASE'+IntToStr(II+1)+'=0');
        Writeln(ffileE,'TAUX'+IntToStr(II+1)+'=0');
        Writeln(ffileE,'VALEUR'+IntToStr(II+1)+'=0');
      end;
    end;
  end;
  Flush(ffileE);
  CloseFile(ffileE);
  result := true;
end;

end.
