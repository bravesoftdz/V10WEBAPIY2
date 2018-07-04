unit UgestionCdeRexel;

interface
Uses HEnt1, HCtrls, Controls, ComCtrls, StdCtrls, ExtCtrls,Ent1,
     SysUtils, Classes, Graphics, Forms, Saisutil, EntGC, AGLInit, 
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fe_Main,
{$ENDIF}
      HmsgBox,
     Messages, Windows,
     UtilPgi,
     UTOB,
     uEntCommun,ParamSoc,
     ShellApi
     ;

function EnvoieCommandeRexel (TOBPiece,TOBAdresses: TOB) : boolean;

implementation
uses IdBaseComponent;

function ParamRexelOk (Fournisseur : string; TOBparamEx : TOB) : boolean;
var QQ : TQuery;
begin
  result := false;
  QQ := OpenSql ('SELECT * FROM BTIERSECHG WHERE BTE_AUXILIAIRE=(SELECT T_AUXILIAIRE FROM TIERS WHERE T_TIERS="'+Fournisseur+'" AND T_NATUREAUXI="FOU")',true,1,'',true);
  if not QQ.eof then
  begin
    TOBparamEx.SelectDB('',QQ);
    result := true;
  end;
  ferme (QQ);
end;

function ConstitueFichierHtm (TOBpiece,TOBAdresse,TOBParamEx: TOB; ListeRef : string) : string;
var TT : String;
    Stream : TextFile ; 
begin
  result := '';
  TT := fileTemp ('.htm');
  if TT = '' then exit;
  result := TT;
  AssignFile(Stream, TT) ;
  ReWrite(Stream) ;
  Writeln(Stream, '<HTML>') ;
  Writeln(Stream, '<HEAD>') ;
  Writeln(Stream, '<TITLE>interface</TITLE>') ;
  Writeln(Stream, '</HEAD>') ;
  Writeln(Stream, '<BODY>') ;
  Writeln(Stream, '<form method="POST" name="Interface">') ;
  Writeln(Stream, '<input type="HIDDEN" name="NumCli" id="NumCli" value="'+TOBParamEx.getString('BTE_NUMCLI')+'">') ;
  Writeln(Stream, '<input type="HIDDEN" name="NomUtil" id="NomUtil" value="'+TOBParamEx.getString('BTE_CODE')+'">') ;
  if TOBParamEx.getBoolean('BTE_FLAGRAZ') then Writeln(Stream, '<input type="HIDDEN" name="FlagRAZ" id="FlagRAZ" value="0">')
                                          else Writeln(Stream, '<input type="HIDDEN" name="FlagRAZ" id="FlagRAZ" value="1">');

  Writeln(Stream, '<input type="HIDDEN" name="LstRefQte" id="LstRefQte" value="'+ListeRef+'">');
  if TOBpiece.getInteger('GP_IDENTIFIANTWOT')<>2 then Writeln(Stream, '<input type="HIDDEN" name="FlagLIV" id="FlagLIV" value="LIV">')
                                                 else Writeln(Stream, '<input type="HIDDEN" name="FlagLIV" id="FlagLIV" value="AGE">');
  Writeln(Stream, '<input type="HIDDEN" name="CodAge" id="CodAge" value="'+TOBParamEx.getString('BTE_AGENCETIERS')+'">');
  Writeln(Stream, '<input type="HIDDEN" name="ChpRS" id="ChpRS" value="'+trim(TOBAdresse.GetString('ADR_LIBELLE'))+'">');
  Writeln(Stream, '<input type="HIDDEN" name="ChpAdr1" id="ChpAdr1" value="'+trim(TOBAdresse.GetString('ADR_ADRESSE1'))+'">');
  Writeln(Stream, '<input type="HIDDEN" name="ChpAdr2" id="ChpAdr2" value="'+trim(TOBAdresse.GetString('ADR_ADRESSE2'))+'">');
  Writeln(Stream, '<input type="HIDDEN" name="ChpAdr3" id="ChpAdr3" value="'+trim(TOBAdresse.GetString('ADR_ADRESSE3'))+'">');
  Writeln(Stream, '<input type="HIDDEN" name="ChpCp" id="ChpCp" value="'+trim(TOBAdresse.GetString('ADR_CODEPOSTAL'))+'">');
  Writeln(Stream, '<input type="HIDDEN" name="ChpVille" id="ChpVille" value="'+trim(TOBAdresse.GetString('ADR_VILLE'))+'">');
  Writeln(Stream, '<input type="HIDDEN" name="RefCde" id="RefCde" value="'+TOBpiece.getString('GP_NATUREPIECEG')+'-'+IntToStr(TOBpiece.getInteger('GP_NUMERO'))+'">');
  Writeln(Stream, '<input type="HIDDEN" name="RefCht" id="RefCht" value="'+TOBpiece.getString('GP_AFFAIRE')+'">');
  Writeln(Stream, '<input type="HIDDEN" name="FlagReliq" id="FlagReliq" value="1">');
  Writeln(Stream, '<input type="HIDDEN" name="DtExp" id="DtExp" value="'+DateToStr(TOBpiece.getDateTime('GP_DATELIVRAISON'))+'">');
  Writeln(Stream, '</form>');
  Writeln(Stream, '<SCRIPT>');
  Writeln(Stream, 'Interface.action="'+TOBParamEx.getString('BTE_WEBADRESSE')+'"; Interface.submit();');
  Writeln(Stream, '</SCRIPT>');
  Writeln(Stream, '</BODY>');
  Writeln(Stream, '</HTML>') ;
  CloseFile(Stream) ;
end;

function ConstitueListemateriaux (TOBpiece,TOBParamEx : TOB; var StModeLiv : string) : string;
var II : integer;
    TOBL : TOB;
    ierror : integer;
begin
  iError := 0;
  result := '';
  for II := 0 to TOBpiece.detail.count -1 do
  begin
    TOBL := TOBpiece.detail[II];
    if TOBL.GetString('GL_TYPELIGNE') <> 'ART' then continue;
    if TOBL.GetString('GL_REFARTTIERS')='' then BEGIN ierror := 1; continue; END;
    if result <> '' then result := result + '|';
    Result := Result + Trim(TOBL.GetString('GL_REFARTTIERS'))+'|'+STRF00(TOBL.GetDouble('GL_QTEFACT'),0);
  end;
  if result = '' then ierror := 2;
  if Ierror = 2 then PGIError ('Aucune référence pour le fournisseur trouvée')
  else if Ierror = 1 then PgiInfo('Certains articles ne sont pas référencés chez le fournisseur');
end;

function DefiniAdresse (TOBPiece,TOBAdresses,TOBAdresse : TOB) : boolean;
  
  procedure getAdresseDepot ( Depot : string; TOBADRESSE : TOB);
  var TOBADD : TOB;
  begin
    TOBADD := TOB.Create ('DEPOTS',nil,-1);
    TRY
      TOBADD.PutValue('GDE_DEPOT',Depot);
      TOBADD.LoadDB (true);
      TOBADRESSE.SetString('ADR_LIBELLE',TOBADD.getvalue('GDE_LIBELLE'));
      TOBADRESSE.SetString('ADR_ADRESSE1',TOBADD.getvalue('GDE_ADRESSE1'));
      TOBADRESSE.SetString('ADR_ADRESSE2',TOBADD.getvalue('GDE_ADRESSE2'));
      TOBADRESSE.SetString('ADR_ADRESSE3',TOBADD.getvalue('GDE_ADRESSE3'));
      TOBADRESSE.SetString('ADR_CODEPOSTAL',TOBADD.getvalue('GDE_CODEPOSTAL'));
      TOBADRESSE.SetString('ADR_VILLE',TOBADD.getvalue('GDE_VILLE'));
    FINALLY
      TOBADD.free;
    END;
  end;


  procedure getAdresseEtablissement ( Etablissement: string; TOBADRESSE : TOB);
  var TOBADD : TOB;
  begin
    TOBADD := TOB.Create ('ETABLISS',nil,-1);
    TRY
      TOBADD.PutValue('ET_ETABLISSEMENT',Etablissement);
      TOBADD.LoadDB (true);
      TOBADRESSE.SetString('ADR_LIBELLE',TOBADD.getvalue('ET_LIBELLE'));
      TOBADRESSE.SetString('ADR_ADRESSE1',TOBADD.getvalue('ET_ADRESSE1'));
      TOBADRESSE.SetString('ADR_ADRESSE2',TOBADD.getvalue('ET_ADRESSE2'));
      TOBADRESSE.SetString('ADR_ADRESSE3',TOBADD.getvalue('ET_ADRESSE3'));
      TOBADRESSE.SetString('ADR_CODEPOSTAL',TOBADD.getvalue('ET_CODEPOSTAL'));
      TOBADRESSE.SetString('ADR_VILLE',TOBADD.getvalue('ET_VILLE'));
    FINALLY
      TOBADD.free;
    END;
  end;

  procedure getAdresseSociete (TOBADRESSE : TOB);
  begin
    TOBADRESSE.SetString('ADR_LIBELLE',GetParamSoc('SO_LIBELLE'));
    TOBADRESSE.SetString('ADR_ADRESSE1',GetParamSoc('SO_ADRESSE1'));
    TOBADRESSE.SetString('ADR_ADRESSE2',GetParamSoc('SO_ADRESSE2'));
    TOBADRESSE.SetString('ADR_ADRESSE3',GetParamSoc('SO_ADRESSE3'));
    TOBADRESSE.SetString('ADR_CODEPOSTAL',GetParamSoc('SO_CODEPOSTAL'));
    TOBADRESSE.SetString('ADR_VILLE',GetParamSoc('SO_VILLE'));
  end;

  procedure GetAdresseChantier (TOBpiece,TOBAdresse : TOB);
  var TOBADD : TOB;
      QQ : TQUery;
  begin
    TOBADD := TOB.Create ('ADRESSES',nil,-1);
    QQ := OpenSql ('SELECT * FROM ADRESSES WHERE ADR_REFCODE="'+TOBPiece.getString('GP_AFFAIRE')+'" AND ADR_TYPEADRESSE="INT"',true,1,'',true);
    TRY
      if not QQ.eof then TOBADD.SelectDB('',QQ); 
      TOBADRESSE.SetString('ADR_LIBELLE',TOBADD.getString('ADR_LIBELLE'));
      TOBADRESSE.SetString('ADR_LIBELLE2',TOBADD.getString('ADR_LIBELLE2'));
      TOBADRESSE.SetString('ADR_ADRESSE1',TOBADD.getString('ADR_ADRESSE1'));
      TOBADRESSE.SetString('ADR_ADRESSE2',TOBADD.getString('ADR_ADRESSE2'));
      TOBADRESSE.SetString('ADR_ADRESSE3',TOBADD.getString('ADR_ADRESSE3'));
      TOBADRESSE.SetString('ADR_CODEPOSTAL',TOBADD.getString('ADR_CODEPOSTAL'));
      TOBADRESSE.SetString('ADR_VILLE',TOBADD.getString('ADR_VILLE'));
    FINALLY
      ferme (QQ);
      TOBADD.free;
    END;
  end;



var Depot,Etablissement : string;
begin
  result := true;
  if TOBpiece.getInteger('GP_IDENTIFIANTWOT') = 0 then
  begin
    // DEPOT
    Depot := TOBPiece.GetValue('GP_DEPOT');
    Etablissement := TOBPiece.GetValue('GP_ETABLISSEMENT');
    if Depot <> '' then
    begin
      getAdresseDepot ( Depot,TOBADresse);
      if (TOBADresse.getString('GPA_LIBELLE')='') and (TOBADresse.getString('ADR_ADRESSE1')='') then
      begin
        PgiError ('Adresse dépot non définie');
        result := false;
      end;
    end else if Etablissement <> '' then
    begin
      getAdresseEtablissement ( Etablissement,TOBADresse);
      if (TOBADresse.getString('GPA_LIBELLE')='') and (TOBADresse.getString('ADR_ADRESSE1')='') then
      begin
        PgiError ('Adresse établissement non défini');
        result := false;
      end;
    end else
    begin
      getAdresseSociete (TOBADresse);
      if (TOBADresse.getString('GPA_LIBELLE')='') and (TOBADresse.getString('ADR_ADRESSE1')='') then
      begin
        PgiError ('Adresse Société non défini');
        result := false;
      end;
    end;
  end else if TOBpiece.getInteger('GP_IDENTIFIANTWOT') = -1 then
  begin
    // CHANTIER
    GetAdresseChantier (TOBpiece,TOBAdresse);
    if (TOBADresse.getString('ADR_LIBELLE')='') and (TOBADresse.getString('ADR_ADRESSE1')='') then
    begin
      PgiError ('Adresse chantier non défini');
      result := false;
    end;
  end else if TOBpiece.getInteger('GP_IDENTIFIANTWOT') = 2 then
  begin
    // Fournisseur
  end;
end;

function EnvoieCommandeRexel (TOBPiece,TOBAdresses : TOB) : boolean;
var ListeRef : string;
     stModeLiv : string;
     TobParamEx,TOBAdresse : TOB;
     HTMLFile : string;
begin
  result := false;
  if TOBpiece.getString('GP_ETATEXPORT') <> 'ATT' then exit;
  TOBParamEx := TOB.Create ('BTIERSECHG',nil,-1);
  TOBADRESSE := TOB.Create ('ADRESSES',nil,-1);
  TRY
    if not ParamRexelOk (TOBPiece.getString('GP_TIERS'),TOBParamEx) then exit;
    //
    ListeRef := ConstitueListemateriaux(TOBPiece,TOBParamEx,StModeLiv);
    if ListeRef = '' then
    begin
      exit;
    end;
    //
    if not DefiniAdresse (TOBPiece,TOBAdresses,TOBAdresse) then exit;
    //
    HTMLFile := ConstitueFichierHtm (TOBpiece,TOBAdresse,TOBParamEx,ListeRef);
    //
    if HTMLFile <> '' then
    begin
      ShellExecute (0, pchar ('open'), pchar (HTMLFile), nil, nil, SW_SHOWNORMAL);
      result := true;
    end;
  FINALLY
    TOBParamEx.free;
    TOBADRESSE.free;
  END;
end;


end.
