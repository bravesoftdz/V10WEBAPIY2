{***********UNITE*************************************************
Auteur  ...... : MB, GH
Créé le ...... : 08/03/2007
Modifié le ... : 12/03/2007
*****************************************************************}
unit UTobTrans;

interface


uses Messages,
  Windows,
  Classes,
  Dialogs,
  Sysutils,
  Forms,
  Controls,
  ComCtrls,
  ExtCtrls,
  Grids,
  StdCtrls,
  Mask,
  Spin,
  ClipBrd,
{$IFDEF VER150} {D7}
  Variants,
  FMTBcd,
{$ENDIF}
  USlkParser,
{$IFNDEF EAGLCLIENT}
  HQry,
  HDB,
  HRichOLE,
  RichEdit,
  HMsgBox,
  DB,
  DBCtrls,
{$IFDEF ODBCDAC}
  odbcconnection,
  odbctable,
  odbcquery,
  odbcdac,
{$ELSE}
{$IFNDEF DBXPRESS}dbtables,
{$ELSE}uDbxDataSet,
{$ENDIF}
{$ENDIF}
  HColComb, //Hstatus,
{$ELSE}
  uhttp,

{$ENDIF}
  Hstatus, //jfd pour initmove dans TOBLoadFromBinStream
  Hent1,
  Hctrls,
  utob;

type
  TOBTrans = class(TOB)
  private
    procedure DeleteBeforeTable;
    procedure GetInsertQuery(var Q : TQuery) ;
    procedure PutParameters(T : TOB ; var Q : TQuery) ;
    function GetQuery: TQuery;
  public
    function BatchInsertDB(DeleteBefore : boolean = false ; newdatabase : String = '' ) : boolean;
  end;

implementation

uses HPanel,
  //TradData,
  TradMini,
{$IFNDEF HVCL}
{$IFDEF EAGLCLIENT}
  HRichOLE,
  HXLSPas,
  HMsgBox,
  uHttp,
  eMul,
{$ELSE}
{$IFDEF FQ12184}
  MajTable, // XP 17-01-2006 FQ 12184
{$ENDIF}
{$IFNDEF BASEEXT}
  uTobDataSet {PMJ530},
{$ENDIF BASEEXT}
{$ENDIF EAGLCLIENT}
  UTXml,
{$IFNDEF EAGL}
  HXLSPas,
{$ENDIF}
{$ENDIF HVCL}
{$IFDEF eAGLServer}
  eSession,
{$IFNDEF BASEEXT}
  UTom {PGR 560f},
{$ENDIF BASEEXT}
{$ELSE}
{$IFNDEF BASEEXT}
  UTom {PGR 560f},
{$ENDIF}
{$ENDIF}

{$IFDEF EAGLCLIENT}
  {AFAIREEAGL}
{$ELSE}
  UGedFiles,
{$IFNDEF SANSAGL}
{$IFNDEF HVCL}
{$IFNDEF eAGLServer}
  //PgiEnv,
{$ENDIF}
{$ENDIF}
{$ENDIF}
{$ENDIF}

{$IFNDEF HVCL}
  UTobXls,
  UxmlUtils,
{$ENDIF}
{$IFDEF UNICODE}
  TntSystem,
{$ENDIF}
  rtfCounter;

////////////////////////////////////////////////////////////////////////////////
// Renvoi un insert avec des paramètres ...
procedure TOBTrans.DeleteBeforeTable();
var
  lenomtable: string;
begin
  if NumTable > 0 then
  begin
    lenomtable := NomTable;
{$IFDEF eaglserver}
    { controle de la véracité de FNumTable par rapport à FNomTable }
    // XP 22.09.2006 Passage en majuscule
    if (UpperCase(NomTable) <> V_PGI.DETABLES[NumTable].nom) then
    begin
      LogEagl('MakeInsertSql : FNumTable=' + IntToStr(NumTable) + ',FNomTable=' + NomTable + ',V_PGI.DETABLES=' +
        V_PGI.DETABLES[NumTable].Nom);
      LeNomTable := PrefixeToTable(ExtractPrefixe(V_PGI.DEChamps[NumTable, 1].Nom));
    end;
{$ENDIF}
    ExecuteSQL('DELETE FROM '+lenomtable);
end;
end;
////////////////////////////////////////////////////////////////////////////////
// Renvoi un insert avec des paramètres ...
procedure TOBTrans.GetInsertQuery(var Q : TQuery );
  procedure AddS(var S: hString; const S1: hString);
  begin
    if S <> '' then
       S := S + ',';
    S := S + S1;
  end;
var
  sValues: hString;
  sChamps: hString;
  iCompteur, m: integer;
  NomChamp: string;
  lenomtable: string;
begin
  sChamps := '';
  sValues := '';

  if NumTable > 0 then
  begin
    lenomtable := NomTable;
{$IFDEF eaglserver}
    { controle de la véracité de FNumTable par rapport à FNomTable }
    // XP 22.09.2006 Passage en majuscule
    if (UpperCase(NomTable) <> V_PGI.DETABLES[NumTable].nom) then
    begin
      LogEagl('MakeInsertSql : FNumTable=' + IntToStr(NumTable) + ',FNomTable=' + NomTable + ',V_PGI.DETABLES=' +
        V_PGI.DETABLES[NumTable].Nom);
      LeNomTable := PrefixeToTable(ExtractPrefixe(V_PGI.DEChamps[NumTable, 1].Nom));
    end;
{$ENDIF}
{$IFDEF eAGLServer}
    m := High(V_PGI.HDeChamps[LookupCurrentSession.SocNum, NumTable]); // nosocref
{$ELSE}
    m := High(V_PGI.DeChamps[NumTable]);
{$ENDIF}

    // XP 21.09.2006 le champ _DTAEMODIF avait l'heure lors de la création de la TOB et cela n'est pas en
    // correspondance avec la méthode PutdataSet/Insert qui reforce la date de modif

    // MB : Attention : plusde UpdateDateModif
    // UpdateDateModif;
    Q.Parameters.Clear ;
    for iCompteur := 1 to m do
    begin
      NomChamp := V_PGI.DEChamps[NumTable, iCompteur].Nom;
      AddS(sChamps, NomChamp);
      // On met le ":" pour que la requête soit paramêtrée
      AddS(sValues, ':'+NomChamp);
      with Q.Parameters.AddParameter do
      begin
          Name := NomChamp ;
          DataType := ChampToDBType(NomChamp) ;
      end;
    end;
    Q.SQL.Text := 'INSERT INTO '+LeNomTable + ' (' + sChamps + ') VALUES (' + sValues + ')';
  end;
end;
////////////////////////////////////////////////////////////////////////////////
procedure TOBTrans.PutParameters(T : TOB ; var Q : TQuery) ;
var
   i : integer ;
begin
   for i:=0 to Q.Parameters.Count -1 do
   begin
      Q.Parameters[i].Value := T.GetValue(Q.Parameters[i].Name);
   end;
end;
////////////////////////////////////////////////////////////////////////////////
function TOBTrans.GetQuery: TQuery;
// voir DBOpenSQL dans majtable !
begin
  // adapter driver pour le changesql
  Result := OpenSQL('SELECT 1 FROM '+NomTable+' WHERE 1=2',FALSE);
end;
////////////////////////////////////////////////////////////////////////////////
function TOBTrans.BatchInsertDB(DeleteBefore : boolean = false ; newdatabase : String = '') : boolean;
{$IFDEF EAGLCLIENT}
// A Rajouter Demain ... La fonction BatchInsert dans AGL ou plugin Métier
var
  Qo: Tob;
begin

  Qo := AppServer.Request('InsertTOB', NomTable, Self, '', '');

  if assigned(Qo) then
  begin
    FLastError := Qo.GetValue('ERROR');
    Result := (FLastError = '') ; // XP 18.05.2006
    Qo.free;
    if (FLastError <> 'FALSE') and (FLastError <> '') then
      raise Exception.Create('BatchInsertDB : ' + FLastError)
  end else Result := False ;

end;
{$ELSE}
var
  i: integer;
  Q : TQuery ;
begin
  result := True;
  try
    if NumTable > 0 then
    begin
      if (Controle) then
      begin
        Q := GetQuery ;
        GetInsertQuery(Q) ;
        if newdatabase <> '' then
           ExecuteSQL('@@USE '+newdatabase);
        if DeleteBefore then
           DeleteBeforeTable ;
      end;
    end;

    for i := 0 to Detail.Count - 1 do
    begin
       PutParameters(Detail[i],Q);
       Result := Result and (Q.ExecSQL = 1) ;
    end;
    if newdatabase <> '' then
       ExecuteSQL('@@USE DB'+V_PGI.NoDossier);
    Ferme(Q);
  except
    on e:exception do
    raise;
  end;
end;
{$ENDIF}
////////////////////////////////////////////////////////////////////////////////

initialization
finalization
end.



