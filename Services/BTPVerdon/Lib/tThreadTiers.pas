unit tThreadTiers;

interface

uses
  Classes
  , UtilBTPVerdon
  , ConstServices
  , uTob
  , CbpMCD
  {$IFDEF MSWINDOWS}
  , Windows
  {$ENDIF}
  ;

type
  ThreadTiers = class(TThread)
  public
    TiersValues  : T_TiersValues;
    LogValues    : T_WSLogValues;
    FolderValues : T_FolderValues;

    constructor Create(CreateSuspended : boolean);
    destructor Destroy; override;

  private
    lTn          : T_TablesName;
    ArrBTPFields : array of string;
    ArrTMPFields : array of string;

//    procedure SetName;
    procedure SetFieldsArray;
    //function GetSqlInsertValues(TobData : TOB) : string;
    function GetSql : string;

  protected
    procedure Execute; override;
  end;

implementation

uses
  CommonTools
  , SysUtils
  , hCtrls
  ;

const
  ArrLen = 25;

{ Important : les méthodes et propriétés des objets de la VCL peuvent uniquement être
  utilisés dans une méthode appelée en utilisant Synchronize, comme : 

      Synchronize(UpdateCaption);

  où UpdateCaption serait de la forme

    procedure ThreadTiers.UpdateCaption;
    begin
      Form1.Caption := 'Mis à jour dans un thread';
    end; }

{$IFDEF MSWINDOWS}
type
  TThreadNameInfo = record
    FType: LongWord;     // doit être 0x1000
    FName: PChar;        // pointeur sur le nom (dans l'espace d'adresse de l'utilisateur)
    FThreadID: LongWord; // ID de thread (-1=thread de l'appelant)
    FFlags: LongWord;    // réservé pour une future utilisation, doit être zéro
  end;
{$ENDIF}

{ ThreadTiers }

(*
procedure ThreadTiers.SetName;
{$IFDEF MSWINDOWS}
var
  ThreadNameInfo: TThreadNameInfo;
{$ENDIF}
begin
{$IFDEF MSWINDOWS}
  ThreadNameInfo.FType := $1000;
  ThreadNameInfo.FName := 'ThreadNameTiers';
  ThreadNameInfo.FThreadID := $FFFFFFFF;
  ThreadNameInfo.FFlags := 0;

  try
    RaiseException( $406D1388, 0, sizeof(ThreadNameInfo) div sizeof(LongWord), @ThreadNameInfo );
  except
  end;
{$ENDIF}
end;
*)

constructor ThreadTiers.Create(CreateSuspended : boolean);
begin
  inherited Create(CreateSuspended);
  FreeOnTerminate := True;
  Priority        := tpNormal;
  lTn             := tnTiers;
end;

destructor ThreadTiers.Destroy;
begin
  inherited;
  TUtilBTPVerdon.AddLog(lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, False, TiersValues.LastSynchro), LogValues, 0);
end;

function ThreadTiers.GetSql: string;
begin
  Result := Format('SELECT %s FROM TIERS WHERE T_DATEMODIF >= "%s" ORDER BY T_AUXILIAIRE', [TUtilBTPVerdon.GetFieldsListFromArray(ArrBTPFields), Tools.CastDateTimeForQry(StrToDatetime(TiersValues.LastSynchro))]);
end;

procedure ThreadTiers.SetFieldsArray;
var
  Cpt : integer;

  procedure AddValues(SuffixBTP, SuffixTMP : string);
  begin
    ArrBTPFields[Cpt] := Format('T_%s'  , [SuffixBTP]);
    ArrTMPFields[Cpt] := Format('TIE_%s', [SuffixTMP]);
  end;

begin
  SetLength(ArrBTPFields, ArrLen);
  SetLength(ArrTMPFields, ArrLen);
  for Cpt :=  Low(ArrBTPFields) to High(ArrBTPFields) do
  begin
    case Cpt of
      0  : AddValues('AUXILIAIRE'  , 'AUXILIAIRE');
      1  : AddValues('COLLECTIF'   , 'COLLECTIF');
      2  : AddValues('NATUREAUXI'  , 'NATUREAUXI');
      3  : AddValues('LIBELLE'     , 'LIBELLE');
      4  : AddValues('DEVISE'      , 'DEVISE');
      5  : AddValues('ADRESSE1'    , 'ADRESSE1');
      6  : AddValues('ADRESSE2'    , 'ADRESSE2');
      7  : AddValues('ADRESSE3'    , 'ADRESSE3');
      8  : AddValues('CODEPOSTAL'  , 'CP');
      9  : AddValues('VILLE'       , 'VILLE');
      10 : AddValues('PAYS'        , 'PAYS');
      11 : AddValues('TELEPHONE'   , 'TELEPHONE');
      12 : AddValues('FAX'         , 'TELEPHONE2');
      13 : AddValues('TELEX'       , 'TELEPHONE3');
      14 : AddValues('EMAIL'       , 'EMAIL');
      15 : AddValues('RVA'         , 'WEBURL');
      16 : AddValues('COMPTATIERS' , 'COMPTA');
      17 : AddValues('FACTURE'     , 'TIERSFACTURE');
      18 : AddValues('PAYEUR'      , 'TIERSPAYEUR');
      19 : AddValues('BLOCNOTE'    , 'BLOCNOTE');
      20 : AddValues('SIRET'       , 'SIRET');
      21 : AddValues('NIF'         , 'NUMINTRACOMM');
      22 : AddValues('SOCIETE'     , 'CODESOCIETE');
      23 : AddValues('DATEMODIF'   , 'DATEMODIF');
      24 : AddValues('DATECREATION', 'DATECREATION');
    end;
  end;
end;

(*
function ThreadTiers.GetSqlInsertValues(TobData : TOB) : string;
var
  Cpt       : integer;
  FieldName : string;
  FieldType : tTypeField;
begin
  for Cpt := 0 to High(ArrBTPFields) do
  begin
    FieldName := ArrBTPFields[Cpt];
    FieldType    := Tools.GetFieldType(FieldName{$IF defined(APPSRV)}, FolderValues.BTPServer, FolderValues.BTPDataBase {$IFEND !APPSRV});
    case FieldType of
      ttfNumeric
      , ttfInt     : Result := Format('%s, %s'    , [Result, TobData.GetString(FieldName)]);
      ttfDate      : Result := Format('%s, ''%s''', [Result, UsDateTime(TobData.GetDateTime(FieldName))]);
      ttfCombo
      , ttfText
      , ttfBoolean
      , ttfMemo    : Result := Format('%s, ''%s''', [Result, TobData.GetString(FieldName)]);
    end;
  end;
  Result := Copy(Result, 2, length(Result));
end;
*)

procedure ThreadTiers.Execute;
var
  TobT      : TOB;
  TobTl     : TOB;
  AdoQryL   : AdoQry;
  Cpt       : integer;
  TIEValue  : string;
  SqlUnlock : string;
  Sql       : string;
begin
//  SetName;
  try
    TUtilBTPVerdon.AddLog(lTn, TUtilBTPVerdon.GetMsgStartEnd(lTn, True, TiersValues.LastSynchro), LogValues, 0);
    TobT := TOB.Create('_TIERS', nil, -1);
    try
      SetFieldsArray;
      TobT.LoadDetailFromSQL(GetSql);
      TUtilBTPVerdon.AddLog(lTn, Format('%s tiers trouvé(s)', [IntToStr(TobT.Detail.Count)]), LogValues, 1);
      if TobT.Detail.Count > 0 then
      begin
        AdoQryL := AdoQry.Create;
        try
          AdoQryL.ServerName  := FolderValues.TMPServer;
          AdoQryL.DBName      := FolderValues.TMPDataBase;
          AdoQryL.PgiDB       := '-';
          AdoQryL.FieldsList  := 'TIE_TRAITE';
          AdoQryL.LogValues   := LogValues;
          AdoQryL.TSLResult.Clear;
          SqlUnlock := 'UPDATE TIERS SET TIE_LOCK = ''N'' WHERE T_AUXILIAIRE IN(';
          for Cpt := 0 to pred(TobT.Detail.Count) do
          begin
            TobTl := TobT.Detail[Cpt];
            // Test si enregistrement existe
            AdoQryL.Request := Format('SELECT %s FROM TIERS WHERE TIE_AUXILIAIRE = ''%s''', [AdoQryL.FieldsList, TobTl.GetString('T_AUXILIAIRE')]);
            AdoQryL.SingleTableSelect;
            if AdoQryL.RecordCount = 1 then // Update si TRAITE = O
            begin
              TIEValue := AdoQryL.TSLResult[0];
              if TIEValue = 'O' then
              begin
                Sql := '';
              end;
            end else
            begin
              Sql := Format('INSERT INTO TIERS (%s) VALUES (%s)'
                            , [  TUtilBTPVerdon.GetSqlInsertFields(tnTiers, ArrTMPFields)
                               , TUtilBTPVerdon.GetSqlInsertValues(tnTiers, FolderValues, TobTl, ArrBTPFields)
                              ]);
            end;
            AdoQryL.RecordCount := 0;
            AdoQryL.Request     := Sql;
//            AdoQryL.InsertUpdate;
          end;
        finally
          AdoQryL.free;
        end;
      end;
    finally
      TUtilBTPVerdon.SetLastSynchro(lTn);
      FreeAndNil(TobT);
    end;
  except
  end;
end;

end.
