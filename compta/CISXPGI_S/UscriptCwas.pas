unit UscriptCwas;

interface
uses Classes, Windows, SysUtils, DB, uDbxDataSet, Variants, ADODB, Forms, StdCtrls,
  Dialogs, FileCtrl, inifiles, menus, Uscript,
  HEnt1, HMsgBox, Hctrls, UDMIMP, UScriptTob, bde, Imptsql;


function BlobToString(Blob: TBlobField): string;
procedure StringToBlob(Value: string; var Field : TBlobField);
procedure SelectRequeteCwas (var FScriptRequete : TScriptRequete; var SQ: TDataSource);

implementation


{$IFDEF EAGLCLIENT}
var
  Cpt: Integer = 1;
procedure ImportAddField(Q: TDataset; Name: string; fType: TFieldType; flen:
  word; required: boolean; calc: boolean; Fieldvisible: boolean);
var
  F: Tfield;
begin
  inc(Cpt);
  case fType of
    ftString: F := TStringField.Create(Q);
    ftSmallint: F := TSmallIntField.Create(Q);
    ftMemo: F := TMemoField.Create(Q);
    ftcurrency: F := TCurrencyField.Create(Q);
    ftBlob: F := TBlobField.Create(Q);
    ftBoolean: F := TBooleanField.Create(Q);
    ftInteger: F := TIntegerField.Create(Q);
    ftTime: F := TDateTimeField.Create(Q);
    ftFloat: F := TFloatField.Create(Q);
    ftDate: F := TDateField.Create(Q);
    ftGraphic: F := TGraphicField.Create(Q);
  else
    F := TField.Create(Q);
  end;
  F.SetFieldType(fType);
  try
    F.Name := Trim(Name);
  except
    F.Name := 'C' + IntToStr(Cpt);
  end;
  F.FieldName := Name;
  if fType <> ftsmallint then
  try
    F.Size := fLen - 1;
  except F.Size := 0;
  end;
  F.Calculated := Calc;
  F.Visible := Fieldvisible;
  F.DataSet := TDataSet(Q);
end;


const
  ftTypes: array[0..2] of TFieldType = (ftString, ftFloat, ftString);

function BlobToString(Blob: TBlobField): string;
var BinStream : TBlobStream;
    StrStream : TStringStream;
    s : string;

begin
BinStream := TBlobStream.Create(Blob, bmRead);
try
    StrStream := TStringStream.Create(s);
    try
        BinStream.Seek(0, soFromBeginning);
        StrStream.CopyFrom(BinStream, BinStream.Size);
        StrStream.Seek(0, soFromBeginning);
        Result := StrStream.DataString;
        StrStream.Free;
    except
        Result := '';
    end;
finally
    BinStream.Free
end;
end;

procedure StringToBlob(Value: string; var Field : TBlobField);
var StrStream : TStringStream;
    BinStream : TBlobStream;
begin
StrStream := TStringStream.Create(Value);
try
    BinStream := TBlobStream.Create(Field, bmWrite);
    try
        StrStream.Seek(0, soFromBeginning);
        BinStream.CopyFrom(StrStream, StrStream.Size);
    finally
        BinStream.Free;
    end;
finally
    StrStream.Free;
end;
end;

procedure SelectRequeteCwas (var FScriptRequete : TScriptRequete; var SQ: TDataSource);
var B                             : TBookmark;
	T                         : TADOTable;
	N                         : Integer;
	aFMTNumber, aNewFMTNumber : FMTNumber;
	FAliasName                : String;
        Q                         : TQuery;

	procedure InitTable;
	begin
    T.Connection := DMImport.DBGLOBALD;
		if FScriptRequete.TypeSortie = tsParadox then
		begin // Sortie Paradox
			T.TableName := ExtractFileName(FScriptRequete.NomFichier);
		end
		else if FScriptRequete.TypeSortie = tsODBC then
		begin // sortie ODBC
			T.TableName := ExtractFileName(FScriptRequete.NomFichier);
			FScriptRequete.DestSortie := tsAliasODBC;
			FScriptRequete.BdeToODBC(Q, T);
			exit;
		end
		else
		begin  // sortie Ascii
                        if (ParamCount > 0) and (GetInfoVHCX.NomFichier <> '') then
                           T.TableName := GetInfoVHCX.NomFichier
                        else
			    T.TableName := ExtractFileName(FScriptRequete.NomFichier);
		end;
	end;
begin
                  Q := nil;
		  DbiGetNumberFormat(aFMTNumber);	// modifiaction du separateur de decimal
		  aNewFMTNumber := aFMTNumber;		//  DecimalSeparator ne marche pas
		  aNewFMTNumber.cDecimalSeparator := '.';	// pour les tables ASCII
		  DbiSetNumberFormat(aNewFMTNumber);
		  T := TTable.Create(nil);
		  try
                          B := Q.GetBookmark;
			  InitTable;
                          if not T.Active then
				  with Q.FieldDefs do
					for N:=0 to Count-1 do
						with Items[N] do
						begin
						  if DataType = ftMemo then
							  T.FieldDefs.Add(Name, ftString, 1024, false)
						  else
							  T.FieldDefs.Add(Name, DataType, Size, false);
						end;

			  Sleep(2000); // suite au pb avec Le fichier .SCH
			  if FScriptRequete.TypeSortie <> tsODBC then
			  begin
				if FScriptRequete.ModeAscii = maDelimite then   // TCurrencyField
				begin
					ModifierSchema(FAliasName, ExtractFileName(FScriptRequete.NomFichier));
				end;
			  end;
  			  try
				T.Open;
				SQ.dataset := nil ;
				while not Q.Eof do
				begin
				  T.Append;
				  for N:=0 to Q.FieldCount-1 do
						T.Fields[N].Value := Q.Fields[N].Value;
				  T.Post;
				  Q.Next;
				end;
			  finally Q.GotoBookmark(B);  end;
			  SQ.dataset := Q;
		  finally
			  Q.FreeBookmark(B);
			  T.Free;
			  DbiSetNumberFormat(aFMTNumber);
		  end;
end;

{$ENDIF}


end.
