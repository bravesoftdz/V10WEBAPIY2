unit Ulog;

interface
uses SysUtils;

procedure ecritLog (const Emplacement : string; const SMessage : string);
implementation

procedure ecritLog (const Emplacement : string; const SMessage : string);
var f : TextFile;
		nomLog : string;
begin
  nomLog := IncludeTrailingPathDelimiter(Emplacement)+'Tdbcom07.log';
  AssignFile(f, nomLog);
  if FileExists(nomLog) then Append(f)
                        else ReWrite(f);
  writeln ( f, SMessage);
  Flush(f);
  CloseFile(f);
end;

end.
