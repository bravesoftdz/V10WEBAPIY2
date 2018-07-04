unit Ulog;

interface
uses SysUtils;

procedure ecritLog (const Emplacement,NomLog : string; const SMessage : string);
implementation

procedure ecritLog (const Emplacement,NomLog : string; const SMessage : string);
var f : TextFile;
		fnomLog : string;
begin
  fnomLog := IncludeTrailingPathDelimiter(Emplacement)+NomLog+'.log';
  AssignFile(f, fnomLog);
  Append (f);
  writeln ( f, SMessage);
  Flush(f);
  CloseFile(f);
end;

end.
