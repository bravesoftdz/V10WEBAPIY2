unit Ulog;

interface
uses SysUtils;

procedure ecritLog (const Emplacement : string; const SMessage : string; Nomfic : string='');
implementation

procedure ecritLog (const Emplacement : string; const SMessage : string; Nomfic : string='');
var f : TextFile;
    NomFile : string;
		nomLog : string;
begin
  if NomFic = '' then NomFile := 'WebService.log' else NomFile := Nomfic;
  nomLog := IncludeTrailingPathDelimiter(Emplacement)+NomFile;

  if not FileExists(nomLog) then
  begin
    AssignFile(f, nomLog); 
    ReWrite(f);
    closeFile(f);
  end;

  AssignFile(f, nomLog);
  Append (f);
  writeln ( f, SMessage);
  Flush(f);
  CloseFile(f);
end;

end.

