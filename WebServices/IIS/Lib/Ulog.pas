unit Ulog;

interface
uses SysUtils;

procedure ecritLog (const Emplacement : string; const SMessage : string; NomFichier : string='');
implementation

procedure ecritLog (const Emplacement : string; const SMessage : string; NomFichier : string='');
var f : TextFile;
		nomLog : string;
    nomFic : string;
begin
  if NomFichier = '' then nomFic := 'WebService.log' else nomFic := NomFichier;
  nomLog := IncludeTrailingPathDelimiter(Emplacement)+'WebService.log';
  AssignFile(f, nomLog);
  Append (f);
  writeln ( f, SMessage);
  Flush(f);
  CloseFile(f);
end;

end.
