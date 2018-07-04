unit UfileCtrl;

interface

uses sysUtils,Windows;

function CopieFichier (nomOrig,Nomdest : string):boolean;
function Creationdir (repertoire : string) : boolean;

implementation


function Creationdir (repertoire : string) : boolean;
begin
  result:= Createdir (repertoire);
end;

function CopieFichier (nomOrig,Nomdest : string):boolean;
var FromF, ToF: file;
    NumRead, NumWritten: Integer;
    Buf: array[1..2048] of Char;
begin
  AssignFile(FromF, NomOrig);
  {$I-}
  Reset(FromF, 1);
  {$I+}

  if IoResult = 0 then
  begin
    AssignFile(ToF, Nomdest);
    {$I-}
    Rewrite(ToF, 1);
    {$I+}
    if IoResult = 0 then
    begin
      result := true;
      repeat
        BlockRead(FromF, Buf, SizeOf(Buf), NumRead);
        BlockWrite(ToF, Buf, NumRead, NumWritten);
      until (NumRead = 0) or (NumWritten <> NumRead);
    end else Result := false;
    CloseFile(ToF);
  end else result := false;
  CloseFile(FromF);
end;

procedure deleteRepertoire (repertoire : string);
var nomfic : string;
    Rec : TSearchRec;
begin
  Nomfic:= Repertoire+'\*.*';
  if FindFirst (Nomfic,faAnyFile,Rec) = 0 then
  begin
    if (rec.name <> '.') and (rec.name <> '..') then deleteFile (PAnsichar(repertoire+'\'+rec.Name));
    while FindNext (Rec) = 0 do
    begin
      if (rec.name <> '.') and (rec.name <> '..') then deleteFile (PAnsiChar(repertoire+'\'+rec.Name));
    end;
  end;
  SysUtils.FindClose (Rec);
  removedir(Repertoire);
end;

end.
