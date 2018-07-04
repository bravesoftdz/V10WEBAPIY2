unit UtilgestionChaine;

interface

uses Windows,
     classes,
     Graphics,
     FileCtrl,
     math;


function IncStringCode (var Code : string ; LongMax : integer) : boolean;

implementation

function IncStringCode (var Code : string ; LongMax : integer) : boolean;

  function IncCharCode (var Code : string; Index : integer) : boolean;
  var CharCode : Char;
  begin
    result := false;
    CharCode := Code[Index];
    if ShortInt(CharCode) = ShortInt('Z') Then
    BEGIN
      ShortInt(CharCode) := ShortInt('0');
      Result := True; // on demande l'incrementation de la base suivante
    END else if ShortInt(CharCode) = ShortInt('9') then
    begin
      ShortInt(CharCode) := ShortInt('A');
    end else
    begin
      Inc(CharCode);
    end;
    Code := Copy(Code,1,Index-1)+CharCode+Copy(Code,Index+1,length(Code));
  end;

var StringMax : string;
    Index : integer;
    ChangeUnite : boolean;
begin
  result := true;
  StringMax := StringOfChar ('Z',LongMax);
  if Code = StringMax then BEGIN result := false; Exit; END;
  Index := LongMax;
  ChangeUnite := False;
  repeat
    ChangeUnite := IncCharCode (Code,Index);
    if ChangeUnite then
    begin
      dec(Index);
    end;
  until (not ChangeUnite) or (Index =0);
end;

end.
 