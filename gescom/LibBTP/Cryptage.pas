unit Cryptage;

interface
uses
	hctrls,
  sysutils,
	windows,
  Classes
  ;

function EnCrypte (const ChaineEntree : string) : string;
function Decrypte (const ChaineEntree : string) : string;

implementation

const CRYPTO = 'I0L1E2T3A4I5T6U7N8E9F0O1I2S3D4A5N6S7L8A9V0I1L2L3E5D6E4F2O1I8E9U0N7P8E9Tj';
      LCRPT = 72;


function Crypt(entree:AnsiString; cle:AnsiString):AnsiString;
var
  I,J,LongEntree,LongCle,VResult : integer;
  VEntree,VCle : ^byte;
  VBuffer: ^AnsiChar;
  buffer : AnsiString;
begin
  LongEntree := length(entree);
  LongCle := length(Cle);

  If (LongEntree > 0) and (LongCle > 0) then begin

    setlength(buffer,LongEntree);

    for I:= 1 to LongEntree do begin
      for J := 1 to LongCle do begin
        VEntree := @Entree[I];
        VCle := @Cle[J];
        VResult := VEntree^ xor VCle^;
        VBuffer := @VResult;
        Buffer[I] := VBuffer^;
      end;
    end;
  end;
  if Buffer <> '' then result := buffer else Result := '';
end;

function Decrypte (const ChaineEntree : string) : string;
begin
  Result := Crypt(ChaineEntree,CRYPTO);
end;

function EnCrypte (const ChaineEntree : string) : string;
begin
  Result := Crypt(ChaineEntree,CRYPTO);
end;


end.
