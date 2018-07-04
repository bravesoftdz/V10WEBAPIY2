unit EclateTra;
//  fait l'eclatement d'un fichier de transfert .tra en fichiers .pgi
//  compatibles avec le programme RecupNeg

interface

uses SysUtils;

function Genere_PGI(NomFicTra : string) : boolean;
function RechercheRang(PrefixeCour : string) : integer;

//  Valeurs à modifier si la structure du .tra change...
const OFFSETIDENT = 4;
      LGIDENT     = 3;

var OpenFic : array of textfile;
    Prefixe : array of string;

implementation

function Genere_PGI(NomFicTra : string) : boolean;
var FicTra, FicPrefixe : textfile;
    stEnreg, stPrefixes, stIdent : string;
    stTemp1 : string;
    ind1, indPrefixe : integer;

begin
SetLength(OpenFic, 1);
SetLength(Prefixe, 1);

//  Chargement des prefixes de tables à traiter pour GesCom S5
AssignFile (FicPrefixe, 'PrefixGCS5');
Reset (FicPrefixe);
while (not EOF(FicPrefixe)) do
    begin
    readln(FicPrefixe, stEnreg);
    stPrefixes := stPrefixes + stEnreg + ';';
    end;
CloseFile(FicPrefixe);

stTemp1 := NomFicTra + '.tra';
AssignFile (FicTra, stTemp1);
Reset (FicTra);

Result := False;
while (not EOF(FicTra)) do
    begin
    Readln(FicTra, stEnreg);
    stIdent := Trim(Copy(stEnreg, OFFSETIDENT, LGIDENT));
    if Pos(stIdent, stPrefixes) <> 0 then
        begin
        indPrefixe := RechercheRang(stIdent);
        if indPrefixe = -1 then
            begin
            if High(OpenFic) > 0 then
                begin
                SetLength(OpenFic, High(OpenFic) + 2);
                SetLength(Prefixe, High(Prefixe) + 2);
                end
                else
                begin
                SetLength(OpenFic, 2);
                SetLength(Prefixe, 2);
                end;
            indPrefixe := High(OpenFic);
            Prefixe[indPrefixe] := stIdent;
            AssignFile(OpenFic[indPrefixe], stIdent + '.pgi');
            Rewrite(OpenFic[indPrefixe]);
            end;
        Writeln(OpenFic[indPrefixe],stEnreg);
        Result := True;
        end;
    end;
CloseFile(FicTra);

//  Fermeture des fichiers .pgi ouverts en cours de traitement
for ind1 := 0 to High(OpenFic) do
    if Prefixe[ind1] <> '' then
        begin
        Flush(OpenFic[ind1]);
        CloseFile(OpenFic[ind1]);
        end;
end;

function RechercheRang(PrefixeCour : string) : integer;
var ind1 : integer;
begin
Result := -1;
for ind1 := 0 to Length(Prefixe) - 1 do
    if Prefixe[ind1] = PrefixeCour then
        begin
        Result := ind1;
        Exit;
        end;
end;

end.
