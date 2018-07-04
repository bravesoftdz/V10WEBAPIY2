unit UtilGPAO;

interface

  function GetNbEnregInFile (Fichier : String) : integer;

implementation

{***********A.G.L.***********************************************
Auteur  ...... : Y. Peuillon
Créé le ...... : 05/11/2003
Modifié le ... : 26/11/2003
Description .. : Renvoi le nombre d'enregistrement du fichier passé en
Suite ........ : paramètre
Mots clefs ... : 
*****************************************************************}
function GetNbEnregInFile (Fichier : String) : integer;
var
  FileIn : TextFile;
  Cpt : integer;
begin
  Cpt := 0;
  // Ouverture du fichier en entrée
  AssignFile(FileIn, Fichier);
  Reset( FileIn );
  try
    // Parcours du fichier en entrée
    while not SeekEof(FileIn) do
    begin
      Readln( FileIn );
      Inc( Cpt );
    end;
  finally
    CloseFile( FileIn);
  end;
  Result := Cpt;
end;

end.
