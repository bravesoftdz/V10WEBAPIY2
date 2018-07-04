{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/06/2007
Modifié le ... :   /  /
Description .. : Unité qui contient l'ensemble des fonctions et procédures
Suite ........ : utilisées pour la gestion des répertoires
Mots clefs ... : PAIE;PGDIR
*****************************************************************}
{
}
unit PGRepertoire;

interface
uses SysUtils, CbpPath;

function VerifieCheminPG (Repertoire : string) : string;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/06/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : PAIE;PGDIR
*****************************************************************}
function VerifieCheminPG (Repertoire : string) : string;
begin
if not DirectoryExists(Repertoire) then
   begin
   Result:= TCbpPath.GetCegidDistriStd;
   if not DirectoryExists(Result) then
      Result:= '';
   end
else
   Result:= Repertoire;
end;

end.
