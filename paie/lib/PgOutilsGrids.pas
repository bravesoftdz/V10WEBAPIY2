{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/01/2008
Modifié le ... :   /  /
Description .. : Unité qui contient l'ensemble des fonctions et procédures
Suite ........ : utilisées pour la gestion des grids
Mots clefs ... : PAIE;GRIDS
*****************************************************************}
unit PgOutilsGrids;

interface

Uses HCtrls;

procedure ValideComboDansGrid (Grille : THGrid; ActualCol, Col, Row : Integer);

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/01/2008
Modifié le ... :   /  /    
Description .. : Procédure permettant la sortie de la cellule d'une grid de 
Suite ........ : type Combo.
Suite ........ : Paramètres :
Suite ........ : Grille : Control sur la grid
Suite ........ : ActualCol : Colonne en cours
Suite ........ : Col : Colonne par défaut
Suite ........ : Row : Ligne par défaut
Mots clefs ... : PAIE;GRIDS
*****************************************************************}
procedure ValideComboDansGrid (Grille : THGrid; ActualCol, Col, Row : Integer);
begin
if (Assigned (Grille.ValCombo)) then
   begin
   if (Copy(Grille.ColFormats[ActualCol], 1, 3)='CB=') then
      begin
      Grille.UpdateCell;
      Grille.Col:= Col;
      Grille.Row:= Row;
      Grille.ValCombo.Hide;
      Grille.ValCombo.Free;
      end;
   end;
end;

end.
