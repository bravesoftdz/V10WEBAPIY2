{***********UNITE*************************************************
Auteur  ...... : CB
Créé le ...... : 26/03/2001
Modifié le ... :
Description .. : Fonction générales au planning
Suite ........ : Notamment gestion des paramètres
Mots clefs ... : PLANNING
*****************************************************************}
unit LibPlanning;

interface

Uses HCtrls, Hplanning, Controls, graphics;

  function DecodeFontStyle(pFontStyle : TFontStyles) : String;
  function EncodeFontStyle(pFontStyle : String) : TFontStyles;

implementation

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 16/03/2002
Modifié le ... :
Description .. : decode le style de font
Suite ........ :
Mots clefs ... :
*****************************************************************}
function DecodeFontStyle(pFontStyle : TFontStyles) : String;
begin

  if fsBold in pFontStyle then
    result := 'fsBold';

  if fsItalic in pFontStyle then
    if result = '' then
      result := 'fsItalic'
    else
      result := result + ';' + 'fsItalic';

  if fsUnderline in pFontStyle then
    if result = '' then
      result := 'fsUnderline'
    else
      result := result + ';' + 'fsUnderline';

  if fsStrikeOut in pFontStyle then
    if result = '' then
      result := 'fsStrikeOut'
    else
      result := result + ';' + 'fsStrikeOut';

end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 16/03/2002
Modifié le ... :
Description .. : encode le style de font
Suite ........ :
Mots clefs ... :
*****************************************************************}
function EncodeFontStyle(pFontStyle : String) : TFontStyles;
var
  vFontStyle : String;
  vStyle : String;

begin

  result := [];
  vFontSTyle := pFontStyle;
  while (vFontSTyle <> '') do
    begin
      vStyle := ReadTokenSt(vFontSTyle);
      if vStyle = 'fsUnderline' then result := result + [fsUnderline]
      else if vStyle = 'fsBold' then result := result + [fsBold]
      else if vStyle = 'fsItalic' then result := result + [fsItalic]
      else if vStyle = 'fsStrikeOut' then result := result + [fsStrikeOut];
    end;

end;

end.
