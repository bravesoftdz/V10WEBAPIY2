unit AglInitDialogs;

interface

uses Forms, sysutils, Graphics, Classes, Controls, M3FP, dialogs,
     hctrls, StdCtrls, HEnt1;//, HRMouliPiece;

implementation

function AglControlChecked(parms: array of variant; nb: integer): variant;
var F: Tform;
begin
  Result := True;
  F := TForm(LongInt(parms[0]));
  if TCheckBox(F.FindComponent(parms[1])).Checked = False then Result := False;
end;

function AglSelectionCouleurText(parms: array of variant; nb: integer): variant;
var T: TColorDialog;
  F: TForm;
  couleur: Tcolor;
begin
  couleur := Tcolor(Parms[1]);
  Result := '';
  F := TForm(LongInt(parms[0]));
  T := TColorDialog.Create(F);
  T.color := couleur;
  T.CustomColors.Clear;
  T.CustomColors.Add('ColorA=' + InttoHex(couleur, 6));
  //  T.CustomColors.Add ('ColorB='+InttoHex (0, 6));
  if T.Execute then
  begin
    Result := IntTostr(T.Color);
    T.Free;
  end;
end;

procedure Ajoutautrescolor(SuiteText: TStrings; F: TFORM);
  procedure TraiteColor(EditC, Coltext: string);
  var editchamp: THEDIT;
    Couleur: Tcolor;
  begin
    EditChamp := THEDIT(F.FindComponent(EditC));
    if (EditChamp <> nil) and (EditChamp.text <> '') then
    begin
      Couleur := TColor(strtoint(EditChamp.text));
      SuiteText.Add(Coltext + IntToHex(Couleur, 6));
    end;
  end;
begin
  TraiteColor('HPP_COULEURSAMEDI', 'ColorC=');
  TraiteColor('HPP_COULDIMANCHE', 'ColorD=');
  TraiteColor('HPP_COULLUNDI', 'ColorE=');
  TraiteColor('HPP_COULMARDI', 'ColorF=');
  TraiteColor('HPP_COULMERCREDI', 'ColorG=');
  TraiteColor('HPP_COULJEUDI', 'ColorH=');
  TraiteColor('HPP_COULVENDREDI', 'ColorI=');
  TraiteColor('HPP_COULVEILJF', 'ColorJ=');
  TraiteColor('HPP_COULJFERIE', 'ColorK=');
end;

function AglSelectionCouleurNew(parms: array of variant; nb: integer): variant;
var
  T: TColorDialog;
  F: TForm;
  Couleur: Tcolor;
begin
  Couleur := TColor(Parms[1]);
  {Correction GF 13/02/03}
  Result := IntToStr(Couleur);
  F := TForm(LongInt(parms[0]));
  T := TColorDialog.Create(F);
  T.color := Couleur;
  T.CustomColors.Clear;
  T.CustomColors.Add('ColorA=' + IntToHex(Couleur, 6));
  T.CustomColors.Add('ColorB=' + IntToHex(0, 6));
  Ajoutautrescolor(T.CustomColors, F);
  if T.Execute then
  begin
    if (T.color <> 0) then Result := IntTostr(T.Color)
    else result := '';
    T.Free;
  end;
  {Correction GF 13/02/03}
  if Result = '-1' then Result := IntToStr(clWhite);
end;

function AglSelectionFonte(parms: array of variant; nb: integer): variant;
var T: TFontDialog;
  F: TForm;
begin
  Result := '';
  // Modele := TFont (Parms[1]);
  F := TForm(LongInt(parms[0]));
  T := TFontDialog.Create(F);
  T.font.Name := THLabel(F.FindComponent('LFONTE')).Font.name;
  if T.Execute then Result := T.Font.Name;
  T.Free;
end;

procedure InitM3Proj;
begin
  RegisterAglFunc('SelectionCouleurText', True, 1, AglSelectionCouleurText);
  RegisterAglFunc('SelectionCouleurNew', True, 1, AglSelectionCouleurNew);
  RegisterAglFunc('SelectionFonte', True, 1, AglSelectionFonte);
  RegisterAglFunc('ControlChecked', True, 1, AglControlChecked);
  //RegisterAglFunc('RecalculPiece', True, 6, ScriptRecalculPiece);
  //RegisterAglProc('HREditionAnomaliesPiece',True,3, ScriptHREditionAnomaliesPiece);
end;

initialization
  InitM3Proj;

end.
