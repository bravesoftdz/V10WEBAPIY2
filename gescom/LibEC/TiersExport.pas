unit TiersExport;

interface
uses GenExport, UTOB, ETransUtil;

type
    TOF_ExportTiers = class(TOF_ExportGen)
    private

    published
      procedure ExportTiaires(Fmt : TTOBFormat);
      function ExporteAussiLeReste(T : TOB) : integer;

    end;


implementation
uses SysUtils, Classes, Forms, dbtables, HCtrls, Mul, UTOF, EntGC, M3FP;

procedure AGLExportTiers(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then TOTOF := TFMul(F).LaTOF
                  else exit;
  if (TOTOF is TOF_ExportTiers) then TOF_ExportTiers(TOTOF).ExportTiaires(DefaultFormat)
                                else exit;
end;

procedure TOF_ExportTiers.ExportTiaires(Fmt : TTOBFormat);
begin
  //ExportGenerik('TIERS', ['T_NATUREAUXI'], Fmt, ExporteAussiLeReste);  // ???????????????????????????
  ExportGenerik('TIERS', ['T_TIERS'], Fmt, ExporteAussiLeReste);
end;

function TOF_ExportTiers.ExporteAussiLeReste(T : TOB) : integer;
var TCon, TTar, TLO : TOB;
    i : integer;
    Q : TQuery;
    CodTiers : String;
begin
  result := 0;
  FExp.SetTTMaxValue(T.Detail.Count+2);

  TCon := TOB.Create('Mama_CONTACT', nil, -1);
  TTar := TOB.Create('Mama_TARIF', nil, -1);
  TLO := TOB.Create('Mama_LIENSOLE', nil, -1);
  for i := 0 to T.Detail.Count-1 do
  begin
    CodTiers := T.Detail[i].GetValue('ETI_AUXILIAIRE');  // La table est convertie en etable
    FExp.SetTTSubText('Traitement des données associées au tiers '+CodTiers);

    Q := OpenSQL('SELECT * FROM CONTACT WHERE C_AUXILIAIRE="'+CodTiers+'"', true);
    TCon.LoadDetailDB('CONTACT', '', '', Q, true);
    Ferme(Q);

// Modif BBI 09/04 suite à nouveaux Tarifs
    DemoulineTarif;

    Q := OpenSQL('SELECT * FROM TARIF WHERE GF_TIERS="'+CodTiers+'" AND GF_NATUREAUXI<>"FOU"', true);
    TTar.LoadDetailDB('TARIF', '', '', Q, true);
    Ferme(Q);

    if LiensOLEQuandTiers then
    begin
       Q := OpenSQL('SELECT * FROM LIENSOLE WHERE LO_TABLEBLOB="T" AND LO_IDENTIFIANT="'+CodTiers+'"', true);
       TLO.LoadDetailDB('LIENSOLE', '', '', Q, true);
       Ferme(Q);
    end;

    FExp.SetTTValue(i+2);
    if FExp.TTCanceled then break;
  end;
  DeRTFizeLiensOLE(TLO);

  if not FExp.TTCanceled then
  begin
    Insert('E', FBaseName, Pos('_', FBaseName)+1);
    FExp.TOBToFormat(TCon, FBaseName+'_CONTACT', true, FTF);
    FExp.TOBToFormat(TTar, FBaseName+'_TARIF', true, FTF);
    if LiensOLEQuandTiers then FExp.TOBToFormat(TLO, FBaseName+'_LIENSOLE', true, FTF);
  end;

  TTar.Free;
  TCon.Free;
  TLO.Free;
end;


initialization
RegisterClasses([TOF_ExportTiers]);
RegisterAGLProc('ExportTiers', true, 0, AGLExportTiers);

end.
