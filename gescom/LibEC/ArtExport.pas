unit ArtExport;

interface
uses GenExport, UTOB, ETransUtil;

type
    TOF_ExportArt = class(TOF_ExportGen)
    private
      TOBTraitement : TOB;
      function CopyTOBDim(T : TOB) : integer;

    public
      FForceFiles : boolean;
      FNoTarif : boolean;

    published
      procedure ExportArtick(Fmt : TTOBFormat);
      function ExporteAussiLeReste(T : TOB) : integer;

      property TobFormat : TTOBFormat read FTF write FTF;

    end;


implementation
uses SysUtils, Classes, Forms, dbtables, HCtrls, Mul, UTOF, EntGC, M3FP;

procedure AGLExportArt(Parms : Array of Variant; Nb : Integer);
var F : TForm;
    TOTOF : TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then TOTOF := TFMul(F).LaTOF
                  else exit;
  if (TOTOF is TOF_ExportArt) then TOF_ExportArt(TOTOF).ExportArtick(DefaultFormat)
                              else exit;
end;

procedure TOF_ExportArt.ExportArtick(Fmt : TTOBFormat);
begin
  ExportGenerik('ARTICLE', ['GA_ARTICLE'], Fmt, ExporteAussiLeReste);
end;

function TOF_ExportArt.CopyTOBDim(T : TOB) : integer;
begin
  result := 0;
  if TOBTraitement.FindFirst(['GDI_TYPEDIM', 'GDI_GRILLEDIM', 'GDI_CODEDIM'],
                             [T.GetValue('GDI_TYPEDIM'), T.GetValue('GDI_GRILLEDIM'), T.GetValue('GDI_CODEDIM')],
                             false) = nil
    then TOB.Create('DIMENSION', TOBTraitement, -1).Dupliquer(T, true, true);
end;

function TOF_ExportArt.ExporteAussiLeReste(T : TOB) : integer;
var //TAL,TLO,TTar,TCC,TCE,TGTA,TNE,TNL : TOB;
    T1, T2 : TOB;
    i, tcount, i_dim : integer;
    Q : TQuery;
    RefArt, CodArt, GrD, CoD, NomFic : String;
const NbAssoc = 8;
begin
  result := 0;
  tcount := T.Detail.Count;
  FExp.SetTTMaxValue(2+(tcount*NbAssoc)+2);

  Insert('E', FBaseName, Pos('_', FBaseName)+1);

  TOBTraitement := TOB.Create('Mama_DIMENSION', nil, -1);
  FExp.SetTTSubText('Export des dimensions associées aux articles...');
  for i := 0 to T.Detail.Count-1 do
  begin
    for i_dim := 1 to 5 do
    begin
      GrD := T.Detail[i].GetValue('EA_GRILLEDIM'+inttostr(i_dim));
      CoD := T.Detail[i].GetValue('EA_CODEDIM'+inttostr(i_dim));

      VH_GC.GCTOBDim.ParcoursTraitement(['GDI_TYPEDIM', 'GDI_GRILLEDIM', 'GDI_CODEDIM'],
                                        ['DI'+inttostr(i_dim), GrD, CoD],
                                        false, CopyTOBDim);
    end;
    FExp.SetTTValue(2+(tcount*0)+i);
    if FExp.TTCanceled then break;
  end;
  NomFic := FBaseName+'_DIM';
  if not FExp.TTCanceled then
    FExp.TOBToFormat(TOBTraitement, NomFic, not FForceFiles, FTF);
  TOBTraitement.Free;

  T1 := TOB.Create('Mama_ARTICLELIE', nil, -1);
  FExp.SetTTSubText('Export des articles liés associés aux articles...');
  for i := 0 to T.Detail.Count-1 do
  begin
    CodArt := T.Detail[i].GetValue('EA_CODEARTICLE');

    Q := OpenSQL('SELECT * FROM ARTICLELIE WHERE GAL_ARTICLE="'+CodArt+'"', true);
    T1.LoadDetailDB('ARTICLELIE', '', '', Q, true);
    Ferme(Q);
    FExp.SetTTValue(2+(tcount*1)+i);
    if FExp.TTCanceled then break;
  end;
  NomFic := FBaseName+'_ARTICLELIE';
  if not FExp.TTCanceled then
    FExp.TOBToFormat(T1, NomFic, not FForceFiles, FTF);
  T1.Free;

  T1 := TOB.Create('Mama_LIENSOLE', nil, -1);
  FExp.SetTTSubText('Export des objets OLE associés aux articles...');
  for i := 0 to T.Detail.Count-1 do
  begin
    RefArt := T.Detail[i].GetValue('EA_ARTICLE');

    Q := OpenSQL('SELECT * FROM LIENSOLE WHERE LO_TABLEBLOB="GA" AND LO_IDENTIFIANT="'+RefArt+'" AND (LO_QUALIFIANTBLOB="PHJ" OR LO_QUALIFIANTBLOB="VIJ")', true);
    T1.LoadDetailDB('LIENSOLE', '', '', Q, true);
    Ferme(Q);
    FExp.SetTTValue(2+(tcount*2)+i);
    if FExp.TTCanceled then break;
  end;
  NomFic := FBaseName+'_LIENSOLE';
  DeRTFizeLiensOLE(T1);
  if not FExp.TTCanceled then
    FExp.TOBToFormat(T1, NomFic, not FForceFiles, FTF);
  T1.Free;

if not FNoTarif then
begin
  T1 := TOB.Create('Mama_TARIF', nil, -1);
  FExp.SetTTSubText('Export des tarifs associés aux articles...');
  for i := 0 to T.Detail.Count-1 do
  begin
    RefArt := T.Detail[i].GetValue('EA_ARTICLE');
// Modif BBI 09/04 suite à nouveaux Tarifs
    DemoulineTarif;

    Q := OpenSQL('SELECT * FROM TARIF WHERE GF_ARTICLE="'+RefArt+'" AND GF_NATUREAUXI<>"FOU"', true);
    T1.LoadDetailDB('TARIF', '', '', Q, true);
    Ferme(Q);
    FExp.SetTTValue(2+(tcount*3)+i);
    if FExp.TTCanceled then break;
  end;
  NomFic := FBaseName+'_TARIF';
  if not FExp.TTCanceled then
    FExp.TOBToFormat(T1, NomFic, not FForceFiles, FTF);
  T1.Free;
end;

  T1 := TOB.Create('Mama_GTRADARTICLE', nil, -1);
  FExp.SetTTSubText('Export des traductions articles associées aux articles...');
  for i := 0 to T.Detail.Count-1 do
  begin
    RefArt := T.Detail[i].GetValue('EA_ARTICLE');

    Q := OpenSQL('SELECT * FROM GTRADARTICLE WHERE GTA_ARTICLE="'+RefArt+'"', true);
    T1.LoadDetailDB('GTRADARTICLE', '', '', Q, true);
    Ferme(Q);
    FExp.SetTTValue(2+(tcount*4)+i);
    if FExp.TTCanceled then break;
  end;
  NomFic := FBaseName+'_GTRADARTICLE';
  if not FExp.TTCanceled then
    FExp.TOBToFormat(T1, NomFic, not FForceFiles, FTF);
  T1.Free;

  T1 := TOB.Create('Mama_NOMENENT', nil, -1);
  T2 := TOB.Create('Mama_NOMENLIG', nil, -1);
  FExp.SetTTSubText('Export des nomenclatures associées aux articles...');
  for i := 0 to T.Detail.Count-1 do
  begin
    RefArt := T.Detail[i].GetValue('EA_ARTICLE');

    Q := OpenSQL('SELECT * FROM NOMENENT WHERE GNE_ARTICLE="'+RefArt+'"', true);
    T1.LoadDetailDB('NOMENENT', '', '', Q, true);
    Ferme(Q);
    FExp.SetTTValue(2+(tcount*5)+i);
    if FExp.TTCanceled then break;
  end;

  for i := 0 to T1.Detail.Count-1 do
  begin
    Q := OpenSQL('SELECT * FROM NOMENLIG WHERE GNL_NOMENCLATURE="'+T1.Detail[i].GetValue('GNE_NOMENCLATURE')+'"', true);
    T2.LoadDetailDB('NOMENLIG', '', '', Q, true);
    Ferme(Q);
    if FExp.TTCanceled then break;
  end;

  if not FExp.TTCanceled then
  begin
    NomFic := FBaseName+'_NOMENENT';
    FExp.TOBToFormat(T1, NomFic, not FForceFiles, FTF);
    NomFic := FBaseName+'_NOMENLIG';
    FExp.TOBToFormat(T2, NomFic, not FForceFiles, FTF);
  end;
  T1.Free;
  T2.Free;

  T1 := TOB.Create('Mama_CONDITIONNEMENT', nil, -1);
  FExp.SetTTSubText('Export des conditionnements associés aux articles...');
  for i := 0 to T.Detail.Count-1 do
  begin
    RefArt := T.Detail[i].GetValue('EA_ARTICLE');

    Q := OpenSQL('SELECT * FROM CONDITIONNEMENT WHERE GCO_ARTICLE="'+RefArt+'"', true);
    T1.LoadDetailDB('CONDITIONNEMENT', '', '', Q, true);
    Ferme(Q);
    FExp.SetTTValue(2+(tcount*6)+i);
    if FExp.TTCanceled then break;
  end;
  NomFic := FBaseName+'_CONDITIONNEMENT';
  if not FExp.TTCanceled then
    FExp.TOBToFormat(T1, NomFic, not FForceFiles, FTF);
  T1.Free;

  T1 := TOB.Create('Mama_CHOIXCOD', nil, -1);
  FExp.SetTTSubText('Export des libellés des familles...');
  Q := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE LIKE "FN%" OR CC_TYPE LIKE "GG%"', true); // Libellés familles
  T1.LoadDetailDB('CHOIXCOD', '', '', Q, true);
  Ferme(Q);
  NomFic := FBaseName+'_CHOIXCOD';
  if not FExp.TTCanceled then
    FExp.TOBToFormat(T1, NomFic, not FForceFiles, FTF);
  T1.Free;

  T1 := TOB.Create('Mama_CHOIXEXT', nil, -1);
  FExp.SetTTSubText('Export des libellés des tables libres...');
  Q := OpenSQL('SELECT * FROM CHOIXEXT WHERE YX_TYPE LIKE "LA%"', true); // Libellés des Tables Libre des tables ARTICLES
  T1.LoadDetailDB('CHOIXEXT', '', '', Q, true);
  Ferme(Q);
  NomFic := FBaseName+'_CHOIXEXT';
  if not FExp.TTCanceled then
    FExp.TOBToFormat(T1, NomFic, not FForceFiles, FTF);
  T1.Free;
end;

(*
function TOF_ExportArt.ExporteAussiLeReste(T : TOB) : integer;
var TAL,TLO,TTar,TCC,TCE,TGTA,TNE,TNL : TOB;
    i, j, i_dim : integer;
    Q : TQuery;
    RefArt, CodArt, GrD, CoD : String;
begin
  result := 0;
  FExp.SetTTMaxValue(T.Detail.Count+2);

  TAL := TOB.Create('Mama_ARTICLELIE', nil, -1);
  TLO := TOB.Create('Mama_LIENSOLE', nil, -1);
  TTar := TOB.Create('Mama_TARIF', nil, -1);
  TOBTraitement := TOB.Create('Mama_DIMENSION', nil, -1);
  TCC := TOB.Create('Mama_CHOIXCOD', nil, -1);
  TCE := TOB.Create('Mama_CHOIXEXT', nil, -1);
  TGTA := TOB.Create('Mama_GTRADARTICLE', nil, -1);
  TNE := TOB.Create('Mama_NOMENENT', nil, -1);
  TNL := TOB.Create('Mama_NOMENLIG', nil, -1);

  for i := 0 to T.Detail.Count-1 do
  begin
    CodArt := T.Detail[i].GetValue('EA_CODEARTICLE');
    RefArt := T.Detail[i].GetValue('EA_ARTICLE');
    FExp.SetTTSubText('Export des données associées à l''article '+CodArt);

    for i_dim := 1 to 5 do
    begin
      GrD := T.Detail[i].GetValue('EA_GRILLEDIM'+inttostr(i_dim));
      CoD := T.Detail[i].GetValue('EA_CODEDIM'+inttostr(i_dim));

      VH_GC.GCTOBDim.ParcoursTraitement(['GDI_TYPEDIM', 'GDI_GRILLEDIM', 'GDI_CODEDIM'],
                                        ['DI'+inttostr(i_dim), GrD, CoD],
                                        false, CopyTOBDim);
    end;

    Q := OpenSQL('SELECT * FROM ARTICLELIE WHERE GAL_ARTICLE="'+CodArt+'"', true);
    TAL.LoadDetailDB('ARTICLELIE', '', '', Q, true);
    Ferme(Q);

    Q := OpenSQL('SELECT * FROM LIENSOLE WHERE LO_TABLEBLOB="GA" AND LO_IDENTIFIANT="'+RefArt+'" AND (LO_QUALIFIANTBLOB="PHJ" OR LO_QUALIFIANTBLOB="VIJ")', true);
    TLO.LoadDetailDB('LIENSOLE', '', '', Q, true);
    Ferme(Q);

    Q := OpenSQL('SELECT * FROM TARIF WHERE GF_ARTICLE="'+RefArt+'" AND GF_NATUREAUXI<>"FOU"', true);
    TTar.LoadDetailDB('TARIF', '', '', Q, true);
    Ferme(Q);

    Q := OpenSQL('SELECT * FROM GTRADARTICLE WHERE GTA_ARTICLE="'+RefArt+'"', true);
    TGTA.LoadDetailDB('GTRADARTICLE', '', '', Q, true);
    Ferme(Q);

    Q := OpenSQL('SELECT * FROM NOMENENT WHERE GNE_ARTICLE="'+RefArt+'"', true);
    TNE.LoadDetailDB('NOMENENT', '', '', Q, true);
    Ferme(Q);

    for j := 0 to TNE.Detail.Count-1 do
    begin
      Q := OpenSQL('SELECT * FROM NOMENLIG WHERE GNL_NOMENCLATURE="'+TNE.Detail[j].GetValue('GNE_NOMENCLATURE')+'"', true);
      TNL.LoadDetailDB('NOMENLIG', '', '', Q, true);
      Ferme(Q);
    end;

    FExp.SetTTValue(i+2);
    if FExp.TTCanceled then break;
  end;

  Q := OpenSQL('SELECT * FROM CHOIXCOD WHERE CC_TYPE LIKE "FN%" OR CC_TYPE LIKE "GG%"', true); // Libellés familles
  TCC.LoadDetailDB('CHOIXCOD', '', '', Q, true);
  Ferme(Q);

  Q := OpenSQL('SELECT * FROM CHOIXEXT WHERE YX_TYPE LIKE "LA%"', true); // Libellés des Tables Libre des tables ARTICLES
  TCE.LoadDetailDB('CHOIXCOD', '', '', Q, true);
  Ferme(Q);

  DeRTFizeLiensOLE(TLO);

  if not FExp.TTCanceled then
  begin
    Insert('E', FBaseName, Pos('_', FBaseName)+1);
    FExp.TOBToFormat(TAL, FBaseName+'_ARTICLELIE', not FForceFiles, FTF);
    FExp.TOBToFormat(TOBTraitement, FBaseName+'_DIM', not FForceFiles, FTF);
    FExp.TOBToFormat(TTar, FBaseName+'_TARIF', not FForceFiles, FTF);
    FExp.TOBToFormat(TLO, FBaseName+'_LIENSOLE', not FForceFiles, FTF);
    FExp.TOBToFormat(TCC, FBaseName+'_CHOIXCOD', not FForceFiles, FTF);
    FExp.TOBToFormat(TCE, FBaseName+'_CHOIXEXT', not FForceFiles, FTF);
    FExp.TOBToFormat(TGTA, FBaseName+'_GTRADARTICLE', not FForceFiles, FTF);
    FExp.TOBToFormat(TNE, FBaseName+'_NOMENENT', not FForceFiles, FTF);
    FExp.TOBToFormat(TNL, FBaseName+'_NOMENLIG', not FForceFiles, FTF);
  end;

  TTar.Free;
  TAL.Free;
  TLO.Free;
  TOBTraitement.Free;
  TCC.Free;
  TCE.Free;
  TGTA.Free;
  TNE.Free;
  TNL.Free;
end;
*)

initialization
RegisterClasses([TOF_ExportArt]);
RegisterAGLProc('ExportArt', true, 0, AGLExportArt);

end.
