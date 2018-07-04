unit CreePieceOle;

interface

uses
  {$IFNDEF EAGLCLIENT}
  eBizUtil,
  {$ENDIF !EAGLCLIENT}
  UTob, HEnt1, HCtrls, SysUtils, GerePiece, M3FP, FactGrp;

function CreeLaPieceDepuisFichier (stFichier : string) : variant;
function CreerLaPieceTob (TobFile : TOB; TobResultat : Tob = nil) : string;

implementation
uses 
   CbpMCD
   ,CbpEnumerator
;

procedure PasseLesChamps (TobOrigine, TobArrivee : Tob; stTable : string = '');
var iIndChamp : integer;
    iTable, iChamp : integer;
    stPrefixe : string;
    NomChamps : string;
  Mcd : IMCDServiceCOM;
  Table     : ITableCOM ;
  FieldList : IEnumerator ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  for iIndChamp := 1000 to 1000 + TobOrigine.ChampsSup.Count - 1 do
  begin
    //TobArrivee.AddChampSupValeur (TobOrigine.GetNomChamp (iIndChamp), TobOrigine.GetString (TobOrigine.GetNomChamp (iIndChamp)));
		nomchamps := stPrefixe + '_' + TobOrigine.GetNomChamp (iIndChamp);
    if assigned(mcd.getField(nomchamps)) then
    begin
      if (mcd.getField(nomchamps).tipe = 'DATE') then
        TobArrivee.AddChampSupValeur (TobOrigine.GetNomChamp (iIndChamp), StrToDate (TobOrigine.GetString (TobOrigine.GetNomChamp (iIndChamp))))
      else if (mcd.getField(nomchamps).tipe = 'DOUBLE') then
        TobArrivee.AddChampSupValeur (TobOrigine.GetNomChamp (iIndChamp), VALEUR (TobOrigine.GetString (TobOrigine.GetNomChamp (iIndChamp))))
      else if (mcd.getField(nomchamps).tipe = 'INTEGER') then
        TobArrivee.AddChampSupValeur (TobOrigine.GetNomChamp (iIndChamp), StrToInt (TobOrigine.GetString (TobOrigine.GetNomChamp (iIndChamp))))
      else
        TobArrivee.AddChampSupValeur (TobOrigine.GetNomChamp (iIndChamp), TobOrigine.GetValue (TobOrigine.GetNomChamp (iIndChamp)));
      end;
  end;
end;

function CreerLaPieceTob (TobFile : TOB; TobResultat : Tob = nil) : string;
{$IFNDEF BTP}
var
  TobData, TobLigneData, TobGQData, TobGQDData, TobPiece, TobLigne, TobGQServi, TobGQDServi : Tob;
  iIndPiece, iIndLigne, iIndGQ, iIndGQD : integer;
{$ENDIF}
begin
{$IFNDEF BTP}
	// on enleve pour l'instant bicose les fonctions appelés ne sont pas Batiment
  // pb de compilation assurée
  TobData := Tob.Create ('', nil, -1);
  if TobFile.Detail.Count > 0 then
  begin
    for iIndPiece := 0 to TobFile.Detail.Count - 1 do
    begin
      TobData.ClearDetail;
      TobPiece := TobFile.Detail[iIndPiece];
      PasseLesChamps (TobPiece, TobData, 'PIECE');
      for iIndLigne := 0 to TobPiece.Detail.Count - 1 do
      begin
        TobLigne := TobPiece.Detail[iIndLigne];
        TobLigneData := Tob.Create ('', TobData, -1);
        PasseLesChamps (TobLigne, TobLigneData, 'LIGNE');
        for iIndGQ := 0 to TobLigne.Detail.Count - 1 do
        begin
          TobGQServi := TobLigne.Detail[iIndGQ];
          TobGQData := Tob.Create ('', TobLigneData, -1);
          PasseLesChamps (TobGQServi, TobGQData, 'DISPO');
          for iIndGQD := 0 to TobGQServi.Detail.Count - 1 do
          begin
            TobGQDServi := TobGQServi.Detail[iIndGQD];
            TobGQDData := Tob.Create ('', TobGQData, -1);
            PasseLesChamps (TobGQDServi, TobGQDData, 'DISPODETAIL');
          end;
        end;
      end;
      if TobData.Detail.Count > 0 then NewCreatePieceFromTob (TobData);
      if TobResultat <> nil then
      begin
        if not TobResultat.FieldExists ('VALIDE') then TobResultat.AddChampSup ('VALIDE', false);
        if TobData.FieldExists ('GP_NUMERO') then
        begin
          TobResultat.SetInteger ('VALIDE', 0);
          TobResultat.AddChampSupValeur ('NUMERO', TobData.GetInteger ('GP_NUMERO'));
        end else TobResultat.SetInteger ('VALIDE', 2);
      end;
      if TobData.FieldExists ('Error') then
        Result := TobData.GetString ('Error')
      else Result := '';
    end;
  end;
  TobData.Free;
{$ENDIF}
end;

function CreeLaPieceDepuisFichier (stFichier : string) : variant;
var
  {$IFDEF EAGLCLIENT}
  TobFile : TOB;
  {$ELSE EAGLCLIENT}
  TobFile : TOB_2;
  {$ENDIF !EAGLCLIENT}
begin
  {$IFDEF EAGLCLIENT}
  TobFile := Tob.Create ('', nil, -1);
  TOBLoadFromFile (stFichier, nil, TobFile);
  //TobFile.SaveToXmlFile (stFichier + '.xml', true, true);
  {$ELSE EAGLCLIENT}
  TobFile := Tob_2.Create ('', nil, -1);
  TobFile.LoadFromXMLFile (stFichier);
  {$ENDIF !EAGLCLIENT}

  Result := CreerLaPieceTob (TobFile);

  TobFile.Free;
end;

end.
