unit PGTobOutils;

interface

uses
  uTob;

  procedure EclateTob(var FromTob, ToTob: Tob; ChampCritere: String; Duplicate : Boolean = True); // ; NoDoublons : Boolean = False

implementation

          //TGestionPresence.
Procedure EclateTob(var FromTob, ToTob : Tob; ChampCritere : String; Duplicate : Boolean = True);  //; NoDoublons : Boolean = False
var
  indexTob : Integer;
  TempTob, FilleTob, PetiteFilleTob : Tob;
  TempValue : String;
  NomTob, NomTobPetiteFille : String;
//  function FindTobInTob(SearchTob, InTob : Tob) : Boolean;
//  var
//    index, indexInTob : Integer;
//    TempInTob : Tob;
//    NomChamp : String;
//  begin
//    result := False;
//    { On parcours les tobs filles de la tob dans laquelle on effectue la recherche }
//    for indexInTob := 0 to InTob.detail.count -1 do
//    begin
//      TempInTob := InTob.detail[indexInTob];
//      { Pour chaque champs (réel) de la tob à chercher }
//      result := True;
//      for index := 1 to SearchTob.NombreChampReel do
//      begin
//        { On compare la valeur du champs avec la valeur du même champs dans la fille courante }
//        NomChamp := SearchTob.GetNomChamp(index);
//        if not InTob.FieldExists(NomChamp) then
//        begin
//          result := False;
//          exit;
//        end;
//        if TempInTob.GetValue(NomChamp) <> SearchTob.GetValue(NomChamp) then
//        begin
//          result := False;
//          break;  //On passe à la fille suivante
//        end;
//        //Sinon on passe au champs suivant
//      end;
//      if result then exit;
//    end;
//  end;
begin
  if not Assigned(ToTob) then
    ToTob := Tob.create(FromTob.NomTable, nil, -1);
  NomTob := ToTob.NomTable;
  if FromTob.Detail.Count >= 1 then
    NomTobPetiteFille := FromTob.Detail[0].NomTable
  else
    NomTobPetiteFille := '';
  for indexTob := FromTob.Detail.Count -1 downto 0 do
  begin
    TempTob := FromTob.Detail[indexTob];
    TempValue := TempTob.GetString(ChampCritere);
    FilleTob := ToTob.FindFirst([ChampCritere], [TempValue], False);
    if not Assigned(FilleTob) then
    begin
      FilleTob := TOB.Create(NomTob, ToTob, -1);
      FilleTob.AddChampSupValeur(ChampCritere, TempValue);
    end;
//    if (not NoDoublons) or (NoDoublons and (not FindTobInTob(TempTob, FilleTob))) then
//    begin
    if Duplicate then
    begin
      PetiteFilleTob := TOB.Create(NomTob, FilleTob, -1);
      PetiteFilleTob.Dupliquer(TempTob, True, True, True);
    end else begin
      TempTob.ChangeParent(FilleTob, -1);
    end;
//    end;
  end;
end;

end.
