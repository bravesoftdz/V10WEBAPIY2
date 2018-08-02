unit UtilBTPVerdon;

interface

type

  T_TablesName = (tnNone, tnChantier, tnDevis, tnLignesBR, tnTiers);

  TUtilBTPVerdon = class (TObject)
  public
    class function GetBTPTableName(Tn : T_TablesName) : string;
    class function GetSynchroTableName(Tn : T_TablesName) : string;
  end;

const
  DBSynchroName = 'VERDON_TAMPON';

implementation

{ TUtilBTPVerdon }

class function TUtilBTPVerdon.GetBTPTableName(Tn : T_TablesName): string;
begin
  case Tn of
    tnChantier : Result := 'AFFAIRE';
    tnDevis    : Result := 'PIECE';
    tnLignesBR : Result := 'LIGNE';
    tnTiers    : Result := 'TIERS';
  else
    Result := '';
  end;
end;

class function TUtilBTPVerdon.GetSynchroTableName(Tn : T_TablesName): string;
begin
  case Tn of
    tnChantier : Result := 'CHANTIER';
    tnDevis    : Result := 'DEVIS';
    tnLignesBR : Result := 'LIGNESBR';
    tnTiers    : Result := 'TIERS';
  else
    Result := '';
  end;
end;

end.
