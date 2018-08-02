unit UtilBTPVerdon;

interface

type

  T_TablesName = (tnNone, tnChantier, tnDevis, tnLignesBR, tnTiers);

  T_TiersValues    = Record
                       FirstExec : boolean;
                       Count     : integer;
                       TimeOut   : integer;
                     end;

  T_ChantierValues = Record
                      FirstExec : boolean;
                      Count     : integer;
                      TimeOut   : integer;
                     end;

  T_DevisValues    = Record
                      FirstExec : boolean;
                      Count     : integer;
                      TimeOut   : integer;
                     end;

  T_LignesBRValues = Record
                      FirstExec : boolean;
                      Count     : integer;
                      TimeOut   : integer;
                     end;

  TUtilBTPVerdon = class (TObject)
  public
    class function GetBTPTableName(Tn : T_TablesName) : string;
    class function GetSynchroTableName(Tn : T_TablesName) : string;
    class function GetMsg(Tn : T_TablesName; Start : boolean) : string;

  end;

const
  DBSynchroName = 'VERDON_TAMPON';

implementation

uses
  CommonTools
  , SysUtils
  ;

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

class function TUtilBTPVerdon.GetMsg(Tn : T_TablesName; Start : boolean) : string;
begin
  Result := Format('%s de traitement de la table %s.', [Tools.iif(Start, 'Début', 'Fin'), TUtilBTPVerdon.GetSynchroTableName(Tn)]);
end;

end.
