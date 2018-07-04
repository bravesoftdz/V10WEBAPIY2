unit IntegDef;

interface

uses classes, SysUtils, IntegGen, Math;

type
  TIntegDefinition = class (TPersistent)
    private
      FType : TIntegTypeLigne;
      FContrainte : string;
    public
      constructor Create;
      procedure SetType (T : TIntegTypeLigne);
      procedure SetContrainte (St : string);
      procedure UpdateContrainte (Pos : integer; St : string);
      function Analyse (St : string) : boolean;
    published
      property TypeLigne : TIntegTypeLigne read FType write SetType;
      property Contrainte : string read FContrainte write SetContrainte;
  end;

implementation

uses DefCompte, DefEcriture, DefJournal, DefMois;

{ TIntegDefinition }

function TIntegDefinition.Analyse(St: string): boolean;
var i : integer;
  bOK : boolean;
begin
  bOK := True;
  for i:=1 to Length(FContrainte) do
  begin
    if FContrainte[i]='#' then // OK si #
    else if ((i > Length(St)) or (FContrainte[i] <> St[i])) then
    begin
      bOK := False;
      break;
    end;
  end;
  Result := bOK;
end;

constructor TIntegDefinition.Create;
begin
  Contrainte := '##################################################'+
              '##################################################'+
              '##################################################'+
              '##################################################'+
              '##################################################'+
              '#####';                                          
end;

procedure TIntegDefinition.SetContrainte(St: string);
begin
  FContrainte := St;
end;

procedure TIntegDefinition.SetType(T: TIntegTypeLigne);
begin
  FType := T;
end;

procedure TIntegDefinition.UpdateContrainte(Pos: integer; St: string);
begin
  Contrainte := Copy (Contrainte,1,Pos-1)+St+Copy (Contrainte,Pos+Length(St),Length(Contrainte)-Pos-Length(St));
end;

end.
