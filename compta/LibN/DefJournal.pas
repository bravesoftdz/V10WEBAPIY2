unit DefJournal;

interface

uses SysUtils,HEnt1, IntegDef, IntegGen;

type
  TIntegDefJournal = class (TIntegDefinition)
    private
    public
      Journal : TIntegZoneDef;
      Libelle : TIntegZoneDef;
      constructor Create;
      procedure Charge (var F : TextFile);
      procedure Enregistre (var F : TextFile);      
      procedure MajValeur (St : string);
      function GetJournal : string;
  end;

implementation

{ TIntegDefJournal }

procedure TIntegDefJournal.Charge(var F: TextFile);
var St, stEntete : string;
begin
  repeat
    Readln (F, St);
    stEntete := IntegExtractEntete (St);
    if stEntete = EX_JPOSJ then Journal.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_JTAILLEJ then Journal.Lg := IntegExtractValeur(St)
    else if stEntete = EX_JPOSL then Libelle.Pos := IntegExtractValeur(St)+1;
  until ((sT='') or (stEntete=EX_JTAILLEL));
  if stEntete = EX_JTAILLEL then Libelle.Lg := IntegExtractValeur(St);
end;

constructor TIntegDefJournal.Create;
begin
  inherited;
  SetType (ilJournal);
end;

procedure TIntegDefJournal.Enregistre(var F: TextFile);
begin
  Writeln (F, EX_JPOSJ+'  '+Format('%.04d',[Journal.Pos-1]));
  Writeln (F, EX_JTAILLEJ+'  '+Format('%.04d',[Journal.Lg]));
  Writeln (F, EX_JPOSL+'  '+Format('%.04d',[Libelle.Pos-1]));
  Writeln (F, EX_JTAILLEL+'  '+Format('%.04d',[Libelle.Lg]));
end;

function TIntegDefJournal.GetJournal: string;
begin
  Result := Journal.Val;
end;

procedure TIntegDefJournal.MajValeur(St: string);
begin
  Journal.Val := Copy (St,Journal.Pos,Journal.Lg);
  Libelle.Val := Copy (St,Libelle.Pos,Libelle.Lg);
end;

end.
