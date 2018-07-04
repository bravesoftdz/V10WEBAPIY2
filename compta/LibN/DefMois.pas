unit DefMois;

interface

uses SysUtils, HEnt1,IntegDef, IntegGen;

type
  TIntegDefMois = class (TIntegDefinition)
    private
    public
      Jour : TIntegZoneDef;
      Mois : TIntegZoneDef;
      Annee : TIntegZoneDef;
      MoisRel : TIntegZoneDef;
      Journal : TIntegZoneDef;
      constructor Create;
      procedure Charge (var F : TextFile);
      procedure Enregistre (var F : TextFile);      
      procedure MajValeur (St : string);
      procedure GetDate (var Ctx : TIntegCtxEcriture);
  end;


implementation

{ TIntegDefMois }

procedure TIntegDefMois.Charge(var F: TextFile);
var St, stEntete : string;
begin
  repeat
    Readln (F, St);
    stEntete := IntegExtractEntete (St);
    if stEntete = EX_MPOSDJ then Jour.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_MPOSDM then Mois.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_MPOSDA then Annee.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_MPOSMR then MoisRel.Lg := IntegExtractValeur(St)
    else if stEntete = EX_MPOSJX then Journal.Pos := IntegExtractValeur(St)+1;
  until ((sT='') or (stEntete=EX_MTAILLEJX));
  if stEntete = EX_MTAILLEJX then Journal.Lg := IntegExtractValeur(St);
end;

constructor TIntegDefMois.Create;
begin
  inherited;
  SetType (ilMois);
end;

procedure TIntegDefMois.Enregistre(var F: TextFile);
begin
  Writeln (F, EX_MPOSDJ+'  '+Format('%.04d',[Jour.Pos-1]));
  Writeln (F, EX_MPOSDM+'  '+Format('%.04d',[Mois.Pos-1]));
  Writeln (F, EX_MPOSDA+'  '+Format('%.04d',[Annee.Pos-1]));
  Writeln (F, EX_MPOSMR+'  '+Format('%.04d',[MoisRel.Pos-1]));
  Writeln (F, EX_MPOSJX+'  '+Format('%.04d',[Journal.Pos-1]));
  Writeln (F, EX_MTAILLEJX+'  '+Format('%.04d',[Journal.Lg]));
end;

procedure TIntegDefMois.GetDate (var Ctx : TIntegCtxEcriture);
begin
  Ctx.Jour := Jour.Val;
  Ctx.Mois := Mois.Val;
  Ctx.Annee := Annee.Val;
  Ctx.MoisRel := MoisRel.Val;
end;

procedure TIntegDefMois.MajValeur(St: string);
begin
  Jour.Val := Valeur(Copy (St,Jour.Pos,2));
  Mois.Val := Valeur(Copy (St,Mois.Pos,2));
  Annee.Val := Valeur(Copy (St,Annee.Pos,2));
  MoisRel.Val := Valeur(Copy (St,MoisRel.Pos,2));
  Journal.Val := Copy (St,Journal.Pos,Journal.Lg);
end;


end.

