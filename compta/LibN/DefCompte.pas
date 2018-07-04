unit DefCompte;

interface

uses SysUtils,uTOB, HEnt1,IntegDef, IntegGen;

type
  TIntegDefCompte = class (TIntegDefinition)
    private
    public
      Compte : TIntegZoneDef;
      Libelle : TIntegZoneDef;
      Debit : TIntegZoneDef;
      Credit : TIntegZoneDef;
      constructor Create; 
      procedure Charge (var F : TextFile);
      procedure Enregistre (var F : TextFile);
      procedure MajValeur (St : string);
      procedure UpdateTOB (T : TOB; Ctx : TIntegContexte);
  end;

implementation

{ TIntegDefCompte }

procedure TIntegDefCompte.Charge(var F: TextFile);
var St, stEntete : string;
begin
  repeat
    Readln (F, St);
    stEntete := IntegExtractEntete (St);
    if stEntete = EX_CPOSNC then Compte.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_CTAILLENC then Compte.Lg := IntegExtractValeur(St)
    else if stEntete = EX_CPOSLC then Libelle.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_CTAILLELC then Libelle.Lg := IntegExtractValeur(St)
    else if stEntete = EX_CPOSDB then Debit.Pos := IntegExtractValeur(St)+1
    else if stEntete = EX_CTAILLEDB then Debit.Lg := IntegExtractValeur(St)
    else if stEntete = EX_CPOSCR then Credit.Pos := IntegExtractValeur(St)+1;
  until ((sT='') or (stEntete=EX_CTAILLECR));
  if stEntete = EX_CTAILLECR then Credit.Lg := IntegExtractValeur(St);
end;

constructor TIntegDefCompte.Create;
begin
  inherited;
  SetType (ilCompte);
end;

procedure TIntegDefCompte.Enregistre(var F: TextFile);
begin
  Writeln (F, EX_CPOSNC+'  '+Format('%.04d',[Compte.Pos-1]));
  Writeln (F, EX_CTAILLENC+'  '+Format('%.04d',[Compte.Lg]));
  Writeln (F, EX_CPOSLC+'  '+Format('%.04d',[Libelle.Pos-1]));
  Writeln (F, EX_CTAILLELC+'  '+Format('%.04d',[Libelle.Lg]));
  Writeln (F, EX_CPOSDB+'  '+Format('%.04d',[Debit.Pos-1]));
  Writeln (F, EX_CTAILLEDB+'  '+Format('%.04d',[Debit.Lg]));
  Writeln (F, EX_CPOSCR+'  '+Format('%.04d',[Credit.Pos-1]));
  Writeln (F, EX_CTAILLECR+'  '+Format('%.04d',[Credit.Lg]));
end;

procedure TIntegDefCompte.MajValeur(St: string);
begin
  Compte.Val := Copy (St,Compte.Pos,Compte.Lg);
  Libelle.Val := Copy (St,Libelle.Pos,Libelle.Lg);
  Debit.Val := Valeur(Copy (St,Debit.Pos,Debit.Lg));
  Credit.Val := Valeur(Copy (St,Credit.Pos,Credit.Lg));
end;

procedure TIntegDefCompte.UpdateTOB(T: TOB; Ctx : TIntegContexte);
var stCompte, stAuxi : string;
begin
  stCompte := Compte.Val; stAuxi := '';
  IntegGetCompteEquiv (Ctx ,stCompte, stAuxi);
  T.AddChampSupValeur('General',stCompte);
  T.AddChampSupValeur('Auxiliaire',stAuxi);
  T.AddChampSupValeur('Libelle',Libelle.Val);
  T.AddChampSupValeur('Debit',Debit.Val);
  T.AddChampSupValeur('Credit',Credit.Val);      
end;

end.
