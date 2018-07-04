unit IntegBase;

interface

uses classes, sysutils, HMsgBox,
      // uses intégrateur
      IntegGen, IntegDef, DefCompte, DefEcriture, DefJournal, DefMois;

type
  TIntegBase = class (TPersistent)
    private
      FCurDef : integer;
      FCurTypeDef : TIntegTypeLigne;
      FFileName : string;
      FDefinition : TList;
      FNbContrainte : integer;
      FTitre : string;
      function ConvertTypeDef (C : Char) : TIntegTypeLigne;
      function TypeDefToChar(IL: TIntegTypeLigne): Char;
      function FindFirstDef(TypeLigne: TIntegTypeLigne): Pointer;
      function FindNextDef : Pointer;
    public
      constructor Create (stFileName : string);
      destructor Destroy; override;
      procedure Charge;
      procedure Enregistre;
      function Analyse (St : string) : TIntegDefinition;
      procedure EnregistreFmtDomix; // Voir pour un autre format (XML ???)
      function GetFileName : string;
      procedure AddDefinition (ADef : Pointer);
      procedure SetTitre (St : string);
      function GetTitre : string;
  end;

implementation

{ TIntegBase }

function TIntegBase.FindFirstDef(TypeLigne : TIntegTypeLigne) : Pointer;
var i : integer;
  P : Pointer;
begin
  P := nil;
  FCurTypeDef := TypeLigne;
  for i:=0 to FDefinition.Count - 1 do
  begin
    case TypeLigne of
      ilCompte :  if TObject(FDefinition.Items[i]) is TIntegDefCompte then
                  begin
                    FCurDef := i;
                    P := FDefinition.Items[i];
                    break;
                  end;
      ilEcriture :  if TObject(FDefinition.Items[i]) is TIntegDefEcriture then
                  begin
                    FCurDef := i;
                    P := FDefinition.Items[i];
                    break;
                  end;
      ilJournal :  if TObject(FDefinition.Items[i]) is TIntegDefJournal then
                  begin
                    FCurDef := i;
                    P := FDefinition.Items[i];
                    break;
                  end;
      ilMois :  if TObject(FDefinition.Items[i]) is TIntegDefMois then
                  begin
                    FCurDef := i;
                    P := FDefinition.Items[i];
                    break;
                  end;
    end;
  end;
  Result := P;
end;

function TIntegBase.FindNextDef: Pointer;
var i : integer;
  P : Pointer;
begin
  P := nil;
  for i:=FCurDef+1 to FDefinition.Count - 1 do
  begin
    case FCurTypeDef of
      ilCompte :  if TObject(FDefinition.Items[i]) is TIntegDefCompte then
                  begin
                    FCurDef := i;
                    P := FDefinition.Items[i];
                    break;
                  end;
      ilEcriture :  if TObject(FDefinition.Items[i]) is TIntegDefEcriture then
                  begin
                    FCurDef := i;
                    P := FDefinition.Items[i];
                    break;
                  end;
      ilJournal :  if TObject(FDefinition.Items[i]) is TIntegDefJournal then
                  begin
                    FCurDef := i;
                    P := FDefinition.Items[i];
                    break;
                  end;
      ilMois :  if TObject(FDefinition.Items[i]) is TIntegDefMois then
                  begin
                    FCurDef := i;
                    P := FDefinition.Items[i];
                    break;
                  end;
    end;
  end;
  Result := P;
end;

procedure TIntegBase.Charge;
var F : TextFile;
  St , stEntete : string;
  ADef : Pointer;
  TypeLigne : TIntegTypeLigne;
begin
  AssignFile(F, FFileName);
  Reset (F);
  while not Eof (F) do
  begin
    Readln (F,St);
    stEntete := Copy (St, 1, EX_LGCONTRAINTE);
    if stEntete = EX_NBCONTRAINTE then
    begin
      FNbContrainte := IntegExtractValeur(St);
    end else if stEntete = EX_TYPELIGNE then
    begin
      TypeLigne := ConvertTypeDef (St[EX_OFFSETVALEUR]);
      case TypeLigne of
        ilCompte : ADef := TIntegDefCompte.Create;
        ilEcriture : ADef := TIntegDefEcriture.Create;
        ilJournal : ADef := TIntegDefJournal.Create;
        ilMois : ADef := TIntegDefMois.Create;
        else ADef := TIntegDefinition.Create;
      end;
      if ADef<> nil then FDefinition.Add(ADef);
    end else if stEntete = EX_CONTRAINTE then TIntegDefinition(ADef).SetContrainte(Copy(St,EX_OFFSETVALEUR,255))
    else if stEntete = EX_DEBUTCOMPTE then  // Première ligne description compte
    begin
      ADef := FindFirstDef(ilCompte);
      if ADef <> nil then TIntegDefCompte(ADef).Charge(F);
      while (ADef <> nil) do
      begin
        ADef := FindNextDef;
        if ADef <> nil then TIntegDefCompte(ADef).Charge(F);
      end;
    end else if stEntete = EX_DEBUTECRITURE then  // Première ligne description écriture
    begin
      ADef := FindFirstDef(ilEcriture);
      if ADef <> nil then TIntegDefEcriture(ADef).Charge(F);
      while (ADef <> nil) do
      begin
        ADef := FindNextDef;
        if ADef <> nil then TIntegDefEcriture(ADef).Charge(F);
      end;
    end else if stEntete = EX_DEBUTJOURNAL then  // Première ligne description journal
    begin
      ADef := FindFirstDef(ilJournal);
      if ADef <> nil then TIntegDefJournal(ADef).Charge(F);
      while (ADef <> nil) do
      begin
        ADef := FindNextDef;
        if ADef <> nil then TIntegDefJournal(ADef).Charge(F);
      end;
    end else if stEntete = EX_DEBUTMOIS then  // Première ligne description journal
    begin
      ADef := FindFirstDef(ilMois);
      if ADef <> nil then TIntegDefMois(ADef).Charge(F);
      while (ADef <> nil) do
      begin
        ADef := FindNextDef;
        if ADef <> nil then TIntegDefMois(ADef).Charge(F);
      end;
    end;
  end;
  CloseFile(F);
end;

function TIntegBase.ConvertTypeDef(C: Char): TIntegTypeLigne;
begin
  case C of
    'c' : Result := ilCompte;
    'e' : Result := ilEcriture;
    'j' : Result := ilJournal;
    'm' : Result := ilMois;
    's' : Result := ilSautCompte;
    'r' : Result := ilRejet;
    else Result := ilNeutre;
  end;
end;

function TIntegBase.TypeDefToChar(IL : TIntegTypeLigne): Char;
begin
  case IL of
    ilCompte : Result := 'c';
    ilEcriture : Result := 'e';
    ilJournal : Result := 'j';
    ilMois : Result := 'm';
    ilSautCompte : Result := 's';
    ilRejet : Result := 'r';
    else Result := #0;
  end;
end;

constructor TIntegBase.Create(stFileName: string);
begin
  FFileName := stFileName;
  FDefinition := TList.Create;
  FNbContrainte := 0;
end;

destructor TIntegBase.Destroy;
begin
  FDefinition.Clear;
  FDefinition.Free;
  inherited;
end;

procedure TIntegBase.EnregistreFmtDomix;
begin
  // Enregistrement de la base au format domix
end;

function TIntegBase.Analyse(St: string) : TIntegDefinition;
var i : integer;
  ADef : TIntegDefinition;
  bTrouve : boolean;
begin
  ADef := nil; bTrouve := False;
  for i:=0 to FDefinition.Count - 1 do
  begin
    ADef := TIntegDefinition(FDefinition.Items[i]);
    if ADef.Analyse(St) then // Contrainte OK
    begin
      bTrouve := True;
      if ADef is TIntegDefCompte then  TIntegDefCompte(ADef).MajValeur (St)
      else if ADef is TIntegDefEcriture then  TIntegDefEcriture(ADef).MajValeur (St)
      else if ADef is TIntegDefJournal then  TIntegDefJournal(ADef).MajValeur (St)
      else if ADef is TIntegDefMois then  TIntegDefMois(ADef).MajValeur (St);
      break;
    end;
  end;
  if bTrouve then Result := ADef
  else Result := nil;
end;

procedure TIntegBase.Enregistre;
var F : TextFile;
  i : integer;
begin
  AssignFile (F, fFileName);
  Rewrite(F);
  Writeln(F,EX_TYPE +'  COMP');
  Writeln(F,EX_TITRE +'  '+FTitre);
  Writeln(F,EX_CODEACTION +'  RIEN');
  Writeln(F,EX_CODETABLE +'  dat');
  Writeln(F,EX_DECOUPE +'  0000');
  Writeln(F,EX_NBDECIMAL +'  0000');
  Writeln(F,EX_TYPERECUP +'  GL');
  Writeln(F,EX_NBCONTRAINTE +'  '+Format('%.04d',[FNbContrainte]));
  for i:=0 to FDefinition.Count - 1 do
  begin
    if TypeDefToChar(TIntegDefinition(FDefinition[i]).TypeLigne)<>#0 then
      Writeln (F, EX_TYPELIGNE+'  '+TypeDefToChar(TIntegDefinition(FDefinition[i]).TypeLigne));
    Writeln (F, EX_NOMATRICE+'  0000');
    Writeln (F, EX_CONTRAINTE+'  '+TIntegDefinition(FDefinition[i]).Contrainte);
  end;
  Writeln (F, EX_DEBUTCOMPTE);
  for i:=0 to FDefinition.Count - 1 do
    if TObject(FDefinition[i]) is TIntegDefCompte then TIntegDefCompte(FDefinition[i]).Enregistre(F);
  Writeln (F, EX_DEBUTMOIS);
  for i:=0 to FDefinition.Count - 1 do
    if TObject(FDefinition[i]) is TIntegDefMois then TIntegDefMois(FDefinition[i]).Enregistre(F);
  Writeln (F, EX_DEBUTJOURNAL);
  for i:=0 to FDefinition.Count - 1 do
    if TObject(FDefinition[i]) is TIntegDefJournal then TIntegDefJournal(FDefinition[i]).Enregistre(F);
  Writeln (F, EX_DEBUTECRITURE);
  for i:=0 to FDefinition.Count - 1 do
    if TObject(FDefinition[i]) is TIntegDefEcriture then TIntegDefEcriture(FDefinition[i]).Enregistre(F);
  CloseFile (F);
end;

function TIntegBase.GetFileName: string;
begin
  Result := FFileName;
end;

procedure TIntegBase.AddDefinition(ADef: Pointer);
begin
  if TObject(ADef) is TIntegDefEcriture then
    FDefinition.Add(TIntegDefECriture(ADef));
  Inc (FNbContrainte,1);
end;

function TIntegBase.GetTitre: string;
begin
  Result := FTitre;
end;

procedure TIntegBase.SetTitre(St: string);
begin
  FTitre := St;
end;

end.
