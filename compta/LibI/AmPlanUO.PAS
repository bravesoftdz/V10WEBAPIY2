{***********UNITE*************************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 18/05/2004
Modifié le ... :   /  /
Description .. : Classe des gestion des unités d'oeuvre
Mots clefs ... :
Suite......... : fq 17569 - mbo - 27/04/2006 - modif date de début amort
                 + réinitialisation du plan d'unités d'oeuvre si modif durée amort
                 ou date deb amort  

*****************************************************************}
unit AmPlanUO;

interface

uses SysUtils
     , Hent1
     , uTOB
     , HCtrls
     , ImOuPlan
     , ImEnt
     ;

type
  TPlanUO = class
    private
      fPlan   : TOB;
      fCodeImmo : string;
      fDuree : integer;
      fTotalUO : double;
      fDateDebut : TDateTime;
      function GetAnterieurUO : double;
      procedure InitParDefaut ( DateDebut : TDateTime; Duree : integer; Quantite : double );
      procedure MajDuree(const Value: integer);
      procedure MajTotalUO(const Value: double);
      procedure MajDateDebut(const Value: TDateTime);
      function  SommeUO : double;
    public
      constructor Create;
      destructor Destroy; override;
      procedure Charge (CodeImmo : string); overload;
      procedure Charge (CodeImmo: string; Duree : integer; TotalUO : double; DateDebut : TDateTime); overload;
      procedure SetCodeImmo ( CodeImmo : string );
      procedure Enregistre;
      procedure MajPlan ( T : TOB );
      function  Valide : boolean;
      procedure Raz;
    published
      property LePlan : TOB read fPlan;
      property AnterieurUO : double read GetAnterieurUO;
      property Duree : integer read fDuree write MajDuree;
      property TotalUO : double read fTotalUO write MajTotalUO;
      property DateDebut : TDateTime read fDateDebut write MajDateDebut;
  end;

implementation

{ TPlanUO }

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 18/05/2004
Modifié le ... :   /  /
Description .. : Chargement de la liste des unités d'oeuvre pour une
Suite ........ : immobilisation
Mots clefs ... :
*****************************************************************}
procedure TPlanUO.Charge(CodeImmo: string; Duree : integer; TotalUO : double; DateDebut : TDateTime);
begin
  fCodeImmo := CodeImmo;
  fDuree := Duree;
  fTotalUO := TotalUO;
  fDateDebut := DateDebut;
  fPlan.LoadDetailDBFromSQL('IMMOUO','SELECT * FROM IMMOUO WHERE IUO_IMMO="'+CodeImmo+'"',False);
end;

procedure TPlanUO.Charge(CodeImmo: string);
var TImmo : TOB;
begin
  fCodeImmo := CodeImmo;
  TImmo := TOB.Create ('IMMO',nil,-1);
  try
    if TImmo.SelectDB('"'+CodeImmo+'"',nil,True) then
    begin
      fPlan.LoadDetailDBFromSQL('IMMOUO','SELECT * FROM IMMOUO WHERE IUO_IMMO="'+CodeImmo+'"',False);
      fDuree := TImmo.GetValue('I_DUREEECO');
      fTotalUO := TImmo.GetValue('I_UNITEOEUVRE');
      //fq 17569 fDateDebut := TImmo.GetValue('I_DATEPIECEA');
      fDateDebut := TImmo.GetValue('I_DATEDEBECO');
    end;
  finally
    TImmo.Free;
  end;
end;

constructor TPlanUO.Create;
begin
  fPlan := TOB.Create ('',nil,-1);
end;

destructor TPlanUO.Destroy;
begin
  fPlan.Free;
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 18/05/2004
Modifié le ... :   /  /
Description .. : Enregistrement dans la base du plan d'unités d'oeuvre
Mots clefs ... :
*****************************************************************}
procedure TPlanUO.Enregistre;
begin
  BeginTrans;
  try
    ExecuteSQL('DELETE FROM IMMOUO WHERE IUO_IMMO="'+fCodeImmo+'"');
    fPlan.SetAllModifie (True);
    fPlan.InsertDB(nil);
    CommitTrans;
  except
    RollBack;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 18/05/2004
Modifié le ... :   /  /
Description .. : Initialise le plan d'unités d'oeuvre en répartissant au prorata
Suite ........ : de la quantité saisie sur chaque période
Mots clefs ... :
*****************************************************************}
function TPlanUO.GetAnterieurUO: double;
var i : integer;
begin
  Result := 0;
  if fPlan = nil then exit;
  if fPlan.Detail.Count > 0 then
  begin
    for i:= 0 to fPlan.Detail.Count - 1 do
    begin
      if (fPlan.Detail[i].GetValue('IUO_DATE') < VHImmo^.EnCours.Deb) then
        Result := Result + fPlan.Detail[i].GetValue('IUO_UNITEOEUVRE')
      else break;
    end;
  end;
end;

procedure TPlanUO.InitParDefaut(DateDebut: TDateTime; Duree: integer;
  Quantite: double);
var DateCalcul, DateDebutEx, DateFinEx : TDateTime;
    PremMois,PremAnnee,NbMois : Word;
    NbMoisTotal : integer;
    TUO : TOB;
    CumulUO, NombreUO : double;
begin
  fPlan.ClearDetail;
  DateCalcul := DateDebut;
  DateDebutEx := iDate1900;
  DateFinEx := iDate1900;
  NbMoisTotal := 0;
  NombreUO := 0;
  CumulUO := 0;
  TUO := nil;
  { Calcul du nombre d'unités d'oeuvre pour chaque période }
  while (NbMoisTotal < Duree ) do
  begin
    GetDatesExercice(DateCalcul,DateDebutEx,DateFinEx);
    NOMBREMOIS(DateCalcul,DateFinEx,PremMois,PremAnnee,NbMois );
    NbMoisTotal := NbMoisTotal + NbMois;
    TUO := TOB.Create('IMMOUO',fPlan,-1);
    TUO.PutValue('IUO_IMMO',fCodeImmo);
    TUO.PutValue('IUO_DATE',DateFinEx);
    NombreUO := Arrondi((Quantite*NbMois)/Duree, V_PGI.OkDecV);
    CumulUO := Arrondi( CumulUO + NombreUO , V_PGI.OkDecV );
    TUO.PutValue('IUO_UNITEOEUVRE', NombreUO);
    DateCalcul := DateFinEx+1;
  end;
  { à la fin , on réajuste pour que le total soit égal à Quantite }
  if TUO <> nil then TUO.PutValue('IUO_UNITEOEUVRE',Arrondi((Quantite-(CumulUO-NombreUO)), V_PGI.OkDecV));
end;

procedure TPlanUO.MajDateDebut(const Value: TDateTime);
var DateCalcul, DateDebutEx, DateFinEx : TDateTime;
    i : integer;
begin
  fDateDebut := Value;
  // mbo fq 17569 if fPlan.Detail.Count = 0 then InitParDefaut( fDateDebut, fDuree, fTotalUO) else
  begin
    InitParDefaut( fDateDebut, fDuree, fTotalUO);
    DateCalcul := fDateDebut;
    DateDebutEx := iDate1900;
    DateFinEx := iDate1900;
    { Calcul du nombre d'unités d'oeuvre pour chaque période }
    for i:=0 to fPlan.Detail.Count - 1 do
    begin
      GetDatesExercice(DateCalcul,DateDebutEx,DateFinEx);
      fPlan.Detail[i].PutValue('IUO_DATE',DateFinEx);
      DateCalcul := DateFinEx+1;
    end;
  end;
end;

procedure TPlanUO.MajDuree(const Value: integer);
var DateCalcul, DateDebutEx, DateFinEx : TDateTime;
    PremMois,PremAnnee,NbMois : Word;
    NbMoisTotal : integer;
    nLigne,nLigneAvant : integer;
    TUO : TOB;
begin
 // mbo fq 17569  if fPlan.Detail.Count = 0 then InitParDefaut( fDateDebut, Value, fTotalUO) else
   begin
    InitParDefaut( fDateDebut, Value, fTotalUO);

    nLigne := 0; nLigneAvant := fPlan.Detail.Count;
    DateCalcul := fDateDebut;
    DateDebutEx := iDate1900;
    DateFinEx := iDate1900;
    NbMoisTotal := 0;
    while (NbMoisTotal < Value ) do
    begin
      GetDatesExercice(DateCalcul,DateDebutEx,DateFinEx);
      NOMBREMOIS(DateCalcul,DateFinEx,PremMois,PremAnnee,NbMois );
      NbMoisTotal := NbMoisTotal + NbMois;
      Inc (nLigne);
      if (nLigne > nLigneAvant) then
      begin
        TUO := TOB.Create ('IMMOUO',fPlan,-1);
        TUO.PutValue('IUO_IMMO',fCodeImmo);
        TUO.PutValue('IUO_DATE',DateFinEx);
        TUO.PutValue('IUO_UNITEOEUVRE', 0);
      end;
      DateCalcul := DateFinEx+1;
    end;
    if (nLigne < nLigneAvant) then
    begin
      while (fPlan.Detail.Count <> nLigne) do
        fPlan.Detail[fPlan.Detail.Count-1].Free;
    end;
  end;
  fDuree := Value;
end;

procedure TPlanUO.MajPlan( T : TOB );
var i : integer;
    Ti : TOB;
begin
  fPlan.ClearDetail;
  for i := 0 to T.Detail.Count - 1 do
  begin
    Ti := TOB.Create ('IMMOUO',fPlan,-1);
    Ti.PutValue('IUO_IMMO',fCodeImmo );
    Ti.PutValue('IUO_DATE',T.Detail[i].GetValue('IUO_DATE') );
    Ti.PutValue('IUO_UNITEOEUVRE',T.Detail[i].GetValue('IUO_UNITEOEUVRE') );
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. : Mise à jour du nombre total d"unités d'oeuvre
Mots clefs ... :
*****************************************************************}
procedure TPlanUO.MajTotalUO(const Value: double);
begin
  if (fPlan.Detail.Count = 0) or (SommeUO = 0 ) then InitParDefaut( fDateDebut, fDuree, Value);
  fTotalUO := Value;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. : Mise à jour du code immo
Mots clefs ... :
*****************************************************************}
procedure TPlanUO.Raz;
begin
  fCodeImmo := '';
  fDuree :=0;
  fTotalUO :=0;
  fDateDebut := iDate1900;
  fPlan.ClearDetail;
end;

procedure TPlanUO.SetCodeImmo(CodeImmo: string);
var i : integer;
begin
  fCodeImmo := CodeImmo;
  for i:=0 to fPlan.Detail.Count - 1 do
    fPlan.Detail[i].PutValue('IUO_IMMO',fCodeImmo);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. : Calcul de la somme totale des unités d'oeuvre saisies
Mots clefs ... :
*****************************************************************}
function TPlanUO.SommeUO: double;
var i : integer;
begin
  Result := 0;
  for i := 0 to fPlan.Detail.Count - 1 do
    Result := Result + fPlan.Detail[i].GetValue('IUO_UNITEOEUVRE');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/12/2004
Modifié le ... :   /  /
Description .. : Contrôle de la validité du plan variable
Suite ........ : - le nombre d'unité d'oeuvre saisi doit être égal au nombre
Suite ........ : total d'unités d'oeuvre du plan
Mots clefs ... :
*****************************************************************}
function TPlanUO.Valide: boolean;
begin
  Result := (Arrondi((FTotalUO - SommeUO), 2) = 0)
end;

end.
