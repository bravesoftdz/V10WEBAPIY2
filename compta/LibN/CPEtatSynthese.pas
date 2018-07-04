unit CPEtatSynthese;

interface

uses
  { Delphi }
  classes,
  sysutils,
  comctrls,
  { AGL }
  hctrls,
  {$IFDEF EAGLCLIENT}
  utileAGL,
  {$ELSE}
  EdtREtat,
  {$ENDIF}
  uTob,
  formule,
  HEnt1,
  { Compta }
  calcole;

type
  TCPModeCalcul = (mcFormule, mcRubrique );

  TCPEtatSynthese = class
    private
      fLaTob : TOB;
      fTobEnCours : TOB;
      fMaquette : string;
      fAvecDetail : boolean;
      fNbColonne : integer;
      fDefColonne : array [1..8] of string;           // Définition de la colonne
      fTypeColonne : array [1..8] of TCPModeCalcul;   // Mode de calcul de la colonne
      fTitreColonne : array [1..8] of string;         // Titre des colonnes
      fColEnCours : integer;
      fAvecDecimales : boolean;
      fAvecMontant0 : boolean;
      function GetVariable(St: string): variant;
      function DecaleLibelle(St: string): string;
      procedure TraiteLigneRubrique ( Tl : TOB );
      procedure TraiteLigneFormule ( Tl : TOB );
      function AjouteLigne ( St : string ; bDetail : boolean ) : TOB;
      function EvalueFormule ( stFormule : string ) : double;
      function FaitParametreCritere : string;
      function FormateValeur ( const Value : double ) : string;
    public
      constructor Create;
      destructor Destroy; override;
      procedure Calcule;
      procedure Imprime;
      procedure ParametreColonne(const NumColonne : integer;const Value: string; const ModeCalcul: TCPModeCalcul; const stTitre : string);
      property Maquette : string read fMaquette write fMaquette;
      property Detail : boolean read fAvecDetail write fAvecDetail;
      property Decimales : boolean read fAvecDecimales write fAvecDecimales;
      property Montant0 : boolean read fAvecMontant0 write fAvecMontant0;
  end;

implementation

{ TCPEtatSynthese }

procedure TCPEtatSynthese.Calcule;
var
  LaMaquette : TStringList;
  i : integer;
  Tl : TOB;
begin
  { Calcul du nombre de colonnes }
  for i := 1 to 8 do
  begin
    if fDefColonne[i]<>'' then Inc (fNbColonne)
    else if ((i mod 2) <> 0) then break;
  end;
  { Chargement de la maquette }
  LaMaquette := TStringList.Create;
  LaMaquette.LoadFromFile(fMaquette);
  LaMaquette.Delete(0);
  LaMaquette.Delete(0);
  for i := 0 to LaMaquette.Count - 1 do
  begin
    Tl := AjouteLigne (  LaMaquette[i] , False);

    if (Tl.GetValue('TYPE')='R') then TraiteLigneRubrique ( Tl )
    else if ((Tl.GetValue('TYPE')='F1') or (Tl.GetValue('TYPE')='F2')) then
      TraiteLigneFormule ( Tl );
  end;
  LaMaquette.Free;
end;

constructor TCPEtatSynthese.Create;

begin
  fAvecDetail := False;
  fLaTob := TOB.Create ( '', nil, -1);
end;

destructor TCPEtatSynthese.Destroy;
begin
  fLaTOB.Free;
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 21/04/2005
Modifié le ... :   /  /
Description .. : Récupére la valeur d'un élément d'une formule
Mots clefs ... : 
*****************************************************************}
function TCPEtatSynthese.GetVariable(St: string): variant;
var
  T : TOB;
  i : integer;
begin
  if Copy(St, 1, 3) = 'COL' then
  begin
    i := ValeurI(St);
    if fTobEnCours <> nil then Result := Valeur(fTobEnCours.GetValue('COLONNE' + IntToStr(i)))
    else result := 0;
  end else
  begin
    T := fLaTob.FindFirst(['NOM'], [St], FALSE);
    if T <> nil then
      Result := Valeur(T.GetValue('COLONNE'+IntToSTr(fColEnCours)))
    else result := 0;
  end;
end;

procedure TCPEtatSynthese.Imprime;
var
  T, LaTobDetail : TOB;
  i : integer;
begin
  if fAvecDetail then
  begin
    LaTobDetail := fLaTob;
    fLaTob := TOB.Create ('',nil,-1);
    for i:= 0 to LaTobDetail.Detail.Count - 1 do
    begin
      if LaTobDetail.Detail[i].GetValue('DETAIL')='-' then
      begin
        T := TOB.Create ('', fLaTob, - 1);
        T.Dupliquer(LaTOBDetail.Detail[i],False,True);
      end;
    end;
    while (LaTobDetail.Detail.Count > 0) do
    begin
      LaTobDetail.Detail[0].PutValue('DETAIL','X');
      LaTobDetail.Detail[0].ChangeParent ( fLaTob, - 1);
    end;
    LaTobDetail.Free;
  end;
  LanceEtatTOB('E','ESY','ESC',fLaTob,True,False,False,nil,'','',False,0, FaitParametreCritere);
end;

function TCPEtatSynthese.DecaleLibelle(St: string): string;
var
  i: integer;
begin
  if (Length(St) > 0) then
  begin
    i := 1;
    while ( (i <= length(St)) and (St[i] = '.') ) do    // Correction SBO 06/05/2004
    begin
      St[i] := ' ';
      Inc(i, 1);
    end;
  end;
  Result := St;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 21/04/2005
Modifié le ... :   /  /
Description .. : Ajout d'une ligne dans la TOB (état) à partir d'une ligne de la
Suite ........ : maquette
Mots clefs ... :
*****************************************************************}
function TCPEtatSynthese.AjouteLigne(St : string ; bDetail : boolean ): TOB;
var
  Tl : TOB;
  stType, stAlign, stLibelle : string;
begin
  Tl := TOB.Create('Une TOB', fLaTOB, -1);
  try
    if bDetail then
    begin
      Tl.AddChampSupValeur('TYPE','D');
      Tl.AddChampSupValeur('ALIGN', '');
      Tl.AddChampSupValeur('STYLE', '');
      Tl.AddChampSupValeur('LIBELLE', ' - '+Copy(St, 1, Pos(':', St) - 1));
      Tl.AddChampSupValeur('NOM', '');
      Tl.AddChampSupValeur('BASE', '');
      Tl.AddChampSupValeur('FORMULE', '');
      Tl.AddChampSupValeur('DETAIL','X');  // Si détail uniquement alors DETAIL = 'X'
    end
    else
    begin
      stType := Trim(UpperCase(ReadTokenPipe(St, '|')));
      stAlign := Trim(UpperCase(ReadTokenPipe(St, '|')));
      if stType='F' then
      begin
        if stAlign = 'D' then stType := 'F1'
        else stType := 'F2';
      end;
      Tl.AddChampSupValeur('TYPE', stType);
      Tl.AddChampSupValeur('ALIGN', stAlign);
      Tl.AddChampSupValeur('STYLE', Trim(UpperCase(ReadTokenPipe(St, '|'))));
      stLibelle := DecaleLibelle(Trim(ReadTokenPipe(St, '|')));
      Tl.AddChampSupValeur('LIBELLE', stLibelle);
      Tl.AddChampSupValeur('NOM', Trim(UpperCase(ReadTokenPipe(St, '|'))));
      Tl.AddChampSupValeur('BASE', Trim(UpperCase(ReadTokenPipe(St, '|'))));
      Tl.AddChampSupValeur('FORMULE', Trim(UpperCase(ReadTokenPipe(St, '|'))));
      Tl.AddChampSupValeur('DETAIL','-');  // Si détail uniquement alors DETAIL = 'X'
    end;
    Tl.AddChampSupValeur('COLONNE1','');
    Tl.AddChampSupValeur('COLONNE2','');
    Tl.AddChampSupValeur('COLONNE3','');
    Tl.AddChampSupValeur('COLONNE4','');
  except
    Tl.Free;
  end;
  Result := Tl;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 21/04/2005
Modifié le ... :   /  /
Description .. : Traitement d'une ligne de la TOB de type formule
Mots clefs ... :
*****************************************************************}
procedure TCPEtatSynthese.TraiteLigneFormule(Tl: TOB);
var
  i : integer;
begin
  for i := 1 to fNbColonne do
  begin
    fColEnCours := i;
    Tl.PutValue('COLONNE'+IntToStr(i), FormateValeur(EvalueFormule (Tl.GetValue('FORMULE'))));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 21/04/2005
Modifié le ... :   /  /
Description .. : Traitement d'une ligne de la TOB de type rubrique
Mots clefs ... :
*****************************************************************}
procedure TCPEtatSynthese.TraiteLigneRubrique(Tl: TOB);
var
  vRubrique : double;
  LesCptes : TStringList;
  i,k : integer;
  Decimales : integer;
begin
  if fAvecDecimales then Decimales := 2
  else Decimales := 0;
  fTobEnCours := Tl;
  if fAvecDetail then  LesCptes := TStringList.Create
  else LesCptes := nil;
  for i := 1 to fNbColonne do
  begin
    if ( fTypeColonne[(2*i)-1] = mcRubrique ) then
      vRubrique := Get_CumulPCL('RUBRIQUE',Tl.GetValue('FORMULE'), 'N', '','', fDefColonne[(2*i)-1], '','-', '-', LesCptes,'')
    else vRubrique := EvalueFormule ( fDefColonne[(2*i)-1] );
    Tl.PutValue('COLONNE'+IntToStr(i),FormateValeur(double(vRubrique)));
    if fAvecDetail then
    begin
      { Ajout des lignes de détail }
      for k := 0 to LesCptes.Count - 1 do
        fTobEnCours := AjouteLigne ( LesCptes[k] , True );
    end;
  end;
  if fAvecDetail then LesCptes.Free;
end;

procedure TCPEtatSynthese.ParametreColonne(const NumColonne : integer;const Value: string; const ModeCalcul: TCPModeCalcul; const stTitre : string);
begin
  fDefColonne[(2*NumColonne)-1] := Value;
  fTypeColonne[(2*NumColonne)-1] := ModeCalcul;
  fTitreColonne[(2*NumColonne)-1] := stTitre;
end;

function TCPEtatSynthese.EvalueFormule( stFormule : string): double;
var
  stMask, stF, stFF : string;
  X : double;
  i : integer;
begin
  stMask := '#,##0.00';
  Delete(StFormule, 1, 1);
  StFormule := '{"' + stMask + '"' + StFormule + '}';
  StF := GFormule(stFormule, GetVariable, nil, 1);
  Result := Valeur (StF);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 21/04/2005
Modifié le ... :   /  /    
Description .. : Constitution des critères à passer à l'édition
Mots clefs ... : 
*****************************************************************}
function TCPEtatSynthese.FaitParametreCritere: string;
var stColonne : string;
    i : integer;
begin
  stColonne := 'NBCOL='+IntToStr(fNbColonne)+'`';
  for i:=1 to fNbColonne do
    stColonne := stColonne + 'TCOL'+IntToStr(i)+'='+fTitreColonne[(2*i)-1]+'`';
  Result :=
    'TITRE=Compte de résultat`'
    +stColonne;
end;

function TCPEtatSynthese.FormateValeur(const Value : double): string;
var
  nDecimales : integer;
begin
  if fAvecDecimales then nDecimales := V_PGI.OkDecV
  else nDecimales := 0;
  if (( Value=0 ) and (not fAvecMontant0)) then Result := ''
  else Result := StrFMontant(Value,13,nDecimales,'',True);
end;

end.
