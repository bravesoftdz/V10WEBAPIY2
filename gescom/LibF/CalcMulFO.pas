{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Fonctions @ pour les MUL
Mots clefs ... : FO
*****************************************************************}
unit CalcMulFO;

interface

uses
  SysUtils,
  {$IFNDEF EAGLCLIENT}
  db,
  {$ENDIF}
  UTob, Hctrls, Hent1;

function FOProcCalcMul(Func, Params, WhereSQL: string; TT: TDataset; Total: boolean): string;

implementation

uses
  MFOSOLDECLIENT_TOF, FOUtil;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 05/12/2003
Modifié le ... : 05/12/2003
Description .. : Affichage du solde du client depuis un MUL
Suite ........ :
Suite ........ : @FOSOLDECLIENT(GP;X)
Mots clefs ... : CB;MUL
*****************************************************************}
function FormuleSoldeClient(Params: string; TT: TDataset): string;
var
  s1, s2: string;
  Montant: double;
  sCode: string;
  Fld: TField;
begin
  Result := '';
  s1 := ReadTokenPipe(Params, ';'); // Préfixe du code Tiers
  s2 := ReadTokenPipe(Params, ';'); // pour tous les établissements
  if s1 = '' then s1 := 'T';
  s1 := s1 + '_TIERS';
  Fld := TT.FindField(s1);
  if Fld <> nil then
  begin
    sCode := Fld.AsString;
    if (sCode <> '') and (sCode <> FODonneClientDefaut) then
    begin
      Montant := FOCalculSoldeClient(sCode, (s2 = 'X'));
      Result := StrFMontant(Montant, 12, V_PGI.OkDecV, '', False);
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/05/2003
Modifié le ... : 26/05/2003
Description .. : Fonctions @ pour les MUL
Mots clefs ... : FO
*****************************************************************}

function FOProcCalcMul(Func, Params, WhereSQL: string; TT: TDataset; Total: boolean): string;
begin
  Result := '';
  if Func = 'FOSOLDECLIENT' then
  begin
    Result := FormuleSoldeClient(Params, TT);
  end;
end;

end.
