{***********UNITE*************************************************
Auteur  ...... : Paie PGI
Créé le ...... : 20/06/2003
Modifié le ... :   /  /
Description .. : Publication des méthodes du script Agl pour utilisation delphi
Mots clefs ... : PAIE
*****************************************************************}
{
PT1 20/06/2003 SB V_42 Modification de procedure en function des gestions associées
PT2 29/09/2005 PH V_60 FQ 11932 Appel fonction publication script gestion LIENOLE
PT3 29/12/2006 GG V_80 FQ 13454 Fonction de l'AGL permettant de ne pas obligé la
    saisie des zéros inutiles à gauche du matricule Salarié utilisé dans
    ABSPAYEES_MUL
PT4 19/09/2007 FC V_80 FQ 14593 Fonction qui permet de tester si on a une habilitation ou non
}
unit AGLInitPaie;

interface
uses M3VM, classes;

implementation

uses HCtrls, HMsgBox, M3FP, SysUtils,
{$IFNDEF EAGLCLIENT}
  Fe_Main, DB, Fiche, Mul, FichList,
{$ELSE}
  UtileAGL, MaineAgl, eFiche, eMul, eFichList,
{$ENDIF}
  GestionAssocie, RubriqueProfil, Hent1, PartEtab,
  Ventil, P5Def, PGOperation, Forms,

  EntPaie, PgOutils2, //PT3

  UtilPgi; // PT2

procedure AGLParamVentil(parms: array of variant; nb: integer);
begin
  ParamVentil(string(parms[0]), string(parms[1]), string(parms[2]), tActionFiche(Parms[3]), Boolean(Parms[4]));
end;

function AglTComboGetValue(parms: array of variant; nb: integer): variant;
var
  C: THValComboBox;
  F: TForm;
begin
  F := TForm(LongInt(parms[0]));
  C := THValComboBox(F.FindComponent(string(parms[1])));
  Result := C.Value;
end;

function AglTComboGetText(parms: array of variant; nb: integer): variant;
var
  C: THValComboBox;
  F: TForm;
begin
  F := TForm(LongInt(parms[0]));
  C := THValComboBox(F.FindComponent(string(parms[1])));
  Result := C.Text;
end;

function AGLLanceGestionAssocie(parms: array of variant; nb: integer): variant;
begin
  result := LanceGestionAssocie(string(parms[0]), string(parms[1]), string(parms[2]), string(parms[3])); //PT1
end;

procedure AGLLancePartEtab(parms: array of variant; nb: integer);
begin
  LancePartEtab(string(parms[0]));
end;

function AGLProfilPaieDetail(parms: array of variant; nb: integer): variant;
begin
  result := ProfilPaieDetail(string(parms[0]), string(parms[1]), string(parms[2])); //PT1
end;

function AGLMaxDateDebutExer(parms: array of variant; nb: integer): variant;
begin
  result := MaxDebutExerSocial;
end;

procedure AGLSupprimeOperation(parms: array of variant; nb: integer);
var F: TForm;
  Query, Liste: TComponent;
begin
  F := TForm(Longint(Parms[0]));
  if F = nil then exit;
  Liste := F.FindComponent('FListe');
  if Liste = nil then exit;
  Query := F.FindComponent('Q');
  if (Query = nil) then exit;
  AppelSuppOpe(Liste, Query, F);
end;

//PT3
function AGLFormateMatricule(parms: array of variant; nb: integer): variant;
var
  St : String;
begin
  St := string(parms[0]);
  result := St;
  if St <> '' then
     if (isnumeric (St) and (VH_PAIE.PgTypeNumSal = 'NUM')) then
        result := ColleZeroDevant (StrToInt(St), 10);
end;
//Fin PT3

procedure AGLPaieConceptPlanPaie(parms: array of variant; nb: integer);
var F: TForm;
begin
  F := TForm(Longint(Parms[0]));
  if F = nil then exit;
  PaieConceptPlanPaie (F);
end;
procedure AGLPaieConceptTabMinPaie(parms: array of variant; nb: integer);
var F: TForm;
begin
  F := TForm(Longint(Parms[0]));
  if F = nil then exit;
  PaieConceptTabMinPaie (F);
end;

//DEB PT4
function AGLPaieHabilitationExiste(parms: array of variant; nb: integer): variant;
begin
  if (Assigned(MonHabilitation)) and (MonHabilitation.LeSQL<>'') then
    result := true
  else
    result := false;
end;
//FIN PT4
//////////////////////////////////////////////////////////////////////////////

procedure initM3Paie();
begin
  RegisterAglProc('ParamVentil', FALSE, 5, AGLParamVentil);
  RegisterAglFunc('TComboGetValue', True, 1, AglTComboGetValue);
  RegisterAglFunc('TComboGetText', True, 1, AglTComboGetText);

  RegisterAglProc('LancePartEtab', FALSE, 4, AGLLancePartEtab);
  RegisterAglFunc('GestionAssocie', FALSE, 4, AGLLanceGestionAssocie); // PT1
  RegisterAglFunc('ProfilPaieDetail', FALSE, 3, AGLProfilPaieDetail); // PT1
  RegisterAglFunc('MaxDateDebutExer', TRUE, 0, AGLMaxDateDebutExer);
  RegisterAglProc('SupprimeOperation', TRUE, 0, AGLSupprimeOperation);
  RegisterAglFunc('FormateMatricule', FALSE, 1, AGLFormateMatricule);   //PT3
  RegisterAglProc('ConceptPlanPaie', TRUE, 0, AGLPaieConceptPlanPaie);
  RegisterAglProc('ConceptTabMinPaie', TRUE, 0, AGLPaieConceptTabMinPaie);
  RegisterAglFunc('HabilitationExiste', TRUE, 0, AGLPaieHabilitationExiste); //PT4

//  AglInitLienOle; // PT2
end;

initialization
  initM3Paie();
finalization
end.

