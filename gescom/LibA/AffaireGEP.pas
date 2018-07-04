{***********UNITE************************************************************************************
Auteur  ...... : A.Buchet
Créé le ...... : 19/09/2007
Modifié le ... : 19/09/2007
Description .. : Fonction GEP de Création d'une affaire avec ajout de lignes d'affaire
Description .. : G_CreateAffaire  (codeaffaire, codetiers: string; TOBAffaire, TobLigneAffaire,TobRetour: TOB)
Mots clefs ... : TOBAffaire champs de la table AFFAIRE si créaton ou mise à jour AFFAIRE
Mots clefs ... : codeaffaire et codetiers si uniquement ajout lignes d'affaire
Mots clefs ... : TobLigneAffaire contient dans chaque fille les champs de la table LIGNE
Resultat ..... : tobretour contient les messages dans champs AFFAIRE - NOMBRELIGNES - LIGNEAFFAIRE - PIECEAFFAIRE
******************************************************************************************************}

unit AffaireGEP;

interface

uses
  Classes,
  UTob,
  Forms,
  HEnt1,
  HCtrls,
  SysUtils,
  M3FP,
  HMsgBox,
  Controls,
  Entgc,
  Ed_tools;

function CreeAffaireDepuisTob (codeaffaire, codetiers: string; TOBAffaire, TobLigneAffaire, TobRetour: TOB): boolean;
function AGLCreeAffaire (parms: array of variant; nb: integer): variant;

implementation
uses gerepiece,
  UtilEAffaire,
  wRapport,
  affaireutil,
  affaireduplic,
  saisutil,
  utilpgi;

const
  texteMessage: array [1..6] of string = (
    {1}'L''affaire a été créée',
    {2}'Aucune affaire correspondante',
    {3}'Nombre de Lignes d''affaire créées :',
    {4}'La ligne d''affaire n''a pas été créée pour le code :',
    {5}'La pièce affaire a été correctement enregistrée',
    {6}'La pièce affaire n''a pas été correctement enregistrée'
    );

function CreerLesLignesAffaireDepuisTob (TOBAffaire, TobLigneAffaire, TobMessage: TOB): Boolean;
begin
end;

function CreeAffaireDepuisTob (codeaffaire, codetiers: string; TOBAffaire, TobLigneAffaire, TobRetour: TOB): boolean;
var
  CodeAff, Aff0, Aff1, Aff2, Aff3, Aff4, Tiers, stRetour: string;
  Tobaff, TobMessage: tob;

begin
  Result := false;
  Tobaff := TOBAffaire;
  TobMessage := tob.create ('LesMessages', nil, -1);
  try
    if Assigned (Tobaff) then
    begin
      stRetour := CreateAffaireFromTob (Tobaff);
      if stRetour = '' then
      begin
        CodeAff := Tobaff.getstring ('AFF_AFFAIRE');
        Result := CodeAff <> '';
        TobMessage.AddChampSupValeur ('AFFAIRE', CodeAff + ' : ' + textemessage [1] );
      end
      else
        TobMessage.AddChampSupValeur ('AFFAIRE', stRetour);
    end
    else
    begin
      CodeAff := CodeAffaire;
      if TeststCleAffaire (CodeAff, Aff0, Aff1, Aff2, Aff3, Aff4, Tiers, false, false, false, true) = 1 then
      begin
        Tobaff := tob.create ('AFFAIRE', nil, -1);
        Tobaff.SelectDB ('"' + CodeAff + '"', nil);
        Result := True;
      end
      else
        TobMessage.AddChampSupValeur ('AFFAIRE', CodeAff + ' : ' + textemessage [2] );
    end;
    if Result and (Assigned (TobLigneAffaire)) and (TobLigneAffaire.Detail.Count > 0) then
      Result := CreerLesLignesAffaireDepuisTob (Tobaff, TobLigneAffaire, TobMessage);
    if assigned (TobRetour) then
      TobRetour.Dupliquer (TobMessage, false, true);

  finally
    if not Assigned (Tobaffaire) then
      freeandnil (Tobaff);
    freeandnil (TobMessage);
  end;
end;

function AGLCreeAffaire (parms: array of variant; nb: integer): variant;
var
  TobAffaire, TobLigneaffaire, TobRetour: Tob;
  StCodeAffaire, StCodetiers: string;
begin
  StCodeAffaire := Parms [0] ;
  StCodetiers := Parms [1] ;
  TobAffaire := nil;
  TobLigneaffaire := nil;
  TobRetour  := nil;
  if isnumeric(Parms [2]) then
    TobAffaire := Tob (LongInt (Parms [2] ));
  if isnumeric(Parms [3]) then
    TobLigneaffaire := Tob (LongInt (Parms [3] ));
  if isnumeric(Parms [4]) then
    TobRetour := Tob (LongInt (Parms [4] ));

  CreeAffaireDepuisTob (StCodeAffaire, StCodetiers, TobAffaire, TobLigneaffaire, TobRetour);
end;

end.
