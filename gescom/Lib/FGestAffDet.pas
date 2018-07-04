unit FGestAffDet;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Factutil,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  Doc_Parser,DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,UserChg,
{$ENDIF}
  StdCtrls, CheckLst, ExtCtrls, HPanel, HTB97,HCtrls,M3FP,ENt1,HEnt1;

procedure ParamDetailOuvrage (var AffichageCourant : integer);

type
  TFParamAffDet = class(TForm)
    DockBottom: TDock97;
    Valide97: TToolbar97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    Panel: THPanel;
    ListeChoix: TCheckListBox;
    procedure FormCreate(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
  private
    procedure remplitListeChoix(ListeChoix: TCheckListBox);
    function EncodeResultat(ListeChoix: TCheckListBox): Integer;
    procedure ActiveChoix(Choix: string);
    procedure PositionneContexte;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

function RecupModeAffLibelles (TypePresent : integer) : string;

{$IFDEF AGL530E}
// DEFINITION DES TYPES DE PRESENTATION DES DETAILS D'OUVRAGES
Const DOU_AUCUN    : integer = 0 ;
      DOU_CODE     : integer = 1 ;
      DOU_LIBELLE  : integer = 2 ;
      DOU_QTE      : integer = 4 ;
      DOU_UNITE    : integer = 8 ;
      DOU_PU       : integer = 16 ;
      DOU_MONTANT  : integer = 32 ;
      DOU_TOUS     : integer = 63 ;
// --------
{$ENDIF}

implementation
var Affichage : Integer;

{$R *.DFM}
function RecupModeAffLibelles (TypePresent : integer) : string;
var LibellesModeAff : string;
BEGIN
if TypePresent = DOU_AUCUN then
   BEGIN
   LibellesModeAff := 'Aucun affichage';
   result := LibellesModeAff;
   exit;
   END;
LibellesModeAff := 'Affichage '+string(chr(10));
if (TypePresent and DOU_CODE) = DOU_CODE then LibellesModeAff := LibellesModeAff + 'du code' + string(chr(10));
if (TypePresent and DOU_LIBELLE) = DOU_LIBELLE  then LibellesModeAff := LibellesModeAff + 'du Libellé' + string(chr(10));
if (typepresent and DOU_QTE) = DOU_QTE  then LibellesModeAff := LibellesModeAff + 'de la quantité' + string(chr(10));
if (typepresent and DOU_UNITE) = DOU_UNITE  then LibellesModeAff := LibellesModeAff + 'de l''unité' + string(chr(10));
if (typepresent and DOU_PU) = DOU_PU then LibellesModeAff := LibellesModeAff + 'du prix unitaire' + string(chr(10));
if (typepresent and DOU_MONTANT) = DOU_MONTANT then LibellesModeAff := LibellesModeAff + 'du Montant Ligne' + string(chr(10));
result := LibellesModeAff;
END;

procedure ParamDetailOuvrage (var AffichageCourant : integer);
var
  FParamAffDet: TFParamAffDet;
begin
affichage := AffichageCourant;
FParamAffDet := TFParamAffDet.create (application);
FParamAffDet.ShowModal ;
AffichageCourant := Affichage;
end;

procedure TFParamAffDet.PositionneContexte;
var Indice : Integer;
    nonactif,critere : string;
begin

if not (ctxAffaire in V_PGI.PgiContexte) then NonActif := 'Montant;Prix unitaire'
                                         else NonActif:='';

repeat
Critere:=Trim(ReadTokenSt(NonActif)) ;
if critere <> '' then
   begin
   for Indice := 0 to ListeChoix.Items.count -1 do
       begin
       if ListeChoix.Items[Indice] = Critere then
          begin
          ListeChoix.items.Delete(Indice);
          break;
          end;
       end;
   end;
until Critere='';
end;

procedure TFParamAffDet.FormCreate(Sender: TObject);
begin
PositionneContexte;
remplitListeChoix (ListeChoix);
end;

Function TFParamAffDet.EncodeResultat (ListeChoix: TCheckListBox): Integer;
var Indice : Integer;
begin
Result := 0 ;
for Indice := 0 to ListeChoix.Items.count -1 do
    begin
    if ListeChoix.checked [Indice]= true then
       begin
       if ListeChoix.Items [Indice] = 'Code' then result := (result + DOU_CODE)
       else if ListeChoix.Items [Indice] = 'Désignation' then result := (result + DOU_LIBELLE)
       else if ListeChoix.Items [Indice] = 'Quantité' then result := (result + DOU_QTE)
       else if ListeChoix.Items [Indice] = 'Unité' then result := (result + DOU_UNITE)
       else if ListeChoix.Items [Indice] = 'Prix unitaire' then result := (result + DOU_PU)
       else if ListeChoix.Items [Indice] = 'Montant' then result := (result + DOU_MONTANT);
       end;
    end;
end;

procedure TFParamAffDet.ActiveChoix (Choix : string);
var Indice  : Integer;
begin
for Indice := 0 to ListeChoix.items.count -1 do
    begin
    if ListeChoix.Items[Indice] = Choix then ListeChoix.checked [indice] := true;
    end;
end;

procedure TFParamAffDet.remplitListeChoix (ListeChoix: TCheckListBox);
var Indice : Integer;
begin
for Indice := 0 to ListeChoix.Items.Count -1 do ListeChoix.checked [0] := false;

if (affichage and DOU_CODE) = DOU_CODE then ActiveChoix ('Code');
if (affichage and DOU_LIBELLE) = DOU_LIBELLE then ActiveChoix ('Désignation');
if (affichage and DOU_QTE) = DOU_QTE then ActiveChoix ('Quantité');
if (affichage and DOU_UNITE) = DOU_UNITE then ActiveChoix ('Unité');
if (affichage and DOU_PU) = DOU_PU then ActiveChoix ('Prix unitaire');
if (affichage and DOU_MONTANT) = DOU_MONTANT then ActiveChoix ('Montant');
end;

procedure TFParamAffDet.BValiderClick(Sender: TObject);
begin
Affichage := EncodeResultat (ListeChoix);
close;
end;

procedure TFParamAffDet.BAbandonClick(Sender: TObject);
begin
close;
end;

end.
