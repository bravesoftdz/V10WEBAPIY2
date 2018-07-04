{ Unité : Source TOF de la FICHE : CPAFFECTEBQPREVI
--------------------------------------------------------------------------------------
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 6.30.001.001   21/02/05    JP     Création de l'unité

--------------------------------------------------------------------------------------}
unit CPAFFECTEBQPREVI_TOF;

interface

uses
  StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, dbtables, FE_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  Vierge, Forms, SysUtils, ComCtrls, UTOF, uTob;

type
  TOF_CPAFFECTEBQPREVI = class (TOF)
    procedure OnUpdate               ; override;
    procedure OnArgument(S : string ); override;
  private
    Banque     : string;
    TobEcr     : TOB;
    PeutFermer : Boolean;
    FCanClose  : TCloseQueryEvent;

    procedure RadioClick(Sender : TObject);
    procedure ECanClose(Sender: TObject; var CanClose: Boolean);
  end ;

procedure CP_AffecteBqPrevi(Arg : string);


implementation

uses
  HMsgBox, HCtrls, HEnt1, ULibEcriture, Constantes, Commun;


{---------------------------------------------------------------------------------------}
procedure CP_AffecteBqPrevi(Arg : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche('CP', 'CPAFFECTEBQPREVI', '', '', Arg);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPAFFECTEBQPREVI.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  w : string;
  b : string;
begin
  inherited;
  b := Trim(GetControlText('EDBANQUE'));

  if (b = '') and TRadioButton(GetControl('RBREMPLIT')).Checked then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez renseigner un compte;W;O;O;O;', '', '');
    PeutFermer := False;
    Exit;
  end;

  {Boucle de mise à jour de la table écriture}
  for n := 0 to TobEcr.Detail.Count - 1 do begin
    {JP 21/07/06 : Modification de la fonction WhereEcritureTob}
    w := WhereEcritureTOB(tsGene, TobEcr.Detail[n], True, False, 'E');
    ExecuteSQL('UPDATE ECRITURE SET E_BANQUEPREVI = "' + b + '", E_TRESOSYNCHRO = "' + ets_BqPrevi + '" WHERE ' + w);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPAFFECTEBQPREVI.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;

  TRadioButton(GetControl('RBREMPLIT')).OnClick := RadioClick;
  TRadioButton(GetControl('RBVIDE'   )).OnClick := RadioClick;

  {Récupération de la banque prévisionnelle sélectionnée dans le Mul}
  Banque  := ReadTokenSt(S);

  {Si on a déjà renseigner la date}
  if Trim(Banque) = '' then begin
    TRadioButton(GetControl('RBREMPLIT')).Checked := True;
    SetFocusControl('EDBANQUE');
  end
  else
    TRadioButton(GetControl('RBVIDE')).Checked := True;

  THEdit(GetControl('EDBANQUE')).DataType := 'TTBANQUECP';
  {JP 29/01/07 : gestion du partage de BanqueCP : Uilisation d'une fonction générique
  THEdit(GetControl('EDBANQUE')).Plus := '(BQ_NATURECPTE = "BQE" OR BQ_NATURECPTE = "" OR BQ_NATURECPTE IS NULL) AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '"';}
  SetPlusBanqueCp(GetControl('EDBANQUE'));

  {Récupération de la sélection d'écritures à modifier}
  TobEcr := TFVierge(Ecran).LaTof.LaTOB;
  {Gestion du Canclose si la banque prévisionnelle n'est pas renseignée}
  FCanClose := TFVierge(Ecran).OnCloseQuery;
  TFVierge(Ecran).OnCloseQuery := ECanClose;
  PeutFermer := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPAFFECTEBQPREVI.RadioClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Ok : Boolean;
begin
  Ok := ((UpperCase(TComponent(Sender).Name) = 'RBREMPLIT') and TRadioButton(Sender).Checked);
  SetControlEnabled('EDBANQUE', Ok);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPAFFECTEBQPREVI.ECanClose(Sender: TObject; var CanClose: Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose := PeutFermer;
  {Je suis obligé de lancer FCanClose que si CanClose est à True, car dans Vierge.Pas
   on fait CanClose := LaTOF.Close sans Tester CanClose}
  if CanClose then
    FCanClose(Sender, CanClose);
  PeutFermer := True;
end;


initialization
  RegisterClasses([TOF_CPAFFECTEBQPREVI]);

end.
