{-------------------------------------------------------------------------------------
   Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
               06/02/02  BT   Création de l'unité
 1.50.000.000  22/04/04  JP   Nouvelle gestion des natures (OnUpdate, ModifierLibelle)
 6.30.001.011  23/05/05  JP   FQ 10254 : modification de la limitation de la période de traitement
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TofFicheSuivi ;

interface

uses StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  {$IFDEF TRCONF}
  uLibConfidentialite,
  {$ELSE}
  UTOF,
  {$ENDIF TRCONF}
  SysUtils, Menus, HCtrls, HEnt1, UTOB, ParamSoc, UProcGen;

type
  {$IFDEF TRCONF}
  TOF_FICHESUIVI = class (TOFCONF)
  {$ELSE}
  TOF_FICHESUIVI = class (TOF)
  {$ENDIF TRCONF}
    procedure OnNew                 ; override;
    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
  protected
    Where: string;
    PopupMenu : TPopUpMenu;
    TobSuivi  : TOB;
    BanqueCpt : THValComboBox;
    Regroupement : THValComboBox;
    DeviseAff : string;
    Nature    : string;

    procedure ModifierLibelle;
    procedure LanceDetail(Par: Char);

    procedure DeviseOnChange         (Sender : TObject);
    procedure CheckClick             (Sender : TObject);
    procedure OnPopupDetailFluxClick (Sender : TObject);
    procedure OnPopupDetailSoldeClick(Sender : TObject);
  end;

procedure TRLanceFiche_FicheSuivi(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

Implementation

uses
   Stat, Commun, TofDetailSuivi, ExtCtrls, HMsgBox, Constantes;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_FicheSuivi(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FICHESUIVI.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  TobSuivi := TOB.Create('0', nil, -1);
  TFStat(Ecran).LaTob := TobSuivi;
  TFStat(Ecran).ModeAlimentation := maTOB;
  Ecran.HelpContext := 150;

  THValComboBox(GetControl('TE_DEVISE')).OnChange := DeviseOnChange;
  BanqueCpt := THValComboBox(GetControl('BANQUECPT'));

  Regroupement := THValComboBox(GetControl('REGROUPEMENT'));
  SetControlVisible('REGROUPEMENT', IsTresoMultiSoc);
  SetControlVisible('TREGROUPEMENT', IsTresoMultiSoc);
  Regroupement.Plus := 'DOS_NODOSSIER ' + FiltreNodossier;

  PopupMenu := TPopUpMenu(GetControl('POPUPMENU'));

  PopupMenu.Items[0].OnClick := OnPopupDetailFluxClick;
  PopupMenu.Items[1].OnClick := OnPopupDetailSoldeClick;

  ADDMenuPop(PopupMenu, '', '');
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_FICHESUIVI.OnClose ;
{---------------------------------------------------------------------------------------}
begin
  TobSuivi.Free;
  Inherited ;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_FICHESUIVI.DeviseOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  DeviseAff := GetControlText('TE_DEVISE');

  AssignDrapeau(TImage(GetControl('IDEV')), DeviseAff);
  if DeviseAff = '' then
    THValComboBox(GetControl('TE_DEVISE')).ItemIndex := 0;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FICHESUIVI.CheckClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if V_PGI.AutoSearch then
    TFStat(Ecran).ChercheClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FICHESUIVI.OnPopupDetailFluxClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  LanceDetail('F');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FICHESUIVI.OnPopupDetailSoldeClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  LanceDetail('S');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_FICHESUIVI.LanceDetail(Par: Char);
{---------------------------------------------------------------------------------------}
var
  CurTob : TOB;
  S      : string;
  n      : Double;
begin
  {15/01/04 : mise en commun des critères entre la fiche de suivi et détail de suivi}
  S := GetControlText('DATEVALEUR') + ';';
  {JP 04/04/03 : FQ 10018 : +1, sinon il manquait une colonne}
  n := StrToDate(GetControlText('DATEVALEUR_')) - StrToDate(GetControlText('DATEVALEUR')) + 1;
  {Dans le détail du suivi, on ne gère qu'un trimestre au maximum (185 jours)}
  if n > 185 then begin
    {FQ 10254 : on limite maintenant à un semestre, soit 185 jours au lieu de 93}
    case HShowMessage('1;Fiche de suivi;La fourchette de dates est supérieure à un semestre.'#13 +
                                       'Souhaitez-vous n''afficher que les 185 premiers jours ?;Q;YNC;Y;Y;Y;', '', '') of
      mrNo,
      mrCancel : begin
                   SetFocusControl('DATEVALEUR_');
                   Exit;
                 end;
      mrYes    : S := S + '185;';
    end;
  end
  else
    S := S + IntToStr(Round(n)) + ';';

  {Paramètre : Choix d'affichage ; Societe ; Devise ; Compte ; Nature}
  S := S + Par + ';' + GetParamSocSecur('SO_SOCIETE', '001'){GetControlText('TE_SOCIETE')} + ';' + DeviseAff + ';';
  CurTob := TFStat(Ecran).TV.CurrentTOB;
  {Puis compte général de la banque}
  if CurTob <> Nil then
	S :=  S + CurTob.GetValue('TE_GENERAL');
  S := S + ';';
  {Enfin : Si par solde -> Tous types de flux !}
  {11/12/2003 : Voilà quelque chose qui m'avait échappé et qui me surprend !!
   if Par = 'F' then} S := S + Nature;
  S := TRLanceFiche_DetailSuivi('TR','TRDETAILSUIVI', '', '', S);
  if S <> '' then
    TFStat(Ecran).ChercheClick; {Rafraichissement avec la Tob modifiée}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 11/02/2002
Modifié le ... :   /  /
Description .. : Constitue le Where à partir des critères automatiques de
Suite ........ : FStat et des autres critères, pour charger la Tob.
Mots clefs ... : WHERE;CRITERES
*****************************************************************}
procedure TOF_FICHESUIVI.OnUpdate ;

  {------------------------------------------------------}
  procedure AddWhere(S: String);
  {------------------------------------------------------}
  begin
    if Where = '' then Where := 'WHERE ' + S
                  else Where := Where + ' AND ' + S;
  end;
const
  {22/04/04 : Ajout de TE_NATURE dans la requête}
  CLAUSEGROUP  = ' GROUP BY BQ_LIBELLE, TE_GENERAL, TE_CODEFLUX, TFT_LIBELLE, SENS, TE_NATURE, TE_DEVISE';
  {La vue contient en plus des champs ci-dessous les champs TE_NATURE et TE_DATEVALEUR}
  CLAUSESELEC  = 'SELECT BQ_LIBELLE, TE_GENERAL, TE_CODEFLUX, SENS, SUM(TE_MONTANT) AS TE_MONTANT, ' +
                 'SUM(TE_MONTANT) AS TE_MONTANT, TFT_LIBELLE, TE_DEVISE, TE_NATURE FROM TRIMPFICHESUIVI ';

var
  Q    : TQuery;
  S    : string;
  sNat : string;
begin

  Inherited ;
  Where := TFStat(Ecran).stSQL; // 'Where (Societe, Devise)'

  {Filtre sur les devises}
  S := GetControlText('TE_DEVISE');
  if S <> '' then
    AddWhere('TE_DEVISE = "' + S + '"');

  {Filtre sur les banques}
  S := GetControlText('BANQUECPT');
  if S <> '' then
    AddWhere('BQ_BANQUE = "' + S + '"');

  S := GetControlText('Regroupement');
  {Pour le moment les groupes de banques ne sont pas gérés}
  if S <> '' then
    AddWhere('TE_NODOSSIER = "' + S + '"');

  {Filtre sur les dates de valeur}
  if (Trim(GetControlText('DATEVALEUR')) <> '/  /') and (Trim(GetControlText('DATEVALEUR_')) <> '/  /') then begin
    AddWhere('TE_DATEVALEUR >= "' + USDateTime(StrToDate(GetControlText('DATEVALEUR'))) + '"');
    AddWhere('TE_DATEVALEUR <= "' + USDateTime(StrToDate(GetControlText('DATEVALEUR_'))) + '"');
  end
  else begin
    HShowMessage('0;Fiche de suivi;Veuillez saisir les dates de valeur !;W;O;O;O', '', '');
    Exit;
  end;

  {Filtre sur la nature}
  sNat := GetControlText('TE_NATURE');
  if (Pos('<<' , sNat) = 0) and (Trim(sNat) <> '') then begin {<<Tous>> ou vide}
    Nature := 'TE_NATURE IN (' + GetClauseIn(sNat) + ')';

    AddWhere(Nature);
    Nature := 'AND ' + Nature; {Pour compléter Where en Détail suivi}
  end
  else
    Nature := '';

  Q := OpenSQL(CLAUSESELEC + Where + CLAUSEGROUP, True);

  TobSuivi.LoadDetailDB('1', '', '', Q, False, True);
  Ferme(Q);
  {Met la nature dans le libellé du flux}
  ModifierLibelle;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_FICHESUIVI.ModifierLibelle;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
begin
  for n := 0 to TobSuivi.Detail.Count - 1 do begin
    T := TobSuivi.Detail[n];
    if T.GetString('TE_NATURE') = na_Realise then
      T.SetString('TFT_LIBELLE', T.GetString('TFT_LIBELLE') + ' (Réal.)')
    else if T.GetString('TE_NATURE') = na_Prevision then
      T.SetString('TFT_LIBELLE', T.GetString('TFT_LIBELLE') + ' (Prév.)')
    else if T.GetString('TE_NATURE') = na_Simulation then
      T.SetString('TFT_LIBELLE', T.GetString('TFT_LIBELLE') + ' (Simu.)');
  end;
end;

{Double clic sur une ligne
{---------------------------------------------------------------------------------------}
procedure TOF_FICHESUIVI.OnNew ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  LanceDetail('F'); // Par flux par défaut
end ;

initialization
  registerclasses ( [ TOF_FICHESUIVI ] ) ;
end.

