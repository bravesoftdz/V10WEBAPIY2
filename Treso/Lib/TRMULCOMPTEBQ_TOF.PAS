{ Unité : Source TOF de la FICHE : TRMULCOMPTEBQ
--------------------------------------------------------------------------------------
    Version  |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
    0.91       23/09/03  JP   Création de l'unité
 1.50.001.001  26/04/04  JP   Contrôle avant suppression que le compte n'est pas utilisé par une écriture
 7.05.001.001  01/08/06  JP   Gestion de la Tréso Multi sociétés
 7.05.001.001  23/10/06  JP   Gestion des filtres Multi sociétés
 7.09.001.007  02/03/07  JP   FQ 19724 : Test sur l'utilisation des comptes dans la Tréso étendu à la compta
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TRMULCOMPTEBQ_TOF ;

interface

uses
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  db, mul, FE_Main,
  {$ELSE}
  eMul, MaineAGL, uTob, 
  {$ENDIF}
  {$IFDEF TRCONF}
  uLibConfidentialite,
  {$ELSE}
  UTOF,
  {$ENDIF TRCONF}
  Forms, HCtrls, HEnt1, HMsgBox, HTB97;

type
  {$IFDEF TRCONF}
  TOF_TRMULCOMPTEBQ = class (TOFCONF)
  {$ELSE}
  TOF_TRMULCOMPTEBQ = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument (S : String ) ; override ;
  private
    procedure ListeDblClick  (Sender : TObject);
    procedure NouveauClick   (Sender : TObject);
    procedure SupprimerClick (Sender : TObject);
    {$IFDEF TRESO}
    procedure NoDossierChange(Sender : TObject);
    {$ENDIF TRESO}
  end ;

procedure TRLanceFiche_BanqueCP(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  BANQUECP_TOM, UtilPgi, Commun, Constantes, Ent1, ParamSoc;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_BanqueCP(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCOMPTEBQ.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  {$IFDEF TRESO}
  {08/08/06 : gestion du multi sociétés}
  SetControlVisible('BQ_NODOSSIER' , IsTresoMultiSoc);
  SetControlVisible('TBQ_NODOSSIER', IsTresoMultiSoc);

  {Gestion des filtres multi sociétés sur banquecp et dossier}
  THEdit(GetControl('BQ_CODE')).Plus := FiltreBanqueCp(THEdit(GetControl('BQ_CODE')).DataType, '', '');
  THValComboBox(GetControl('BQ_NODOSSIER')).Plus := '';
  {JP 28/11/06 : Ajout du XX_WHERE POUR filtrer sur les bons dossiers}
  SetControlText('XX_WHERE', 'BQ_NODOSSIER = "' + V_PGI.NoDossier + '"');
  if IsTresoMultiSoc then begin
    THValComboBox(GetControl('BQ_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
    SetControlText('XX_WHERE', 'BQ_NODOSSIER ' + FiltreNodossier);
  end
  else if EstTablePartagee('BANQUECP') then
    SetControlText('XX_WHERE', 'BQ_NODOSSIER = "' + V_PGI.NoDossier + '"');
  {Filtre sur le dossier courants}
  THValComboBox(GetControl('BQ_NODOSSIER')).OnChange := NoDossierChange;
  {$ELSE}
  SetControlVisible('BQ_NODOSSIER', False);
  SetControlVisible('TBQ_NODOSSIER', False);
  {09/10/06 : Gestion du filtre sur les dossiers en compta}
  SetControlText('BQ_NODOSSIER', V_PGI.NoDossier);
  THEdit(GetControl('BQ_CODE')).Plus := 'BQ_NODOSSIER = "' + V_PGI.NoDossier + '"';
  SetControlText('XX_WHERE', 'BQ_NODOSSIER = "' + V_PGI.NoDossier + '"');
  {$ENDIF TRESO}
  Ecran.HelpContext := 500006;
  TFMul(Ecran).FListe.OnDblClick := ListeDblClick;
  TFMul(Ecran).Binsert.OnClick   := NouveauClick;
  TToolbarButton97(GetControl('BDELETE')).OnClick := SupprimerClick;
end;

{$IFDEF TRESO}
{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCOMPTEBQ.NoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  THEdit(GetControl('BQ_CODE')).Plus := FiltreBanqueCp(THEdit(GetControl('BQ_CODE')).DataType, '', GetControlText('BQ_NODOSSIER'));
  SetControlText('BQ_CODE', '');
end;
{$ENDIF TRESO}

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCOMPTEBQ.ListeDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  FicheBanqueCP(GetField('BQ_GENERAL'), taModif, 1, GetField('BQ_NODOSSIER'));
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCOMPTEBQ.NouveauClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRESO}
  {02/08/06 : il est assez difficile de mettre en place ce test sur toutes les bases en même temps.
              le test sera effectué dans la TOM de BANQUECP, lorsque le dossier aura été choisi}
  {$ELSE}
  if EstMultiSoc then begin
    if not ExisteSql('SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE IN ("BQE", "DIV") AND ' +
                     'G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + V_PGI.NoDossier +
                     '") AND G_GENERAL IN (SELECT CLS_GENERAL FROM CLIENSSOC)') then begin
      HShowMessage('0;' + Ecran.Caption + ';Tous les comptes généraux de type bancaire ou courant sont utilisés.'#13 +
                   'Pour pouvoir créer un nouveau compte, il faut d''abord créer un nouveau compte général.;I;O;O;O;', '', '');
      Exit;
    end;
  end
  else
  if not ExisteSql('SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE = "BQE" AND ' +
                   'G_GENERAL NOT IN (SELECT BQ_GENERAL FROM BANQUECP )') then begin
    HShowMessage('0;' + Ecran.Caption + ';Tous les comptes généraux de type bancaire sont utilisés.'#13 +
                 'Pour pouvoir créer un nouveau compte, il faut d''abord créer un nouveau compte général.;I;O;O;O;', '', '');
    Exit;
  end;
  {$ENDIF TRESO}
  FicheBanqueCP('', taCreat, 1, GetField('BQ_NODOSSIER'));

  TFMul(Ecran).BCherche.Click;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULCOMPTEBQ.SupprimerClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  SQL  : string;
  n    : Integer;
  NoOk : Boolean;
begin
  {JP 02/03/04 : Reprise de la fonction et mise en place de la sélection}
  if TFMul(Ecran).FListe.NbSelected = 0 then begin
    {Aucun élément sélectionné, on sort}
    HShowMessage('1;' + Ecran.Caption + ';Veuillez sélectionner un enregistrement.;W;O;O;O;', '', '');
    Exit;
  end;

  NoOk := False;
  if HShowMessage('0;' + Ecran.Caption + ';Êtes-vous sûr de vouloir supprimer les éléments sélectionnés ?;I;YNC;N;N;', '', '') = mrYes then begin
    {On boucle sur la sélection et on les supprime un à un}
    for n := 0 to TFMul(Ecran).FListe.NbSelected - 1 do begin
      TFMul(Ecran).FListe.GotoLeBookmark(n);
      {JP 26/04/04 : Ajout d'un test sur les écritures de trésorerie, pour éviter les suppression intempestives
       JP 02/03/07 : FQ 19724 : extension du contrôle à la compta et surtout gestion du Multi-dossier Tréso}
      SQL := 'DELETE FROM BANQUECP WHERE BQ_CODE = "' + GetField('BQ_CODE') + '"';
      if EstComptaTreso then begin
        if not ExisteSQL('SELECT TE_GENERAL FROM ' + GetTableDossier(GetParamSocSecur('SO_TRBASETRESO', ''), 'TRECRITURE') +
                         ' WHERE TE_GENERAL = "' + GetField('BQ_CODE') + '"') then begin
          ExecuteSQL(SQL);
        end
        else
          NoOk := True;
      end
      else
        ExecuteSQL(SQL);
    end;
    TFMul(Ecran).BCherche.Click;
  end;
  {JP 26/04/04 : si certains comptes ne peuvent être supprimés, on affiche un message d'avertissement}
  if NoOk then
    HShowMessage('2;' + Ecran.Caption + ';Certains comptes étaient utilisés pour des écritures de trésorerie : ils ne peuvent être supprimés;I;O;O;O;', '', '');
end;

initialization
  RegisterClasses ( [ TOF_TRMULCOMPTEBQ ] ) ;

end.
