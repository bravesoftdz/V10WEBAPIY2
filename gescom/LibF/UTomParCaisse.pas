{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 18/06/2001
Modifié le ... : 30/06/2004
Description .. : TOM de la table PARCAISSE : paramètrage des caisses
Mots clefs ... : FO
*****************************************************************}
unit UTomParCaisse;

interface
uses
  Classes, forms, Controls, sysutils, HCtrls, HEnt1, UTOM, UTob, HTB97,
  {$IFDEF EAGLCLIENT}
  Maineagl, eFichList, eMul,
  {$ELSE}
  Fe_Main, db, dbtables, FichList, Mul,
  {$ENDIF}
  LookUp, ParamSoc, LicUtil, M3FP, HMsgBox, UtilGC;

type
  TOM_ParCaisse = class(TOM)
  private
    CLavierOrg: string;
    AffSurPrt: boolean; // l'afficheur utilise le même port que l'imprimante
    TiroirSurPrt: boolean; // le tiroir utilise le même port que l'imprimante
    procedure MetEnErreur(NomChamp: string; NoMsg: integer);
    function ChampObligatoire(NomChamp: string; NoMsg: integer): boolean;
    function GetControlEnabled(NomChamp: string): Boolean;
    procedure ActiveChamp(NomChamp: string; Active: Boolean; ValDefOui: string = ''; ValDefNon: string = '');
    procedure ActiveBtnTest;
    {$IFDEF TOXCLIENT}
    function VerifHeure(NomChamp: string): Boolean;
    procedure MajHeureTox;
    {$ENDIF}
    function VerifLookUp(NomChamp: string; NoMsg: Integer): Boolean;
    function ControleIntegrite: Boolean;
    procedure ReCopieCaisse;
    procedure ReCopieClavier(CaisseDest, CaisseOrg: string);
    function DeleteClavier(Caisse: string): Boolean;
    procedure OnChangeEtablissement;
    procedure OnChangeCliSaisie;
    procedure OnChangeCliReprise;
    procedure OnChangePrtType;
    procedure OnChangePrtPort;
    procedure OnChangePrtParam(NomChamp: string);
    procedure OnChangeTpeType;
    procedure OnChangeTiroirType;
    procedure OnChangeTiroirPort;
    procedure OnChangeTiroirParam;
    procedure OnChangePrtFiscalType;
    procedure OnChangeSignPadType;
    procedure OnChangeVendeur;
    procedure OnChangeModeDetaxe;
    procedure OnChangeRemSaisie;
    procedure OnChangeRemPourMax;
    procedure OnChangeRemPourMax2;
    procedure OnChangeRemPourMax3;
    procedure OnChangeDemSaisie;
    procedure OnChangeCtrlCaisse;
    procedure OnChangeGereFondCaisse;
    procedure OnChangeGereRemiseBnq;
    procedure OnChangeAfficheur;
    procedure OnChangeAffType;
    procedure OnChangeAffPort;
    procedure OnChangeAffParam;
    procedure OnChangeClavierEcran;
    procedure OnChangeImpFmtTic;
    procedure OnChangeToxAppel1;
    procedure DecryptePwdRem;
    procedure ModeleNonParametrable;
    procedure ClickBtnParamTck;
    procedure ParamTicket(Sender: TObject);
    {$IFNDEF FOS5}                        
    procedure ParticularitesBO;
    {$ENDIF}
    {$IFDEF GESCOM}
    procedure ParticularitesGESCOM;
    {$ENDIF}
    procedure FormatDateConnexion;
    procedure InitParamPeripherique;
    procedure MajParamPeripherique;
  public
    procedure OnArgument(stArgument: string); override;
    procedure OnClose; override ;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override ;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnCancelRecord; override;
  end;

implementation
uses
  EntGC, Ent1, GCMzsUtil,
  {$IFDEF TOXCLIENT}
  FOToxUtil,
  {$ENDIF}
  MC_Admin, FODefi, FOUtil;

const
  // Libellés des messages
  TexteMessage: array[1..21] of string = (
    {1}'L''intitulé de la caisse doit être renseigné !',
    {2}'Le code client par défaut doit être renseigné !',
    {3}'Le code de l''établissement de rattachement doit être renseigné !',
    {4}'Le code du dépôt de rattachement doit être renseigné',
    {5}'Le mode de paiement du rendu monnaie doit être renseigné.',
    {6}'Le mode de paiement des écarts de change doit être renseigné.',
    {7}'L''opération de caisse des écarts de caisse doit être renseignée.',
    {8}'L''opération de caisse du fonds de caisse doit être renseignée.',
    {9}'La devise du mode de paiement doit être des ',
    {10}'L''heure d''appel est incorrecte.',
    {11}'Le % de remise est incorrect.',
    {12}'Le mot de passe ne correspond pas à la confirmation saisie.',
    {13}'L''opération de caisse de remise en banque doit être renseignée.',
    {14}'Le code client par défaut est incorrect.',
    {15}'Le code vendeur par défaut est incorrect.',
    {16}'L''opération de caisse des écarts de caisse est incorrecte.',
    {17}'L''opération de caisse du fond de caisse est incorrecte.',
    {18}'L''opération de caisse de remise en banque est incorrecte.',
    {19}'Le nombre de jours pour annuler un ticket est incorrect.',
    {20}'Il existe déjà une caisse avec le même intitulé.',
    {21}'' // JTR - eQualité 11295. A ne pas utiliser
    );

  ///////////////////////////////////////////////////////////////////////////////////////
  //  MetEnErreur : met un champ en erreur et affiche le message associé
  ///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.MetEnErreur(NomChamp: string; NoMsg: integer);
begin
  SetFocusControl(NomChamp);
  LastError := NoMsg;
  LastErrorMsg := TexteMessage[LastError];
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ChampObligatoire : vérifie si un champ est renseigné
///////////////////////////////////////////////////////////////////////////////////////

function TOM_ParCaisse.ChampObligatoire(NomChamp: string; NoMsg: integer): boolean;
begin
  Result := (GetField(NomChamp) <> '');
  if Result = False then MetEnErreur(NomChamp, NoMsg);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  GetControlEnabled : indique si un contrôle est accessible
///////////////////////////////////////////////////////////////////////////////////////

function TOM_ParCaisse.GetControlEnabled(NomChamp: string): Boolean;
var Ctrl: TControl;
begin
  Result := False;
  Ctrl := GetControl(NomChamp);
  if Ctrl <> nil then Result := Ctrl.Enabled;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ActiveChamp : active ou désactive un champ et son libellé
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.ActiveChamp(NomChamp: string; Active: Boolean; ValDefOui: string = ''; ValDefNon: string = '');
begin
  SetControlEnabled(NomChamp, Active);
  SetControlEnabled('T' + NomChamp, Active);
  if Active then
  begin
    if (ValDefOui <> '') and (GetField(NomChamp) = '') then SetField(NomChamp, ValDefOui);
  end else
  begin
    if (ValDefNon = '') then SetField(NomChamp, FOChampValeurVide(NomChamp))
    else SetField(NomChamp, ValDefNon);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ActiveBtnTest : Active les boutons de test des périphériques
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.ActiveBtnTest;
begin
  // Active le bouton de test du TPE
  ActiveChamp('BTESTTPE', (GetField('GPK_TPETYPE') <> ''), '');
  // Active le bouton de test de l'imprimante caisse
  ActiveChamp('BTESTPRT', (GetField('GPK_PRTTYPE') <> ''), '');
  // Active le bouton de test du tiroir caisse
  ActiveChamp('BTESTTIROIR', (GetField('GPK_TIROIRTYPE') <> ''), '');
  // Active le bouton de test de l'afficheur externe
  ActiveChamp('BTESTAFF', (GetField('GPK_AFFTYPE') <> ''), '');
  // Active le bouton de test de l'imprimante fiscale
  ActiveChamp('BTESTPRTFISC', (GetField('GPK_PRTFISCTYPE') <> ''), '');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  VerifHeure : vérifie une heure
///////////////////////////////////////////////////////////////////////////////////////
{$IFDEF TOXCLIENT}

function TOM_ParCaisse.VerifHeure(NomChamp: string): Boolean;
var Heure, hh, mm, ss: string;
begin
  Result := True;
  Heure := GetField(NomChamp);
  if (Trim(Heure) = '') or (Heure = '  :  :  ') or (Heure = '  :  :') then
  begin
    SetField(NomChamp, '');
    Exit;
  end;
  hh := ReadTokenPipe(Heure, ':');
  if (hh = '') or (not IsNumeric(hh)) or (StrToInt(hh) < 0) or (StrToInt(hh) > 23) then Result := False;
  mm := ReadTokenPipe(Heure, ':');
  if (mm = '') or (not IsNumeric(mm)) or (StrToInt(mm) < 0) or (StrToInt(mm) > 59) then Result := False;
  ss := ReadTokenPipe(Heure, ':');
  if (ss = '') or (not IsNumeric(ss)) or (StrToInt(ss) < 0) or (StrToInt(ss) > 59) then Result := False;
  if not Result then MetEnErreur(NomChamp, 10);
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 31/12/2002
Modifié le ... : 31/12/2002
Description .. : Enregistrement de la configuration du ToxServeur si des
Suite ........ : sites Tox exiestent et si on modifie la caisse de référence
Suite ........ : sur le FO
Mots clefs ... :
*****************************************************************}
{$IFDEF TOXCLIENT}

procedure TOM_ParCaisse.MajHeureTox;
begin
  {$IFDEF FOS5}
  if (FOVerifSiteTox) and (GetField('GPK_CAISSE') = Trim(GetParamSoc('SO_GCFOCAISREFTOX'))) then
  begin
    FOToxServeurMajHeure(GetField('GPK_TOXAPPEL1'), GetField('GPK_TOXAPPEL2'));
  end;
  {$ENDIF}
end;
{$ENDIF}

///////////////////////////////////////////////////////////////////////////////////////
//  VerifLookUp : vérifie l'existence d'une valeur saisie dans un lookup
///////////////////////////////////////////////////////////////////////////////////////

function TOM_ParCaisse.VerifLookUp(NomChamp: string; NoMsg: Integer): Boolean;
var CC: TControl;
begin
  Result := True;
  CC := GetControl(NomChamp);
  if CC <> nil then
  begin
    Result := LookUpValueExist(CC);
    if Result = False then MetEnErreur(NomChamp, NoMsg);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ControleIntegrite : vérifie l'intégrité des données
///////////////////////////////////////////////////////////////////////////////////////

function TOM_ParCaisse.ControleIntegrite: Boolean;
var Pour1, Pour2, Pour3: Integer;
  Pwd1, Pwd2, sSql: string;
begin
  Result := False;
  // L'intitulé de la caisse doit être renseigné et être unique
  if not ChampObligatoire('GPK_LIBELLE', 1) then Exit;
  sSql := 'SELECT GPK_CAISSE FROM PARCAISSE WHERE GPK_LIBELLE="' + GetField('GPK_LIBELLE') + '"'
    + ' AND GPK_CAISSE<>"' + GetField('GPK_CAISSE') + '"';
  if ExisteSQL(sSql) then
  begin
    MetEnErreur('GPK_LIBELLE', 20);
    Exit;
  end;
  // Le code client par défaut doit être renseigné
  if (not ChampObligatoire('GPK_TIERS', 2)) or (not VerifLookUp('GPK_TIERS', 14)) then Exit;
  // Le code de l'établissement de rattachement doit être renseigné
  if not ChampObligatoire('GPK_ETABLISSEMENT', 3) then Exit;
  // Le code du dépôt de rattachement doit être renseigné
  if not ChampObligatoire('GPK_DEPOT', 4) then Exit;
  // le vendeur doit être correct
  if (GetField('GPK_VENDREPRISE') = 'CAI') or (GetField('GPK_VENDDEFAUT') <> '') then
  begin
    if not VerifLookUp('GPK_VENDDEFAUT', 15) then Exit;
  end;
  // Le mode de paiement du rendu monnaie doit être renseigné
  if not ChampObligatoire('GPK_MDPRENDU', 5) then Exit;
  // Le mode de paiement des écarts de change doit être renseigné
  if not ChampObligatoire('GPK_MDPECART', 6) then Exit;
  // L'opération de caisse des écarts de caisse doit être renseigné
  if GetField('GPK_CTRLCAISSE') = 'X' then
  begin
    if (not ChampObligatoire('GPK_CTRLECART', 7)) or (not VerifLookUp('GPK_CTRLECART', 16)) then Exit;
  end;
  // L'opération de caisse du fond de caisse doit être renseigné
  if GetField('GPK_GEREFONDCAISSE') = 'X' then
  begin
    if (not ChampObligatoire('GPK_FDCAISSE', 8)) or (not VerifLookUp('GPK_FDCAISSE', 17)) then Exit;
  end;
  // L'opération de caisse des remises en banque doit être renseigné
  if GetField('GPK_GEREREMISEBNQ') = 'X' then
  begin
    if (not ChampObligatoire('GPK_OPREMISEBNQ', 13)) or (not VerifLookUp('GPK_OPREMISEBNQ', 18)) then Exit;
  end;
  {$IFDEF TOXCLIENT}
  // Contrôle les heures des appels
  if not VerifHeure('GPK_TOXAPPEL1') then Exit;
  if not VerifHeure('GPK_TOXAPPEL2') then Exit;
  {$ENDIF}
  // Contrôle des % maximun de remise
  Pour1 := GetField('GPK_REMPOURMAX');
  Pour2 := GetField('GPK_REMPOURMAX2');
  if (Pour2 <> 0) and (Pour2 <= Pour1) then
  begin
    MetEnErreur('GPK_REMPOURMAX2', 11);
    Exit;
  end;
  Pour3 := GetField('GPK_REMPOURMAX3');
  if (Pour3 <> 0) and (Pour3 <= Pour2) then
  begin
    MetEnErreur('GPK_REMPOURMAX3', 11);
    Exit;
  end;
  if Pour2 > 0 then
  begin
    Pwd1 := GetControlText('PWDREM2');
    Pwd2 := GetControlText('CONFPWDREM2');
    if Pwd1 <> Pwd2 then
    begin
      MetEnErreur('PWDREM2', 12);
      Exit;
    end;
  end;
  if Pour3 > 0 then
  begin
    Pwd1 := GetControlText('PWDREM3');
    Pwd2 := GetControlText('CONFPWDREM3');
    if Pwd1 <> Pwd2 then
    begin
      MetEnErreur('PWDREM3', 12);
      Exit;
    end;
  end;
  // contrôle de la durée d'annulation d'un ticket
  Pour1 := GetField('GPK_NBJANNULTIC');
  if (Pour1 < 0) or (Pour1 > 999) then
  begin
    MetEnErreur('GPK_NBJANNULTIC', 19);
    Exit;
  end;
  if CommercialFerme(GetControlText('GPK_VENDDEFAUT')) then // JTR - eQualité 11295
  begin
    MetEnErreur('GPK_VENDDEFAUT', 21);
    Exit;
  end;
  Result := True;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 31/12/2002
Modifié le ... : 31/12/2002
Description .. : Pas de paramétrage des modèles de tickets
Mots clefs ... :
*****************************************************************}

procedure TOM_ParCaisse.ModeleNonParametrable;
var Ind: integer;
begin
  if Ecran = nil then Exit;
  for Ind := 0 to Ecran.ComponentCount - 1 do
  begin
    if (Ecran.Components[Ind] is TToolbarButton97) and
      (copy(TToolbarButton97(Ecran.Components[Ind]).Name,1,11) = 'BPARAMETAT_') then
    begin
      TToolbarButton97(Ecran.Components[Ind]).Visible := False;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JTR
Créé le ...... : 10/09/2004
Modifié le ... :
Description .. : eQualité 11752
Suite ........ : Gestion du click sur les boutons de paramètrage des tickets
Mots clefs ... :
*****************************************************************}
procedure TOM_ParCaisse.ClickBtnParamTck;
var Ind: integer;
begin
  if Ecran = nil then Exit;
  for Ind := 0 to Ecran.ComponentCount - 1 do
  begin
    if (Ecran.Components[Ind] is TToolbarButton97) and
      (copy(TToolbarButton97(Ecran.Components[Ind]).Name,1,11) = 'BPARAMETAT_') then
    begin
      TToolbarButton97(Ecran.Components[Ind]).OnClick := ParamTicket;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JTR
Créé le ...... : 10/09/2004
Modifié le ... :
Description .. : eQualité 11752
Suite ........ : Gestion du click sur les boutons de paramètrage des tickets
Mots clefs ... :
*****************************************************************}
procedure TOM_ParCaisse.ParamTicket(Sender: TObject);
var LeNom, Lequel, TT, Range : string;
    LeCombo : THValComboBox;
    Action : TActionFiche;
begin
  LeNom := 'GPK' + copy(TToolBarButton97(Sender).name,pos('_',TToolBarButton97(Sender).name),length(TToolBarButton97(Sender).name));
  LeCombo := THValComboBox(GetControl(LeNom));
  Action := taModif;
  Lequel := Lecombo.Value;
  TT := LeCombo.DataType;
  Range := LeCombo.Plus;
  FODispatchParamModele(0,Action,Lequel,TT,Range);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 31/12/2002
Modifié le ... : 31/12/2002
Description .. : Particularités du BO
Mots clefs ... :
*****************************************************************}
{$IFNDEF FOS5}

procedure TOM_ParCaisse.ParticularitesBO;
begin
  // Les boutons de test ne sont pas disponibles
  SetControlVisible('BTESTTPE', False);
  SetControlVisible('BTESTPRT', False);
  SetControlVisible('BTESTTIROIR', False);
  SetControlVisible('BTESTAFF', False);
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 31/12/2002
Modifié le ... : 31/12/2002
Description .. : Particularités de la GESCOM
Mots clefs ... :
*****************************************************************}
{$IFDEF GESCOM}

procedure TOM_ParCaisse.ParticularitesGESCOM;
var
  Ctrl : TControl ;
begin
  Ctrl := GetControl('GOUV');
  if Ctrl <> nil then
  begin
    Ctrl.Visible := False;
    SetControlProperty('GCTRLCAIS', 'Top', Ctrl.Top);
  end ;
  SetControlVisible('GPK_GEREEVENT', False);
  SetControlVisible('BPARAMFFO', False);
  SetControlVisible('GDETAXE', False);
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 31/12/2002
Modifié le ... : 31/12/2002
Description .. : Construction du format d'affichage de la date de connexion
Mots clefs ... :
*****************************************************************}

procedure TOM_ParCaisse.FormatDateConnexion;
var Fmt: string;
begin
  Fmt := '"' + TraduireMemoire('En cours d''utilisation depuis le')
    + '" dddddd "' + TraduireMemoire('à') + '" hh:mm';
  SetControlProperty('GPK_DATECONNEXION', 'DisplayFormat', Fmt);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 02/04/2003
Modifié le ... : 02/04/2003
Description .. : On propose les paramètres de communication de
Suite ........ : l'imprimante pour les autres périphériques
Mots clefs ... :
*****************************************************************}

procedure TOM_ParCaisse.InitParamPeripherique;
var Stg, PortAff, PortTiroir, PortPrt: string;
begin
  PortPrt := Trim(GetField('GPK_PRTPORT'));
  PortAff := Trim(GetField('GPK_AFFPORT'));
  if (GetField('GPK_AFFTYPE') <> '') and (GetField('GPK_PRTTYPE') <> '') then
  begin
    if PortAff = '' then
    begin
      PortAff := GetField('GPK_PRTPORT');
      SetField('GPK_AFFPORT', PortAff);
      SetControlText('GPK_AFFPORT', PortAff);
    end;
    Stg := Trim(GetField('GPK_AFFPARAM'));
    if (Stg = '') or ((PortAff <> '') and (PortAff = PortPrt)) then
    begin
      Stg := GetField('GPK_PRTBAUDS') + ';' + GetField('GPK_PRTPARITE') + ';'
        + GetField('GPK_PRTNBBIT') + ';' + GetField('GPK_PRTSTOPBIT') + ';'
        + GetField('GPK_PRTCTRLFLUX') + ';' + GetField('GPK_PRTMODEBIDI');
      SetField('GPK_AFFPARAM', Stg);
    end;
  end;
  // l'afficheur utilise le même port que l'imprimante
  AffSurPrt := ((PortAff <> '') and (PortAff = PortPrt));

  PortTiroir := Trim(GetField('GPK_TIROIRPORT'));
  if (GetField('GPK_TIROIRTYPE') <> '') and (GetField('GPK_PRTTYPE') <> '') then
  begin
    if PortTiroir = '' then
    begin
      PortTiroir := GetField('GPK_PRTPORT');
      SetField('GPK_TIROIRPORT', PortTiroir);
    end;
    Stg := Trim(GetField('GPK_TIROIRPARAM'));
    if (Stg = '') or ((PortTiroir <> '') and (PortTiroir = PortPrt)) then
    begin
      Stg := GetField('GPK_PRTBAUDS') + ';' + GetField('GPK_PRTPARITE') + ';'
        + GetField('GPK_PRTNBBIT') + ';' + GetField('GPK_PRTSTOPBIT') + ';'
        + GetField('GPK_PRTCTRLFLUX') + ';' + GetField('GPK_PRTMODEBIDI');
      SetField('GPK_TIROIRPARAM', Stg);
    end;
  end;
  // le tiroir utilise le même port que l'imprimante
  TiroirSurPrt := ((PortTiroir <> '') and (PortTiroir = PortPrt));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 02/04/2003
Modifié le ... : 02/04/2003
Description .. : Mise en forme des paramètres de l'afficheur et du tiroir
Mots clefs ... :
*****************************************************************}

procedure TOM_ParCaisse.MajParamPeripherique;
var Stg: string;
begin
  Stg := '';
  if GetField('GPK_AFFTYPE') <> '' then
  begin
    Stg := GetControlText('GPK_AFFBAUDS') + ';' + GetControlText('GPK_AFFPARITE') + ';'
      + GetControlText('GPK_AFFNBBIT') + ';' + GetControlText('GPK_AFFSTOPBIT') + ';'
      + GetControlText('GPK_AFFCTRLFLUX') + ';' + GetControlText('GPK_AFFMODEBIDI');
  end;
  SetField('GPK_AFFPARAM', Stg);
  Stg := '';
  if GetField('GPK_TIROIRTYPE') <> '' then
  begin
    Stg := GetControlText('GPK_TIRBAUDS') + ';' + GetControlText('GPK_TIRPARITE') + ';'
      + GetControlText('GPK_TIRNBBIT') + ';' + GetControlText('GPK_TIRSTOPBIT') + ';'
      + GetControlText('GPK_TIRCTRLFLUX') + ';' + GetControlText('GPK_TIRMODEBIDI');
  end;
  SetField('GPK_TIROIRPARAM', Stg);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 31/12/2002
Modifié le ... : 31/12/2002
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOM_ParCaisse.OnArgument(stArgument: string);
var Fmt: string;
begin
  inherited;
  // traitement des particularités des différents produits
  {$IFNDEF TOXCLIENT}
  //SetControlVisible('BFOBO', False);
  SetControlVisible('PCOMM', False);
  {$ELSE}
   {$IFDEF GCGC}
  SetControlVisible('PCOMM', False); // JTR - eQualité 11344
   {$ENDIF GCGC}
  {$ENDIF TOXCLIENT}
  {$IFDEF GESCOM}
  ParticularitesGESCOM;
  {$ENDIF}
  {$IFDEF EAGLCLIENT}
  ModeleNonParametrable;
  {$ELSE}
  ClickBtnParamTck;
  {$ENDIF EAGLCLIENT}
  {$IFNDEF FOS5}
  ParticularitesBO;
  {$ENDIF}
  // format d'affichage de la date de connexion
  FormatDateConnexion;
  // format de saisie des montants
  Fmt := StrfMask(V_PGI.OkDecV, '', True);
  SetControlProperty('GPK_SEUILCLIOBLIG', 'DisplayFormat', Fmt);
  // en mode consultation
  if (Ecran <> nil) and (Ecran is TFFicheListe) and
    (TFFicheListe(Ecran).TypeAction = taConsult) then
  begin
    SetControlEnabled('BDUPLIQUE', False);
    SetControlEnabled('BCLAVIER', False);
    SetControlEnabled('BPARAMFFO', False);
    ModeleNonParametrable;
  end;
  // ancien mode de contrôle d'accès sur le % maximum de remise
  SetControlVisible('GREMPOUR', (ctxFO in V_PGI.PGIContexte));
  SetControlVisible('GREMPOUR_OLD', not (ctxFO in V_PGI.PGIContexte));
  if V_PGI.DrawXP then
  begin
    SetControlProperty('BTESTPRT', 'DisplayMode', dmGlyphOnly);
    SetControlProperty('BTESTAFF', 'DisplayMode', dmGlyphOnly);
    SetControlProperty('BTESTTIROIR', 'DisplayMode', dmGlyphOnly);
    SetControlProperty('BTESTTPE', 'DisplayMode', dmGlyphOnly);
  end;
  SetActiveTabSheet('PGENERAL');
end;

procedure TOM_PARCAISSE.OnClose;
var Qry : TQuery;
    Qte : integer;
begin
  inherited;
  // Test si fond de caisse paramétré et ParamSoc Ok
  Qry := OpenSQL('SELECT COUNT(*) AS QTE FROM PARCAISSE WHERE GPK_GEREFONDCAISSE="X"',true);
  if not Qry.Eof then
  begin
    Qte := Qry.FindField('QTE').AsInteger;
    if (Qte > 0) and ((GetParamsoc('SO_GCVRTINTERNE')='') or (GetParamsoc('SO_GCCAISSGAL')='')) then
      PgiBox('Le paramètrage général pour la gestion du fond de caisse est incomplet.');
  end;
  Ferme(Qry);
end;


///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeField
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeField(F: TField);
begin
  inherited;
  if F.FieldName = 'GPK_ETABLISSEMENT' then OnChangeEtablissement else
    if F.FieldName = 'GPK_CLISAISIE' then OnChangeCliSaisie else
    if F.FieldName = 'GPK_CLIREPRISE' then OnChangeCliReprise else
    if F.FieldName = 'GPK_PRTTYPE' then OnChangePrtType else
    if F.FieldName = 'GPK_PRTPORT' then OnChangePrtPort else
    if (F.FieldName = 'GPK_PRTBAUDS') or (F.FieldName = 'GPK_PRTCTRLFLUX') or
    (F.FieldName = 'GPK_PRTSTOPBIT') or (F.FieldName = 'GPK_PRTPARITE') or
    (F.FieldName = 'GPK_PRTNBBIT') or (F.FieldName = 'GPK_PRTMODEBIDI') then OnChangePrtParam(F.FieldName) else
    if F.FieldName = 'GPK_TPETYPE' then OnChangeTpeType else
    if F.FieldName = 'GPK_TIROIRTYPE' then OnChangeTiroirType else
    if F.FieldName = 'GPK_TIROIRPORT' then OnChangeTiroirPort else
    if F.FieldName = 'GPK_TIROIRPARAM' then OnChangeTiroirParam else
    if F.FieldName = 'GPK_PRTFISCTYPE' then OnChangePrtFiscalType else
    if F.FieldName = 'GPK_SIGNPADTYPE' then OnChangeSignPadType else
    if (F.FieldName = 'GPK_VENDSAISIE') or (F.FieldName = 'GPK_VENDSAISLIG') then OnChangeVendeur else
    if F.FieldName = 'GPK_MODEDETAXE' then OnChangeModeDetaxe else
    if F.FieldName = 'GPK_REMSAISIE' then OnChangeRemSaisie else
    if F.FieldName = 'GPK_REMPOURMAX' then OnChangeRemPourMax else
    if F.FieldName = 'GPK_REMPOURMAX2' then OnChangeRemPourMax2 else
    if F.FieldName = 'GPK_REMPOURMAX3' then OnChangeRemPourMax3 else
    if F.FieldName = 'GPK_DEMSAISIE' then OnChangeDemSaisie else
    if F.FieldName = 'GPK_CTRLCAISSE' then OnChangeCtrlCaisse else
    if F.FieldName = 'GPK_GEREFONDCAISSE' then OnChangeGereFondCaisse else
    if F.FieldName = 'GPK_GEREREMISEBNQ' then OnChangeGereRemiseBnq else
    if F.FieldName = 'GPK_AFFICHEUR' then OnChangeAfficheur else
    if F.FieldName = 'GPK_AFFTYPE' then OnChangeAffType else
    if F.FieldName = 'GPK_AFFPORT' then OnChangeAffPort else
    if F.FieldName = 'GPK_AFFPARAM' then OnChangeAffParam else
    if F.FieldName = 'GPK_CLAVIERECRAN' then OnChangeClavierEcran else
    if F.FieldName = 'GPK_IMPFMTTIC' then OnChangeImpFmtTic else
    if F.FieldName = 'GPK_TOXAPPEL1' then OnChangeToxAppel1;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 31/12/2002
Modifié le ... : 31/12/2002
Description .. : OnUpdateRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_ParCaisse.OnUpdateRecord;
var
  Pwd, sCaisse: string;
  Visu: boolean;
begin
  inherited;
  // Mise en forme des paramètres de l'afficheur et du tiroir
  MajParamPeripherique;
  // Contrôle de l'intégrité des données
  if not ControleIntegrite then Exit;
  // Recopie du paramétrage du clavier de la caisse choisie
  if ClavierOrg <> '' then
  begin
    sCaisse := GetField('GPK_CAISSE');
    // recopie du pavé principal
    RecopieClavier(sCaisse, ClavierOrg);
    // recopie du 2ème pavé
    RecopieClavier(FOAlphaCodeNumeric(sCaisse), FOAlphaCodeNumeric(ClavierOrg));
    ClavierOrg := '';
  end;
  // Enregistrement de la configuration du ToxServeur
  {$IFDEF TOXCLIENT}
  MajHeureTox;
  {$ENDIF}
  // Cryptage des mots de passe
  Pwd := CryptageSt(GetControlText('PWDREM2'));
  SetField('GPK_PWDREM2', Pwd);
  Pwd := CryptageSt(GetControlText('PWDREM3'));
  SetField('GPK_PWDREM3', Pwd);
  // L'accès au paramétrage du clavier est activé
  if (Ecran <> nil) and (Ecran is TFFicheListe) then
  begin
    Visu := (TFFicheListe(Ecran).TypeAction <> taConsult);
    SetControlEnabled('BCLAVIER', Visu);
    SetControlEnabled('BPARAMFFO', Visu);
  end;
  // Active les boutons de test des périphériques
  ActiveBtnTest;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 18/06/2001
Modifié le ... : 18/06/2001
Description .. : OnAfterUpdateRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_PARCAISSE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 31/12/2002
Modifié le ... : 31/12/2002
Description .. : OnLoadRecord
Mots clefs ... :
*****************************************************************}

procedure TOM_ParCaisse.OnLoadRecord;
var Visu: boolean;
begin
  inherited;
  ClavierOrg := '';
  Visu := (GetField('GPK_INUSE') = 'X');
  SetControlVisible('GPK_DATECONNEXION', Visu);
  if Visu then FormatDateConnexion;

  Visu := (GetField('GPK_DATEINTEGR') <> iDate1900);
  SetControlVisible('GPK_DATEINTEGR', Visu);
  SetControlVisible('TGPK_DATEINTEGR', Visu);

  SetControlVisible('BLIBERE', Visu);
  if (ctxMode in V_PGI.PGIContexte) and not (VH_GC.GCMultiDepots) then
  begin
    SetControlVisible('GPK_DEPOT', False);
    SetControlVisible('TGPK_DEPOT', False);
  end;
  // Décryptage des mots de passe
  DecryptePwdRem;
  // L'accès au paramétrage du clavier est inactivé en création d'une caisse
  if (Ecran <> nil) and (Ecran is TFFicheListe) then
  begin
    if TFFicheListe(Ecran).TypeAction = taConsult then
      Visu := False
    else
      Visu := (DS.State <> dsInsert);
    SetControlEnabled('BCLAVIER', Visu);
    SetControlEnabled('BPARAMFFO', Visu);
  end;
  // si besoin on propose les paramètres de communication de l'imprimante pour les autres périphériques
  InitParamPeripherique;
  // Active les boutons de test des périphériques
  ActiveBtnTest;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnNewRecord
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnNewRecord;
var Nature, CodeEtat, Titre: string;
begin
  inherited;
  ClavierOrg := '';
  // On se positionne sur l'onglet général
  SetActiveTabSheet('PGENERAL');
  SetControlVisible('GPK_DATEINTEGR', False);
  SetControlVisible('TGPK_DATEINTEGR', False);
  // Valeurs par défaut de l'établissement et du dépôt de rattachement qui a le même code
  // que l'établissement dans le contexte de la mode.
  SetField('GPK_ETABLISSEMENT', VH^.EtablisDefaut);
  if ctxMode in V_PGI.PGIContexte then SetField('GPK_DEPOT', VH^.EtablisDefaut)
  else SetField('GPK_DEPOT', VH_GC.GCDepotDefaut);
  // On utilise les prix de vente détail (TTC).
  SetField('GPK_APPELPRIXTIC', 'TTC');
  // On propose d'afficher le détail des familles ou le détail des dimensions
  SetField('GPK_INFOARTICLE', '902');
  SetField('GPK_INFOARTDIM', '901');
  // On propose les modèles par défaut
  FODonneCodeEtat(efoTicket, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPFMTTIC', 'T');
  SetField('GPK_IMPMODTIC', CodeEtat);
  FODonneCodeEtat(efoTicketX, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODTICX', CodeEtat);
  FODonneCodeEtat(efoTicketZ, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODTICZ', CodeEtat);
  FODonneCodeEtat(efoRecapVend, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODRECAPV', CodeEtat);
  FODonneCodeEtat(efoListeRegle, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODLISREG', CodeEtat);
  FODonneCodeEtat(efoCheque, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODCHEQUE', CodeEtat);
  FODonneCodeEtat(efoBonAchat, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODBONACH', CodeEtat);
  FODonneCodeEtat(efoBonArrhes, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODBONARR', CodeEtat);
  FODonneCodeEtat(efoBonAvoir, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODBONAV', CodeEtat);
  FODonneCodeEtat(efoTransfert, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODBONTID', CodeEtat);
  FODonneCodeEtat(efoCommande, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODBONCDE', CodeEtat);
  FODonneCodeEtat(efoReception, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODBONREC', CodeEtat);
  FODonneCodeEtat(efoRemiseBq, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODREMIS', CodeEtat);
  FODonneCodeEtat(efoFdCaisse, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODFDC', CodeEtat);
  FODonneCodeEtat(efoStatFam, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODSTAFAM', CodeEtat);
  FODonneCodeEtat(efoStatRem, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODSTAREM', CodeEtat);
  FODonneCodeEtat(efoListeArtVendu, True, Nature, CodeEtat, Titre);
  SetField('GPK_IMPMODLISART', CodeEtat);
  // Affichage du prix en grille
  SetField('GPK_AFFPRIXTIC', '002');
  SetFocusControl('GPK_CAISSE')
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnDeleteRecord
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnDeleteRecord;
var
  sCaisse: string;
begin
  sCaisse := GetField('GPK_CAISSE');
  // suppression du pavé principal
  DeleteClavier(sCaisse);
  // suppression du 2ème pavé
  DeleteClavier(FOAlphaCodeNumeric(sCaisse));
  inherited;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnCancelRecord
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_PARCAISSE.OnCancelRecord;
begin
  inherited;
  OnChangeEtablissement;
  OnChangeCliSaisie;
  OnChangeCliReprise;
  OnChangePrtType;
  OnChangePrtPort;
  OnChangeTpeType;
  InitParamPeripherique;
  OnChangeTiroirType;
  OnChangeTiroirPort;
  OnChangeTiroirParam;
  OnChangeVendeur;
  OnChangeModeDetaxe;
  OnChangeRemSaisie;
  OnChangeRemPourMax;
  OnChangeRemPourMax2;
  OnChangeRemPourMax3;
  OnChangeDemSaisie;
  OnChangeCtrlCaisse;
  OnChangeGereFondCaisse;
  OnChangeGereRemiseBnq;
  OnChangeAfficheur;
  OnChangeAffType;
  OnChangeAffPort;
  OnChangeAffParam;
  OnChangeClavierEcran;
  OnChangeImpFmtTic;
  OnChangeToxAppel1;
  DecryptePwdRem;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeEtablissement : Modification de l'établissement
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeEtablissement;
var Stg, sLst: string;
  QQ: TQuery;
begin
  if ctxMode in V_PGI.PGIContexte then
  begin
    if VH_GC.GCMultiDepots then
    begin
      Stg := 'SELECT ET_DEPOTLIE FROM ETABLISS WHERE ET_ETABLISSEMENT="' + GetField('GPK_ETABLISSEMENT') + '"';
      QQ := OpenSQL(Stg, True);
      if not QQ.EOF then
      begin
        sLst := QQ.FindField('ET_DEPOTLIE').AsString;
        Stg := FOFabriqueSQLIN(sLst, 'GDE_DEPOT', False, False, False);
      end;
      Ferme(QQ);
      if Stg <> '' then Stg := Stg + ' AND ';
      Stg := Stg + 'GDE_SURSITE="X"';
      SetControlProperty('GPK_DEPOT', 'Plus', Stg);
    end else
    begin
      // En mono-dépôt, le dépôt est égal au code établissement
      SetField('GPK_DEPOT', GetField('GPK_ETABLISSEMENT'));
    end;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeCliSaisie : Modification de l'indicateur de saisie du client
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeCliSaisie;
var Active: Boolean;
begin
  Active := (GetField('GPK_CLISAISIE') = 'X');
  ActiveChamp('GPK_CLIREPRISE', Active, 'X', 'X');
  ActiveChamp('GPK_CLISAISIENOM', Active);
  ActiveChamp('GPK_SEUILCLIOBLIG', Active);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeCliReprise : Modification de l'indicateur de reprise du client
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeCliReprise;
var Active: Boolean;
begin
  Active := (GetField('GPK_CLIREPRISE') = 'X');
  ActiveChamp('GPK_PREPOSGRID', Active);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangePrtType : Modification du type d'imprimante
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangePrtType;
var
  Active: Boolean;
  sDispositif, sParam, sPort: string;
begin
  sDispositif := GetField('GPK_PRTTYPE');
  // Activation du port de l'imprimante avec COM1 par défaut
  sParam := DefautTPV(sDispositif, mcPrtCheques);
  sPort := ReadTokenST(sParam);
  sPort := Copy(sPort, 1, 2) + Copy(sPort, 4, 1);
  if sPort = '' then sPort := 'CO1';
  
  Active := ((sDispositif <> '') and (sPort <> 'OPS'));
  ActiveChamp('GPK_PRTPORT', Active, sPort);
  // Activation des périphériques associés (tiroir, afficheur)
  if not Active then
  begin
    if AffSurPrt then
    begin
      AffSurPrt := False;
      OnChangeAffPort;
    end;
    if TiroirSurPrt then
    begin
      TiroirSurPrt := False;
      OnChangeTiroirPort;
    end;
  end;
  ///ActiveChamp('GPK_TIROIRTYPE', Active, '') ;
  ///ActiveChamp('GPK_AFFTYPE', Active, '') ;
  // Activation du choix impression des chèques pour les imprimantes tri fonctions.
  Active := UsesTPV(sDispositif, mcPrtCheques);
  ActiveChamp('GPK_REMPLIRCHQ', Active, '-');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangePrtPort : Modification du port de l'imprimante
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangePrtPort;
var Stg: string;
  Active: Boolean;
begin
  // valeurs par défaut 9600 bps, None, 8, 1, DTS/DSR
  Stg := GetField('GPK_PRTPORT');
  Active := ((Stg = 'CO1') or (Stg = 'CO2') or (Stg = 'CO3') or (Stg = 'CO4'));
  if (Active) and (AffSurPrt) then SetField('GPK_AFFPORT', Stg);
  if (Active) and (TiroirSurPrt) then SetField('GPK_TIROIRPORT', Stg);
  ActiveChamp('GPK_PRTBAUDS', Active, '007');
  ActiveChamp('GPK_PRTPARITE', Active, '001');
  ActiveChamp('GPK_PRTNBBIT', Active, '004');
  ActiveChamp('GPK_PRTSTOPBIT', Active, '001');
  ActiveChamp('GPK_PRTCTRLFLUX', Active, '003');
  ActiveChamp('GPK_PRTMODEBIDI', Active, '-');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangePrtParam : Modification des paramètres de l'imprimante
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangePrtParam(NomChamp: string);
var Stg: string;
begin
  if (not AffSurPrt) and (not TiroirSurPrt) then Exit;

  Stg := GetField(NomChamp);
  if NomChamp = 'GPK_PRTBAUDS' then
  begin
    if AffSurPrt then SetControlText('GPK_AFFBAUDS', Stg);
    if TiroirSurPrt then SetControlText('GPK_TIRBAUDS', Stg);
  end else
    if NomChamp = 'GPK_PRTCTRLFLUX' then
  begin
    if AffSurPrt then SetControlText('GPK_AFFCTRLFLUX', Stg);
    if TiroirSurPrt then SetControlText('GPK_TIRCTRLFLUX', Stg);
  end else
    if NomChamp = 'GPK_PRTSTOPBIT' then
  begin
    if AffSurPrt then SetControlText('GPK_AFFSTOPBIT', Stg);
    if TiroirSurPrt then SetControlText('GPK_TIRSTOPBIT', Stg);
  end else
    if NomChamp = 'GPK_PRTPARITE' then
  begin
    if AffSurPrt then SetControlText('GPK_AFFPARITE', Stg);
    if TiroirSurPrt then SetControlText('GPK_TIRPARITE', Stg);
  end else
    if NomChamp = 'GPK_PRTNBBIT' then
  begin
    if AffSurPrt then SetControlText('GPK_AFFNBBIT', Stg);
    if TiroirSurPrt then SetControlText('GPK_TIRNBBIT', Stg);
  end else
    if NomChamp = 'GPK_PRTMODEBIDI' then
  begin
    if AffSurPrt then SetControlChecked('GPK_AFFMODEBIDI', (Stg = 'X'));
    if TiroirSurPrt then SetControlChecked('GPK_TIRMODEBIDI', (Stg = 'X'));
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeTpeType : Modification du type de TPE
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeTpeType;
var
  Active: Boolean;
  TpeType: string;
begin
  // Activation du port du TPE avec COM2 à 1200 bps par défaut
  TpeType := GetField('GPK_TPETYPE');
  Active := ((TpeType = '001') or (TpeType = '002'));
  ActiveChamp('GPK_TPEPORT', Active, 'CO2');
  ActiveChamp('GPK_TPEBAUDS', Active, '004');
  Active := (TpeType = '101');
  ActiveChamp('GPK_TPEREPERTOIRE', Active);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeTiroirType : Modification du type de tiroir
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeTiroirType;
var
  Active: boolean;
  sDispositif, sParam, sPort: string;
begin
  sDispositif := GetField('GPK_TIROIRTYPE');
  // Activation du port du tiroir avec COM1 par défaut
  sParam := DefautTPV(sDispositif, mcTiroir);
  sPort := ReadTokenST(sParam);
  sPort := Copy(sPort, 1, 2) + Copy(sPort, 4, 1);
  if sPort = '' then sPort := 'CO1';

  Active := ((sDispositif <> '') and (sPort <> 'OPS'));
  ActiveChamp('GPK_TIROIRPORT', Active, sPort);
  ActiveChamp('GPK_TIROIRPIN5', Active, '-');
  Active := (sDispositif <> '');
  ActiveChamp('GPK_TIROIRESP', Active, '-');
  Active := ((sDispositif <> '') and (GetField('GPK_CTRLCAISSE') = 'X'));
  ActiveChamp('GPK_OUVTIROIRFIN', Active, '-');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeTiroirPort : Modification du port du tiroir
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeTiroirPort;
var Stg: string;
  Active: boolean;
begin
  // valeurs par défaut 9600 bps, None, 8, 1, DTS/DSR
  Stg := GetField('GPK_TIROIRPORT');
  if (Stg <> '') and (Stg = GetField('GPK_PRTPORT')) then
  begin
    TiroirSurPrt := True;
    SetControlEnabled('GPK_TIROIRPARAM', False);
    InitParamPeripherique;
  end else
  begin
    TiroirSurPrt := False;
    Active := ((Stg = 'CO1') or (Stg = 'CO2') or (Stg = 'CO3') or (Stg = 'CO4'));
    ActiveChamp('GPK_TIROIRPARAM', Active, '007;001;004;001;003;-;');
    OnChangeTiroirParam;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeTiroirParam : Modification des paramètres de communication du tiroir
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeTiroirParam;
var sParam, Stg: string;
  Active: boolean;
begin
  Active := GetControlEnabled('GPK_TIROIRPARAM');
  sParam := GetField('GPK_TIROIRPARAM');
  ActiveChamp('GPK_TIRBAUDS', Active);
  SetControlText('GPK_TIRBAUDS', ReadTokenST(sParam));
  ActiveChamp('GPK_TIRPARITE', Active);
  SetControlText('GPK_TIRPARITE', ReadTokenST(sParam));
  ActiveChamp('GPK_TIRNBBIT', Active);
  SetControlText('GPK_TIRNBBIT', ReadTokenST(sParam));
  ActiveChamp('GPK_TIRSTOPBIT', Active);
  SetControlText('GPK_TIRSTOPBIT', ReadTokenST(sParam));
  ActiveChamp('GPK_TIRCTRLFLUX', Active);
  SetControlText('GPK_TIRCTRLFLUX', ReadTokenST(sParam));
  ActiveChamp('GPK_TIRMODEBIDI', Active);
  Stg := ReadTokenST(sParam);
  SetControlChecked('GPK_TIRMODEBIDI', (Stg = 'X'));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangePrtFiscalType : Modification du type d'imprimante fiscale
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangePrtFiscalType;
var
  Active: boolean;
  sDispositif: string;
begin
  sDispositif := GetField('GPK_PRTFISCTYPE');
  // Activation du port de l'imprimante avec COM1 par défaut
  Active := (sDispositif <> '');
  ActiveChamp('GPK_PRTFISCPORT', Active, 'CO1');
  ActiveChamp('GPK_PRTFISCENPLUS', Active, '-');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeSignPadType : Modification du type de terminal de signature
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeSignPadType;
var
  Active: boolean;
  sDispositif: string;
begin
  sDispositif := GetField('GPK_SIGNPADTYPE');
  // Activation des options du terminal
  Active := (sDispositif <> '');
  ActiveChamp('GPK_SIGNERAPRIORI', Active, '-');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeVendeur : Modification de l'indicateur de saisie du vendeur
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeVendeur;
var Active: Boolean;
begin
  Active := ((GetField('GPK_VENDSAISIE') = 'X') or (GetField('GPK_VENDSAISLIG') = 'X'));
  ActiveChamp('GPK_VENDOBLIG', Active);
  ActiveChamp('GPK_VENDREPRISE', Active);
  if (GetField('GPK_VENDSAISIE') = 'X') and (GetField('GPK_VENDSAISLIG') = '-') then
  begin
    ActiveChamp('GPK_VENDLIGREP', False, '');
    SetField('GPK_VENDLIGREP', 'ENT');
  end else ActiveChamp('GPK_VENDLIGREP', Active);
  ActiveChamp('GPK_VENDDEFAUT', Active);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeModeDetaxe : Modification du mode de gestion de la détaxe
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeModeDetaxe;
var Active: boolean;
  Stg: string;
begin
  Stg := GetField('GPK_MODEDETAXE');
  Active := ((Stg <> '') and (Stg <> '000'));
  ActiveChamp('GPK_CLIOBLIGDETAXE', Active);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeRemSaisie : Modification de l'indicateur de saisie de la remise
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeRemSaisie;
var Active: Boolean;
begin
  Active := (GetField('GPK_REMSAISIE') = 'X');
  ActiveChamp('GPK_REMAFFICH', Active);
  ActiveChamp('GPK_MAJORPRX', Active);
  ActiveChamp('GPK_REMARRONDI', Active);
  ActiveChamp('GPK_REMPIEDLIG', Active);
  ActiveChamp('GPK_DEMSAISIE', Active);
  ActiveChamp('GPK_REMPOURMAX', Active);
  if (Active) and (GetField('GPK_REMPOURMAX') = 0) and (GetField('GPK_REMPOURMAX2') = 0) then
    SetControlProperty('GPK_REMPOURMAX', 'Value', 100);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeRemPourMax : Modification du % maximun de remise niveau 1
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeRemPourMax;
var Active: Boolean;
  Pourcent: Integer;
begin
  Pourcent := GetField('GPK_REMPOURMAX');
  Active := ((Pourcent < 100) and GetControlEnabled('GPK_REMPOURMAX'));
  ActiveChamp('GPK_REMPOURMAX2', Active);
  if not Active then SetControlProperty('GPK_REMPOURMAX2', 'Value', 0);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeRemPourMax2 : Modification du % maximun de remise niveau 2
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeRemPourMax2;
var Active: Boolean;
  Pourcent: Integer;
begin
  Pourcent := GetField('GPK_REMPOURMAX2');
  Active := (Pourcent > 0);
  ActiveChamp('PWDREM2', Active);
  ActiveChamp('CONFPWDREM2', Active);
  ActiveChamp('GPK_REMPOURMAX3', Active);
  if not Active then
  begin
    SetControlText('PWDREM2', '');
    SetControlText('CONFPWDREM2', '');
    SetControlProperty('GPK_REMPOURMAX3', 'Value', 0);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeRemPourMax3 : Modification du % maximun de remise niveau 3
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeRemPourMax3;
var Active: Boolean;
  Pourcent: Integer;
begin
  Pourcent := GetField('GPK_REMPOURMAX3');
  Active := (Pourcent > 0);
  ActiveChamp('PWDREM3', Active);
  ActiveChamp('CONFPWDREM3', Active);
  if not Active then
  begin
    SetControlText('PWDREM3', '');
    SetControlText('CONFPWDREM3', '');
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeDemSaisie : Modification de l'indicateur de saisie de la démarque
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeDemSaisie;
var Active: Boolean;
begin
  Active := (GetField('GPK_DEMSAISIE') = 'X');
  ActiveChamp('GPK_DEMOBLIG', Active);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeCtrlCaisse : Modification de l'indicateur de saisie du contrôle de caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeCtrlCaisse;
var Active: Boolean;
begin
  Active := (GetField('GPK_CTRLCAISSE') = 'X');
  ActiveChamp('GPK_CTRLAVEUGLE', Active);
  ActiveChamp('GPK_MDPCTRLCAIS', Active);
  ActiveChamp('GPK_GEREFONDCAISSE', Active);
  ActiveChamp('GPK_CTRLECART', Active);
  ActiveChamp('GPK_CTRLPIECBIL', Active);
  ActiveChamp('GPK_GEREREMISEBNQ', Active);
  Active := ((GetField('GPK_TIROIRTYPE') <> '') and (GetField('GPK_CTRLCAISSE') = 'X'));
  ActiveChamp('GPK_OUVTIROIRFIN', Active, '-');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeGereFondCaisse : Modification de l'indicateur de saisie du contrôle de caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeGereFondCaisse;
var Active: Boolean;
begin
  Active := (GetField('GPK_GEREFONDCAISSE') = 'X');
  ActiveChamp('GPK_MDPFDCAIS', Active);
  ActiveChamp('GPK_FDCAISSE', Active);
  ActiveChamp('GPK_MODIFFDCAIS', Active);
  ActiveChamp('GPK_REPRISEFDCAIS', Active);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeGereRemiseBnq : Modification de l'indicateur de saisie du contrôle de caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeGereRemiseBnq;
var Active: Boolean;
begin
  Active := (GetField('GPK_GEREREMISEBNQ') = 'X');
  ActiveChamp('GPK_OPREMISEBNQ', Active);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeAfficheur : Modification de l'existence de l'afficheur interne
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeAfficheur;
var Active: Boolean;
begin
  Active := (GetField('GPK_AFFICHEUR') = 'X');
  ActiveChamp('GPK_AFFINVERSE', Active);
  if not Active then Active := (GetField('GPK_AFFTYPE') <> '');
  ActiveChamp('GPK_AFFMESG', Active);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeAffType : Modification du type d'afficheur externe
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeAffType;
var
  Active: boolean;
  sDispositif, sParam, sPort: string;
begin
  sDispositif := GetField('GPK_AFFTYPE');
  // Activation du port de l'afficheur avec COM1 par défaut
  sParam := DefautTPV(sDispositif, mcAfficheur);
  sPort := ReadTokenST(sParam);
  sPort := Copy(sPort, 1, 2) + Copy(sPort, 4, 1);
  if sPort = '' then sPort := 'CO1';

  Active := ((sDispositif <> '') or (GetField('GPK_AFFICHEUR') = 'X'));
  ActiveChamp('GPK_AFFMESG', Active);
  Active := ((GetField('GPK_AFFTYPE') <> '') and (sPort <> 'OPS'));
  ActiveChamp('GPK_AFFPORT', Active, sPort);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeAffPort : Modification du port de l'afficheur externe
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeAffPort;
var Stg: string;
  Active: boolean;
begin
  // valeurs par défaut 9600 bps, None, 8, 1, DTS/DSR
  Stg := GetField('GPK_AFFPORT');
  if (Stg <> '') and (Stg = GetField('GPK_PRTPORT')) then
  begin
    AffSurPrt := True;
    SetControlEnabled('GPK_AFFPARAM', False);
    InitParamPeripherique;
  end else
  begin
    AffSurPrt := False;
    Active := ((Stg = 'CO1') or (Stg = 'CO2') or (Stg = 'CO3') or (Stg = 'CO4'));
    ActiveChamp('GPK_AFFPARAM', Active, '007;001;004;001;003;-;');
    OnChangeAffParam;
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeAffParam : Modification des paramètres de communication de l'afficheur externe
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeAffParam;
var sParam, Stg: string;
  Active: boolean;
begin
  Active := GetControlEnabled('GPK_AFFPARAM');
  sParam := GetField('GPK_AFFPARAM');
  ActiveChamp('GPK_AFFBAUDS', Active);
  SetControlText('GPK_AFFBAUDS', ReadTokenST(sParam));
  ActiveChamp('GPK_AFFPARITE', Active);
  SetControlText('GPK_AFFPARITE', ReadTokenST(sParam));
  ActiveChamp('GPK_AFFNBBIT', Active);
  SetControlText('GPK_AFFNBBIT', ReadTokenST(sParam));
  ActiveChamp('GPK_AFFSTOPBIT', Active);
  SetControlText('GPK_AFFSTOPBIT', ReadTokenST(sParam));
  ActiveChamp('GPK_AFFCTRLFLUX', Active);
  SetControlText('GPK_AFFCTRLFLUX', ReadTokenST(sParam));
  ActiveChamp('GPK_AFFMODEBIDI', Active);
  Stg := ReadTokenST(sParam);
  SetControlChecked('GPK_AFFMODEBIDI', (Stg = 'X'));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeClavierEcran : Modification de l'existence d'un pavé tactile
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeClavierEcran;
var Active: Boolean;
  DefVal: string;
begin
  if GetField('GPK_CLAVIERECRAN') = 'X' then
  begin
    Active := True;
    {$IFDEF GESCOM}
    DefVal := 'X';
    {$ELSE}
    DefVal := '-';
    {$ENDIF}
  end else
  begin
    Active := False;
    DefVal := '';
  end;
  ActiveChamp('GPK_TOOLBAR', Active, DefVal, 'X');
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeImpFmtTic : Modification du format d'impression du ticket de caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeImpFmtTic;
var Stg, Nature, Format: string;
  Active: Boolean;
begin
  Active := (GetField('GPK_IMPFMTTIC') <> '');
  ActiveChamp('GPK_IMPMODTIC', Active);
  ActiveChamp('GPK_NBEXEMPTIC', Active);
  if Active then
  begin
    Format := GetField('GPK_IMPFMTTIC');
    Nature := '';
    if Format = 'T' then Nature := NATUREMODTICKET else
      if Format = 'E' then Nature := NATUREMODTICETAT else
      if Format = 'L' then Nature := NATUREMODTICDOC;
    Stg := 'MO_NATURE="' + Nature + '" AND MO_TYPE="' + Format + '"';
    SetControlProperty('GPK_IMPMODTIC', 'Plus', Stg);
  end;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  OnChangeToxAppel1 : Modification de l'heure du 1er appel
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.OnChangeToxAppel1;
{$IFDEF TOXCLIENT}
var Active: Boolean;
  {$ENDIF}
begin
  {$IFDEF TOXCLIENT}
  Active := ((GetField('GPK_TOXAPPEL1') <> '') and (GetField('GPK_TOXAPPEL1') <> '  :  :  '));
  ActiveChamp('GPK_TOXAPPEL2', Active);
  {$ENDIF}
end;

///////////////////////////////////////////////////////////////////////////////////////
//  DecryptePwdRem : Décryptage des mots de passe
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.DecryptePwdRem;
var Pwd: string;
begin
  Pwd := DeCryptageSt(GetField('GPK_PWDREM2'));
  SetControlText('PWDREM2', Pwd);
  SetControlText('CONFPWDREM2', Pwd);
  Pwd := DeCryptageSt(GetField('GPK_PWDREM3'));
  SetControlText('PWDREM3', Pwd);
  SetControlText('CONFPWDREM3', Pwd);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  DeleteClavier : supprime le paramétrage du clavier d'une caisse
///////////////////////////////////////////////////////////////////////////////////////

function TOM_ParCaisse.DeleteClavier(Caisse: string): Boolean;
begin
  Result := False;
  if Caisse = '' then Exit;
  Result := (ExecuteSQL('DELETE FROM CLAVIERECRAN WHERE CE_CAISSE="' + Caisse + '"') >= 0);
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ReCopieClavier : recopie le paramétrage du clavier d'une caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.ReCopieClavier(CaisseDest, CaisseOrg: string);
var TOBC, TOBD, TOBL: TOB;
  Ind: Integer;
begin
  if (CaisseDest = '') or (CaisseOrg = '') then Exit;
  DeleteClavier(CaisseDest);
  TOBC := TOB.Create('Clavier', nil, -1);
  TOBC.LoadDetailDB('CLAVIERECRAN', '"' + CaisseOrg + '"', '', nil, False);
  TOBD := TOB.Create('Clavier', nil, -1);
  for Ind := 0 to TOBC.Detail.Count - 1 do
  begin
    TOBL := TOB.Create('CLAVIERECRAN', TOBD, -1);
    TOBL.PutValue('CE_CAISSE', CaisseDest);
    TOBL.PutValue('CE_NUMERO', TOBC.Detail[Ind].GetValue('CE_NUMERO'));
    TOBL.PutValue('CE_PAGE', TOBC.Detail[Ind].GetValue('CE_PAGE'));
    TOBL.PutValue('CE_CAPTION', TOBC.Detail[Ind].GetValue('CE_CAPTION'));
    TOBL.PutValue('CE_COULEUR', TOBC.Detail[Ind].GetValue('CE_COULEUR'));
    TOBL.PutValue('CE_POLICE', TOBC.Detail[Ind].GetValue('CE_POLICE'));
    TOBL.PutValue('CE_IMAGE', TOBC.Detail[Ind].GetValue('CE_IMAGE'));
    TOBL.PutValue('CE_TOUCHE', TOBC.Detail[Ind].GetValue('CE_TOUCHE'));
    TOBL.PutValue('CE_TEXTE', TOBC.Detail[Ind].GetValue('CE_TEXTE'));
    TOBL.PutValue('CE_ALLERA', TOBC.Detail[Ind].GetValue('CE_ALLERA'));
  end;
  TOBC.Free;
  TOBD.InsertDB(nil, True);
  TOBD.Free;
end;

///////////////////////////////////////////////////////////////////////////////////////
//  ReCopieCaisse : recopie le paramétrage d'une caisse
///////////////////////////////////////////////////////////////////////////////////////

procedure TOM_ParCaisse.ReCopieCaisse;
///////////////////////////////////////////////////////////////////////////////////////
  procedure CopieChamp(TOBC: TOB; Champ: string; ChampVide: Boolean);
  var
    Valeur: variant;
  begin
    if (not ChampVide) or (FOTestChampVide(Champ, GetField(Champ))) then
    begin
      Valeur := TOBC.GetValue(Champ);
//      if (VarIsEmpty(Valeur)) or (VarIsNull(Valeur)) then // JTR - Enlever VarIsNull
      if VarIsEmpty(Valeur) then
        Valeur := FOChampValeurVide(Champ);
      SetField(Champ, Valeur);
    end;
  end;
  ///////////////////////////////////////////////////////////////////////////////////////
var
  Retour, CaisseOrg, Pwd, sCaisse: string;
  ChampVide: Boolean;
  TOBOrg: TOB;
begin
  Retour := AGLLanceFiche('MFO', 'COPIECAISSE', '', '', GetControlText('GPK_CAISSE'));
  if Retour = '' then Exit;
  if not (DS.State in [dsEdit, dsInsert]) then DS.Edit;
  // GPK_CAISSE : caisse d'origine
  CaisseOrg := ReadTokenSt(Retour);
  TOBOrg := TOB.Create('PARCAISSE', nil, -1);
  TOBOrg.SelectDB('"' + CaisseOrg + '"', nil);
  if TOBOrg.GetValue('GPK_CAISSE') = '' then Exit;
  // CHXCHPVIDE : recopie uniquement des champs vides
  ChampVide := (ReadTokenSt(Retour) = 'X');
  // CHXGENERAL : recopie du paramétrages des généralités
  if ReadTokenSt(Retour) = 'X' then
  begin
    CopieChamp(TOBOrg, 'GPK_LIBELLE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_ETABLISSEMENT', ChampVide);
    CopieChamp(TOBOrg, 'GPK_DEPOT', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TIERS', ChampVide);
    CopieChamp(TOBOrg, 'GPK_FERME', ChampVide);
    CopieChamp(TOBOrg, 'GPK_CLISAISIE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_CLIREPRISE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_CLISAISIENOM', ChampVide);
    CopieChamp(TOBOrg, 'GPK_SEUILCLIOBLIG', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PREPOSGRID', ChampVide);
    CopieChamp(TOBOrg, 'GPK_LIGNESUIVAUTO', ChampVide);
    CopieChamp(TOBOrg, 'GPK_GERETICKETATT', ChampVide);
    CopieChamp(TOBOrg, 'GPK_APPELPRIXTIC', ChampVide);
    CopieChamp(TOBOrg, 'GPK_NBJANNULTIC', ChampVide);
  end;
  // CHXPERIPH : recopie du paramétrage des périphériques
  if ReadTokenSt(Retour) = 'X' then
  begin
    CopieChamp(TOBOrg, 'GPK_PRTTYPE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PRTPORT', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PRTBAUDS', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PRTNBBIT', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PRTSTOPBIT', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PRTPARITE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PRTCTRLFLUX', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PRTMODEBIDI', ChampVide);
    CopieChamp(TOBOrg, 'GPK_REMPLIRCHQ', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TIROIRTYPE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TIROIRPORT', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TIROIRPARAM', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TIROIRPIN5', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TIROIRESP', ChampVide);
    CopieChamp(TOBOrg, 'GPK_AFFTYPE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_AFFPORT', ChampVide);
    CopieChamp(TOBOrg, 'GPK_AFFPARAM', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TPETYPE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TPEPORT', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TPEBAUDS', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TPEREPERTOIRE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_CLAVIERPISTE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PRTFISCTYPE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PRTFISCPORT', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PRTFISCENPLUS', ChampVide);
    CopieChamp(TOBOrg, 'GPK_SIGNPADTYPE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_SIGNERAPRIORI', ChampVide);
    InitParamPeripherique;
  end;
  // CHXPREF : recopie du paramétrage des préférences
  if ReadTokenSt(Retour) = 'X' then
  begin
    CopieChamp(TOBOrg, 'GPK_VENDSAISIE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_VENDSAISLIG', ChampVide);
    CopieChamp(TOBOrg, 'GPK_VENDOBLIG', ChampVide);
    CopieChamp(TOBOrg, 'GPK_VENDREPRISE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_VENDLIGREP', ChampVide);
    CopieChamp(TOBOrg, 'GPK_VENDDEFAUT', ChampVide);
    CopieChamp(TOBOrg, 'GPK_MDPRENDU', ChampVide);
    CopieChamp(TOBOrg, 'GPK_MDPECART', ChampVide);
    CopieChamp(TOBOrg, 'GPK_LIAISONREG', ChampVide);
    CopieChamp(TOBOrg, 'GPK_MODEDETAXE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_CLIOBLIGDETAXE', ChampVide);
  end;
  // CHXREM : recopie du paramétrage des remises
  if ReadTokenSt(Retour) = 'X' then
  begin
    CopieChamp(TOBOrg, 'GPK_REMSAISIE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_REMAFFICH', ChampVide);
    CopieChamp(TOBOrg, 'GPK_REMARRONDI', ChampVide);
    CopieChamp(TOBOrg, 'GPK_DEMSAISIE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_DEMOBLIG', ChampVide);
    CopieChamp(TOBOrg, 'GPK_MAJORPRX', ChampVide);
    CopieChamp(TOBOrg, 'GPK_REMPIEDLIG', ChampVide);
    CopieChamp(TOBOrg, 'GPK_AFFPRIXTIC', ChampVide);
    CopieChamp(TOBOrg, 'GPK_REMPOURMAX', ChampVide);
    CopieChamp(TOBOrg, 'GPK_REMPOURMAX2', ChampVide);
    CopieChamp(TOBOrg, 'GPK_REMPOURMAX3', ChampVide);
    CopieChamp(TOBOrg, 'GPK_PWDREM2', ChampVide);
    Pwd := DeCryptageSt(GetField('GPK_PWDREM2'));
    SetControlText('PWDREM2', Pwd);
    SetControlText('CONFPWDREM2', Pwd);
    CopieChamp(TOBOrg, 'GPK_PWDREM3', ChampVide);
    Pwd := DeCryptageSt(GetField('GPK_PWDREM3'));
    SetControlText('PWDREM3', Pwd);
    SetControlText('CONFPWDREM3', Pwd);
  end;
  // CHXAFFICH : recopie du paramétrage de l'affichage
  if ReadTokenSt(Retour) = 'X' then
  begin
    CopieChamp(TOBOrg, 'GPK_AFFICHEUR', ChampVide);
    CopieChamp(TOBOrg, 'GPK_AFFINVERSE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_AFFMESG', ChampVide);
    CopieChamp(TOBOrg, 'GPK_CLAVIERECRAN', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TOOLBAR', ChampVide);
    CopieChamp(TOBOrg, 'GPK_INFOARTICLE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_INFOARTDIM', ChampVide);
    CopieChamp(TOBOrg, 'GPK_AFFPHOTOART', ChampVide);
    CopieChamp(TOBOrg, 'GPK_INFOS', ChampVide);
  end;
  // CHXIMP : recopie du paramétrage de l'impression
  if ReadTokenSt(Retour) = 'X' then
  begin
    CopieChamp(TOBOrg, 'GPK_IMPFMTTIC', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODTIC', ChampVide);
    CopieChamp(TOBOrg, 'GPK_NBEXEMPTIC', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODCHEQUE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODTICX', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODTICZ', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODRECAPV', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODLISREG', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODSTAFAM', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODSTAREM', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODLISART', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODBONARR', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODBONAV', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODBONACH', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODBONTID', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODBONCDE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODBONREC', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODREMIS', ChampVide);
    CopieChamp(TOBOrg, 'GPK_IMPMODFDC', ChampVide);
    CopieChamp(TOBOrg, 'GPK_NBEXEMPBON', ChampVide);
  end;
  // CHXCLAVIER : recopie du paramétrage du clavier
  if ReadTokenSt(Retour) = 'X' then
  begin
    CopieChamp(TOBOrg, 'GPK_PARAMSCE', False);
    ClavierOrg := '';
    if DS.State = dsInsert then
    begin
      ClavierOrg := CaisseOrg;
    end else
    begin
      sCaisse := GetField('GPK_CAISSE');
      // recopie du pavé principal
      RecopieClavier(sCaisse, CaisseOrg);
      // recopie du 2ème pavé
      RecopieClavier(FOAlphaCodeNumeric(sCaisse), FOAlphaCodeNumeric(CaisseOrg));
    end;
  end;
  // CHXCOMM : recopie du paramétrage des communications
  if ReadTokenSt(Retour) = 'X' then
  begin
    CopieChamp(TOBOrg, 'GPK_TOXAPPEL1', ChampVide);
    CopieChamp(TOBOrg, 'GPK_TOXAPPEL2', ChampVide);
  end;
  // CHXJOURS : recopie du paramétrage des ouvertures/fermetures des journées
  if ReadTokenSt(Retour) = 'X' then
  begin
    CopieChamp(TOBOrg, 'GPK_GEREFONDCAISSE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_MDPFDCAIS', ChampVide);
    CopieChamp(TOBOrg, 'GPK_FDCAISSE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_MODIFFDCAIS', ChampVide);
    CopieChamp(TOBOrg, 'GPK_REPRISEFDCAIS', ChampVide);
    CopieChamp(TOBOrg, 'GPK_CTRLCAISSE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_CTRLAVEUGLE', ChampVide);
    CopieChamp(TOBOrg, 'GPK_CTRLPIECBIL', ChampVide);
    CopieChamp(TOBOrg, 'GPK_OUVTIROIRFIN', ChampVide);
    CopieChamp(TOBOrg, 'GPK_CTRLECART', ChampVide);
    CopieChamp(TOBOrg, 'GPK_MDPCTRLCAIS', ChampVide);
    CopieChamp(TOBOrg, 'GPK_GEREREMISEBNQ', ChampVide);
    CopieChamp(TOBOrg, 'GPK_OPREMISEBNQ', ChampVide);
    CopieChamp(TOBOrg, 'GPK_GEREEVENT', ChampVide);
  end;
  TOBOrg.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : FOCopieCaisse = recopie le paramétrage d'une caisse
Suite ........ : depuis le script d'une fiche
Suite ........ :  - Parms[0] = Fiche
Mots clefs ... : FO
*****************************************************************}

procedure FOCopieCaisse(parms: array of variant; nb: integer);
var FF: TForm;
  OM: TOM;
begin
  FF := TForm(Longint(Parms[0]));
  if FF is TFFicheListe then OM := TFFicheListe(FF).OM else Exit;
  if OM is TOM_ParCaisse then TOM_ParCaisse(OM).ReCopieCaisse else Exit;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 21/09/2001
Modifié le ... : 21/09/2001
Description .. : FOModificationParLot = modification en série des champs
Suite ........ : d'une table.
Suite ........ :  - Parms[0] = Fiche.
Suite ........ :  - Parms[1] = Nature de la fiche.
Suite ........ :  - Parms[2] = Nom de la fiche pour modifier les éléments un
Suite ........ : par un.
Suite ........ :  - Parms[3] = Nom du champ clé.
Suite ........ :  - Parms[4] = Paramètres.
Mots clefs ... : FO
*****************************************************************}

procedure FOAGLModificationParLot(parms: array of variant; nb: integer);
var FF: TForm;
  TheModifLot: TO_ModifParLot;
begin
  FF := TForm(Longint(Parms[0]));
  if (FF = nil) or not (FF is TFMul) then Exit;
  if (TFMul(FF).FListe.NbSelected = 0) and (not TFMul(FF).FListe.AllSelected) then
  begin
    PGIBox('Aucun élément sélectionné', TFMul(FF).Caption);
    Exit;
  end;
  TheModifLot := TO_ModifParLot.Create;
  TheModifLot.F := TFMul(FF).FListe;
  TheModifLot.Q := TFMul(FF).Q;
  TheModifLot.Titre := TFMul(FF).Caption;
  TheModifLot.Nature := string(parms[1]);
  TheModifLot.FicheAOuvrir := string(parms[2]);
  TheModifLot.FCode := string(parms[3]);
  //TheModifLot.TableName := string(parms[4]); // JTR - Fonction en CWAS
  ModifieEnSerie(TheModifLot, string(parms[4]));
end;

///////////////////////////////////////////////////////////////////////////////////////
//  Initialization
///////////////////////////////////////////////////////////////////////////////////////
initialization
  RegisterClasses([TOM_ParCaisse]);
  RegisterAglProc('FOCopieCaisse', TRUE, 0, FOCopieCaisse);
  RegisterAglProc('FOModificationParLot', TRUE, 5, FOAGLModificationParLot);

end.
