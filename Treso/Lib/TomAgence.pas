{-------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
1.50.001.001  26/04/04   JP    Ajout de OnDeleteRecord : on teste s'il n'y a pas de comptes
                               rattachés à cette agence avant d'autoriser la destruction
6.00.017.001  07/10/04   JP    Création d'un état plutôt que le PRT_AGENCE pour pouvoir afficher
                               les différents libellés
7.00.001.006  16/05/06   JP    FQ 10349 : Avant suppression, test sur toutes les bases d'un
                               éventuel regroupement l'utilisation de l'agence
7.05.001.001  03/10/06   JP    Bloquage de la suppression de l'agence comptes courants en multi sociétés
--------------------------------------------------------------------------------------}
unit TomAgence ;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL, CPCODEPOSTAL_TOF, UtileAGL, eFichList, UTob,
  {$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  FE_Main , CodePost, EdtREtat, FichList, HDB,
  {$ENDIF}
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  StdCtrls, Controls, Classes, SysUtils, HCtrls, HEnt1, UTOM,
  Windows;

type
  TOM_AGENCE = class (TOM)
    procedure OnArgument(S : string); override;
    procedure OnDeleteRecord;         override;
  protected
    Contact : string;

    procedure ContactEnter            (Sender : TObject);
    procedure ContactExit             (Sender : TObject);
    procedure CodeAgenceOnKeyPress    (Sender : TObject; var Key: Char);
    procedure CalendrierOnClick       (Sender : TObject);
    procedure CodePostalOnElipsisClick(Sender : TObject);
    procedure BanqueOnChange          (Sender : TObject);
    procedure ImpressionClick         (Sender : TObject);
  end;

procedure TRLanceFiche_Agence(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

Implementation

uses
  ParamSoc, Commun, Constantes, TomCalendrier, HTB97, UtilPGI;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_Agence(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_AGENCE.OnArgument ( S: String ) ;
{---------------------------------------------------------------------------------------}
begin
  Inherited ;
  Ecran.HelpContext := 500005;
  THEdit(GetControl('TRA_AGENCE'           )).OnKeyPress     := CodeAgenceOnKeyPress;
  THEdit(GetControl('TRA_GUICHET'          )).OnKeyPress     := CodeAgenceOnKeyPress;
  TToolBarButton97(GetControl('BCALENDRIER')).OnClick        := CalendrierOnClick;
  {$IFDEF EAGLCLIENT}
  THEdit(GetControl('TRA_INTERLOCUTEUR'    )).OnEnter        := ContactEnter;
  THEdit(GetControl('TRA_INTERLOCUTEUR'    )).OnExit         := ContactExit;
  THEdit(GetControl('TRA_CODEPOSTAL'       )).OnElipsisClick := CodePostalOnElipsisClick;
  THEdit(GetControl('TRA_VILLE'            )).OnElipsisClick := CodePostalOnElipsisClick;
  THValComboBox(GetControl('TRA_BANQUE'    )).OnClick        := BanqueOnChange;
  {$ELSE}
  THDBEdit(GetControl('TRA_INTERLOCUTEUR'  )).OnEnter        := ContactEnter;
  THDBEdit(GetControl('TRA_INTERLOCUTEUR'  )).OnExit         := ContactExit;
  THDBEdit(GetControl('TRA_CODEPOSTAL'     )).OnElipsisClick := CodePostalOnElipsisClick;
  THDBEdit(GetControl('TRA_VILLE'          )).OnElipsisClick := CodePostalOnElipsisClick;
  THDBValComboBox(GetControl('TRA_BANQUE'  )).OnClick        := BanqueOnChange;
  {$ENDIF}
  TFFicheListe(Ecran).BImprimer.OnClick := ImpressionClick;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_AGENCE.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {JP 03/10/06 : Interdiction de la suppression de la banque des comptes courants si Tréso multi sociétés}
  if EstMultiSoc and (GetParamSocSecur('SO_TRBASETRESO', '') <> '') and
     (VarToStr(GetField('TRA_AGENCE')) = CODECOURANTS) then begin
    LastError := 1;
    LastErrorMsg := TraduireMemoire('Cette agence bancaire est obligatoire pour la gestion des comptes courants : '#13 +
                                    'vous ne pouvez pas la supprimer.');
  end
  {26/04/04 : on s'assure que l'agence en cours de suppression n'est pas utilisée par un compte
   16/05/06 : FQ 10349 : gestion du multisoc car pour le moment BANQUECP n'est pas partagée}
  else if PresenceMS('BANQUECP', 'BQ_AGENCE', VarToStr(GetField('TRA_AGENCE'))) then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Cette agence bancaire est référencée par un compte bancaire, vous ne pouvez pas la supprimer.');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_AGENCE.ImpressionClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  LanceEtat('E', 'ECT', 'AGE', True, False, False, nil, '', 'Liste des agences', False);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_AGENCE.ContactEnter(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  Contact := GetControlText('TRA_INTERLOCUTEUR');
end;

{---------------------------------------------------------------------------------------}
procedure TOM_AGENCE.ContactExit(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  if Contact <> GetControlText('TRA_INTERLOCUTEUR') then begin
    Q := OpenSQL('SELECT C_FAX, C_TELEPHONE, C_CIVILITE, C_RVA FROM CONTACT WHERE C_TYPECONTACT="TRE" ' +
                 'AND C_NUMEROCONTACT = "' + GetControlText('TRA_INTERLOCUTEUR') + '"', True);
    if not Q.EOF then begin
      SetField('TRA_FAX', Q.Fields[0].AsString);
      SetField('TRA_TELEPHONE', Q.Fields[1].AsString);
      SetField('TRA_CIVILITE', Q.Fields[2].AsString);
      SetField('TRA_EMAIL', Q.Fields[3].AsString);
    end;
    Ferme(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_AGENCE.CodeAgenceOnKeyPress(Sender: TObject; var Key: Char);
{---------------------------------------------------------------------------------------}
begin
  {JP 05/08/03 : Ce code est nécessaire pour la version CWAS. En effet mettre '00000' comme EditMask
                 dans decla est interprété en CWAS comme le formatage d'une date !!!!}
  if AnsiUpperCase(TComponent(Sender).Name) = 'TRA_GUICHET' then begin
    if not (Key in ['0'..'9', chr(VK_BACK), chr(VK_RIGHT), chr(VK_LEFT), chr(VK_DELETE)]) then
      Key := #0;
  end else
    ValidCodeOnKeyPress(Key); {Validation du code agence}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_AGENCE.CalendrierOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_Calendrier('TR','TRCALENDRIER', '', GetControlText('TRA_CODECAL'), 'ACTION=MODIFICATION');
  {$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('TRA_CODECAL')).ReLoad;
  {$ELSE}
  THDBValComboBox(GetControl('TRA_CODECAL')).ReLoad;
  {$ENDIF}

end;

{---------------------------------------------------------------------------------------}
procedure TOM_AGENCE.CodePostalOnElipsisClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  if THEdit(Sender).Name = 'TRA_CODEPOSTAL' then
    VerifCodePostal(DS, THEdit(GetControl('TRA_CODEPOSTAL')), THEdit(GetControl('TRA_VILLE')), True)
  else
    VerifCodePostal(DS, THEdit(GetControl('TRA_CODEPOSTAL')), THEdit(GetControl('TRA_VILLE')), False);
  {$ELSE}
  if THDBEdit(Sender).Name = 'TRA_CODEPOSTAL' then
    VerifCodePostal(DS, THDBEdit(GetControl('TRA_CODEPOSTAL')), THDBEdit(GetControl('TRA_VILLE')), True)
  else
    VerifCodePostal(DS, THDBEdit(GetControl('TRA_CODEPOSTAL')), THDBEdit(GetControl('TRA_VILLE')), False);
  {$ENDIF}
end;

{Maj du code bancaire
{---------------------------------------------------------------------------------------}
procedure TOM_AGENCE.BanqueOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT PQ_ETABBQ FROM BANQUES WHERE PQ_BANQUE = "' + THValComboBox(Sender).Value + '"', True);
  if not Q.EOF then
    SetControlText('TRA_ETABBQ', Q.Fields[0].AsString);
  Ferme(Q);
end;

initialization
  RegisterClasses ( [ TOM_AGENCE ] ) ;

end.

