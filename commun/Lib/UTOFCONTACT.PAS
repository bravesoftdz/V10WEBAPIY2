{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/06/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CONTACT ()
Mots clefs ... : TOF;CONTACT
*****************************************************************}
unit UTOFCONTACT;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbtables {BDE}, {$ELSE}uDbxDataSet, {$ENDIF}
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  SaisieList,
  uTableFiltre,
  uTob,
  HTB97,
  ParamSoc
{$IFDEF GRC}
  , EntRT
{$ENDIF}
  ;

type
  TOF_CONTACT = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
    procedure AfficheContact(Sender: TObject);
  private
    TF: TTableFiltre;
    ModifLot: boolean;
    CodeAuxi: string;
    BInsert: TToolbarButton97;
    Action: string;
    type_tiers: string;
    StTiers: string; // renseigné qd appel des contacts depuis GCTIERS (rend les boutons GRC actifs)
  end;

implementation
uses
  wCommuns,
  TiersUtil,
  UtilPGI
  ;

procedure TOF_CONTACT.OnNew;
begin
  inherited;
end;

procedure TOF_CONTACT.OnDelete;
begin
  inherited;
end;

procedure TOF_CONTACT.OnUpdate;
begin
  inherited;
  if not ModifLot then TF.RefreshLignes;
  if ((ctxGRC in V_PGI.PGIContexte) and (StTiers <> '') and ((type_tiers = 'CLI') or (type_tiers = 'PRO')))
{$IFDEF CTI}
          //or (ContexteCti)
{$ENDIF}
  then
  begin
    if TF.TOBFiltre.detail.count > 0 then
    begin
      SetControlEnabled('BACTION', True);
      SetControlEnabled('BCOMPLGRC', True);
    end
    else
    begin
      SetControlEnabled('BACTION', False);
      SetControlEnabled('BCOMPLGRC', False);
    end;
  end;
end;

procedure TOF_CONTACT.OnLoad;
var
  Lequel: string;
  T: TOB;
  Trouve: Boolean;
begin
  inherited;
  Lequel := TFsaisieList(ecran).FLequel;
  if lequel <> '' then
  begin
    T := TF.TobFiltre.FindFirst(['C_NUMEROCONTACT'], [ValeurI(Lequel)], false);
    if T <> nil then
      TF.SelectRecord(TF.TobFiltre.Detail.IndexOf(T) + 1);
  end;

  SetControlVisible('PENTREPRISE', False);
  SetControlVisible('PPARTICULIER', False);
  trouve := ExisteSQL('SELECT T_PARTICULIER FROM TIERS WHERE T_AUXILIAIRE = "' + CodeAuxi + '" AND T_PARTICULIER = "X"');
  if (trouve = False) then
  begin
    SetControlVisible('PENTREPRISE', True);
    SetControltext('TC_TELEPHONE', '&Tél.');
    SetControltext('TC_FAX', 'F&AX');
    { mng 12/10/2004 : je suis obligée de mettre tous les champs du panel, pour le bon fonctionnement
      de la fonction VerifierChampsObligatoires qui teste sur la visibilité des champs }
    SetControlVisible('C_LIPARENT', False);
    //mcd 14/03/2006 saisie systématique date naissance SetControlVisible('C_JOURNAIS',False);
    //mcd 14/03/2006 saisie systématique date naissance SetControlVisible('C_MOISNAIS',False);
    //mcd 14/03/2006 saisie systématique date naissance SetControlVisible('C_ANNEENAIS',False);
    //mcd 14/03/2006 saisie systématique date naissance SetControlVisible('C_SEXE',False);
  end
  else
  begin
    SetControlVisible('PPARTICULIER', True);
    { mng 12/10/2004 : je suis obligée de mettre tous les champs du panel, pour le bon fonctionnement
      de la fonction VerifierChampsObligatoires qui teste sur la visibilité des champs }
    SetControlVisible('C_SERVICE', False);
    SetControlVisible('C_FONCTION', False);
    //mcd 14/03/2006 saisie systématique info mailing SetControlVisible('C_EMAILING',False);
    //mcd 14/03/2006 saisie systématique info mailing SetControlVisible('C_PUBLIPOSTAGE',False);
    SetControlVisible('C_SERVICECODE', False);
    SetControlVisible('C_FONCTIONCODEE', False);
  end;
  if Action = 'CREATION' then //GPAO appel contacts depuis fiche adresse (tablette YYCONTACTIERS)
    Binsert.Click;
  if TF.TOBFiltre.detail.count = 0 then
  begin
    SetControlEnabled('BACTION', False);
    SetControlEnabled('BCOMPLGRC', False);
  end;
  //YMO 12/12/2006 Suppression : FQ18526
{$IFDEF COMPTA}
  SetControlVisible('BCOMPLGRC', False);
{$ENDIF}

{$IFDEF TRESO}
  SetControlVisible('BCOMPLGRC', False);
{$ENDIF}
end;

procedure TOF_CONTACT.OnArgument(S: string);
var x: integer;
  Range: string;
begin
  inherited;
type_tiers := '';
  StTiers := '';
  type_tiers := GetArgumentValue(S, 'TYPE2');
{$IFDEF GRC}
  if (((type_tiers = 'CLI') or (type_tiers = 'PRO')) and VH_RT.RTCreatContacts) or
    ((type_tiers = 'FOU') and (VH_RT.RFCreatContacts)) then TFSaisieList(ecran).TypeAction := TaCreat;
{$ENDIF}
  if TFSaisieList(ecran).TypeAction = TaModif then TFSaisieList(ecran).TypeAction := TaCreat;

  TF := TFSaisieList(Ecran).LeFiltre;
  {JP 05/01/04 : chez moi TF.WhereTable = '', d'où une requête incorrecte}
  if Trim(TF.WhereTable) = '' then
    TF.WhereTable := 'WHERE C_FERME="-"'
  else
    TF.WhereTable := TF.WhereTable + ' and C_FERME="-"';
  TToolbarButton97(GetControl('BALLCONTACT')).OnClick := AfficheContact;
  x := pos('MODIFLOT', S);
  ModifLot := x <> 0;

  Action := GetArgumentValue(S, 'ACTION');
  CodeAuxi := TiersAuxiliaire(GetArgumentValue(S, 'TIERS'));
  if CodeAuxi = '' then
  begin
    Range := TFsaisieList(ecran).FRange;
    CodeAuxi := ReadTokenSt(Range); // Type de contact
    CodeAuxi := ReadTokenSt(Range); // code Auxiliaire
  end;
  BInsert := TToolbarButton97(GetControl('BINSERT'));
  StTiers := GetArgumentValue(S, 'TIERS');
  if (GetParamSocSecur('SO_RTGESTINFOS006', False) = true) then
    TF.LaGridListe := 'RTCONTACTSCP';
end;

procedure TOF_CONTACT.OnClose;
begin
  inherited;
end;

procedure TOF_CONTACT.OnDisplay();
begin
  inherited;
end;

procedure TOF_CONTACT.OnCancel();
begin
  inherited;
  
end;

procedure TOF_Contact.AfficheContact(Sender: TObject);
begin
  if pos('and C_FERME="-"', TF.WhereTable) <> 0 then
  begin
    TF.WhereTable := FindEtReplace(TF.WhereTable, 'and C_FERME="-"', '', True);
    TToolbarButton97(GetControl('BALLCONTACT')).Hint := TraduireMemoire('Contacts non fermés');
  end
  else
  begin
    TF.WhereTable := TF.WhereTable + ' and C_FERME="-"';
    TToolbarButton97(GetControl('BALLCONTACT')).Hint := TraduireMemoire('Tous les contacts');
  end;
  TF.RefreshLignes;
end;

initialization
  registerclasses([TOF_CONTACT]);
end.
