{***********UNITE*************************************************
Auteur  ...... : GF
Créé le ...... : 13/03/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : HREVENEMENT (HREVENEMENT)
Mots clefs ... : TOM;HREVENEMENT
*****************************************************************}
unit UTOMBTEVENEMENT;

interface
Uses StdCtrls,
     Controls,
     Classes,
  {$IFDEF EAGLCLIENT}
     MaineAGL,
     EFiche,
     UtileAGL,
  {$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     Fe_main,
  {$ENDIF}
     MsgUtil,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOM,
     UTob ;

type
  TOM_HREVENEMENT = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;

  private

	NomForm	: String;

  //Gestion des boutons
  BCOULEUR	: TToolbarButton97;

  LCOULEUR	: THEdit;

  procedure BCouleur_OnClick(Sender: TObject);

  end ;

implementation

procedure TOM_HREVENEMENT.OnNewRecord;
begin
  inherited;

  Setfield('HEV_COULEUR',IntToStr(0));
  AfficheCouleur(LCOULEUR, THEdit(GetControl('HEV_COULEUR')));

  Setfield('HEV_DATEDEBUT', V_PGI.DateEntree);
  Setfield('HEV_DATEFIN', V_PGI.DateEntree);

end;

procedure TOM_HREVENEMENT.OnDeleteRecord;
begin
  inherited;
end;

procedure TOM_HREVENEMENT.OnUpdateRecord;
begin
  inherited;

  // Test désignation non vide
  if GetField('HEV_EVENEMENT') = '' then
     begin
     AfficheErreur(NomForm, '1', 'Fiche évènementiel');
     exit;
     end;

  // Test dates valides
  if GetField('HEV_DATEDEBUT') > GetField('HEV_DATEFIN') then
     begin
     SetFocusControl('HEV_DATEDEBUT');
     AfficheErreur(NomForm, '2', 'Fiche évènementiel');
     exit;
     end;

  SetField('HEV_COULEUR', LCOULEUR.Color)

end;

procedure TOM_HREVENEMENT.OnAfterUpdateRecord;
begin
  inherited;
end;

procedure TOM_HREVENEMENT.OnLoadRecord;
begin
  inherited;

  // Chargement Couleur
  AfficheCouleur(LCOULEUR,THEdit(GetControl('HEV_COULEUR')));

end;

procedure TOM_HREVENEMENT.BCOULEUR_OnClick(Sender: TObject);
begin

  SelColorNew(LCOULEUR,THEdit(GetControl('HEV_COULEUR')), TForm(Ecran));

end;

procedure TOM_HREVENEMENT.OnChangeField(F: TField);
begin
  inherited;
end;

procedure TOM_HREVENEMENT.OnArgument(S: string);
begin
  inherited;

  NomForm 	:= ecran.Name;

  BCouleur 	:= TToolbarButton97(ecran.FindComponent('BCOULEUR'));
  BCouleur.onclick := BCouleur_OnClick;

  LCouleur 	:= THEdit(ecran.FindComponent('LCOULEUR'));

end;

procedure TOM_HREVENEMENT.OnClose;
begin
  inherited;
end;

procedure TOM_HREVENEMENT.OnCancelRecord;
begin
  inherited;
end;

initialization
  registerclasses([TOM_HREVENEMENT]);
end.
