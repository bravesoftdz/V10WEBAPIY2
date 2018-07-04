{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/10/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TDI_GESTIONCOMPTES_JDC ()
Mots clefs ... : TOF;TDI_GESTIONCOMPTES_JDC
*****************************************************************}
unit TDI_GESTIONCOMPTES_JDC_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbtables, {$ELSE}uDbxDataSet, {$ENDIF}
  mul,
{$ELSE}
  eMul,
//  uTob,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  UTob,
  HTB97, //TToolbarButton
//  Tdi_LibFonction; // chargerlisteEmetteur
  Paie_cjdc_lib; // chargerlisteEmetteur

type
  TOF_TDI_GESTIONCOMPTES_JDC = class(TOF)

    Grille: THGrid;
    GrilleU: THGrid;

    Combo_Emetteur: THValComboBox;
    Combo_Emetteur1: THValComboBox;
    Combo_EmetteurU: THValComboBox;
    Combo_Teleprocedures: THValComboBox;
    Combo_Utilisateurs: THValComboBox;
    Combo_Comptes_JDC: THValComboBox;

    OngletComptes: TTabSheet;
    OngletAffectations: TTabSheet;

    Compte: TEdit;
    MotdePasse: TEdit;

    BVALIDER: TToolbarButton97;
    BVALIDER1: TToolbarButton97;
    BINSERT: TToolbarButton97;
    BINSERT1: TToolbarButton97;
    BDELETE: TToolbarButton97;
    BDELETE1: TToolbarButton97;
    BFERME: TToolbarButton97;

    SaisieModif: TGroupBox;
    LCOMPTE: THLABEL;
    LTELEPROCEDURE: THLABEL;
    LMOTDEPASSE: THLABEL;
    LUTILISATEUR: THLABEL;
    LCOMPTEU: THLABEL;

    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
    procedure OnEnterOngletComptes(Sender: TObject);
    procedure OnEnterOngletAffectations(Sender: TObject);
    procedure SetStatutSaisieCompte(Statut: boolean);
    procedure SetStatutSaisieAffectation(Statut: boolean);

    procedure OnNouveauCompte(Sender: TObject);
    procedure OnValiderCompte(Sender: TObject);
    procedure OnChangeEmetteur(Sender: TObject);
    procedure OnChangeEmetteur1(Sender: TObject);
    procedure InitialiserGrilleComptes(Sender: TObject);
    procedure OnClickGrille(Sender: TObject);
    procedure OnSupprimer(Sender: TObject);

    procedure OnNouvelleAffectation(Sender: TObject);
    procedure OnSupprimerAffectation(Sender: TObject);
    procedure OnClickAffectations(Sender: TObject);
    procedure OnChangeEmetteurU(Sender: TObject);
    procedure OnValiderAffectation(Sender: TObject);

  private
      sNomEmetteur: string;
  public
    sColonne: string;
    sColonneSuite: string;
    J, JC: Tob;
    Q: TQuery;
//    sNomEmetteur: string;
    sSiretEmetteur: string;
    sNomUtilisateur: string;
    sCodeUtilisateur: string;
    sNomconnexion: string;
    ValuesProc: HTStrings; // values combo téléprocédures   @@ MF
    ItemsEmetteur1: HTStrings; // values combo émetteur @@ MF
    EmetteurSocial  : string; // @@ MF
    SiretEmetteurSocial : string; // @@ MF
    Social : boolean; // @@ MF
  end;

function IdentifiantWS(NomConnexion, MotDePasse: string): boolean;


implementation

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnNew;
begin
  inherited;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnDelete;
begin
  inherited;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnUpdate;
begin
  inherited;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnLoad;
begin
  inherited;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnArgument(S: string);
// d @@ MF
var
  st      : string;
// f @@ MF
begin
  inherited;
// d @@ MF
    Social := False;
    st := (trim(S));
    if (ReadTokenst(S) = 'SOCIAL') then
      Social := True;
    EmetteurSocial := ReadTokenst(S);
    SiretEmetteurSocial := ReadTokenst(S);
    if (Social) then
      SetcontrolVisible ('AFFECTATION', false);
// f @@ MF
//  ComboEmetteur := THVALCOMBOBOX(GetControl('EMETTEUR'));
//  if assigned(ComboEmetteur) then

  GRILLE := THGRID(GetControl('GRILLE'));
  GRILLE.OnClick := OnClickGrille;

  GRILLEU := THGRID(GetControl('GRILLEU'));
  GRILLEU.OnClick := OnClickAffectations;

  Combo_Emetteur := THVALCOMBOBOX(GetControl('COMBO_EMETTEUR'));
  Combo_Emetteur1 := THVALCOMBOBOX(GetControl('COMBO_EMETTEUR1'));
  Combo_EmetteurU := THVALCOMBOBOX(GetControl('COMBO_EMETTEURU'));
  Combo_Teleprocedures := THVALCOMBOBOX(GetControl('COMBO_TELEPROCEDURES'));
  ValuesProc := Combo_teleprocedures.Values;

  Combo_Utilisateurs := THVALCOMBOBOX(GetControl('COMBO_UTILISATEURS'));
  Combo_Comptes_JDC := THVALCOMBOBOX(GetControl('COMBO_COMPTES_JDC'));

  BVALIDER := TToolbarButton97(GetControl('BVALIDER'));
  if assigned(BVALIDER) then
    BVALIDER.OnClick := OnValiderCompte;
  BVALIDER1 := TToolbarButton97(GetControl('BVALIDER1'));
  if assigned(BVALIDER1) then
    BVALIDER1.OnClick := OnValiderAffectation;

  BINSERT := TToolbarButton97(GetControl('BINSERT'));
  BINSERT1 := TToolbarButton97(GetControl('BINSERT1'));
  BDELETE := TToolbarButton97(GetControl('BDELETE'));
  BDELETE.OnClick := OnSupprimer;

  BDELETE1 := TToolbarButton97(GetControl('BDELETE1'));
  BDELETE1.OnClick := OnSupprimerAffectation;

  BFERME := TToolbarButton97(GetControl('BFERME'));

  LCOMPTE := THLABEL(GetControl('LCOMPTE'));
  LTELEPROCEDURE := THLABEL(GetControl('LTELEPROCEDURE'));
  LMOTDEPASSE := THLABEL(GetControl('LMOTDEPASSE'));
  LUTILISATEUR := THLABEL(GetControl('LUTILISATEUR'));
  LCOMPTEU := THLABEL(GetControl('LCOMPTEU'));



  OngletComptes := TTabSheet(GetControl('COMPTESJDC'));
  if assigned(OngletComptes) then
    OngletComptes.OnShow := OnEnterOngletComptes;
  OngletAffectations := TTabSheet(GetControl('AFFECTATION'));
  if assigned(OngletAffectations) then
    OngletAffectations.OnShow := OnEnterOngletAffectations;

  if assigned(BINSERT) then
    BINSERT.OnClick := OnNouveauCompte;
  if assigned(BINSERT1) then
    BINSERT1.OnClick := OnNouvelleAffectation;

  Compte := TEdit(GetControl('Compte'));
  MotdePasse := TEdit(GetControl('MOTDEPASSE'));


  if assigned(Combo_Emetteur) then
  begin
    Combo_Emetteur.OnChange := OnChangeEmetteur;
// d @@ MF
    if (Social) then
      ChargerListeEmetteurs(Combo_Emetteur, EmetteurSocial)
    else
// f @@ MF
      ChargerListeEmetteurs(Combo_Emetteur, '<Tous>');
  end;
  if assigned(Combo_Emetteur1) then
  begin
    Combo_Emetteur1.OnChange := OnChangeEmetteur1;
// d @@ MF
    if (Social) then
      ChargerListeEmetteurs(Combo_Emetteur1, EmetteurSocial)
    else
// f @@ MF
      ChargerListeEmetteurs(Combo_Emetteur1, '');
    ItemsEmetteur1 := Combo_Emetteur1.Items;
  end;
  //

  sColonne := 'NJC_NOMEMETTEUR; NJC_NOMCONNEXION; NOM_LISTEPROCEDURE; NJC_LISTEPROCEDURE;';
  sColonneSuite := 'NOM_UTILISATEUR; YX_LIBELLE; YX_CODE;';
 //
// d @@ MF
  if (not Social) then
  begin
    if assigned(Combo_EmetteurU) then
    begin
      Combo_EmetteurU.OnChange := OnChangeEmetteurU;
      ChargerListeEmetteurs(Combo_EmetteurU, '');
      Combo_EmetteurU.Text := '';
      OnChangeEmetteurU(nil);
    end;
  end;
// f @@ MF

  SaisieModif := TGroupBox(GetControl('SAISIEMODIF'));

  MotDePasse.Text := '';
  Combo_Emetteur.Text := '';
  Combo_EmetteurU.Text := '';
  Combo_Teleprocedures.Text := '';
  MotdePasse.Text := '';

  InitialiserGrilleComptes(nil);
  OnChangeEmetteur(nil);
  OnEnterOngletComptes(nil);
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnClose;
begin
  inherited;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnDisplay();
begin
  inherited;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnCancel();
begin
  inherited;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnEnterOngletComptes(Sender: TObject);
begin
  inherited;
  BInsert.Visible := TRUE;
  BInsert1.Visible := FALSE;
  // seront réactivés sur sélection ou saisie
  BValider.Visible := FALSE;
  BValider1.Visible := FALSE;
  BDelete.Visible := FALSE;
  BDelete1.Visible := FALSE;
  SetStatutSaisieCompte(FALSE);
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnEnterOngletAffectations(Sender: TObject);
begin
  inherited;
  BInsert1.Visible := TRUE;
  BInsert.Visible := FALSE;
  // seront réactivés sur sélection ou saisie
  BValider.Visible := FALSE;
  BValider1.Visible := FALSE;
  BDelete.Visible := FALSE;
  BDelete1.Visible := FALSE;
  SetStatutSaisieAffectation(False);
end;



procedure TOF_TDI_GESTIONCOMPTES_JDC.SetStatutSaisieCompte(Statut: boolean);
begin

  SaisieModif.Visible := Statut;
  Compte.Visible := Statut;
  Combo_Emetteur1.Visible := Statut;
  Combo_TeleProcedures.Visible := Statut;
  MotDePasse.Visible := Statut;
  LCOMPTE.Visible := Statut;
  LTELEPROCEDURE.Visible := Statut;
  LMOTDEPASSE.Visible := Statut;
  BValider.Visible := Statut;
  BDELETE.Visible := Statut;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.SetStatutSaisieAffectation(Statut: boolean);
begin
  Combo_Utilisateurs.Enabled := Statut;
  Combo_Utilisateurs.Visible := Statut;
  Combo_Comptes_JDC.Visible := Statut;
  LUTILISATEUR.Visible := Statut;
  LCOMPTEU.Visible := Statut;
  BValider1.Visible := Statut;
  BDELETE1.Visible := FALSE;  // en création masquer suppression
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnNouveauCompte(Sender: TObject);
begin
  SetStatutSaisieCompte(TRUE);
  Combo_Emetteur1.Enabled := TRUE; // permettre saisie

  BValider.Visible := TRUE;
  Compte.Text := '';
  Compte.SetFocus;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnNouvelleAffectation(Sender: TObject);
begin
  SetStatutSaisieAffectation(TRUE);
  Combo_Utilisateurs.setFocus;
  BValider1.Visible := TRUE;
  BDELETE1.Visible := FALSE;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnValiderCompte(Sender: TObject);
var
  i, nbLignes: integer;
  sChampIn: string;
  TobCur: Tob;
  Statut: boolean;
  sSql, sSiret: string;
begin
// si ligne existante, modifier dans la TOB et rafraichir
// sinon ajouter dans la tob et rafraichir.
  Statut := TRUE;
  sSql := Combo_Emetteur1.Value; //requete SQL AND ...="......"
  sSiret := ReadTokenPipe(sSql, '=');
  sSiret := sSql;
  sSiret := stringReplace(sSiret, '"', '', [rfReplaceAll]);
  sSiret := stringReplace(sSiret, ' ', '', [rfReplaceAll]);
  if (Social) then  sSiret :=  SiretEmetteurSocial; // @@ MF
  if (sSiret = '') then
  begin
    Statut := FALSE;
    PgiInfo(' Vous devez choisir le site émetteur titulaire', 'Enregistrement d''un compte JDC');
    Combo_Emetteur.SetFocus;
  end;
  if IdentifiantWS(Compte.Text, MotDePasse.Text) = FALSE then
  begin
    Statut := FALSE;
    Compte.SetFocus;
  end;

  if (Statut = TRUE) then
  begin
    i := 0;
    nbLignes := J.Detail.Count;

    while i < nbLignes do
    begin
      TobCur := J.Detail[i];
      sChampIn := TobCur.GetValue('NJC_NOMCONNEXION');
      if sChampIn = Compte.Text then break
      else
        inc(i);
    end;

    if i = nbLignes then
    begin
     // non trouvé ajouter
      TobCur := Tob.create('NJDC_COMPTES', J, -1);
//      TobCur.AddChampSup('NJC_INTITULE', FALSE);
//      TobCur.AddChampSup('NJC_NOMCONNEXION', FALSE);
//      TobCur.AddChampSup('NJC_PASSWORD', FALSE);
      TobCur.AddChampSup('NJC_NOMEMETTEUR', FALSE);

      TobCur.PutValue('NJC_INTITULE', sSiret);
      TobCur.PutValue('NJC_NOMCONNEXION', Compte.Text);
      TobCur.PutValue('NJC_PASSWORD', MotDePasse.Text);
      TobCur.PutValue('NJC_NOMEMETTEUR', Combo_Emetteur.Text);
      TobCur.PutValue('NJC_LISTEPROCEDURE', Combo_Teleprocedures.value);
      TobCur.InsertDB(nil);
      J.PutGridDetail(GRILLE, False, False, 'NJC_NOMEMETTEUR; NJC_NOMCONNEXION; NOM_LISTEPROCEDURE;', True);
    end
    else
    begin
     // modification
      TobCur.PutValue('NJC_INTITULE', sSiret);
      TobCur.PutValue('NJC_NOMCONNEXION', Compte.Text);
      TobCur.PutValue('NJC_PASSWORD', MotDePasse.Text);

      sSql := 'UPDATE NJDC_COMPTES SET';
      sSql := sSql + ' NJC_INTITULE="' + sSiret + '",';
      sSql := sSql + ' NJC_PASSWORD="' + MotDePasse.Text + '",';
      sSql := sSql + ' NJC_LISTEPROCEDURE="' + Combo_Teleprocedures.Value + '"';
      sSql := sSql + ' WHERE NJC_NOMCONNEXION="' + Compte.Text + '"';
      ExecuteSql(sSql);
//   TobCur.UpdateDB(FALSE);   pas compris pourquoi ne fonctione pas
      TobCur.PutValue('NJC_NOMEMETTEUR', Combo_Emetteur.Text);
      J.PutGridDetail(GRILLE, False, False, 'NJC_NOMEMETTEUR; NJC_NOMCONNEXION; NOM_LISTEPROCEDURE;', True);
    end;

    Grille.setFocus;

    Compte.Text := '';
    MotDePasse.Text := '';
    if statut = TRUE then
    begin
      InitialiserGrilleComptes(nil); //rustique mais ça marche
      SetStatutSaisieCompte(FALSE);
      BValider.Visible := FALSE;
    end;
  end;
end;


procedure TOF_TDI_GESTIONCOMPTES_JDC.OnValiderAffectation(Sender: TObject);
var TobCur: Tob;
  sChampIn: string;
  i, nbLignes: integer;
  Values : HtStrings; // @@ MF
begin
  sNomUtilisateur := Combo_Utilisateurs.Text;
  sNomConnexion := Combo_Comptes_JDC.Text;
  sCodeUtilisateur := Combo_Utilisateurs.value;
  values := Combo_Utilisateurs.values;
  
  nbLignes := JC.Detail.count;
  i := 0;
  while i < nbLignes do
  begin
    TobCur := JC.Detail[i];
    sChampIn := TobCur.GetValue('NOM_UTILISATEUR');
    if sChampIn = sNomUtilisateur then
    begin
      i := -1;
      break;
    end;
    inc(i);
  end;
  if (i = -1) then
  begin
    //PgiInfo('a déjà une connexion de définie', 'L''utilisateur ' + SNomUtilisateur);
    ExecuteSql('DELETE ##DP##.CHOIXEXT WHERE YX_TYPE="TDC" AND YX_CODE="' + sSiretEmetteur + sCodeUtilisateur + '"');
    TobCur.PutValue('YX_LIBELLE', sNomConnexion);
    EcrireTabletteConfig('TDC', sSiretEmetteur + sCodeUtilisateur, sNomConnexion, '', '');
    // pour affichage
    JC.PutGridDetail(GRILLEU, false, false, 'NOM_UTILISATEUR; YX_LIBELLE;', TRUE);
  end
  else
  begin
   // ajouter dans les connexions
    if (sNomUtilisateur <> '') and (sNomConnexion <> '') then
    begin
      nbLignes := JC.Detail.count;
      TobCur := Tob.Create('CHOIXEXT', JC, -1);
      TobCur.AddChampSup('NOM_UTILISATEUR', FALSE);
      TobCur.PutValue('YX_TYPE', 'TDC');
      TobCur.PutValue('YX_CODE', sSiretEmetteur + sCodeUtilisateur);
      TobCur.PutValue('NOM_UTILISATEUR', sNomUtilisateur);
      TobCur.PutValue('YX_LIBELLE', sNomConnexion);
{
   pas possiblr d'utiliser la TOB en forçant ecriture en DB0
      TobCur.InsertOrUpdateDB(False);
}
      EcrireTabletteConfig('TDC', sSiretEmetteur + sCodeUtilisateur, sNomConnexion, '', '');
    // pour affichage
      JC.PutGridDetail(GRILLEU, false, false, 'NOM_UTILISATEUR; YX_LIBELLE;', TRUE);
//    TobCur.free;
    end
    else
      PgiInfo('Vous devez choisir un utilisateur et un compte JDC', 'Saisie incomplète');
  end;
 // ajouter une ligne dans la grille
  SetStatutSaisieAffectation(FALSE);
end;

function IdentifiantWS(NomConnexion, MotDePasse: string): boolean;
begin
  Result := FALSE;
  if (NomConnexion <> '') and (MotDepasse <> '') then
  begin
    if pos('@', NomConnexion) <> 0 then
      PgiInfo('ne pas mettre @....', 'dans le nom de connexion')
    else
      Result := True;
  end
  else
    PgiInfo('Voir: Paramétrage du site, Onglet Web Services', 'Le nom de connexion ou (et) le mot de passe est absent');

end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnChangeEmetteur(Sender: TObject);
begin
  sNomEmetteur := Combo_Emetteur.text;
  InitialiserGrilleComptes(nil);
  SetStatutSaisieCompte(FALSE);
  Grille.setFocus;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnChangeEmetteur1(Sender: TObject);
begin
  sNomEmetteur := Combo_Emetteur.text;
end;


procedure TOF_TDI_GESTIONCOMPTES_JDC.OnChangeEmetteurU(Sender: TObject);
var
  Q, Q1: Tquery;
  sSql, ListeGroupes, ListeCodeUser,unUtilisateur: string;
  i, nbLignes: integer;
  TobCur: Tob;
begin
  // charger la liste des utilisateurs
  SetStatutSaisieAffectation(FALSE);

  if assigned(Combo_Utilisateurs) then
  begin
    Combo_Utilisateurs.Items.Clear;
    Combo_Utilisateurs.Values.Clear;
    Combo_Comptes_JDC.Items.Clear;
    sNomEmetteur := Combo_EmetteurU.text;
    sSql := Combo_EmetteurU.Value; //requete SQL AND ...="......"
    sSiretEmetteur := ReadTokenPipe(sSql, '=');
    sSiretEmetteur := sSql;
    sSiretEmetteur := stringReplace(sSiretEmetteur, '"', '', [rfReplaceAll]);
    sSiretEmetteur := stringReplace(sSiretEmetteur, ' ', '', [rfReplaceAll]);

  // requete sur la tablette TDG pour récupérer les groupes
    sSql := 'SELECT YX_CODE FROM ##DP##.CHOIXEXT WHERE YX_TYPE= "TDG" AND YX_CODE LIKE "%' + sSiretEmetteur + '"';

    Q := OpenSql(sSql, true);
    ListeGroupes := '(';
    while not Q.eof do
    begin
      ListeGroupes := ListeGroupes + '"' + copy(Q.FindField('YX_CODE').asString, 1, 3) + '",';
      Q.next;
    end;
    Ferme(Q);
    ListeGroupes := ListeGroupes + ')';
    ListeGroupes := stringReplace(ListeGroupes, ',)', ')', [rfReplaceAll]);

    if ListeGroupes <> '()' then
    begin
   // requete sur la table USERCONF pour récupérer les codes utilisateur
   // des groupes rattachés à l'émetteur
      sSql := 'SELECT UCO_USER FROM ##DP##.USERCONF WHERE UCO_GROUPECONF IN ' + ListeGroupes;
      Q := openSql(sSql, true);
      ListeCodeUser := '(';

      while not Q.eof do
      begin
      	unUtilisateur :=Q.FindField('UCO_USER').asString ;
      	if pos (unUtilisateur, ListeCodeUser) = 0 then
        	ListeCodeUser := ListeCodeUser + '"' +unUtilisateur  + '",';
        Q.next;
      end;
      Ferme(Q);
      ListeCodeUser := ListeCodeUser + ')';
      ListeCodeUser := stringReplace(ListeCodeUser, ',)', ')', [rfReplaceAll]);

      sSql := 'SELECT US_UTILISATEUR, US_LIBELLE FROM UTILISAT WHERE US_UTILISATEUR IN ' + ListeCodeUser;
    end
    else
     // prendre tous les utilisateurs
      sSql := 'SELECT US_UTILISATEUR, US_LIBELLE FROM UTILISAT ';


    Q := OpenSql(sSql, true);
    while not Q.eof do
    begin
      Combo_Utilisateurs.Items.Add(Q.FindField('US_LIBELLE').asString);
      Combo_Utilisateurs.Values.Add(Q.FindField('US_UTILISATEUR').asString);
      Q.next;
    end;
    Ferme(Q);
  end;

  if assigned(Combo_Comptes_JDC) then
  begin
    sSql := 'SELECT NJC_NOMCONNEXION FROM NJDC_COMPTES WHERE NJC_INTITULE ="' + sSiretEmetteur + '"';

    Q := OpenSQL(sSql, True); // liste par émetteur/utilisateur des connexions
    while not Q.EOF do
    begin
      Combo_Comptes_JDC.Items.Add(Q.FindField('NJC_NOMCONNEXION').asString);
      Combo_Comptes_JDC.Values.Add(Q.FindField('NJC_NOMCONNEXION').asString);
      Q.Next;
    end;
    Ferme(Q);
  end;
  // iitialiser la grille Connexions
  // tablette TDC
  // YX_CODE=SIRET+ NOM-UTILISATEUR
  sSql := 'SELECT YX_CODE, YX_LIBELLE FROM ##DP##.CHOIXEXT WHERE YX_TYPE="TDC" AND  ';
  sSql := sSql + 'YX_CODE LIKE "' + sSiretEmetteur + '%"';

  JC := TOB.Create('CHOIXEXT', nil, -1);
  Q := OpenSQL(sSql, True); // liste par émetteur/utilisateur des connexions
  // parcourir la tob pour corriger certains champs
  i := 0;
  if not Q.EOF then
  begin
    JC.LoadDetailDB('LIGNES_CONNEXION', '', '', Q, False, True);
    nbLignes := JC.Detail.Count;

    while i < nbLignes do
    begin
      TobCur := JC.Detail[i];
      //YX_CODE =SIRET+ code utilisateur
      TobCur.AddChampSup('NOM_UTILISATEUR', FALSE);
      sSql := 'SELECT US_LIBELLE FROM UTILISAT WHERE US_UTILISATEUR="';
      sSql := sSql + copy(TobCur.GetValue('YX_CODE'), 15, 3) + '"';
      Q1 := OpenSql(sSQl, True);
      if not Q1.eof then
        TobCur.PutValue('NOM_UTILISATEUR', Q1.FindField('US_LIBELLE').asString) // Groupe
      else
        TobCur.PutValue('NOM_UTILISATEUR', 'Inconnu:' + copy(TobCur.GetValue('YX_CODE'), 15, 3)); // Groupe
      Ferme(Q1);
      i := i + 1;
    end;
  end;
  JC.PutGridDetail(GRILLEU, FALSE, FALSE, sColonneSuite, True);
  Ferme(Q);
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.InitialiserGrilleComptes(Sender: TObject);

var
  i, nbLignes: integer;
  TobCur: Tob;
  E, TobCurE: TOB;
//@@ MF  sNomEmetteur, sSiret, GuidPer: string;
  sSiret, GuidPer: string;
  Q1: Tquery;
  sSql: string;
  IndexProc: integer; // pour menu téléprocédures
begin

  if assigned(J) then J.free;
  J := TOB.Create('NJDC_COMPTES', nil, -1);

  E := TOB.create('NTDI_EMETTEUR', nil, -1);
  Q := OpenSql(' SELECT * FROM NTDI_EMETTEUR', TRUE);
  if not Q.EOF then
    E.LoadDetailDB('NTDI_EMETTEUR', '', '', Q, False, True);
  FERME(Q);
  // AND NTM_SIRETEMETTEUR IN("...","...)
  // AND NTM_SIRETEMETTEUR = "....."
  sSql := Combo_Emetteur.Value;
//pgiInfo (sSql, 'Trace:' + V_PGI.User);

  if pos('=', sSql) <> 0 then
  begin
    sSiret := ReadTokenPipe(sSql, '=');
    sSiret := sSql;
    sSiret := stringReplace(sSiret, '"', '', [rfReplaceAll]);
    sSiret := stringReplace(sSiret, ' ', '', [rfReplaceAll]);
    Q := OpenSql(' SELECT * FROM NJDC_COMPTES WHERE NJC_INTITULE ="' + sSiret + '"', TRUE);
  end
  else
  begin
    sSql := stringReplace(sSql, 'AND NTM_SIRETEMETTEUR ', '', [rfIgnoreCase]);
// d @@ MF
    if (social) then
      sSql := '='+SiretEmetteurSocial;
// f @@ MF
    Q := OpenSql(' SELECT * FROM NJDC_COMPTES WHERE NJC_INTITULE ' + sSql, TRUE);
  end;
  if not Q.EOF then
  begin
    J.LoadDetailDB('LIGNES', '', '', Q, False, True);
    // le champ NJC_INTITULE contient un Siret Emetteur; Ajouter NJC_NOMEMETTEUR
    // et mettre dans NJC_INTITULE le libellé
    i := 0;
    nbLignes := J.Detail.Count;

    while i < nbLignes do
    begin
      TobCur := J.Detail[i];
      sSiret := TobCur.getValue('NJC_INTITULE');
      TobCurE := E.FindFirst(['NEM_SIRET'], [sSiret], FALSE);

      // mettre un libellé en clair pour les téléprocédures
      IndexProc := ValuesProc.IndexOf(TobCur.getValue('NJC_LISTEPROCEDURE'));
      TobCur.AddChampSup('NOM_LISTEPROCEDURE', FALSE);
      if IndexProc >= 0 then
        TobCur.PutValue('NOM_LISTEPROCEDURE', Combo_Teleprocedures.Items[IndexProc])
      else
        TobCur.PutValue('NOM_LISTEPROCEDURE', '');

//@@ MF      sNomEmetteur := '';

      if TobCurE <> nil then
      begin
        guidPer := TobCurE.Getvalue('NEM_GUIDPER');
        Q1 := OpenSql('SELECT ANN_NOM1 FROM ANNUAIRE WHERE ANN_GUIDPER ="' + GuidPer + '"', TRUE);
        if not Q1.EOF then
          sNomEmetteur := Q1.FindField('ANN_NOM1').asString;
        Ferme(Q1);
      end;
      TobCur.AddChampSupvaleur('NJC_NOMEMETTEUR', sNomEmetteur);
      inc(i);
    end;
  end;
  if assigned(Grille) then
    J.PutGridDetail(GRILLE, False, False, sColonne, True);
  FERME(Q);
  E.free;
end;


procedure TOF_TDI_GESTIONCOMPTES_JDC.OnClickGrille(Sender: TObject);
var
  TobCur: Tob;
  Ligne: integer;
  CodeProc: string;
begin
 // trouver la ligne courante
  Ligne := Grille.Row;
  if (J = nil) or (Ligne > J.Detail.Count) then Exit; // ligne vide

  TobCur := TOB(Grille.Objects[0, Ligne]);
//  Combo_Emetteur1.Text := TobCur.getValue('NJC_NOMEMETTEUR');
  Combo_Emetteur1.ItemIndex := ItemsEmetteur1.IndexOf(TobCur.getValue('NJC_NOMEMETTEUR'));
  ;
  Combo_Emetteur1.Enabled := FALSE; // on ne peut pas corriger
  Compte.Text := TobCur.getValue('NJC_NOMCONNEXION');
  MotDePasse.Text := TobCur.getValue('NJC_PASSWORD');

  CodeProc := TobCur.getValue('NJC_LISTEPROCEDURE');
  if CodeProc <> '' then
    Combo_Teleprocedures.ItemIndex := strToInt(CodeProc)
  else
    Combo_Teleprocedures.ItemIndex := 0;

  SetStatutSaisieCompte(TRUE);
  Compte.SetFocus;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnSupprimer(Sender: TObject);
var
  TobCur: Tob;
  Ligne: integer;
  LeNomConnexion: string;
begin
 // trouver la ligne courante
  Ligne := Grille.Row;
  if (J = nil) or (Ligne > J.Detail.Count) then Exit; // ligne ou tableau vide

  TobCur := TOB(Grille.Objects[0, Ligne]);
  LeNomConnexion := TobCur.getValue('NJC_NOMCONNEXION');
  if PGIAsk(LeNomConnexion, 'Supprimer le compte ?') = mrYes then
  begin
    TobCur.free;
    ExecuteSql('DELETE NJDC_COMPTES WHERE NJC_NOMCONNEXION="' + LeNomConnexion + '"');
    J.PutGridDetail(GRILLE, False, False, sColonne, True);
//  Intitule.text := '';
    Compte.Text := '';
    MotDePasse.Text := '';
    SetStatutSaisieCompte(FALSE);
  // purger éventuellement la tablette TDC
    ExecuteSql('DELETE ##DP##.CHOIXEXT WHERE YX_TYPE="TDC" AND YX_LIBELLE="' + LeNomConnexion + '"');
  end;
  Grille.setFocus;
end;

procedure TOF_TDI_GESTIONCOMPTES_JDC.OnSupprimerAffectation(Sender: TObject);
var
  TobCur: Tob;
  Ligne: integer;
  Cle, sUtilisateur: string;
begin
 // trouver la ligne courante
  Ligne := GrilleU.Row;
  if (JC = nil) or (Ligne > JC.Detail.Count) then Exit; // ligne ou tableau vide
  TobCur := TOB(GrilleU.Objects[0, Ligne]);
  if TobCur <> nil then
  begin
    Cle := TobCur.getValue('YX_CODE');
    sUtilisateur := TobCur.getValue('NOM_UTILISATEUR');
    if PGIAsk('Utilisateur:'+sUtilisateur, 'Supprimer l''affectation ?') = mrYes then
    begin
    	TobCur.free;
    	ExecuteSql('DELETE ##DP##.CHOIXEXT WHERE YX_CODE="' + Cle + '"');
    	JC.PutGridDetail(GrilleU, False, False, sColonneSuite, True);
    end
  end;

  Combo_Utilisateurs.Text := '';
  Combo_Comptes_JDC.Text := '';
  SetStatutSaisieAffectation(FALSE);
  BDELETE1.Visible := FALSE;
  GrilleU.setFocus;
end;


procedure TOF_TDI_GESTIONCOMPTES_JDC.OnClickAffectations(Sender: TObject);
var
  TobCur: Tob;
  Ligne: integer;
  ValuesUtilisateurs, ValuesComptes: HtStrings; //@@ MF
  index: integer;
  CodeUtilisateur: string;
begin
 // trouver la ligne courante
  Ligne := GrilleU.Row;
  if (JC = nil) or (Ligne > JC.Detail.Count) then Exit; // ligne ou tableau vide

  TobCur := TOB(GrilleU.Objects[0, Ligne]);
  SetStatutSaisieAffectation(TRUE);
  ValuesUtilisateurs := Combo_Utilisateurs.Values;
  ValuesComptes := Combo_Comptes_JDC.Values;
  Combo_Utilisateurs.Enabled := FALSE; // interdire la saisie
  ComBo_Comptes_JDC.SetFocus;

  CodeUtilisateur := TobCur.getValue('YX_CODE');
  CodeUtilisateur := copy(CodeUtilisateur, 15, 3);
  Index := ValuesUtilisateurs.indexOf(CodeUtilisateur);
  Combo_Utilisateurs.ItemIndex := Index;
  sNomUtilisateur := Combo_Utilisateurs.Text;

  Index := ValuesComptes.indexOf(TobCur.getValue('YX_LIBELLE'));
  Combo_Comptes_JDC.ItemIndex := Index;

  sNomConnexion := Combo_Comptes_JDC.Text;

  BDELETE1.Visible := TRUE;
end;



initialization
  registerclasses([TOF_TDI_GESTIONCOMPTES_JDC]);
end.

