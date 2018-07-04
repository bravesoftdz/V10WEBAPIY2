{***********UNITE*************************************************
Auteur  ...... : Franck VAUTRAIN
Créé le ...... : 26/03/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTTYPERES ()
Mots clefs ... : TOF;BTTYPERES
*****************************************************************}
Unit UtofBtTypeRes;

Interface

Uses UTOF,
     StdCtrls,
     Controls,
     Classes,
     HTB97,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}
     dbtables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF}
     Fiche,
     FE_Main,
{$ELSE}
     eFiche,
     MaineAGL,
{$ENDIF}
     UTob,
     Lookup,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox;

Type
  TOF_BTTYPERES = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Argument : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private

    fAction         : TActionFiche;

    BTR_TYPRES			: THEdit;
    BTR_FAMRES			: THEdit;
    BTR_LIBELLE			: THEdit;
    BTR_LIBELLE2		: THEdit;

    BTR_SURBOOKING	: TCheckBox;
    BTR_GEREPLANNING: TCheckBox;

    TobTypeRes 			: Tob;

    Typres          : String;
    FamRes          : String;

    procedure AfficheErreur;
    procedure ChargeEcran;
    procedure ControleChamp(Champ, Valeur: String);
    procedure ControleCritere(Valeur: String);
    procedure FamResOnElipsisClick(SEnder: TObject);
    procedure FamResOnExit(SEnder: TObject);
    procedure GestionEcran;
    procedure LectureFamilleRes(FamRes: string);
    procedure TypResOnExit(Sender: TObject);

    Function ControlSousFamille : Boolean;

  end;

const
	// libellés des messages
  TexteMessage: array[1..4] of string  = (
  {1}         'Le code Sous-type Ressource doit être renseigné',
  {2}         'Le code type Ressource doit être renseigné',
  {3}         'Le libelle ne peut être à blanc',
  {4}         'Le sous-type est affecté à une ressource, suppression interdite !'
  );

Implementation

procedure TOF_BTTYPERES.OnNew ;
begin
  Inherited ;


end ;

procedure TOF_BTTYPERES.OnDelete ;
Var StrSQL  : String;
    QQ      : TQuery;
begin
  Inherited ;

  //Vérification si cette sous famille pas sur une ressource
  StrSQL := 'SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_CHAINEORDO="' + TypRes;
  StrSQL := StrSQL + '" AND ARS_TYPERESSOURCE="' + Famres + '"';
  QQ := OpenSQL(StrSQL, true);

  if Not QQ.Eof then
     Begin
     //Message d'erreur
     LastError := 4;
     AfficheErreur;
     Ferme(QQ);
     Exit;
     end;

  Ferme(QQ);

  //si modification chargment de la tob Modele de Tache
  ExecuteSQL('DELETE FROM BTTYPERES WHERE BTR_TYPRES ="' + Typres + '" AND BTR_FAMRES="' + Famres + '"');

  Ecran.Close;

end;

procedure TOF_BTTYPERES.OnUpdate ;
begin
  Inherited ;

  //Controle de la cohérence des zones saisies
  if BTR_TYPRES.text = '' then
     begin
     LastError := 1;
     AfficheErreur;
     BTR_TYPRES.SetFocus;
     exit;
     end;

  if BTR_FAMRES.text = '' then
     begin
     LastError := 2;
     BTR_FAMRES.SetFocus;
     exit;
     end;

  if BTR_LIBELLE.text = '' then
     begin
     LastError := 3;
     BTR_LIBELLE.SetFocus;
     exit;
     end;

  //chargement de la TOB
  TobTypeRes.putvalue('BTR_TYPRES', BTR_TYPRES.text);
  TobTypeRes.putvalue('BTR_FAMRES', BTR_FAMRES.text);
  TobTypeRes.putvalue('BTR_LIBELLE', BTR_LIBELLE.text);
  TobTypeRes.putvalue('BTR_LIBELLE2', BTR_LIBELLE2.text);

  if BTR_GEREPLANNING.Checked then
     TobTypeRes.putvalue('BTR_GEREPLANNING', 'X')
  else
     TobTypeRes.putvalue('BTR_GEREPLANNING', '-');

  if BTR_SURBOOKING.Checked then
     TobTypeRes.putvalue('BTR_SURBOOKING', 'X')
  else
     TobTypeRes.putvalue('BTR_SURBOOKING', '-');

  TobTypeRes.InsertOrUpdateDb(False);

end ;

procedure TOF_BTTYPERES.OnLoad ;
begin

  Inherited ;

  SetControlProperty('BtDelete', 'Visible', False);

  if fAction = taCreat then
     Begin
     Exit;
     end;

  if not ControlSousFamille then
     Begin
     fAction := TaCreat;
     Exit;
     end;

  ChargeEcran;

end ;

Procedure TOF_BTTYPERES.ChargeEcran;
Begin

  SetControlProperty('BtDelete', 'Visible', True);

  //chargement de l'écran
  BTR_TYPRES.Text := TypRes;
  BTR_TYPRES.Enabled := false;

  BTR_FAMRES.text := FamRes;
  LectureFamilleRes(FamRes);

  BTR_LIBELLE.Text  := TobTypeRes.GetString('BTR_LIBELLE');
  BTR_LIBELLE2.text := TobTypeRes.GetString('BTR_LIBELLE2');

  SetControlText('BTR_GEREPLANNING', TobTypeRes.GetString('BTR_GEREPLANNING'));
  SetControlText('BTR_SURBOOKING', TobTypeRes.GetString('BTR_SURBOOKING'));

end;

Function TOF_BTTYPERES.ControlSousFamille : Boolean;
Var StrSQL  : String;
    QQ      : TQuery;
Begin

  Result := True;

  //si modification chargment de la tob Modele de Tache
  StrSQL := 'SELECT * FROM BTTYPERES WHERE BTR_TYPRES ="' + Typres + '" AND BTR_FAMRES="' + Famres + '"';
  QQ := OpenSql (StrSQL,true);

  if QQ.eof then
     Result := False
  else
     TobTypeRes.selectDB('',QQ);

  ferme (QQ);

end;

procedure TOF_BTTYPERES.OnArgument (Argument : String ) ;
var Critere : String;
    Valeur  : String;
    Champ   : String;
    X       : integer;
begin
  Inherited ;

  //Récupération valeur de argument
	Critere:=(Trim(ReadTokenSt(Argument)));

  while (Critere <> '') do
    begin
      if Critere <> '' then
      begin
        X := pos (':', Critere) ;
        if x = 0 then
          X := pos ('=', Critere) ;
        if x <> 0 then
        begin
          Champ := copy (Critere, 1, X - 1) ;
          Valeur := Copy (Critere, X + 1, length (Critere) - X) ;
        	ControleChamp(champ, valeur);
				end
      end;
      ControleCritere(Critere);
      Critere := (Trim(ReadTokenSt(Argument)));
    end;

  GestionEcran;

end;

Procedure TOF_BTTYPERES.GestionEcran;
Begin

  //création de la Tob Modele de taches
  TobTypeRes := Tob.create('BTTYPERES', nil, -1);

  //Initialisation des zones d'écran...
  BTR_TYPRES := THedit(getcontrol('BTR_TYPRES'));
  BTR_TYPRES.OnExit := TypResOnExit;

  BTR_FAMRES := THedit(getcontrol('BTR_FAMRES'));
  BTR_FAMRES.OnExit := FamResOnExit;
  BTR_FAMRES.OnElipsisClick := FamResOnElipsisClick;

  BTR_LIBELLE := THedit(getcontrol('BTR_LIBELLE'));
  BTR_LIBELLE2 := THedit(getcontrol('BTR_LIBELLE2'));

  BTR_GEREPLANNING := TCheckBox(getcontrol('BTR_GEREPLANNING'));
  BTR_SURBOOKING := TCheckBox(getcontrol('BTR_SURBOOKING'));

  Ecran.caption := 'Sous-types de ressources';
  UpdateCaption(Ecran);

end;

procedure  TOF_BTTYPERES.FamResOnElipsisClick(Sender: TObject);
Begin

  SetControlProperty('BTR_FAMRES', 'Plus', '');

  LookupCombo(BTR_FAMRES);

  if GetControlText('BTR_FAMRES')='' then exit;

  //Lecture de l'adresse sélectionnée et affichage des informations
  LectureFamilleRes(BTR_FAMRES.Text);

end;

Procedure TOF_BTTYPERES.TypResOnExit(Sender: TObject);
Begin

  //Lecture de l'adresse sélectionnée et affichage des informations
  if ControlSousFamille then ChargeEcran;

end;

Procedure TOF_BTTYPERES.FamResOnExit(Sender: TObject);
Begin

  //Lecture de l'adresse sélectionnée et affichage des informations
  LectureFamilleRes(BTR_FAMRES.Text);

end;


Procedure TOF_BTTYPERES.ControleChamp(Champ : String;Valeur : String);
Begin

  if champ = 'TYPRES' then TypRes := Valeur;
  if champ = 'FAMRES' then FamRes := Valeur;

  if Champ = 'ACTION' then
	   Begin
     if Valeur = 'CREATION' Then
        fAction := TaCreat
     else if Valeur = 'MODIFICATION' then
        fAction := TaModif
     else
	      fAction := TaConsult
  end;

end;

Procedure TOF_BTTYPERES.ControleCritere(Valeur : String);
Begin

end;

procedure TOF_BTTYPERES.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPERES.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPERES.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTTYPERES.LectureFamilleRes(FamRes: string);
var Req       : String;
    TobFamRes : Tob;
begin

  if FamRes = '' then exit;

  SetControlText('LBTR_FAMRES', '');

  //Lecture du Type ACtion Evenement sélectionné et affichage des informations
  req := 'SELECT HFR_LIBELLE FROM HRFAMRES WHERE HFR_FAMRES="' + FamRes + '"';

  TobFamRes := Tob.Create('#HRFAMRES',Nil, -1);
  TobFamRes.LoadDetailFromSQL(Req);
  if TobFamRes.Detail.Count > 0 then
     Setcontroltext('LBTR_FAMRES', TobFamRes.detail[0].GetString('HFR_LIBELLE'))
  else
     Setcontroltext('LBTR_FAMRES', '');

  TobFamRes.free;

end;

Procedure TOF_BTTYPERES.AfficheErreur;
Var Titre : string;
Begin

  Titre := TraduireMemoire('Sous-type Ressource');

  if LastError <> 0 then
	   LastErrorMsg := TraduireMemoire(TexteMessage[LastError])
  Else
	   LastErrorMsg := TraduireMemoire(LastErrorMsg);

  PGIBox(LastErrorMsg,Titre);

  LastError := 0;

end;

Initialization
  registerclasses ( [ TOF_BTTYPERES ] ) ;
end.

