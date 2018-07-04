{***********UNITE*************************************************
Auteur  ...... : Régis ROHAULT
Créé le ...... : 15/11/2001
Modifié le ... : 21/01/2002
Description .. : Source TOM de la TABLE : REGROUPEMENT
Suite ........ : (REGROUPEMENT)
Mots clefs ... : TOM;REGROUPEMENT
*****************************************************************}
Unit TomRegroupement ;

Interface

Uses StdCtrls, Controls, Classes, db, forms, sysutils, dbTables, ComCtrls,
	 HCtrls, HEnt1, UTOM, Fiche, FichList,
	 Htb97, Commun ;

Type
  TOM_REGROUPEMENT = Class (TOM)
	procedure OnNewRecord                ; override ;
	procedure OnDeleteRecord             ; override ;
	procedure OnUpdateRecord             ; override ;
	procedure OnAfterUpdateRecord        ; override ;
	procedure OnLoadRecord               ; override ;
	procedure OnChangeField ( F: TField) ; override ;
	procedure OnArgument ( S: String )   ; override ;
	procedure OnClose                    ; override ;
	procedure OnCancelRecord             ; override ;
	private
	// Perso
	Compte 		: TListBox;
	GroupeCompte: TListBox;
	Societe		: THValComboBox ;
	Devise		: THValComboBox ;
	BAjout 		: TToolBarButton97;
	BSuppr		: TToolBarButton97;
	GroupeModified: Boolean;
	procedure CodeRegOnKeyPress(Sender: TObject; var Key: Char);
	procedure ListBoxOnDblClick(Sender: TObject);
	procedure MoveCompte(Src: TListBox);
	procedure BAjoutOnclick		(Sender: Tobject);
	procedure BSupprOnclick		(Sender: Tobject);
	procedure SocDevOnChange	(Sender: TObject);
	procedure LoadGroupe;
end ;

Implementation

uses
	HDb, ExtCtrls;

const
	Table = 'REGGENERAUX'; // Table détail de REGROUPEMENT


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 21/01/2002
Modifié le ... :   /  /
Description .. : Rempli la liste des comptes pour une société et une devise
Mots clefs ... : COMPTE;SOCIETE;DEVISE
*****************************************************************}
procedure TOM_REGROUPEMENT.SocDevOnChange(Sender : TObject);
var	Q: TQuery;
begin
  GroupeCompte.Clear;
  Compte.Clear;
  Q := OpenSQL('Select BQ_GENERAL from BANQUECP where BQ_DEVISE="' + Devise.Value +
				'" and BQ_SOCIETE="' + Societe.Value + '"', True);
  while not Q.Eof do
  begin
	Compte.Items.Add(Q.Fields[0].AsString);
	Q.Next;
  end;
  Ferme(Q);

  AssignDrapeau(TImage(GetControl('IDEV')), Devise.Value);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 21/01/2002
Modifié le ... :   /  /
Description .. : Déplace les comptes sélectionnés d'une liste vers l'autre
Mots clefs ... : LISTE;GROUPE
*****************************************************************}
procedure TOM_REGROUPEMENT.MoveCompte(Src: TListBox);
var	Dest: TListBox;
	I: Integer;
begin
	Inherited;
	Ds.Edit;
	SetField('TRE_CODEREG', GetField('TRE_CODEREG')); // Place en mode modif
	GroupeModified := True;

	if Src = Compte then
		Dest := GroupeCompte
	else
		Dest := Compte;
	//BeginUpdate
	for I := Src.Items.Count-1 downto 0 do
		if Src.Selected[I] then
		begin
			Dest.Items.Add(Src.Items[I]);
			Src.Items.Delete(I);
		end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 21/01/2002
Modifié le ... :  /  /
Description .. : Reharge la liste des comptes et transfère dans la deuxième
Suite ........ : liste ceux appartenant au groupe
Mots clefs ... : LISTE;GROUPE
*****************************************************************}
procedure TOM_REGROUPEMENT.LoadGroupe;
var Q : TQuery ;
	str : String;
	I: Integer;
begin
  str := GetField('TRE_CODEREG');
  if str <> '' then  // Pas en création (OnNew fait un OnLoad)
  begin
	Q := OpenSQL( 'Select TGR_GENERAL,TGR_SOCIETE from ' + table +
				' where TGR_CODEREG="' + str + '"', True);
	Societe.Value := Q.Fields[1].AsString; // Même société pour un groupe !
	//Devise.Value := TGR_DEVISE
	while not Q.Eof do // Transfère les éléments d'un groupe dans la liste
	begin
		str := Q.Fields[0].AsString; //('TGR_GENERAL');
		I := Compte.Items.IndexOf(str);
		if I >= 0 then
		begin
			GroupeCompte.Items.Add(str);
			Compte.Items.Delete(I);
		end;
		// else le compte n'existe plus !
		Q.Next;
	end;
	Ferme(Q);
  end;
  GroupeModified := False;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 21/01/2002
Modifié le ... :   /  /
Description .. : Gère le passage entre les listes
Mots clefs ... : LISTE;GROUPE
*****************************************************************}
procedure TOM_REGROUPEMENT.BAjoutOnclick(Sender : Tobject);
begin
  Inherited;
  MoveCompte(Compte);
end;

procedure TOM_REGROUPEMENT.BSupprOnclick(Sender : Tobject);
begin
  Inherited;
  MoveCompte(GroupeCompte);
end ;

procedure TOM_REGROUPEMENT.ListBoxOnDblClick(Sender: TObject);
begin
  Inherited;
  MoveCompte(TListBox(Sender));
end;


procedure TOM_REGROUPEMENT.CodeRegOnKeyPress(Sender: TObject; var Key: Char);
begin
  Inherited;
  ValidCodeOnKeyPress(Key);
end;


procedure TOM_REGROUPEMENT.OnNewRecord ;
var	I: Integer;
begin
  Inherited ;
  Societe.ItemIndex := 0; // Init avec première société !
  I := Devise.Values.IndexOf('EUR');
  if I < 0 then	I := 0; // Première monnaie si pas d'Euro !
  Devise.ItemIndex := I;
  SocDevOnChange(Nil); // Compte doit avoir une liste complète !
end ;

procedure TOM_REGROUPEMENT.OnDeleteRecord ;
begin
  Inherited ;
  ExecuteSQL('Delete from ' + Table + ' where TGR_CODEREG="' + GetField('TRE_CODEREG') + '"');
end ;

procedure TOM_REGROUPEMENT.OnUpdateRecord ;
begin
  Inherited ;
  if GroupeCompte.Items.Count = 0 then // Au moins un (2 !?) élément dans le groupe
  begin
	TrShowMessage(Ecran.Caption, 9, '', '');
	Abort;
  end;
end ;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 21/01/2002
Modifié le ... :   /  /    
Description .. : Remplace l'enregistrement de REGGENERAUX 
Suite ........ : correspondant à celui de REGROUPEMENT
Mots clefs ... : TABLE DETAIL
*****************************************************************}
procedure TOM_REGROUPEMENT.OnAfterUpdateRecord ;
var I: Integer;
	Code: String;
begin
  Inherited ;
  if GroupeModified then
  try
	BeginTrans;
	OnDeleteRecord;
	Code := GetField('TRE_CODEREG');
	for I := 0 to GroupeCompte.Items.Count-1 do
		ExecuteSQL('Insert into ' + Table + ' (TGR_CODEREG,TGR_GENERAL,TGR_SOCIETE,TGR_DEVISE) values ("'+
				Code +'","'+ GroupeCompte.Items[I] +'","'+ Societe.Value +'","'+ Devise.Value +'")');

	CommitTrans;
	GroupeModified := False;
  except
	RollBack;
  end;
end ;

procedure TOM_REGROUPEMENT.OnLoadRecord ;
begin
  Inherited ;
  LoadGroupe;
end ;

procedure TOM_REGROUPEMENT.OnCancelRecord ;
begin
  Inherited ;
  if GroupeModified then
	LoadGroupe;	// Restaure les listes
end ;


procedure TOM_REGROUPEMENT.OnArgument ( S: String ) ;
begin
  Inherited ;
	THDBEdit(GetControl('TRE_CODEREG')).OnKeyPress := CodeRegOnKeyPress;
    DS.Filter := ' CC_TYPE="TRG"' ;
    DS.Filtered:=TRUE;

	Devise		:= THValComboBox	(GetControl('TRE_DEVISE')) ;
	Societe		:= THValComboBox	(GetControl('SOCIETE')) ;
	Compte 		:= TListBox			(GetControl('COMPTE'));
	GroupeCompte := TListBox		(GetControl('GROUPECOMPTE')) ;
	BAjout 		:= TToolBarButton97	(GetControl('BAJOUT'));
	BSuppr 		:= TToolBarButton97	(GetControl('BSUPPR'));

	BAjout.Onclick		:= BAjoutOnclick ;
	BSuppr.Onclick		:= BSupprOnclick ;
	Societe.OnChange  := SocDevOnChange ;
	Devise.OnChange   := SocDevOnChange ;
	Devise.SetFocus ; // ?
	Compte.OnDblClick := ListBoxOnDblClick;
	GroupeCompte.OnDblClick := ListBoxOnDblClick;
end ;

procedure TOM_REGROUPEMENT.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_REGROUPEMENT.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOM_REGROUPEMENT ] ) ;
end.
