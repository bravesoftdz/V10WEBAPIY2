{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... : 22/05/2003
Description .. : Source TOF d'une 'fiche reliée à une table avec liste'
Suite ........ : premierement elle marche
Suite ........ : ensuite elle est bi-compatible !
Mots clefs ... : UTOFFILVIE;TOFFILVIE;FIL;VIE;TOF;GENERIQUE
*****************************************************************}

Unit UTOFFILVIE;

//================================================================================
// Interface
//================================================================================
Interface

Uses
{$IFDEF EAGLCLIENT}
	MaineAGL,
	UtileAGL,
{$ELSE}
	FE_Main,
	db,
	{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
	PrintDBG,
{$ENDIF}
	vierge,
	classes,
	windows,
	sysutils,
	Controls,
	StdCtrls,
	HCtrls,
	HTB97,
	HMsgBox,
	UTOF,
	UTOB,
	HEnt1
	;

//==================================================
// Definition de class
//==================================================
Type
	TOFFILVIE = Class (TOF)
	public
	//===============//
	//   INTERFACE   //
	//===============//
		// Valeur a initializer OBLIGATOIREMENT (rien ne fonctionne sans ....)
		TableName,           // nom de la table
		CodeName,            // nom du champs a affiché comme 'code' de la table
		LibelleName,         // nom du champs a affiché comme 'Libelle' de la table
		// valeur calculé si besoin ....
		Prefixe,             // Prefixe des champs de la table ...
		// Valeur optionnel ....
		AbregeName,          // Nom du champs abregé a setté avec les 17 premier caracteres du 'libellé'*
		PListeName : string; // nom de la liste pour impression

		// les retours de boolean doivent etre 'true' si tt vas bien et que l'on continue, 'false' sinon !
		// les fonction 'befor' sont les test d'avnt requetes
		// les fonction 'after' sont comprisent dans les transaction ... (autre traitement DB 'client')

		procedure Init(S : String)           ; virtual ; abstract ; // init de la class

		procedure InitNew(var new : TOB)     ; virtual ; // init d'un nouvel enreg
		procedure AfterLoadRecord            ; virtual ; // Load d'un enreg (avant affichage)
		procedure AfterDisplayRecord         ; virtual ; // Load d'un enreg (apres affichage)
		function  BeforDelete : boolean      ; virtual ; // peut on effacé ??
		procedure AfterDelete(ToDel : TOB)   ; virtual ; // traitement post delete
		function  BeforUpdate : boolean      ; virtual ; // peut on updaté ??
		procedure AfterUpdate(ToMod : TOB)   ; virtual ; // traitement post update

		function  VerifExitCode : boolean    ; virtual ; // verif a la sortie du code
		function  VerifExitLibelle : boolean ; virtual ; // verif a la sortie du libelle

		function  AddWhereCritere : string   ; virtual ; // critere suplementaire pour remplissage
	//===============//
	// PLUS  TOUCHER //
	//===============//
	protected
		// mode consultation ??
		ModeConsultation : boolean;

		// control pour la fille
		FListe : THGrid;
		Query : TOB;

		procedure OnNew                  ; override ;
		procedure OnDelete               ; override ;
		procedure OnUpdate               ; override ;
		procedure OnLoad                 ; override ;
		procedure OnArgument(S : String) ; override ;
		procedure OnDisplay              ; override ;
		procedure OnClose                ; override ;
		procedure OnCancel               ; override ;
	private
		// caption de la fenetre
		Caption : string;

		// boolean d'etape du traitement
		OnCreat,OnRead,OnUndo : boolean;

		// nouvel enreg ??
		IsNew : boolean;

		// colonne de trie
		SortColumn : string;

		// btn de parcour
		First,Prev,Next,Last : TToolbarbutton97;

		// evenement
		procedure OnClickMonValider(Sender : TObject);
		procedure OnClickBImprimer(Sender : TObject);

		Procedure OnRowExitFListe(Sender : TObject ; ou : Longint ; var Cancel : Boolean ; Chg : Boolean);
		procedure OnSortedFListe(Sender : TObject);

		Procedure OnClickFirst(Sender : TObject);
		Procedure OnClickPrev(Sender : TObject);
		Procedure OnClickNext(Sender : TObject);
		Procedure OnClickLast(Sender : TObject);

		procedure OnExitCode(Sender: TObject);
		procedure OnExitLibelle(Sender: TObject);

		procedure FormKeyDown(Sender : TObject ; var Key : Word ; Shift : TShiftState);

		// autre fonction
		procedure ReverseTOB;
		function  SaveDB(row : integer) : boolean;
		function  IsModified(row : integer) : boolean;
	end;

//================================================================================
// Implementation
//================================================================================
Implementation

//==================================================
// Definition des Constant
//==================================================
Const Messages : Array[0..9] of String =(
		{00}    'Vous devez renseigner un code.',
		{01}    'Vous devez renseigner un libellé.',
		{02}    'Le code que vous avez saisi existe déjà. Vous devez le modifier.',
		{03}    'Confirmez-vous la suppression de l''enregistrement ?',
		{04}    'La Suppression n''a pu être effectuée pour des raison technique ...',
		{05}    'Voulez-vous enregister les modification ?',
		{06}    'La modification n''a pu être effectuée pour des raison technique ...',
		{07}    'PARAMETRE DEVELOPPEUR',
		{08}    'La classe n''a pas été correctement initialisée ... ça va planté ...',
		{09}    'La liste à été trié selon une colonne inconue'
		);


//==================================================
// Evenements Virtuel de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.InitNew(var new : TOB);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	new.InitValeurs();
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.AfterLoadRecord;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.AfterDisplayRecord;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function  TOFFILVIE.BeforDelete : boolean;
begin
	result := false;
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;
	result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.AfterDelete(ToDel : TOB);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function  TOFFILVIE.BeforUpdate : boolean;
begin
	result := false;
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;
	result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.AfterUpdate(ToMod : TOB);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function  TOFFILVIE.VerifExitCode : boolean;
var
	code : string;
begin
	result := false;
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	if (OnUndo) then exit;      // UNDO : alors pas de test ;)
	if (OnCreat) then exit;     // en cour de creation ... repositionnement sur le code !

	code := GetControlText(CodeName);

	// code obligatoire
	if (code = '') then
	begin
		PGIBox(TraduireMemoire(Messages[0]), TraduireMemoire(Ecran.Caption));
		SetFocusControl(CodeName);
		exit;
	end;
	// recherche dans la liste
	if (IsNew) then
	begin
		if (ExisteSQL('SELECT ' + CodeName + ' FROM ' + TableName + ' WHERE ' + CodeName + '="' + code + '" AND ' + AddWhereCritere)) then
		begin
			PGIBox(TraduireMemoire(Messages[2]),TraduireMemoire(Ecran.Caption));
			SetFocusControl(CodeName);
			exit;
		end;
	end;

	FListe.Cells[1,Fliste.row] := GetControlText(CodeName);

	result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function  TOFFILVIE.VerifExitLibelle : boolean;
begin
	result := false;
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	if (OnUndo) then exit;      // UNDO : alors pas de test ;)
	if (OnCreat) then exit;     // en cour de creation ... repositionnement sur le code !

	// Libelle Obligatoire
	if (GetControlText(LibelleName) = '') then
	begin
		PGIBox(TraduireMemoire(Messages[1]), TraduireMemoire(Ecran.Caption));
		SetFocusControl(LibelleName);
		Exit;
	end;

	// abregé ???
	if ((not (AbregeName = '')) and (GetControlText(AbregeName) = '')) then SetControlText(AbregeName,copy(GetControlText(LibelleName),1,17));

	// mise a jour de la liste
	FListe.Cells[2,Fliste.row] := GetControlText(LibelleName);

	result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function  TOFFILVIE.AddWhereCritere : string;
begin
	result := '';
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnNew;
var
	cancel : boolean;
	tmp : TOB;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	Inherited;

	// sauvegarde ??
	OnRowExitFliste(nil,FListe.Row,cancel,false);

	// passe en creation
	OnCreat := true;

	// en mode creation ??
	if (IsNew = true) then exit;

	// Ajout d'une ligne dans la Grille
	if (Query.Detail.Count >= (FListe.RowCount-1)) then FListe.RowCount := FListe.RowCount+1
	else if (Query.Detail.Count = 0) then
	else exit;

	//tob fille sup !
	tmp := TOB.Create(TableName,Query,-1);

	// Positionnement + evt
	if (FListe.RowCount-1 = 1) then OnRowExitFListe(nil,FListe.RowCount-1,cancel,false) // un seul enreg ... celui k'on créé
	else FListe.gotoRow(FListe.RowCount-1);                                             // plusieur enreg ....

	// set de la variable de creation
	IsNew := true;

	// Positionnement
	SetControlEnabled(CodeName,true);
	SetFocusControl(CodeName);

	// disable des ctrl
	SetControlEnabled('bInsert',false);
	SetControlEnabled('bDelete',false);

	// affichage des valeur
	InitNew(tmp);
	tmp.PutEcran(ecran);

	// passe en creation
	OnCreat := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnDelete;
var
	ligne : integer;
	ToDel : TOB;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	Inherited;

	ligne := FListe.row;
	IsNew := false;

	if (not (PGIAsk(TraduireMemoire(Messages[3]),Ecran.Caption) = mrYes)) then exit;
	if (not BeforDelete) then exit;

	try
		BeginTrans;

		// reaffichage
		if ((ligne = 1) and (ligne+1 >= FListe.RowCount)) then OnNew        // on etait sur la 1er ligne et y'en a pas d'autre => nouvelle ligne
		else if (ligne+1 >= FListe.RowCount) then FListe.GotoRow(ligne-1)   // on etait sur la derniere ligne => on revien d'une ligne
		else Query.Detail[FListe.Row].PutEcran(ecran);                      // on etait sur une ligne et y'en a d'autre => recup des donne de la ligne suivante

		// suppression de la liste actuel
		ToDel := TOB.Create('A Effacer',nil,-1);
		Query.Detail[ligne-1].ChangeParent(ToDel,-1);

		// effacement de la liste
		ToDel.DeleteDB;
		// effacement de la grille
		FListe.DeleteRow(ligne);

		// fonction client de delete
		AfterDelete(ToDel.Detail[0]);

		// effacage de la TOB
		FreeAndNil(ToDel);

		// commit
		CommitTrans;
	except
		RollBack;
		OnLoad;
		PGIBox(TraduireMemoire(Messages[4]),TraduireMemoire(Ecran.Caption));
	end;

	// disable des ctrl
	First.Enabled := (not (Fliste.row = 1));
	Prev.Enabled := (not (FListe.Row = 1));
	Next.Enabled := (not (FListe.row = (FListe.RowCount - 1)));
	Last.Enabled := (not (FListe.row = (FListe.RowCount - 1)));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnUpdate;
var
	tempcode : string;
	i : integer;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	Inherited;

	if (IsNew) then
	begin
		SaveDB(FListe.row-1);

		tempcode := GetControlText(CodeName);
		// retrie de la liste
		Query.Detail.Sort(SortColumn);
		if (FListe.SortDesc) then ReverseTOB;
		// reaffichage
		Query.PutGridDetail(FListe,false,false,CodeName + ';' + LibelleName,true);
		// reset sur l'enreg en question
		for i := 0 to Query.Detail.Count -1 do
		begin
			if (Query.Detail[i].GetString(CodeName) = tempcode) then
			begin
				if (FListe.RowCount > i+1) then FListe.Row := i+1;
				// disable des ctrl
				First.Enabled := (not (Fliste.row = 1));
				Prev.Enabled := (not (FListe.Row = 1));
				Next.Enabled := (not (FListe.row = (FListe.RowCount - 1)));
				Last.Enabled := (not (FListe.row = (FListe.RowCount - 1)));
				// break ...
				break;
			end;
		end;
	end
	else
	begin
		if (IsModified(FListe.row-1)) then SaveDB(FListe.row-1);
	end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnLoad;
var
	stQuery : string;
	where : string;
	cancel : boolean;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	Inherited;

	OnRead := true;

	// recup des enreg
	stQuery := 'SELECT * FROM ' + TableName;                            // select de base
	where := AddWhereCritere;                                           // recup les condition sup
	if (not (where = '')) then stQuery := stQuery + ' WHERE ' + where;  // ajout des condition sup
	stQuery := stQuery + ' ORDER BY ' + CodeName;                       // order by le code

	Query := TOB.Create(TableName,nil,-1);
	Query.LoadDetailDBFromSQL(TableName,stQuery);

	if (Query.Detail.Count > 0) then
	begin
		// maj de l'affichage
		cancel := false;
		Query.PutGridDetail(FListe,false,false,CodeName + ';' + LibelleName,true);
		OnRowExitFListe(nil,0,cancel,true);
	end
	// créé un new :)
	else OnNew;

	OnRead := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnArgument(S : String);
var
	s1 : string;
	bvalider,monvalider : TToolBarButton97;
begin
	Inherited;

	// mode ...
	IsNew := false;
	OnRead := true;
	OnUndo := false;
	OnCreat := false;

	// svg du caption
	Caption := Ecran.Caption;

	// Récupération des paramètres de lancement
	s1 := uppercase(s);
	ModeConsultation := pos('CONSULTATION',ReadTokenSt(S1)) > 0; // Action = ...

	// init de la class dérivé !
	Init(S);
	if ((TableName = '') or (CodeName = '') or (LibelleName = '')) then
	begin
		PGIBox(TraduireMemoire(Messages[8]),TraduireMemoire(Messages[7]));
		PostQuitMessage(1);
		exit;
	end;
	if (Prefixe = '') then Prefixe := copy(CodeName,1,pos('_',CodeName));
	SortColumn := CodeName;

	// possibilité d'impression
	if (PListeName = '') then SetControlVisible('BImprimer',false)
	else SetControlVisible('BImprimer',true);
	TToolbarbutton97(GetControl('BImprimer',true)).OnClick := OnClickBImprimer;

	// recup des 4 btn de nav
	First := TToolbarbutton97(GetControl('FIRST',true));
	Prev := TToolbarbutton97(GetControl('PREV',true));
	Next := TToolbarbutton97(GetControl('NEXT',true));
	Last := TToolbarbutton97(GetControl('LAST',true));
	First.OnClick := OnClickFirst;
	Prev.OnClick := OnClickPrev;
	Next.OnClick := OnClickNext;
	Last.OnClick := OnClickLast;

	// recup du btn valider ;)
	bvalider := TToolBarButton97(GetControl('bValider',true));
	monvalider := TToolBarButton97(GetControl('MONVALIDER',true));
	monvalider.OnClick := OnClickMonValider;
	monvalider.Top := bvalider.Top;
	monvalider.Left := bvalider.Left;

	// recup de la grille
	FListe := THGrid(GetControl('FLISTE',true));
	FListe.OnRowExit := OnRowExitFListe;
	FListe.OnSorted := OnSortedFListe;

	// recup des ctrls
	THEdit(GetControl(CodeName,true)).OnExit := OnExitCode;
	THEdit(GetControl(LibelleName,true)).OnExit := OnExitLibelle;

	// Control des touches
	TFVierge(Ecran).OnKeyDown := FormKeyDown;

	// si consultation .....
	if (ModeConsultation) then
	begin
		// les btns
		SetControlEnabled('bDefaire',false);
		SetControlEnabled('bInsert',false);
		SetControlEnabled('bDelete',false);
		SetControlEnabled('bValider',false);
		// la panel
		SetControlEnabled('EPANEL',false)
	end;
	// enable des btn de nav ...
	First.Enabled := false;
	Prev.Enabled := false;
	Next.Enabled := false;
	Last.Enabled := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnClose;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	Inherited;

	FreeAndNil(Query);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnDisplay;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnCancel;
var
	cancel : boolean;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	Inherited;

	OnUndo := true;

	// reaffichage
	Query.Detail[FListe.row-1].PutEcran(ecran);
	Query.Detail[FListe.row-1].PutLigneGrid(FListe,FListe.row,false,false,CodeName + ';' + LibelleName);

	// if IsNew !!!
	if (IsNew) then
	begin
		// si un seul enreg
		if (Query.Detail.Count = 1) then exit // normalement ca arrive jamais ... le btn est disable
		else
		begin
			Fliste.DeleteRow(FListe.row);
			IsNew := false;

			// affichage du nvl enreg
			Query.Detail[FListe.row-1].PutEcran(ecran);
			Query.Detail[FListe.row-1].PutLigneGrid(FListe,FListe.row,false,false,CodeName + ';' + LibelleName);
			OnRowExitFListe(nil,FListe.row,cancel,true);
		end;
	end;

	// disable des ctrl
	SetControlEnabled(CodeName,False);
	SetControlEnabled('bInsert',true);
	SetControlEnabled('bDelete',true);

	OnUndo := false;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnClickMonValider(Sender: TObject);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	OnUpdate;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnClickBImprimer(Sender : TObject);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	if (not (PListeName = '')) then PrintDBGrid({$IFNDEF EAGLCLIENT}nil,nil,{$ENDIF}Caption,PListeName,AddWhereCritere);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOFFILVIE.OnRowExitFListe(Sender : TObject ; ou : Longint ; var Cancel : Boolean ; Chg : Boolean);
var
	ToUpdate : boolean;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	ToUpdate := false;

	// est modifié ??
	if (not (OnRead)) then ToUpdate := (IsModified(ou-1) or IsNew);
	// maj si necessaire
	if (ToUpdate) then
	begin
		if (PGIAsk(TraduireMemoire(Messages[5]),Ecran.Caption) = mrYes) then Cancel := not SaveDB(ou-1)
		else OnCancel;
	end;
	// svg raté ??
	if (Cancel) then exit;

	// disable des ctrl
	First.Enabled := (not (Fliste.row = 1));
	Prev.Enabled := (not (FListe.Row = 1));
	Next.Enabled := (not (FListe.row = (FListe.RowCount - 1)));
	Last.Enabled := (not (FListe.row = (FListe.RowCount - 1)));

	// appele de la fonction client;
	AfterLoadRecord;

	// put des valeur sur l'ecran
	if (not (FListe.Row = 0)) then Query.Detail[FListe.Row-1].PutEcran(ecran);
	Ecran.Caption := Caption + ' ' + FListe.Cells[1,FListe.Row] + ' ' + FListe.Cells[2,FListe.Row];
	UpdateCaption(Ecran);

	// appele de la fonction client
	AfterDisplayRecord;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOFFILVIE.OnSortedFListe(Sender : TObject);
var
	tempcode : string;
	i : integer;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	// recup de la colonne
	if (FListe.SortedCol = 1) then SortColumn := CodeName
	else if (FListe.SortedCol = 2) then SortColumn := LibelleName
	else
	begin
		PGIBox(TraduireMemoire(Messages[9]),TraduireMemoire(Ecran.Caption));
		exit;
	end;

	// svg de l'enreg en cours
	tempcode := GetControlText(CodeName);

	// sort ...
	Query.Detail.Sort(SortColumn);
	if (FListe.SortDesc) then ReverseTOB;

	// reaffichage du comptenu de la tob
	Query.Detail[FListe.Row-1].PutEcran(ecran);

	// reset sur l'enreg en question
	for i := 0 to Query.Detail.Count -1 do
	begin
		if (Query.Detail[i].GetString(CodeName) = tempcode) then
		begin
			if (FListe.RowCount > i+1) then FListe.GotoRow(i+1);
			// disable des ctrl
			First.Enabled := (not (Fliste.row = 1));
			Prev.Enabled := (not (FListe.Row = 1));
			Next.Enabled := (not (FListe.row = (FListe.RowCount - 1)));
			Last.Enabled := (not (FListe.row = (FListe.RowCount - 1)));
			// break ...
			break;
		end;
	end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOFFILVIE.OnClickFirst(Sender : TObject);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	// deplacement dans la grille
	FListe.gotoRow(1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOFFILVIE.OnClickPrev(Sender : TObject);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	// deplacement dans la grille
	if (FListe.row - 1 > 0) then FListe.gotoRow(FListe.row - 1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOFFILVIE.OnClickNext(Sender : TObject);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	// deplacement dans la grille
	if (FListe.Row + 1 < FListe.RowCount) then FListe.gotoRow(FListe.Row + 1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOFFILVIE.OnClickLast(Sender : TObject);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	// deplacement dans la grille
	FListe.gotoRow(FListe.RowCount - 1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnExitCode(Sender : TObject);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	VerifExitCode;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.OnExitLibelle(Sender : TObject);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	VerifExitLibelle;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.FormKeyDown(Sender : TObject ; var Key : Word ; Shift : TShiftState);
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	case key of
		{VK_N} 78 : if (Shift = [ssCtrl]) then begin key := 0; OnNew; end;       // Nouveau
		{VK_F} 70 : if (Shift = [ssCtrl]) then ;
		VK_F3 : // precedant ou premier
			begin
				key := 0;
				if (Shift = [ssCtrl]) then OnClickFirst(nil)
				else if (Shift = []) then OnClickPrev(nil);
			end;
		VK_F4 : // suivant ou dernier
			begin
				key := 0;
				if (Shift = [ssCtrl]) then OnClickLast(nil)
				else if (Shift = []) then OnClickNext(nil);
			end;
		VK_F10 : begin key := 0; OnUpdate; end;                                  // Enregistrement
		VK_DELETE : if (Shift = [ssCtrl]) then begin key := 0; OnDelete; end;    // Supprimer
	end;
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOFFILVIE.ReverseTOB();
var
	TempQuery : TOB;
	i : integer;
begin
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	TempQuery := TOB.Create(TableName,nil,-1);
	for i := Query.Detail.Count - 1 downto 0 do Query.Detail[i].ChangeParent(TempQuery,-1);
	FreeAndNil(Query);
	Query := TempQuery;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOFFILVIE.SaveDB(row : integer) : boolean;
begin
	result := false;
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	// code et libellé obligatoire .... redone here
	if (not VerifExitCode) then exit;
	if (not VerifExitLibelle) then exit;

	// test 'client'
	if (not BeforUpdate) then exit;

	try
		BeginTrans;

		// recup des valeur dans la TOB
		Query.Detail[row].GetEcran(ecran);

		// test pour l'abregé ....
		if ((not (AbregeName = '')) and (GetControlText(AbregeName) = '')) then Query.Detail[row].PutValue(AbregeName,copy(GetControlText(LibelleName),1,17));

		// update or insert ....
		if (IsNew) then result := Query.Detail[row].InsertDB(nil,false)
		else result := Query.Detail[row].UpdateDB(false,false);

		// fin de creation ...
		IsNew := false;

		// fonction client de update
		AfterUpdate(Query.Detail[row]);

		// commit
		CommitTrans;
	except
		RollBack;
		PGIBox(TraduireMemoire(Messages[6]),TraduireMemoire(Ecran.Caption));
	end;

	// recup des valeur sur l'ecran
	Query.Detail[row].PutEcran(ecran);
	Query.Detail[row].PutLigneGrid(FListe,row+1,false,false,CodeName + ';' + LibelleName);

	// disable des ctrl
	SetControlEnabled(CodeName,False);
	SetControlEnabled('bInsert',true);
	SetControlEnabled('bDelete',true);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOFFILVIE.IsModified(row : integer) : boolean;
var
	i : integer;
	tc : TControl;
	tq : TOB;
begin
	result := false;
	if ((Prefixe = '') or (TableName = '') or (CodeName = '') or (LibelleName = '')) then exit;

	result := true;
	tq := Query.Detail[row];

	// modifier le row en cours ???
	if (IsNew) then exit
	else
	begin
		for i := 0 to ecran.ComponentCount-1 do
		begin
			tc := TControl(ecran.Components[i]);

			if (pos(Prefixe,tc.Name) = 1) then
			begin
				if (UpperCase(tc.ClassName) = 'THEDIT') then
				begin
					case (THedit(tc).OpeType) of
						OtString : if (not (THedit(tc).Text = tq.GetValue(tc.name))) then exit;
						OtReel   : if (not (StrToInt(THedit(tc).Text) = tq.GetValue(tc.name))) then exit;
						OtDate   : if (not (StrToDate(THedit(tc).Text) = tq.GetValue(tc.name))) then exit;
					end;
				end
				else if (UpperCase(tc.ClassName) = 'TCHECKBOX') then
				begin
					if (not (TCheckBox(tc).Checked = (tq.GetValue(tc.name) = 'X'))) then exit;
				end
				else if (UpperCase(tc.ClassName) = 'THSPINEDIT') then
				begin
					if (not (THSpinEdit(tc).value = tq.GetValue(tc.name))) then exit;
				end
				else if (UpperCase(tc.ClassName) = 'THVALCOMBOBOX') then
				begin
					if (THValComboBox(tc).ItemIndex < 0) then if (not (tq.GetValue(tc.name) = '')) then exit else continue;
					if (not (THValComboBox(tc).values[THValComboBox(tc).ItemIndex] = tq.GetValue(tc.name))) then exit;
				end
				else if (UpperCase(tc.ClassName) = 'THMULTIVALCOMBOBOX') then
				begin
					if (not (THMultiValComboBox(tc).value = tq.GetValue(tc.name))) then exit;
				end;
			end;
		end;
	end;

	result := false;
end;

(*
=> Fiche minimum a prevoir .....
* y ajouté les champs du code et du libellé au minimum !

//======== ENTETE ========\\
Nature=CP
Code=UTOFFILVIE
Libelle=Generique pour gestion des FIL
Ancetre=VIE
//======== FORMES ========\\
inherited UTOFFILVIE: TFVierge
  Left = 223
  Top = 151
  HelpContext = 1196100
  VertScrollBar.Range = 0
  ActiveControl = FLISTE
  BorderStyle = bsSingle
  Caption = '???'
  ClientHeight = 329
  ClientWidth = 525
  OldCreateOrder = True
  Position = poDesigned
  TOFName = 'UTOFFILVIE'
  PixelsPerInch = 96
  TextHeight = 13
  inherited Dock971: TDock97
	Top = 294
	Width = 525
	inherited PBouton: TToolWindow97
	  ClientWidth = 269
	  ClientAreaWidth = 269
	  FullSize = False
	  Resizable = False
	  inherited BValider: TToolbarButton97
		Left = 173
		Enabled = False
		Visible = False
	  end
	  inherited BFerme: TToolbarButton97
		Left = 205
	  end
	  inherited HelpBtn: TToolbarButton97
		Left = 237
	  end
	  inherited bDefaire: TToolbarButton97
		Visible = True
	  end
	  inherited Binsert: TToolbarButton97
		Visible = True
	  end
	  inherited BDelete: TToolbarButton97
		Visible = True
	  end
	  inherited BImprimer: TToolbarButton97
		Left = 141
	  end
	  object MONVALIDER: TToolbarButton97
		Left = 173
		Top = 2
		Width = 28
		Height = 27
		Hint = 'Valider'
		Default = True
		Flat = False
		Glyph.Data = {
		  42020000424D4202000000000000420000002800000010000000100000000100
		  1000030000000002000000000000000000000000000000000000007C0000E003
		  00001F0000000042004200420042004200420042004200420042004200420042
		  0042004200420042004200420042100010000042004200420042004200420042
		  0042004200420042004200421000000200021000004200420042004200420042
		  0042004200420042004210000002000200020002100000420042004200420042
		  0042004200420042100000020002000200020002000210000042004200420042
		  0042004200421000000200020002E00300020002000200021000004200420042
		  004200420042000200020002E0030042E0030002000200021000004200420042
		  004200420042E0030002E003004200420042E003000200020002100000420042
		  0042004200420042E00300420042004200420042E00300020002000210000042
		  00420042004200420042004200420042004200420042E0030002000200021000
		  004200420042004200420042004200420042004200420042E003000200020002
		  1000004200420042004200420042004200420042004200420042E00300020002
		  00021000004200420042004200420042004200420042004200420042E0030002
		  000200021000004200420042004200420042004200420042004200420042E003
		  0002000210000042004200420042004200420042004200420042004200420042
		  E003000200020042004200420042004200420042004200420042004200420042
		  0042E0030042}
	  end
	end
	object PBouton1: TToolWindow97
	  Left = 273
	  Top = 0
	  ClientHeight = 31
	  ClientWidth = 248
	  Caption = 'Barre outils fiche'
	  ClientAreaHeight = 31
	  ClientAreaWidth = 248
	  DockPos = 273
	  Resizable = False
	  TabOrder = 1
	  object FIRST: TToolbarButton97
		Left = 3
		Top = 2
		Width = 60
		Height = 27
		Alignment = taRightJustify
		Flat = False
		Glyph.Data = {
		  C2010000424DC20100000000000042000000280000000F0000000C0000000100
		  1000030000008001000000000000000000000000000000000000007C0000E003
		  00001F0000001863186318631042104210421863186318631863186318631042
		  1042186300001863FF7F00000000000010421863186318631863186300000000
		  1042186300001863FF7F00000000000010421863186318631863000000000000
		  1042186300001863FF7F00000000000010421863186318630000000000000000
		  1042186300001863FF7F00000000000010421863186300000000000000000000
		  1042186300001863FF7F00000000000010421863000000000000000000000000
		  1042186300001863FF7F0000000000001042FF7F000000000000000000000000
		  1042186300001863FF7F0000000000001042FF7FFF7F00000000000000000000
		  1042186300001863FF7F00000000000010421863FF7FFF7F0000000000000000
		  1042186300001863FF7F000000000000104218631863FF7FFF7F000000000000
		  1042186300001863FF7F0000000000001042186318631863FF7FFF7F00000000
		  1042186300001863FF7FFF7FFF7FFF7F18631863186318631863FF7FFF7FFF7F
		  186318630000}
	  end
	  object PREV: TToolbarButton97
		Left = 64
		Top = 2
		Width = 60
		Height = 27
		Alignment = taRightJustify
		Flat = False
		Glyph.Data = {
		  42020000424D4202000000000000420000002800000010000000100000000100
		  1000030000000002000000000000000000000000000000000000007C0000E003
		  00001F0000001863186318631863186318631863186318631863186318631863
		  1863186318631863186318631863186318631863186318631863186318631863
		  1863186318631863186318631863186318631863186318631863104210421863
		  1863186318631863186318631863186318631863186318630000000010421863
		  1863186318631863186318631863186318631863186300000000000010421863
		  1863186318631863186318631863186318631863000000000000000010421863
		  1863186318631863186318631863186318630000000000000000000010421863
		  1863186318631863186318631863186300000000000000000000000010421863
		  1863186318631863186318631863FF7F00000000000000000000000010421863
		  1863186318631863186318631863FF7FFF7F0000000000000000000010421863
		  18631863186318631863186318631863FF7FFF7F000000000000000010421863
		  186318631863186318631863186318631863FF7FFF7F00000000000010421863
		  1863186318631863186318631863186318631863FF7FFF7F0000000010421863
		  18631863186318631863186318631863186318631863FF7FFF7FFF7F18631863
		  1863186318631863186318631863186318631863186318631863186318631863
		  1863186318631863186318631863186318631863186318631863186318631863
		  186318631863}
	  end
	  object NEXT: TToolbarButton97
		Left = 125
		Top = 2
		Width = 60
		Height = 27
		Alignment = taRightJustify
		Flat = False
		Glyph.Data = {
		  42020000424D4202000000000000420000002800000010000000100000000100
		  1000030000000002000000000000000000000000000000000000007C0000E003
		  00001F0000001863186318631863186318631863186318631863186318631863
		  1863186318631863186318631863186318631863186318631863186318631863
		  1863186318631863186318631863186310421042186318631863186318631863
		  1863186318631863186318631863FF7F00000000104218631863186318631863
		  1863186318631863186318631863FF7F00000000000010421863186318631863
		  1863186318631863186318631863FF7F00000000000000001042186318631863
		  1863186318631863186318631863FF7F00000000000000000000104218631863
		  1863186318631863186318631863FF7F00000000000000000000000010421863
		  1863186318631863186318631863FF7F00000000000000000000000018631863
		  1863186318631863186318631863FF7F00000000000000000000186318631863
		  1863186318631863186318631863FF7F00000000000000001863186318631863
		  1863186318631863186318631863FF7F00000000000018631863186318631863
		  1863186318631863186318631863FF7F00000000186318631863186318631863
		  1863186318631863186318631863FF7FFF7F1863186318631863186318631863
		  1863186318631863186318631863186318631863186318631863186318631863
		  1863186318631863186318631863186318631863186318631863186318631863
		  186318631863}
	  end
	  object LAST: TToolbarButton97
		Left = 187
		Top = 2
		Width = 60
		Height = 27
		Alignment = taRightJustify
		Flat = False
		Glyph.Data = {
		  C2010000424DC20100000000000042000000280000000F0000000C0000000100
		  1000030000008001000000000000000000000000000000000000007C0000E003
		  00001F0000001863186310421042186318631863186318631042104210421042
		  1863186300001863FF7F000000001042186318631863FF7F0000000000001042
		  1863186300001863FF7F000000000000104218631863FF7F0000000000001042
		  1863186300001863FF7F000000000000000010421863FF7F0000000000001042
		  1863186300001863FF7F000000000000000000001042FF7F0000000000001042
		  1863186300001863FF7F00000000000000000000000010420000000000001042
		  1863186300001863FF7F000000000000000000000000FF7F0000000000001042
		  1863186300001863FF7F000000000000000000001863FF7F0000000000001042
		  1863186300001863FF7F000000000000000018631863FF7F0000000000001042
		  1863186300001863FF7F000000000000186318631863FF7F0000000000001042
		  1863186300001863FF7F000000001863186318631863FF7F0000000000001042
		  1863186300001863FF7FFF7F18631863186318631863FF7FFF7FFF7FFF7F1863
		  186318630000}
	  end
	end
  end
  object LPANEL: THPanel [1]
	Left = 273
	Top = 0
	Width = 252
	Height = 294
	Align = alRight
	FullRepaint = False
	TabOrder = 1
	BackGroundEffect = bdFlat
	ColorShadow = clWindowText
	ColorStart = clBtnFace
	TextEffect = tenone
	object FLISTE: THGrid
	  Left = 1
	  Top = 1
	  Width = 250
	  Height = 292
	  Align = alClient
	  ColCount = 3
	  DefaultRowHeight = 18
	  RowCount = 2
	  Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goColSizing, goRowSelect]
	  TabOrder = 0
	  SortedCol = -1
	  Titres.Strings = (
		''
		'Code'
		'Libellé')
	  Couleur = False
	  MultiSelect = False
	  TitleBold = True
	  TitleCenter = True
	  ColCombo = 0
	  SortEnabled = False
	  SortRowExclude = 0
	  TwoColors = False
	  AlternateColor = 13224395
	  DBIndicator = True
	  ColWidths = (
		10
		75
		145)
	end
  end
  object EPANEL: THPanel [2]
	Left = 0
	Top = 0
	Width = 273
	Height = 294
	Align = alClient
	FullRepaint = False
	TabOrder = 2
	BackGroundEffect = bdFlat
	ColorShadow = clWindowText
	ColorStart = clBtnFace
	TextEffect = tenone
  end
end
//======== SCRIPT ========\\

*)

end.

