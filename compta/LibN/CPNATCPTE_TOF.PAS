{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPNATCPTE ()
Mots clefs ... : TOF;CPNATCPTE
*****************************************************************}

Unit CPNATCPTE_TOF;

//================================================================================
// Interface
//================================================================================
Interface

Uses
{$IFDEF EAGLCLIENT}
	MaineAGL,
{$ELSE}
	FE_Main,
	db,
	dbtables,
{$ENDIF}
//	UTOFFILVIE,
	UTOB,
  uTOF,
	Hctrls,
	hmsgbox,
	HEnt1,
	Ent1,
	Forms,
	controls,
	Classes,
	Sysutils,
	LibChpLi,
	CPNATCPTECODELG_TOF
	;

//==================================================
// Fonctions d'ouverture de la fiche
//==================================================
Procedure FicheNatCpte(Q : TQuery ; Code,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
Procedure AGLFicheNatCpte(Q : TQuery ; Code,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
procedure CPLanceFiche_SaisieTableLibre(S : String);

//==================================================
// Definition de class
//==================================================
Type
//	TOF_CPNATCPTE = Class (TOFFILVIE)
	TOF_CPNATCPTE = Class (TOF)
	private
		CodeTableLibre : String;    // Code de la table a saisir
		CodeEnreg : String;         // Code de l'enregistrement a afficher
		LongueurCode : Integer;     // Nombre de caractère utilisé pour le code
		NewLongueur : Boolean;      // Indicateur de saisie de la taille du code

	public

		procedure InitNew(var new : TOB)   ; override ; // init d'un nouvel enreg
		function  BeforDelete : boolean    ; override ; // peut on effacé ??
		procedure AfterDelete(ToDel : TOB) ; override ; // traitement post delete

		function  VerifExitCode : boolean  ; override ; // verif a la sortie du code

		function  AddWhereCritere : string ; override ; // critere suplementaire

    protected
    		procedure OnArgument(S : String) ; override ;
	private
		procedure RecupereLongueur;
	end;

//================================================================================
// Implementation
//================================================================================
Implementation

//==================================================
// Definition des Constant
//==================================================
Const MessageListe : Array[1..1] of String =(
		{01}    'Voulez-vous mettre à jour les comptes référençant ce code de table libre ?'
		);

//==================================================
// Fonctions d'ouverture de la fiche
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure FicheNatCpte(Q : TQuery ; Code,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
Var
	Param : String;
begin
	Case Comment of
		taCreat,taCreatOne : Param := 'ACTION=CREATION';
		taModif            : Param := 'ACTION=MODIFICATION';
	else Param := 'ACTION=CONSULTATION';
	end;

	Param := Param + ';' + Code + ';' + Compte;

	CPLanceFiche_SaisieTableLibre(Param);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure AGLFicheNatCpte(Q : TQuery ; Code,Compte : String ; Comment : TActionFiche ; QuellePage : Integer);
Var
	Param : String;
begin
	Case Comment of
		taCreat,taCreatOne : Param := 'ACTION=CREATION';
		taModif            : Param := 'ACTION=MODIFICATION';
	else Param := 'ACTION=CONSULTATION';
	end;

	Param := Param + ';' + Code + ';' + Compte;

	CPLanceFiche_SaisieTableLibre(Param);
END;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_SaisieTableLibre(S : String);
var
	Args, Arg1, Arg2 : String;
begin
	// Récupération des paramètres de lancement
	Args := S;
	ReadTokenSt(S); // Action = ...
	Arg1 := ReadTokenSt(S); // code table
	Arg2 := ReadTokenSt(S); // lequel

	AGLLanceFiche('CP','CPNATCPTE2',Arg1,Arg2,Args);
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
procedure TOF_CPNATCPTE.InitNew(var new : TOB);
begin
	Inherited;

	new.PutValue('NT_TYPECPTE',CodeTableLibre);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPNATCPTE.BeforDelete : boolean;
begin
	result := Inherited BeforDelete;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 25/11/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNATCPTE.AfterDelete(ToDel : TOB);
var
	code : string;
	Table,Champ : string;
begin
	// Effacement des champs référençant l'enregistrement à effacer
	code := ToDel.GetValue('NT_NATURE');

	if (code <> '') then
	begin
		Case CodeTableLibre[1] of
			'G' : begin Table := 'GENERAUX'; Champ := 'G_TABLE'+Copy(CodeTableLibre,3,1); end;
			'T' : begin Table := 'TIERS'; Champ := 'T_TABLE'+Copy(CodeTableLibre,3,1); end;
			'S' : begin Table := 'SECTION'; Champ := 'S_TABLE'+Copy(CodeTableLibre,3,1); end;
			'B' : begin Table := 'BUDGENE'; Champ := 'BG_TABLE'+Copy(CodeTableLibre,3,1); end;
			'D' : begin Table := 'BUDSECT'; Champ := 'BS_TABLE'+Copy(CodeTableLibre,3,1); end;
			'E' : begin Table := 'ECRITURE'; Champ := 'E_TABLE'+Copy(CodeTableLibre,3,1); end;
			'A' : begin Table := 'ANALYTIQ'; Champ := 'Y_TABLE'+Copy(CodeTableLibre,3,1); end;
			'U' : begin Table := 'BUDECR'; Champ := 'BE_TABLE'+Copy(CodeTableLibre,3,1); end;
			'I' : begin Table := 'IMMO'; Champ := 'I_TABLE'+Copy(CodeTableLibre,3,1); end;
		end;

		if (ExisteSql('SELECT ' + Champ + ' FROM ' + Table + ' WHERE ' + Champ + '="' + Code + '"')) then
		begin
			if (PGIAsk(TraduireMemoire(MessageListe[1]),TraduireMemoire('Paramétrage des zones libres')) = mrYes) then
			begin
				ExecuteSql('UPDATE ' + Table + ' SET ' + Champ + '=""' + ' WHERE ' + Champ + '="' + Code + '"');
			end;
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
function TOF_CPNATCPTE.VerifExitCode : boolean;
var
	code : string;
begin
	result := Inherited VerifExitCode;
	if (not result) then exit;

	// get du code
	code := GetControlText('NT_NATURE');
	if (not (code = '')) then
	begin
		// completion du code
		if Length(code) > LongueurCode then code := copy(code,1,LongueurCode);
		if Length(code) < LongueurCode then code := BourreLaDoncSurLaTable(CodeTableLibre,code);
		SetControlText('NT_NATURE',code);
	end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BOY
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPNATCPTE.AddWhereCritere : string;
begin
	result := Inherited AddWhereCritere + 'NT_TYPECPTE="' + CodeTableLibre + '" ';
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 21/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPNATCPTE.RecupereLongueur;
var
	i,j : Integer;
begin
	LongueurCode := 0;

	// Doit-on saisir la longueur du code ?
	NewLongueur := (not (Presence('NATCPTE','NT_TYPECPTE', CodeTableLibre)));
	if ((NewLongueur) and (not ModeConsultation)) then
	begin
		// Propose la saisie et mets à jour VH
		CPLanceFiche_CodeLongueur(CodeTableLibre);
	end;

	// Récupération de la longueur du code dans VH
	case CodeTableLibre[1] Of
		'G' : i := 1;
		'T' : i := 2;
		'S' : i := 3;
		'B' : i := 4;
		'D' : i := 5;
		'E' : i := 6;
		'A' : i := 7;
		'U' : i := 8;
		'I' : i := 9;
	else i := 0;
	end;

	if (i > 0) then
	begin
		j := StrToInt(Copy(CodeTableLibre,2,2))+1;
		LongueurCode := VH^.LgTableLibre[i,j];
	end;

	// si nouvelle longueur, MAJ du code si celui-ci est trop grand
	if (NewLongueur) then
	begin
		if (LongueurCode > 1) and (Length(CodeEnreg) > LongueurCode) then SetControlText('NT_NATURE',Copy(CodeEnreg,1,LongueurCode));
		NewLongueur := false;
	end;
end;

//================================================================================
// Initialization
//================================================================================
procedure TOF_CPNATCPTE.OnArgument(S: String);
begin
  inherited;
	NewLongueur := False;
	LongueurCode := 0;

	// init des parametre de la class parente
	TableName := 'NATCPTE';
	CodeName := 'NT_NATURE';
	LibelleName := 'NT_LIBELLE';
	PListeName := 'PRT_NATCPTE';

	// Récupération des paramètres de lancement
	ReadTokenSt(S); // Action = ...
	CodeTableLibre := ReadTokenSt(S);
	CodeEnreg := ReadTokenSt(S);

	// Initialisation interface
	if (Ecran <> nil) then
	Begin
{$IFNDEF CCS3}
		PersoChampsLibres(CodeTableLibre,TWinControl(GetControl('TInfos1'))) ;
		PersoChampsLibres(CodeTableLibre,TWinControl(GetControl('TInfos2'))) ;
{$ENDIF}
		// Si rien de visible dans les onglets d'infos, on les cache
		if (not GetControlVisible('TInfos1')) and (not GetControlVisible('TInfos2'))
		then SetControlVisible('Pages',False) ;
	End;

	// Récupérer ou saisie de la longueur max du code à saisir
	RecupereLongueur;
	SetControlProperty('NT_NATURE','MaxLength',LongueurCode);
end;

Initialization
	registerclasses([TOF_CPNATCPTE]);
end.
