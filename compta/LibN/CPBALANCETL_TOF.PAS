{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 03/05/2002
Modifié le ... : 26/08/2002
Description .. : Source TOF de la FICHE : CPBALANCETL ()
Suite ........ : 
Mots clefs ... : TOF;CPBALANCETL;BALANCE
*****************************************************************}
Unit CPBALANCETL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Graphics,
{$IFDEF EAGLCLIENT}
  Maineagl, eQRS1, uTOB,
{$ELSE}
  Fe_main, QRS1, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
{$ENDIF}
     sysutils,
     ComCtrls,
     HCtrls,
     UTOF,
     TofMeth ;

procedure CPLanceFiche_EtatBalanceSurTables;
procedure CPLanceFiche_EtatBalanceSurTablesM;

Type
  TOF_CPBALANCETL = Class (TOF_Meth)
  private
    CptG1, CptG2: THEdit;
    Exo, Etab, Devise: THValComboBox ;
    Rupture1, Rupture2, Rupture3, Rupture4 : THValComboBox ;
    RuptureType : THRadioGroup;
    Date1, Date2 : THEdit;
    DateD, DateF: TDatetime;
    DevPivot: THEdit;
    pFiltres	:	TTabSheet;
    inNbRupture : Integer; // Nombre de rupture à gérer suivant la série
    stComboPlus : String; // Critère supplèmentaire pour affichage dans les combos
    procedure CompteOnExit(Sender: TObject) ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    procedure InitFiche ;
    function GenererRequete: String;
    function GenererRegroupementLib:String;
    procedure AfficherFiltrePourTable(stTableIdx: String);
    procedure AfficherFiltres;
    procedure AffectGroup ( inTag : Integer ) ;
    procedure ChangeGroup ( inTag : Integer ) ;
    procedure RuptureOnClick(Sender: TObject);
    procedure RuptureOnEnter(Sender: TObject);
    procedure RuptureTypeOnClick(Sender: TObject);
    function  GetMinMaxCompte(stTable, stCol, stDefaut : String) : String;
  public
    procedure OnLoad                   ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  ULibExercice,
  {$ENDIF MODENT1}
  HQry,   // RecupWhereCritere
  HEnt1,
  Ent1;



//==============================================================================
procedure CPLanceFiche_EtatBalanceSurTables;
begin
	AGLLanceFiche('CP','CPBALANCETL','','','');
end;
//==============================================================================
procedure CPLanceFiche_EtatBalanceSurTablesM;
begin
	AGLLanceFiche('CP','CPBALANCETLM','','','');
end;
//==============================================================================

procedure TOF_CPBALANCETL.OnArgument (S : String ) ;
var
	inIdx : Integer;
begin
  Inherited ;

  with TFQRS1(Ecran) do begin
    // Paramétrage de l'état ???
    ParamEtat := (*(not (ctxPCL in V_PGI.PGIContexte)) and*) ( V_PGI.FSuperviseur );
    // Composants
    Exo				:=	THValComboBox(GetControl('E_EXERCICE')) ;
    CptG1			:=	THEdit(GetControl('E_GENERAL')) ;
    CptG2			:=	THEdit(GetControl('E_GENERAL_')) ;
    Date1			:=	THEdit(GetControl('E_DATECOMPTABLE')) ;
    Date2			:=	THEdit(GetControl('E_DATECOMPTABLE_')) ;
    Etab			:=	THValComboBox(GetControl('E_ETABLISSEMENT')) ;
    Devise		:=	THValComboBox(GetControl('E_DEVISE')) ;
    DevPivot	:=	THEdit(GetControl('DEVPIVOT'));
    Rupture1  :=	THValComboBox(GetControl('RUPTURE1')) ;
    Rupture2  :=	THValComboBox(GetControl('RUPTURE2')) ;
    Rupture3  :=	THValComboBox(GetControl('RUPTURE3')) ;
    Rupture4  :=	THValComboBox(GetControl('RUPTURE4')) ;
		RuptureType := THRadioGroup(GetControl('RUPTURE')) ;
		pFiltres	:=	TTabSheet(GetControl('FILTRES')) ;
  end;
  if Exo<>nil then begin
    Exo.OnChange:=ExoOnChange ;
    Exo.Value:=VH^.Entree.Code ;
  end ;
  if CptG1<>nil then CptG1.OnExit:=CompteOnExit ;
  if CptG2<>nil then CptG2.OnExit:=CompteOnExit ;
  if Date1<>nil then Date1.OnExit:=DateOnExit ;
  if Date2<>nil then Date2.OnExit:=DateOnExit ;
  if (Etab<>nil) and (Etab.ItemIndex=-1) then Etab.ItemIndex := 0 ;
  if Devise<>nil then Devise.ItemIndex := 0;
  if (Devise<>nil) and (DevPivot<>nil) then
    DevPivot.Text:=Devise.Items[Devise.Values.IndexOf(V_PGI.DevisePivot)];
	// Evènement sur la sélection des ruptures
  Rupture1.OnClick := RuptureOnClick;
  Rupture1.OnEnter := RuptureOnEnter;
  Rupture2.OnClick := RuptureOnClick;
  Rupture2.OnEnter := RuptureOnEnter;
  Rupture3.OnClick := RuptureOnClick;
  Rupture3.OnEnter := RuptureOnEnter;
  Rupture4.OnClick := RuptureOnClick;
  Rupture4.OnEnter := RuptureOnEnter;
	RuptureType.OnClick := RuptureTypeOnClick;
  //Paramètrage Série
	SetControlVisible('Rupture2', V_PGI.LaSerie > S3 );
	SetControlVisible('TRupture2', V_PGI.LaSerie > S3 );
	SetControlVisible('Rupture3', V_PGI.LaSerie > S5 );
	SetControlVisible('TRupture3', V_PGI.LaSerie > S5 );
	SetControlVisible('Rupture4', V_PGI.LaSerie > S5 );
	SetControlVisible('TRupture4', V_PGI.LaSerie > S5 );
	Case V_PGI.LaSerie of
  	S7 :  begin
          inNbRupture := 4;
          stComboPlus := '';
          end;
  	S5 :  begin
          inNbRupture := 2;
          stComboPlus := 'AND CC_CODE <= "E01"';
          end;
    else  begin
          inNbRupture := 1;
          stComboPlus := 'AND CC_CODE <= "E00"';
          end;
	end;
	// Initialisation des champs
  for inIdx := 1 to inNbRupture do SetControlProperty('RUPTURE'+IntToStr(inIdx),'Plus',stComboPlus);
  InitFiche;
  TFQRS1(Ecran).Pages.ActivePage:=TFQRS1(Ecran).Pages.Pages[0];
end ;

procedure TOF_CPBALANCETL.OnLoad;
begin
  inherited;
  // Préparation des champs pour la création de la requête par QRS1
  // Permet d'utiliser les mécanismes auto de création du where
  { FQ 19383 BVE 21.05.07 }
  // SetControlText('XX_WHERE','E_QUALIFPIECE="N"');
  { END FQ 19383}
  SetControlText('XX_VARIABLE',GenererRegroupementLib);
end;

procedure TOF_CPBALANCETL.OnUpdate;
begin
  inherited;
  // La requête créé par QRS1 est écrasée
	TFQRS1(Ecran).WhereSQL := GenererRequete;

	// Auto-remplissage si comptes auxiliaires non renseignés
	if (Trim(CptG1.Text) = '') then CptG1.Text := GetMinMaxCompte('GENERAUX', 'MIN(G_GENERAL)','0');
	if (Trim(CptG2.Text) = '') then CptG2.Text := GetMinMaxCompte('GENERAUX', 'MAX(G_GENERAL)','ZZZZZZZZZZZZZZZZZ');

	if GetCheckBoxState('Blocnote') = cbChecked
  	then SetControlText('AvecBlocnote','X')
    else SetControlText('AvecBlocnote','');
end;

//==============================================================================

procedure TOF_CPBALANCETL.CompteOnExit(Sender: TObject);
begin
	if (THEdit(Sender)=CptG1) or (THEdit(Sender)=CptG2) then
  	DoCompteOnExit(THEdit(Sender), CptG1, CptG2);
end;

procedure TOF_CPBALANCETL.DateOnExit(Sender: TObject);
begin
	DoDateOnExit(THEdit(Sender), Date1, Date2, DateD, DateF);
end;

procedure TOF_CPBALANCETL.ExoOnChange(Sender: TObject);
var ExoDate: TExoDate;
    MM,AA,NbMois: word;
BEGIN
  DoExoToDateOnChange(Exo, Date1, Date2);
  DateD := StrToDate(Date1.Text);
  DateF := StrToDate(Date2.Text);
  ExoDate.Deb := StrToDate(Date1.Text);
  ExoDate.Fin := StrToDate(Date2.Text);
  NOMBREPEREXO(ExoDate,MM,AA,NbMois) ;
  if NbMois > 12 then
    Date2.Text := DateToStr(FinDeMois(PlusMois(ExoDate.Deb, 11)));
end;

procedure TOF_CPBALANCETL.InitFiche;
// Initialisation des champs / libellés / champs cachés de la fiche
var
	texte : String;
	i : Integer;
begin
  SetControlChecked('Blocnote', False);
  SetControlText('AvecBlocnote', '');
  SetControlText('XX_VARIABLET1', ''); SetControlText('XX_VARIABLEL1', '');
  SetControlText('XX_VARIABLET2', ''); SetControlText('XX_VARIABLEL2', '');
  SetControlText('XX_VARIABLET3', ''); SetControlText('XX_VARIABLEL3', '');
  SetControlText('XX_VARIABLET4', ''); SetControlText('XX_VARIABLEL4', '');
  SetcontrolText('XX_VARIABLE', '');
	// init des combo
	Rupture1.ItemIndex := 0; RuptureOnClick(TObject(Rupture1));
  Rupture2.ItemIndex := 0;
  Rupture3.ItemIndex := 0;
  Rupture4.ItemIndex := 0;
	// Nom des tables
  For i:= 0 to (inNbRupture - 1) do
  	begin
    texte := RechDom('TTTABLESLIBRESECR','E0' + IntToStr(i), False);
		SetControlText('TE_TABLE' + IntToStr(i) , texte);
  	end;
end;

function TOF_CPBALANCETL.GenererRequete: String;
// Construit la requête envoyer à l'état, le nombre de champ et de jointures peut
//  changer en fonction du niveau de regroupement
Var
	stChamps, stJointures : String;
begin
	// Liste des champs utilisé pour le regroupement et jointures
  if Rupture1.value <> '' then
  	begin
    stChamps := 'RUP1.NT_NATURE, RUP1.NT_Libelle, ';
    stJointures := 'Left join NATCPTE RUP1 on E_Table' + Copy(Rupture1.value, 3, 1) +
                    ' = RUP1.NT_NATURE AND RUP1.NT_TYPECPTE = "E0' + Copy(Rupture1.value, 3, 1) + '" ' ;
  	end;
  if Rupture2.value <> '' then
  	begin
		stChamps := stChamps + 'RUP2.NT_NATURE, RUP2.NT_Libelle, ';
		stJointures := stJointures + 'Left join NATCPTE RUP2 on E_Table' + Copy(Rupture2.value, 3, 1) +
    															' = RUP2.NT_NATURE AND RUP2.NT_TYPECPTE = "E0' + Copy(Rupture2.value, 3, 1) + '" ' ;
    end;
  if Rupture3.ItemIndex > 0 then
  	begin
		stChamps := stChamps + 'RUP3.NT_NATURE, RUP3.NT_Libelle, ';
		stJointures := stJointures + 'Left join NATCPTE RUP3 on E_Table' + Copy(Rupture3.value, 3, 1) +
    															' = RUP3.NT_NATURE AND RUP3.NT_TYPECPTE = "E0' + Copy(Rupture3.value, 3, 1) + '" ' ;
    end;
  if Rupture4.ItemIndex > 0 then
  	begin
		stChamps := stChamps + 'RUP4.NT_NATURE, RUP4.NT_Libelle, ';
		stJointures := stJointures + 'Left join NATCPTE RUP4 on E_Table' + Copy(Rupture4.value, 3, 1) +
    															' = RUP4.NT_NATURE AND RUP4.NT_TYPECPTE = "E0' + Copy(Rupture4.value, 3, 1) + '" ' ;
    end;
	// Contruction Requete
  Result := 'SELECT ' + stChamps + 'E_general, G_LIBELLE, Sum(E_Debit) D, Sum(E_Credit) C'
  				+' FROM Ecriture Left join Generaux on E_General = G_General '
          + stJointures
          + RecupWhereCritere(TPageControl(GetControl('PAGES')))
					+' GROUP BY ' + stChamps + 'E_general, G_Libelle'
          +' ORDER BY ' + stChamps + 'E_general, G_Libelle';
end;

function TOF_CPBALANCETL.GenererRegroupementLib: String;
// Construit la liste des libellés des tables utilisées pour le regroupement
begin
	Result := '';
  if Rupture1.value <> '' then
  	Result := Rupture1.items[Rupture1.ItemIndex];
  if Rupture2.ItemIndex > 0 then
  	begin
  	if Result <> '' then Result := Result + ', ';
    Result := Result + Rupture2.items[Rupture2.ItemIndex];
    end;
  if Rupture3.ItemIndex > 0 then
  	begin
  	if Result <> '' then Result := Result + ', ';
    Result := Result + Rupture3.items[Rupture3.ItemIndex];
    end;
  if Rupture4.ItemIndex > 0 then
  	begin
  	if Result <> '' then Result := Result + ', ';
    Result := Result + Rupture4.items[Rupture4.ItemIndex];
    end;
end;

procedure TOF_CPBALANCETL.AfficherFiltrePourTable(stTableIdx: String);
// Rend visible ou non les champs de filtres concernant la table libre n°inTableIdx
begin
	SetControlVisible('TE_TABLE'+stTableIdx, True);
	SetControlVisible('TE_TABLE'+stTableIdx+'_', True);
	SetControlVisible('E_TABLE'+stTableIdx, True);
	SetControlVisible('E_TABLE'+stTableIdx+'_', True);
	SetControlText('E_TABLE'+stTableIdx,'');
	SetControlText('E_TABLE'+stTableIdx+'_','');
end;

procedure TOF_CPBALANCETL.AfficherFiltres;
// Gère l'affichage des filtres dsur les tables libres en fonction de la
//  sélection des regroupements,
var
	inIdx : Integer;
begin
	// rendre invisible tous les filtres
  For inIdx:=0 to pFiltres.ControlCount-1 do
    pFiltres.Controls[inIdx].visible := False ;
	// afficher ceux sélectionner pour regroupement
  if Rupture1.value <> '' then
		AfficherFiltrePourTable(Copy(Rupture1.value, 3, 1));
  if Rupture2.value <> '' then
    AfficherFiltrePourTable(Copy(Rupture2.value, 3, 1));
  if Rupture3.value <> '' then
    AfficherFiltrePourTable(Copy(Rupture3.value, 3, 1));
  if Rupture4.value <> '' then
    AfficherFiltrePourTable(Copy(Rupture4.value, 3, 1));

end;

Procedure TOF_CPBALANCETL.AffectGroup ( inTag : Integer ) ;
var
  StPlus, StValue, StText, StTemp : string;
  inIdx : integer;
BEGIN
  StPlus := stComboPlus;
  StValue := string (THValComboBox (GetControl('RUPTURE'+IntToStr(inTag))).Value);
  StText := string (THValComboBox (GetControl('RUPTURE'+IntToStr(inTag))).Text);
  For inIdx := 1 to inNbRupture do
    BEGIN
    if inIdx = inTag then continue;
    StTemp := string (THValComboBox (GetControl('RUPTURE'+IntToStr(inIdx))).Value);
    If StTemp <> '' then StPlus := StPlus + ' AND CC_CODE <>"' + StTemp + '"';
    END;
  THValComboBox (GetControl('RUPTURE'+IntToStr(inTag))).Plus := StPlus;
  if StValue<>''
  	then THValComboBox (GetControl('RUPTURE'+InttoStr(inTag))).Value := StValue
    else THValComboBox (GetControl('RUPTURE'+InttoStr(inTag))).ItemIndex := 0;
END;


Procedure TOF_CPBALANCETL.ChangeGroup ( inTag : Integer ) ;
var
    stValue : string;
    inIdx : integer;
    thvcbRupture : THValComboBox;
BEGIN
  thvcbRupture := THValComboBox (GetControl('RUPTURE'+InttoStr(inTag)));
	StValue := string (thvcbRupture.Value);
	if StValue = '' then
    BEGIN
    THEdit (GetControl('XX_VARIABLET'+InttoStr(inTag))).Text := '';
    THEdit (GetControl('XX_VARIABLEL'+InttoStr(inTag))).Text := '';
    For inIdx := inTag + 1 to inNbRupture do
        BEGIN
        THValComboBox (GetControl('RUPTURE'+InttoStr(inIdx))).Enabled := True;
        THValComboBox (GetControl('RUPTURE'+InttoStr(inIdx))).Value := '';
        THValComboBox (GetControl('RUPTURE'+InttoStr(inIdx))).Enabled := False;
        THValComboBox (GetControl('RUPTURE'+InttoStr(inIdx))).Color := clBtnFace;
        THEdit (GetControl('XX_VARIABLET'+InttoStr(inIdx))).Text := '';
        THEdit (GetControl('XX_VARIABLEL'+InttoStr(inIdx))).Text := '';
        END;
    END
	else
    BEGIN
    if inTag < inNbRupture then
      BEGIN
      THValComboBox (GetControl('RUPTURE'+InttoStr(inTag + 1))).Enabled := True;
      THValComboBox (GetControl('RUPTURE'+InttoStr(inTag + 1))).Color := clWindow;
      END;
    THEdit (GetControl('XX_VARIABLET'+InttoStr(inTag))).Text := 'E_TABLE' + Copy(stValue,3,1);
    THEdit (GetControl('XX_VARIABLEL'+InttoStr(inTag))).Text := thvcbRupture.Items[thvcbRupture.ItemIndex];
    END;
END;

procedure TOF_CPBALANCETL.RuptureOnClick(Sender: TObject);
// A la sélection / modification d'un axe de regroupement :
//  - Mise à jour des champs XX_ utilisé par l'état,
//  - Gestion accessibilité en cascade,
//  - Exclusion de la table sélectionnée dans les autres combos pour éviter multi-sélection
//  - Affichage ou non des filtres associés aux tables sélectionnées,
Var
  thvcbRupture : THValComboBox;
begin
  thvcbRupture := THValComboBox(TControl(Sender));
	// Mise à jour des champs XX_ utilisé par l'état, + Gestion accessibilité en cascade
  ChangeGroup(thvcbRupture.Tag);
	// Exclusion de la table sélectionnée dans les autres combos pour éviter multi-sélection
  AffectGroup(thvcbRupture.Tag);
	// Affichage ou non des filtres associés aux tables sélectionnées,
	AfficherFiltres;
end;

procedure TOF_CPBALANCETL.RuptureOnEnter(Sender: TObject);
// A la sélection / modification d'un axe de regroupement :
//  - Exclusion de la table sélectionnée dans les autres combos pour éviter multi-sélection
Var
  thvcbRupture : THValComboBox;
begin
  thvcbRupture := THValComboBox(TControl(Sender));
	// Mise à jour des champs XX_ utilisé par l'état, + Gestion accessibilité en cascade
  AffectGroup(thvcbRupture.Tag);
	// Gestion d'affichage des filtres
end;

procedure TOF_CPBALANCETL.RuptureTypeOnClick(Sender: TObject);
begin
	SetControlText('RUPTURETYPE',RuptureType.value);
end;

function TOF_CPBALANCETL.GetMinMaxCompte(stTable, stCol, stDefaut: String): String;
var
	Q : TQuery;
begin
 	Q := OpenSQL('SELECT ' + stCol +' CODE FROM ' + stTable,True);
  if not Q.Eof
  	then Result := Q.FindField('CODE').asString
	  else Result := stDefaut;
  Ferme(Q);
end;

Initialization
  registerclasses ( [ TOF_CPBALANCETL ] ) ;
end.
