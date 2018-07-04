{***********UNITE*************************************************
Auteur  ...... : Compta
Créé le ...... : 30/07/2002
Modifié le ... : 02/10/2006
Description .. : Source TOF de la FICHE : CPJALCENTRAL ()
Mots clefs ... : TOF;CPJALCENTRAL
*****************************************************************}
Unit CPJALCENTRAL_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MainEAgl, utob, eQRS1, 
{$ELSE}
     db, dbtables, Fe_Main, QRS1,
{$ENDIF}
     forms, sysutils, ComCtrls, Spin, Ent1,
     HCtrls, HEnt1, UTOF,
     Filtre,
     hqry ;

procedure CPLanceFiche_JournalCentral;

Type
  TOF_CPJALCENTRAL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
  private
  	Pages : TPageControl;
  	Affichage : THRadioGroup;
    Exercice, Etab, Devise : THValComboBox;
    QualifPiece : THMultiValcomboBox;
    JournalDe, JournalA, DateComptaDe, DateComptaA, JalExcept : THEdit;
    ANouveaux, Reports : TCheckBox;
    DateDebut, DateFin : TDatetime;
  	AvecAffichage, AvecQualifPiece, AvecReport, AvecAN, XXWhere : THEdit;
    procedure DeviseChanged(Sender : TObject);
    procedure JournalOnExit(Sender: TObject) ;
    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    function  GenererRequeteJAL : String;
    procedure RemplirEDTBALANCE;
    function  GenererRequeteEtat : String;
		function  GetMinMaxJournal(stTable, stCol, stDefaut : String) : String;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  HMsgBox;

//==============================================================================
procedure CPLanceFiche_JournalCentral;
begin
	AGLLanceFiche('CP','CPJALCENTRAL','','','');
end;
//==============================================================================

procedure TOF_CPJALCENTRAL.OnNew ;
begin
  Inherited ;
	// pages standards
	Exercice.Value				:= VH^.Entree.Code ;
  ExoOnChange(nil);
	QualifPiece.value 		:= 'N';
  // pages avancés
	Etab.ItemIndex 				:= 0;
	Devise.ItemIndex 			:= 0;
  DeviseChanged(nil);
	// Pas le choix de l'état
	SetControlEnabled('FETAT', False);
end ;

procedure TOF_CPJALCENTRAL.OnUpdate ;
begin
  Inherited ;
  RemplirEDTBALANCE; // Remplissage table temporaire
	TFQRS1(Ecran).WhereSQL := GenererRequeteEtat; // Mise en place requête finale
end ;

procedure TOF_CPJALCENTRAL.OnLoad ;
var stTypeEcr : String;
begin
  Inherited ;
  XXWhere.Text := '';
  DateDebut := StrToDate(DateComptaDe.Text);
  DateFin := StrToDate(DateComptaA.Text);

	// Auto-remplissage si jorunaux non renseignés
	if (Trim(JournalDe.Text) = '') then JournalDe.Text := GetMinMaxJournal('JOURNAL', 'MIN(J_JOURNAL)','0');
	if (Trim(JournalA.Text) = '') then JournalA.Text := GetMinMaxJournal('JOURNAL', 'MAX(J_JOURNAL)','ZZZ');

	// libellés pour les types d'écritures...
  AvecQualifPiece.Text := '';
	stTypeEcr := QualifPiece.Value;
	if stTypeEcr <> '' then
    while stTypeEcr <> '' do
	    AvecQualifPiece.Text := AvecQualifPiece.Text
      			+ RechDom(QualifPiece.DataType,ReadTokenSt(stTypeEcr),False)
            + ' '
  else
  	AvecQualifPiece.Text := '<<Tous>>';

	// Devise
  if (Devise.ItemIndex > 0) and (Affichage.value <> 'EURO')
    then AvecAffichage.Text := Devise.Items[Devise.ItemIndex]
		else AvecAffichage.Text := 'Euro';

	// Option Reports
	if ANouveaux.checked then AvecAN.Text := 'Oui' else AvecAN.Text := 'Non';

	// Option Reports
	if Reports.checked then AvecReport.Text := 'X' else AvecReport.Text := '';

end ;

procedure TOF_CPJALCENTRAL.OnArgument (S : String ) ;
var
	i : Integer;
begin
  Inherited ;
{$IFDEF EAGLCLIENT}
	TFQRS1(Ecran).FNomFiltre	:=	GetLeNom(TFQRS1(Ecran).Name); // Correction eAGL pour filtres a oter en 5.4.5
  ChargeFiltre(TFQRS1(Ecran).FNomFiltre,THComboBox(GetControl('FFILTRES')),TPageControl(GetControl('PAGES')));
{$ELSE}
{$ENDIF}
  Pages := TPageControl(GetControl('Pages'));
	// Contrôles critères et options - Page standard
  JournalDe				:= THEdit(GetControl('JOURNALDE'));
  JournalA	 			:= THEdit(GetControl('JOURNALA'));
	Exercice				:= THValComboBox(Getcontrol('E_EXERCICE'));
  DateComptaDe		:= THEdit(GetControl('E_DATECOMPTABLE'));
  DateComptaA			:= THEdit(GetControl('E_DATECOMPTABLE_'));
  QualifPiece			:= THMultiValComboBox(Getcontrol('E_QUALIFPIECE'));
	// Contrôles critères et options - Page avancés
	Etab						:= THValComboBox(Getcontrol('E_ETABLISSEMENT'));
  Devise					:= THValComboBox(Getcontrol('E_DEVISE'));
  JalExcept				:= THEdit(GetControl('JALEXCEPT'));
  Affichage	 			:= THRadioGroup(GetControl('AFFICHAGE'));
	Reports					:= TCheckBox(GetControl('REPORTS'));
	ANouveaux				:= TCheckBox(GetControl('ANOUVEAUX'));

  // Contrôles invisibles utilisés dans l'état
	AvecQualifPiece	:= THEdit(GetControl('AVECQUALIFPIECE'));
  AvecAN					:= THEdit(GetControl('AVECANOUVEAUX'));
  AvecAffichage		:= THEdit(GetControl('AVECAFFICHAGE'));
  AvecReport	    := THEdit(GetControl('AVECREPORT'));
  XXWhere 				:= THEdit(GetControl('XX_WHERE'));
	// Evènements
	JournalDe.OnExit 				:= JournalOnExit;
	JournalA.OnExit 				:= JournalOnExit;
	DateComptaDe.OnExit  		:= DateOnExit;
	DateComptaA.OnExit 	 		:= DateOnExit;
  Exercice.onChange				:= ExoOnChange;
	Devise.OnChange					:= DeviseChanged;
  // Ordre des pages
  for i := 0 to Pages.PageCount - 1 do
    Pages.Pages[i].PageIndex := Pages.Pages[i].Tag;
end ;

procedure TOF_CPJALCENTRAL.DateOnExit(Sender: TObject);
begin
  if (THEdit(Sender).Name = DateComptaDe.Name) and ((StrToDate(DateComptaDe.Text) < DateDebut) or (StrToDate(DateComptaDe.Text) > DateFin))
	  then DateComptaDe.Text := DateToStr(DateDebut);
  if (THEdit(Sender).Name = DateComptaA.Name) and ((StrToDate(DateComptaA.Text) < DateDebut) or (StrToDate(DateComptaA.Text) > DateFin)) then
    DateComptaA.Text := DateToStr(DateFin);
  if (THEdit(Sender).Name = DateComptaDe.Name) and (StrToDate(DateComptaDe.Text) > StrToDate(DateComptaA.Text)) then
    DateComptaA.Text := DateComptaDe.Text
  else if (THEdit(Sender).Name = DateComptaA.Name) and (StrToDate(DateComptaDe.Text) > StrToDate(DateComptaA.Text)) then
    DateComptaDe.Text := DateComptaA.Text;
end;

procedure TOF_CPJALCENTRAL.ExoOnChange(Sender: TObject);
begin
  ExoToDates(Exercice.Value, DateComptaDe, DateComptaA);
  DateDebut := StrToDate(DateComptaDe.Text);
  DateFin := StrToDate(DateComptaA.Text);
end;

procedure TOF_CPJALCENTRAL.JournalOnExit(Sender: TObject);
begin
	if Sender = nil then Exit;
  if (THEdit(Sender).Name = JournalDe.Name) and (JournalDe.Text > JournalA.Text) and (JournalA.Text <> '')
  	then JournalA.Text := JournalDe.Text
	else
   	if (THEdit(Sender).Name = JournalA.Name) and (JournalDe.Text > JournalA.Text)
    then JournalDe.Text := JournalA.Text;
end;

procedure TOF_CPJALCENTRAL.DeviseChanged(Sender: TObject);
begin
	if Devise.ItemIndex > 0  then Affichage.Enabled := True
  else
  	begin
    Affichage.Value := 'EURO';
	  Affichage.Enabled := False;
    end;
end;

function TOF_CPJALCENTRAL.GenererRequeteJAL: String;
var
	listeJalExcept : String;
begin
// SELECT -> Champs d'origine
	result := result + ' SELECT "' + V_PGI.User + '", J_JOURNAL, J_LIBELLE';
// Table Journal
  result := result + ' FROM JOURNAL';
// Conditions du WHERE (sur les comptes uniquement donc)
  // Comptes de ... à
  result := result + ' WHERE J_JOURNAL >="' + JournalDe.Text + '"';
  result := result + ' AND J_JOURNAL <="' + JournalA.Text + '"';
  // Nature
  result := result + ' AND J_JOURNAL<>"&#@"';
  result := result + ' AND J_NATUREJAL<>"CLO"';
  result := result + ' AND J_NATUREJAL<>"ODA"';
  result := result + ' AND J_NATUREJAL<>"ANA"';
	if not ANouveaux.Checked then
 		result := result + ' AND J_NATUREJAL<>"ANO"';
  // On complète les conditions : ---> Jounaux d'exception
  if JalExcept.Text <> '' then
  	begin
    listeJalExcept := FindEtReplace(JalExcept.Text,',',';', True);
		listeJalExcept := FindEtReplace(listeJalExcept,';','","', True);
  	result := result + ' AND J_JOURNAL NOT IN ("' + listeJalExcept + '")';
    end;
// Group by
	result := result + ' GROUP BY J_JOURNAL, J_LIBELLE';
end;

procedure TOF_CPJALCENTRAL.RemplirEDTBALANCE;
var	stReq : String;
begin
	// Effacement des enregistrements present pour le user
	ExecuteSQL('DELETE FROM CEDTBALANCE WHERE CED_USER = "' + V_PGI.User + '"');
	// Création des enregistrements
	stReq := 'INSERT INTO CEDTBALANCE (CED_USER, CED_COMPTE, CED_LIBELLE) ' + GenererRequeteJAL;
	ExecuteSQL(stReq);
end;

function TOF_CPJALCENTRAL.GenererRequeteEtat: String;
begin
// Champs
	if Affichage.Value = 'EURO'
  	then Result := 'SELECT CED_COMPTE JOURNAL, CED_LIBELLE LIBELLE, Sum(E_Debit) D, Sum(E_Credit) C'
	  else Result := 'SELECT CED_COMPTE JOURNAL, CED_LIBELLE LIBELLE, Sum(E_DebitDev) D, Sum(E_CreditDev) C';
// Table Compte
  result := result + ' FROM CEDTBALANCE';
// Table Ecriture
  result := result + ' LEFT OUTER JOIN ECRITURE ON CED_COMPTE = E_JOURNAL ';
  // Critères Ecritures (inclus dans Left join car jointure externe)
  result := result + FindEtReplace(RecupWhereCritere(Pages),'WHERE','AND', False);
// Group by
	result := result + ' GROUP BY CED_COMPTE, CED_LIBELLE';
// Order by
	result := result + ' ORDER BY CED_COMPTE, CED_LIBELLE';
end;

function TOF_CPJALCENTRAL.GetMinMaxJournal(stTable, stCol, stDefaut: String): String;
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
  registerclasses ( [ TOF_CPJALCENTRAL ] ) ;
end.
