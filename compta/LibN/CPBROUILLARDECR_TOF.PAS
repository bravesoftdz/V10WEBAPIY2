{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 18/11/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBROUILLARDECR ()
Mots clefs ... : TOF;CPBROUILLARDECR;FPD2867 
*****************************************************************}
Unit CPBROUILLARDECR_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
     Fe_main,
     QRS1,
{$else}
     eMul,
     uTob,
     Maineagl,
     eQRS1,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HMsgBox,
     UTOF,
     TofMeth,
     uLibExercice,
     HCompte,
     SAISUTIL;

  procedure CPLanceFiche_CPBROUILLARDECR(TypeEdition: Char);

Type
  TOF_CPBROUILLARDECR = Class (TOF_Meth)
  private
    FTypeEdition: Char;

    {Critères}
    Exercice:       THValComboBox;
    DateComptaDe:   THEdit;
    DateComptaA:    THEdit;
    JournalDe:      THEdit;
    JournalA:       THEdit;
    NumeroPieceDe:  THEdit;
    NumeroPieceA:   THEdit;
    {Critères avancés}
    RefInterne:       THEdit;
    EcrValide:        TCheckBox;
    EcrANouveau:      TCheckBox;
    Devise:           THValcomboBox;
    {Options d'édition}
    ChoixMonnaie:     THRadioGroup;
    ChkSymboleDevise: TCheckBox;
    ChkTriDateComp:   TCheckBox;
    ChkTriJournal:    TCheckBox;
    {Dev}
    TitreEtat:        THEdit;
    SymboleDevise:    THEdit;
    AvecChoixMonnaie: THEdit;
    NumExercice:      THEdit;

    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    procedure JournalOnExit(Sender: TObject) ;
    procedure ChoixMonnaieChange(Sender: TObject);
    procedure ChkTriDateCompOnClick(Sender: TObject);
    procedure ChkTriJournalOnClick(Sender: TObject);

    procedure RecupCritEdt;
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Hent1,
  Ent1;

procedure CPLanceFiche_CPBROUILLARDECR(TypeEdition: Char);
begin
  {TypeEdition = R: Revision, S: Simulation, N: Courant}
  AGLLanceFiche('CP','CPBROUILECR_QR1','','',TypeEdition);
end;

procedure TOF_CPBROUILLARDECR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDECR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDECR.OnUpdate ;
var
  ClauseWhere: String;
  ClauseOrder: String;
begin
  Inherited ;

  {Clause Where}
  ClauseWhere := '';
  if RefInterne.Text<>'' then
    ClauseWhere := ClauseWhere + ' And Upper(E_REFINTERNE) like "'+TraduitJoker(RefInterne.Text)+'"';

  if EcrValide.State <> CbGrayed then begin
    if EcrValide.Checked then
      ClauseWhere := ClauseWhere + ' And E_VALIDE = "X"'
    else
      ClauseWhere := ClauseWhere + ' And E_VALIDE = "-"';
  end;

  if EcrANouveau.Checked then
    ClauseWhere := ClauseWhere + ' And (E_ECRANOUVEAU="N" or E_ECRANOUVEAU="H" or E_ECRANOUVEAU="OAN")'
  else
    ClauseWhere := ClauseWhere + ' And E_ECRANOUVEAU="N"';

  if (FTypeEdition in ['N', 'R', 'S']) then
    ClauseWhere  := ClauseWhere + ' And E_QUALIFPIECE="'+FTypeEdition+'"';

  {Clause Order by}
  if ChkTriJournal.Checked then
    ClauseOrder := ' Order by E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE'
  else if ChkTriDateComp.Checked then
    ClauseOrder := ' Order by E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE'
  else
    ClauseOrder := ' Order by E_NUMEROPIECE, E_NUMLIGNE, E_JOURNAL, E_DATECOMPTABLE';

  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL +
    ' ' + ClauseWhere + ' ' + ClauseOrder;
end ;

procedure TOF_CPBROUILLARDECR.OnLoad ;
var
  RDev: RDevise;
begin
  Inherited ;

  SymboleDevise.Text := '';
  if ChkSymboleDevise.Checked then begin
    if (ChoixMonnaie.ItemIndex = 0) or (Devise.Value=V_PGI.DevisePivot) then begin
      SymboleDevise.Text := ' '+V_PGI.SymbolePivot;
    end
    else begin
      RDev.Code := Devise.Value;
      GetInfosDevise(RDev);
      SymboleDevise.Text := ' '+RDev.Symbole;
    end;
  end;
  RecupCritEdt;
end ;

procedure TOF_CPBROUILLARDECR.OnArgument (S : String ) ;
begin
  Inherited ;

  if Length(S) = 0 then
    FTypeEdition := ' '
  else
    FTypeEdition := S[1];

  Pages          := TPageControl(Getcontrol('PAGES', true));
  {Critères}
  Exercice       := THValComboBox(Getcontrol('COMBOEXERCICE', true));
  DateComptaDe   := THEdit(GetControl('E_DATECOMPTABLE', true));
  DateComptaA    := THEdit(GetControl('E_DATECOMPTABLE_', true));
  JournalDe      := THEdit(GetControl('E_JOURNAL', true));
  JournalA       := THEdit(GetControl('E_JOURNAL_', true));
  NumeroPieceDe  := THEdit(GetControl('E_NUMEROPIECE', true));
  NumeroPieceA   := THEdit(GetControl('E_NUMEROPIECE_', true));
  {Critères avancés}
  RefInterne     := THEdit(Getcontrol('REFINTERNE', true));
  EcrValide      := TCheckBox(Getcontrol('ECRITUREVALIDE', true));
  EcrANouveau    := TCheckBox(Getcontrol('ECRANOUVEAU', true));
  Devise         := THValcomboBox(Getcontrol('E_DEVISE', true));
  {Options d'édition}
  ChoixMonnaie      := THRadioGroup(Getcontrol('CHOIXMONNAIE', true));
  ChkSymboleDevise  := TCheckBox(Getcontrol('CHKSYMBOLEDEVISE', true));
  ChkTriDateComp    := TCheckBox(Getcontrol('TRIDATECOMPTABLE', true));
  ChkTriJournal     := TCheckBox(Getcontrol('TRIJOURNAL', true));
  {Dev}
  AvecChoixMonnaie  := THEdit(Getcontrol('AVECCHOIXMONNAIE', true));
  NumExercice       := THEdit(Getcontrol('E_EXERCICE', true));
  SymboleDevise     := THEdit(Getcontrol('SYMBOLEDEVISE', true));
  TitreEtat         := THEdit(Getcontrol('TITREETAT', true));

  DateComptaDe.OnExit   := DateOnExit;
  DateComptaA.OnExit    := DateOnExit;
  JournalDe.OnExit      := JournalOnExit;
  JournalA.OnExit       := JournalOnExit;
  Exercice.onChange     := ExoOnChange;
  ChoixMonnaie.OnClick  := ChoixMonnaieChange;
  ChkTriDateComp.OnClick:= ChkTriDateCompOnClick;
  ChkTriJournal.OnClick := ChkTriJournalOnClick;   

  CInitComboExercice(Exercice);       // Init de la combo Exercice en Relatif
  if (ComboEtab.ItemIndex = -1) and (ComboEtab.Items.Count > 0) then
    ComboEtab.ItemIndex := 0;

  Devise.ItemIndex := 0;
  ChoixMonnaieChange(ChoixMonnaie);

  if (CtxPCl in V_PGI.PgiContexte) and (VH^.CPExoRef.Code <>'') then
    Exercice.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
  else
    Exercice.Value := CExerciceVersRelatif(VH^.Entree.Code) ;
     
  case FTypeEdition of
   'R' : begin      {Revision}
         Ecran.Caption     := TraduireMemoire('Brouillard des écritures de révision');
         TFQRS1(Ecran).FNomFiltre:='BROUILREV' ;
         Ecran.HelpContext := 7670000 ;
         EcrValide.Visible := False;
         EcrValide.State   := cbGrayed;
         end;
   'S' : begin      {Simulation}
         Ecran.Caption     := TraduireMemoire('Brouillard des écritures de simulation');
         TFQRS1(Ecran).FNomFiltre:='BROUILSIM' ;
         Ecran.HelpContext := 7289000 ;
         EcrValide.Visible := False;
         EcrValide.State   := cbGrayed;
         end;
   'N' : begin      {Courant}
         Ecran.Caption     := TraduireMemoire('Brouillard des écritures courantes');
         TFQRS1(Ecran).FNomFiltre:='BROUILNOR' ;
         Ecran.HelpContext := 7268000 ;
         end;
  end;
  UpdateCaption(Ecran);
  TitreEtat.Text := Ecran.Caption;
end ;

procedure TOF_CPBROUILLARDECR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDECR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDECR.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDECR.ChoixMonnaieChange(Sender: TObject);
begin
  {Choix monnaie possible uniquement si une devise est sélectionnée}
  if (Devise.ItemIndex=0) or (Devise.Value=V_PGI.DevisePivot) then begin
    if ChoixMonnaie.ItemIndex <> 0 then
      HShowMessage('0;'+Ecran.Caption+';Vous devez d''abord sélectionner une devise particulière.;E;O;O;O;','','');
    ChoixMonnaie.ItemIndex := 0;
    end;
  AvecChoixMonnaie.Text := ChoixMonnaie.Value;
end;

procedure TOF_CPBROUILLARDECR.DateOnExit(Sender: TObject);
var
  DateD: TDateTime;
  DateF: TDateTime;
begin
  DateD := StrToDate(DateComptaDe.Text);
  DateF := StrToDate(DateComptaA.Text);

  DoDateOnExit(THEdit(Sender), DateComptaDe, DateComptaA, DateD, DateF);
end;

procedure TOF_CPBROUILLARDECR.ExoOnChange(Sender: TObject);
begin
  CExoRelatifToDates(Exercice.Value, DateComptaDe, DateComptaA);
  NumExercice.Text := CRelatifVersExercice(Exercice.Value);
end;

procedure TOF_CPBROUILLARDECR.JournalOnExit(Sender: TObject);
begin
  DoJalOnExit(THEdit(Sender), JournalDe, JournalA);
end;

{Voir unité UTILEDT}
procedure PositionneFourchetteST(TC1,TC2 : THEdit; tt: TZoomTable) ;
var
  St:          String;
  Q:           TQuery;
begin
  if (TC1.Text='') And (TC2.Text='') then begin
    Case CaseFic(tt) Of
      fbGene : St:='SELECT MIN(G_GENERAL), Max(G_GENERAL) FROM GENERAUX WHERE G_FERME="-" ' ;
      fbAux : St:='SELECT MIN(T_AUXILIAIRE), Max(T_AUXILIAIRE) FROM TIERS WHERE T_FERME="-" ' ;
      fbJal : St:='SELECT MIN(J_JOURNAL), Max(J_JOURNAL) FROM JOURNAL WHERE J_FERME="-" ' ;
      fbAxe1..fbAxe5 : St:='SELECT MIN(S_SECTION), Max(S_SECTION) FROM SECTION WHERE S_FERME="-" ' ;
      fbBudGen : St:='SELECT MIN(BG_BUDGENE), Max(BG_BUDGENE) FROM BUDGENE WHERE BG_FERME="-" ' ;
      fbBudJal : St:='SELECT MIN(BJ_BUDJAL), Max(BJ_BUDJAL) FROM BUDJAL WHERE BJ_FERME="-" ' ;
      fbBudSec1..fbBudSec5 : St:='SELECT MIN(BS_BUDSECT), Max(BS_BUDSECT) FROM BUDSECT WHERE BS_FERME="-" ' ;
      fbNatCpt : St:='SELECT MIN(NT_NATURE), Max(NT_NATURE) FROM NATCPTE WHERE NT_SOMMEIL="-" ' ;
      end;
    St:=St+RecupWhere(tt) ;

    {**************************************************************************}
    {Il faut modifier l'unité HCompte, la fonction RecupWhere car la nature
    d'un jnl de vente est VTE et pas VEN}
    {**************************************************************************}
    if tt = tzJvente then
      St := StringReplace(St, '"VEN"', '"VTE"', []);
    {**************************************************************************}
    {**************************************************************************}
    Q:=OpenSQL(St,TRUE) ;
    if not Q.EOF then begin
      TC1.Text := Q.Fields[0].AsString;
      TC2.Text := Q.Fields[1].AsString;
      end;
    Ferme(Q);
    end;
end;

procedure TOF_CPBROUILLARDECR.RecupCritEdt;
begin
  PositionneFourchetteST(JournalDe, JournalA, tzJournal);

  If NumeroPieceDe.Text = '' then
    NumeroPieceDe.Text := '0';
  If NumeroPieceA.Text = '' then
    NumeroPieceA.Text := '999999999';
end;

procedure TOF_CPBROUILLARDECR.ChkTriDateCompOnClick(Sender: TObject);
begin
  if ChkTriDateComp.Checked then
    ChkTriJournal.Checked := False;
end;

procedure TOF_CPBROUILLARDECR.ChkTriJournalOnClick(Sender: TObject);
begin
  if ChkTriJournal.Checked then
    ChkTriDateComp.Checked := False;
end;

Initialization
  registerclasses ( [ TOF_CPBROUILLARDECR ] ) ; 
end.

