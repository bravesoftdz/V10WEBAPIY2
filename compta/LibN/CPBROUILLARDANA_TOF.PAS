{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 28/11/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBROUILLARDANA ()
Mots clefs ... : TOF;CPBROUILLARDANA
*****************************************************************}
Unit CPBROUILLARDANA_TOF ;

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


  procedure CPLanceFiche_CPBROUILLARDANA;

Type
  TOF_CPBROUILLARDANA = Class (TOF_Meth)
  private
    {Critères}
    Axe:            THValComboBox;
    Exercice:       THValComboBox;
    DateComptaDe:   THEdit;
    DateComptaA:    THEdit;
    JournalDe:      THEdit;
    JournalA:       THEdit;
    {Critères avancés}
    RefInterne:       THEdit;
    NumeroPieceDe:    THEdit;
    NumeroPieceA:     THEdit;
    TitreEtat:        THEdit;
    {Options d'édition}
    ChkTriDateComp:   TCheckBox;
    ChkTriJournal:    TCheckBox;
    {Dev}
    NumExercice:        THEdit;

    procedure ExoOnChange(Sender: TObject) ;
    procedure DateOnExit(Sender: TObject) ;
    procedure JournalOnExit(Sender: TObject) ;
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


procedure CPLanceFiche_CPBROUILLARDANA;
begin
  AGLLanceFiche('CP','CPBROUILANA_QR1','','','');
end;

procedure TOF_CPBROUILLARDANA.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDANA.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDANA.OnUpdate ;
var
  ClauseWhere: String;
  ClauseOrder: String;
begin
  Inherited ;

  {Clause Where}
  ClauseWhere := '';
  if RefInterne.Text<>'' then
    ClauseWhere := ClauseWhere + ' And Upper(Y_REFINTERNE) like "'+TraduitJoker(RefInterne.Text)+'"';

  ClauseWhere  := ClauseWhere +
    ' And Y_TYPEANALYTIQUE="X"' +
    ' And Y_QUALIFPIECE="N"';

  {Clause Order by}
  if ChkTriJournal.Checked then
    ClauseOrder := ' Order by Y_JOURNAL, Y_NUMEROPIECE, Y_NUMLIGNE'
  else if ChkTriDateComp.Checked then
    ClauseOrder := ' Order by Y_DATECOMPTABLE, Y_JOURNAL, Y_NUMEROPIECE, Y_NUMLIGNE'
  else
    ClauseOrder := ' Order by Y_NUMEROPIECE, Y_NUMLIGNE, Y_JOURNAL, Y_DATECOMPTABLE';

  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL +
    ' ' + ClauseWhere + ' ' + ClauseOrder;
end ;

procedure TOF_CPBROUILLARDANA.OnLoad ;
begin
  Inherited ;

  RecupCritEdt;
end ;

procedure TOF_CPBROUILLARDANA.OnArgument (S : String ) ;
begin
  Inherited ;

  Pages          := TPageControl(Getcontrol('PAGES', true));
  {Critères}
  Axe            := THValComboBox(Getcontrol('Y_AXE', true));
  Exercice       := THValComboBox(Getcontrol('COMBOEXERCICE', true));
  DateComptaDe   := THEdit(GetControl('Y_DATECOMPTABLE', true));
  DateComptaA    := THEdit(GetControl('Y_DATECOMPTABLE_', true));
  JournalDe      := THEdit(GetControl('Y_JOURNAL', true));
  JournalA       := THEdit(GetControl('Y_JOURNAL_', true));
  {Critères avancés}
  RefInterne     := THEdit(Getcontrol('REFINTERNE', true));
  NumeroPieceDe  := THEdit(GetControl('Y_NUMEROPIECE', true));
  NumeroPieceA   := THEdit(GetControl('Y_NUMEROPIECE_', true));
  {Options d'édition}
  ChkTriDateComp    := TCheckBox(Getcontrol('TRIDATECOMPTABLE', true));
  ChkTriJournal     := TCheckBox(Getcontrol('TRIJOURNAL', true));
  {Dev}
  NumExercice       := THEdit(Getcontrol('Y_EXERCICE', true));
  TitreEtat         := THEdit(Getcontrol('TITREETAT', true));

  DateComptaDe.OnExit   := DateOnExit;
  DateComptaA.OnExit    := DateOnExit;
  JournalDe.OnExit      := JournalOnExit;
  JournalA.OnExit       := JournalOnExit;
  Exercice.onChange     := ExoOnChange;
  ChkTriDateComp.OnClick:= ChkTriDateCompOnClick;
  ChkTriJournal.OnClick := ChkTriJournalOnClick;

  CInitComboExercice(Exercice);       // Init de la combo Exercice en Relatif
  if (ComboEtab.ItemIndex = -1) and (ComboEtab.Items.Count > 0) then
    ComboEtab.ItemIndex := 0;

  if (Axe.ItemIndex = -1) and (Axe.Items.Count > 0) then
    Axe.ItemIndex := 0;

  if (CtxPCl in V_PGI.PgiContexte) and (VH^.CPExoRef.Code <>'') then
    Exercice.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
  else
    Exercice.Value := CExerciceVersRelatif(VH^.Entree.Code) ;

  Ecran.Caption     := TraduireMemoire('Brouillard des écritures analytiques');
  Ecran.HelpContext := 7385000 ;
  UpdateCaption(Ecran);
  TitreEtat.Text := Ecran.Caption;
end ;

procedure TOF_CPBROUILLARDANA.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDANA.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDANA.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDANA.ChkTriDateCompOnClick(Sender: TObject);
begin
  if ChkTriDateComp.Checked then
    ChkTriJournal.Checked := False;
end;

procedure TOF_CPBROUILLARDANA.ChkTriJournalOnClick(Sender: TObject);
begin
  if ChkTriJournal.Checked then
    ChkTriDateComp.Checked := False;
end;

procedure TOF_CPBROUILLARDANA.DateOnExit(Sender: TObject);
var
  DateD: TDateTime;
  DateF: TDateTime;
begin
  DateD := StrToDate(DateComptaDe.Text);
  DateF := StrToDate(DateComptaA.Text);

  DoDateOnExit(THEdit(Sender), DateComptaDe, DateComptaA, DateD, DateF);
end;

procedure TOF_CPBROUILLARDANA.ExoOnChange(Sender: TObject);
begin
  CExoRelatifToDates(Exercice.Value, DateComptaDe, DateComptaA);
  NumExercice.Text := CRelatifVersExercice(Exercice.Value);
end;

procedure TOF_CPBROUILLARDANA.JournalOnExit(Sender: TObject);
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

procedure TOF_CPBROUILLARDANA.RecupCritEdt;
begin
  PositionneFourchetteST(JournalDe, JournalA, tzJAna);

  If NumeroPieceDe.Text = '' then
    NumeroPieceDe.Text := '0';
  If NumeroPieceA.Text = '' then
    NumeroPieceA.Text := '999999999';
end;

Initialization
  registerclasses ( [ TOF_CPBROUILLARDANA ] ) ; 
end.

