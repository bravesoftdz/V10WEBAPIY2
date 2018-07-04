{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/12/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPJALCPTGENE ()
Mots clefs ... : TOF;CPJALCPTGENE
*****************************************************************}
Unit CPJALCPTGENE_TOF ;

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

  procedure CPLanceFiche_CPJALCPTGENE;

Type
  TOF_CPJALCPTGENE = Class (TOF_Meth)
  private
    {Critères}
    JournalDe:      THEdit;
    JournalA:       THEdit;
    Exercice:       THValComboBox;
    TypeEcrNormal:  TCheckBox;
    TypeEcrPrev:    TCheckBox;
    TypeEcrSitu:    TCheckBox;
    DateComptaDe:   THEdit;
    DateComptaA:    THEdit;
    {Critères avancés}
    Devise:           THValcomboBox;
    NumeroPieceDe:    THEdit;
    NumeroPieceA:     THEdit;
    RefInterne:       THEdit;
    EcrValide:        TCheckBox;
    EcrRevision:      TCheckBox;
    EcrANouveau:      TCheckBox;
    {Options d'édition}
    ChoixMonnaie:     THRadioGroup;
    ChkSymboleDevise: TCheckBox;
    {Dev}
    TitreEtat:        THEdit;
    AvecChoixMonnaie: THEdit;
    SymboleDevise:    THEdit;
    TypeEcriture:     THEdit;

    procedure ExoOnChange(Sender: TObject) ;
    procedure ChoixMonnaieChange(Sender: TObject);

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
 

procedure CPLanceFiche_CPJALCPTGENE;
begin
  AGLLanceFiche('CP','CPJALCPTGENE_QR1','','','');
end;

procedure TOF_CPJALCPTGENE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPJALCPTGENE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPJALCPTGENE.OnUpdate ;
var
  ClauseWhere: String;
  ClauseQuali: String;
begin
  Inherited ;

  ClauseQuali := '';
  if TypeEcrNormal.Checked then
    ClauseQuali := ClauseQuali + ' Or E_QUALIFPIECE="N"';
  if TypeEcrPrev.Checked then
    ClauseQuali := ClauseQuali + ' Or E_QUALIFPIECE="P"';
  if TypeEcrSitu.Checked then
    ClauseQuali := ClauseQuali + ' Or E_QUALIFPIECE="S" Or E_QUALIFPIECE="U"';
  if ClauseQuali = '' then
    ClauseQuali := ' And E_QUALIFPIECE="N"'
  else
  begin
    System.Delete(ClauseQuali, 1, Length(' Or '));  {Retire le 1er OR}
    ClauseQuali := '('+ClauseQuali+')';
  end;

  ClauseWhere := ' And J_JOURNAL<>"&#@" And J_NATUREJAL<>"CLO"' +
                 ' And J_NATUREJAL<>"ODA" And J_NATUREJAL<>"ANA"';
  if RefInterne.Text<>'' then
    ClauseWhere := ClauseWhere + ' And UPPER(E_REFINTERNE) like "'+TraduitJoker(RefInterne.Text)+'" ';

  if not EcrANouveau.Checked then
    ClauseWhere := ClauseWhere + ' And J_NATUREJAL<>"ANO"';

  if EcrRevision.State = cbGrayed then
    ClauseWhere := ClauseWhere + ' And ('+ClauseQuali+' Or E_QUALIFPIECE="R")'
  else if EcrRevision.Checked then
    ClauseWhere := ClauseWhere + ' And E_QUALIFPIECE="R"'
  else if not EcrRevision.Checked then
    ClauseWhere := ClauseWhere + ' And '+ClauseQuali;

  ClauseWhere := ClauseWhere + ' And E_QUALIFPIECE<>"C"';
  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL+' '+ClauseWhere;
end ;

procedure TOF_CPJALCPTGENE.OnLoad ;
var
  RDev: RDevise;
  TypeEcr:      String;
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

  {si aucun type d'écriture n'est choisi, je force le type normal}
  if (not TypeEcrNormal.Checked) and (not TypeEcrPrev.Checked) and
     (not TypeEcrSitu.Checked) and (not EcrRevision.Checked) then
    TypeEcrNormal.Checked := True;

  TypeEcr := '';
  if TypeEcrNormal.Checked then
    TypeEcr := TypeEcr+', '+TraduireMemoire('Normal');
  if TypeEcrPrev.Checked then
    TypeEcr := TypeEcr+', '+TraduireMemoire('Prévision');
  if TypeEcrSitu.Checked then
    TypeEcr := TypeEcr+', '+TraduireMemoire('Situation et Simulation');
  if Length(TypeEcr) > 1 then
    System.Delete(TypeEcr, 1, Length(', '));  {Retire la 1ere virgule}
  TypeEcriture.Text := TypeEcr;

  RecupCritEdt;
end ;

procedure TOF_CPJALCPTGENE.OnArgument (S : String ) ;
begin
  Inherited ;

  Pages          := TPageControl(Getcontrol('PAGES', true));
  {Critères}
  JournalDe      := THEdit(GetControl('E_JOURNAL', true));
  JournalA       := THEdit(GetControl('E_JOURNAL_', true));
  Exercice       := THValComboBox(Getcontrol('E_EXERCICE', true));
  TypeEcrNormal  := TCheckBox(GetControl('NORMAL', true));
  TypeEcrPrev    := TCheckBox(GetControl('PREVISION', true));
  TypeEcrSitu    := TCheckBox(GetControl('SITUATION', true));
  DateComptaDe   := THEdit(GetControl('E_DATECOMPTABLE', true));
  DateComptaA    := THEdit(GetControl('E_DATECOMPTABLE_', true));
  {Critères avancés}
  Devise         := THValcomboBox(Getcontrol('E_DEVISE', true));
  NumeroPieceDe  := THEdit(GetControl('E_NUMEROPIECE', true));
  NumeroPieceA   := THEdit(GetControl('E_NUMEROPIECE_', true));
  RefInterne     := THEdit(Getcontrol('REFERENCEINTERNE', true));
  EcrValide      := TCheckBox(Getcontrol('E_VALIDE', true));
  EcrRevision    := TCheckBox(Getcontrol('INCECREV', true));
  EcrANouveau    := TCheckBox(Getcontrol('INCJALANO', true));
  {Options d'édition}
  ChoixMonnaie      := THRadioGroup(Getcontrol('CHOIXMONNAIE', true));
  ChkSymboleDevise  := TCheckBox(Getcontrol('CHKSYMBOLEDEVISE', true));
  {Dev}
  AvecChoixMonnaie  := THEdit(Getcontrol('AVECCHOIXMONNAIE', true));
  SymboleDevise     := THEdit(Getcontrol('SYMBOLEDEVISE', true));
  TitreEtat         := THEdit(Getcontrol('TITREETAT', true));
  TypeEcriture      := THEdit(Getcontrol('TYPEECRITURE', true));

  Exercice.onChange    := ExoOnChange;
  ChoixMonnaie.OnClick := ChoixMonnaieChange;

  if (ComboEtab.ItemIndex = -1) and (ComboEtab.Items.Count > 0) then
    ComboEtab.ItemIndex := 0;

  Devise.ItemIndex := 0;
  ChoixMonnaieChange(ChoixMonnaie);

  if (CtxPCl in V_PGI.PgiContexte) and (VH^.CPExoRef.Code <>'') then
    Exercice.Value := VH^.CPExoRef.Code
  else
    Exercice.Value := VH^.Entree.Code;

  Ecran.Caption := TraduireMemoire('Journal périodique par compte général');
  Ecran.HelpContext := 7409000;
  UpdateCaption(Ecran);
  TitreEtat.Text := Ecran.Caption;
end ;

procedure TOF_CPJALCPTGENE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPJALCPTGENE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPJALCPTGENE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPJALCPTGENE.ChoixMonnaieChange(Sender: TObject);
begin
  {Choix monnaie possible uniquement si une devise est sélectionnée}
  if (Devise.ItemIndex=0) or (Devise.Value=V_PGI.DevisePivot) then begin
    if ChoixMonnaie.ItemIndex <> 0 then
      HShowMessage('0;'+Ecran.Caption+';Vous devez d''abord sélectionner une devise particulière.;E;O;O;O;','','');
    ChoixMonnaie.ItemIndex := 0;
    end;
  AvecChoixMonnaie.Text := ChoixMonnaie.Value;
end;

procedure TOF_CPJALCPTGENE.ExoOnChange(Sender: TObject);
begin
  ExoToDates(Exercice.Value, DateComptaDe, DateComptaA);
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

procedure TOF_CPJALCPTGENE.RecupCritEdt;
begin
  PositionneFourchetteST(JournalDe, JournalA, tzJournal);

  If NumeroPieceDe.Text = '' then
    NumeroPieceDe.Text := '0';
  If NumeroPieceA.Text = '' then
    NumeroPieceA.Text := '999999999';
end;

Initialization
  registerclasses ( [ TOF_CPJALCPTGENE ] ) ; 
end.

