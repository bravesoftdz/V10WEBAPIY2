{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 07/12/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPJALOUVCLO ()
Mots clefs ... : TOF;CPJALOUVCLO
*****************************************************************}
Unit CPJALOUVCLO_TOF ;

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
     SAISUTIL,
     HTB97,
     AGLInit,       // TheData
     CritEdt,       // ClassCritEdt
     UtilEdt;       // MajEditionLegal

  procedure CPLanceFiche_CPJALOUVCLO(TypeEdition: Char; Argument: String='');
  procedure LanceEdLegaleJo(Apercu, Couleur: Boolean);

Type
  TOF_CPJALOUVCLO = Class (TOF_Meth)
  private
    FTypeEdition: Char;
    FTriParJnl:       Boolean;

    {Critères}
    Exercice:       THValComboBox;
    DateComptaDe:   THEdit;
    DateComptaA:    THEdit;
    NumeroPieceDe:  THEdit;
    NumeroPieceA:   THEdit;
    {Critères avancés}
    RefInterne:       THEdit;
    NouvFct:          TCheckBox;
    Devise:           THValcomboBox;
    {Options d'édition}
    ChoixMonnaie:     THRadioGroup;
    ChkLibCompte:     TCheckBox;
    {Dev}
    TitreEtat:        THEdit;
    AvecChoixMonnaie: THEdit;
    JournalDe:        THEdit;
    JournalA:         THEdit;
    EdtObligatoire:   THEdit;
    NumeroExercice:   THEdit;
    E_DATECOMPTABLE_    : THEdit ;
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
    procedure ChargementCritEdt        ; override ;
  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Hent1,
  Ent1;


procedure LanceEdLegaleJo(Apercu, Couleur: Boolean);
var Argument: String;
begin
  if Apercu then
    Argument := 'X;'
  else
    Argument := '-;';

  if Couleur then
    Argument := Argument + 'X;'
  else
    Argument := Argument + '-;';

  CPLanceFiche_CPJALOUVCLO('L', Argument);
end;

procedure CPLanceFiche_CPJALOUVCLO(TypeEdition: Char; Argument: String='');
begin
  AGLLanceFiche('CP','CPJALOUVCLO_QR1','','',TypeEdition+';'+Argument);
end;
 

procedure TOF_CPJALOUVCLO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPJALOUVCLO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPJALOUVCLO.OnUpdate ;
var
  ClauseWhere: String;
  ClauseOrder: String;
begin
  Inherited ;

  {Clause Where}
  ClauseWhere := '';
  if RefInterne.Text<>'' then
    ClauseWhere := ClauseWhere + ' And Upper(E_REFINTERNE) like "'+TraduitJoker(RefInterne.Text)+'"';

  case FTypeEdition of
    'L': begin    {Légale}
           ClauseWhere := ClauseWhere + ' and E_QUALIFPIECE<>"C" ';
         end;
    'O': begin    {Ouverture}
           ClauseWhere := ClauseWhere + ' And E_QUALIFPIECE="N" And (E_ECRANOUVEAU="OAN" or E_ECRANOUVEAU="H")';
         end;
    'F': begin    {Fermeture}
           ClauseWhere := ClauseWhere + ' And E_QUALIFPIECE="C" And E_ECRANOUVEAU="C"';
         end;
  end ;

  case FTypeEdition of
    'L': begin    {Légale}
           ClauseWhere := ClauseWhere + ' AND J_NATUREJAL<>"CLO" AND J_NATUREJAL<>"ODA" AND J_NATUREJAL<>"ANA"'+
                                        ' AND J_NATUREJAL<>"ANO"';
         end;
    'O': begin    {Ouverture}
           ClauseWhere := ClauseWhere + ' AND J_NATUREJAL="ANO"';
         end;
    'F': begin    {Fermeture}
           ClauseWhere := ClauseWhere + ' AND J_NATUREJAL="CLO"';
         end;
  end;

  {Clause Order by}
  if NouvFct.Checked then begin
    if FTriParJnl then
      ClauseOrder := ' Order By E_JOURNAL,E_EXERCICE,E_PERIODE,E_NUMEROPIECE,E_DATECOMPTABLE,E_NUMLIGNE,E_NUMECHE'
    else
      ClauseOrder := ' Order By E_JOURNAL,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE';
  end
  else begin
    if FTriParJnl then
      ClauseOrder := ' Order By E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE'
    else
      ClauseOrder := ' Order By E_JOURNAL,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE';
  end;

  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL +
    ' ' + ClauseWhere + ' ' + ClauseOrder;

  If FTypeEdition = 'L' then        {Légale}
    MajEditionLegal('JLD',NumeroExercice.Text,DateComptaDe.Text,DateComptaA.Text);

  TitreEtat.Text := Ecran.Caption+' '+FProvisoire.Text; // YMO 21/11/2006 FQ18786 - Gestion BOI (Norme NF)
end ;

procedure TOF_CPJALOUVCLO.OnLoad ;
begin
  Inherited ;

  RecupCritEdt;

   {$IFNDEF GCGC}
  // YMO 21/11/2006 FQ18786 - Gestion BOI (Norme NF)
  FDateFinEdition := StrToDate(E_DateComptable_.Text);
  {$ENDIF}
end ;

procedure TOF_CPJALOUVCLO.OnArgument (S : String ) ;
var
  Argument: String;
begin
  Inherited ;

  FTypeEdition := ReadTokenSt(S)[1];
  Argument := S;

  Pages          := TPageControl(Getcontrol('PAGES', true));
  {Critères}
  Exercice       := THValComboBox(Getcontrol('E_EXERCICE', true));
  DateComptaDe   := THEdit(GetControl('E_DATECOMPTABLE', true));
  DateComptaA    := THEdit(GetControl('E_DATECOMPTABLE_', true));
  NumeroPieceDe  := THEdit(GetControl('E_NUMEROPIECE', true));
  NumeroPieceA   := THEdit(GetControl('E_NUMEROPIECE_', true));
  {Critères avancés}
  RefInterne     := THEdit(Getcontrol('REFINTERNE', true));
  NouvFct        := TCheckBox(Getcontrol('NVFCT', true));
  Devise         := THValcomboBox(Getcontrol('E_DEVISE', true));
  {Options d'édition}
  ChoixMonnaie      := THRadioGroup(Getcontrol('CHOIXMONNAIE', true));
  ChkLibCompte      := TCheckBox(Getcontrol('CHKLIBELLECPT', true));
  {Dev}
  AvecChoixMonnaie  := THEdit(Getcontrol('AVECCHOIXMONNAIE', true));
  TitreEtat         := THEdit(Getcontrol('TITREETAT', true));
  JournalDe         := THEdit(GetControl('E_JOURNAL', true));
  JournalA          := THEdit(GetControl('E_JOURNAL_', true));
  EdtObligatoire    := THEdit(GetControl('EDTOBLIGATOIRE', true));
  NumeroExercice    := THEdit(GetControl('NUMEXERCICE', true));

  E_DATECOMPTABLE_  := THEdit(GetControl('E_DATECOMPTABLE_')) ;

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

  FTriParJnl := False;
  case FTypeEdition of
   'O' : begin      {Ouverture}
         Ecran.Caption     := TraduireMemoire('Journal d''ouverture');
         TFQRS1(Ecran).FNomFiltre:='OUVERTURE';
         Ecran.HelpContext := 7757000;
         end;
   'F' : begin      {Fermeture}
         Ecran.Caption     := TraduireMemoire('Journal de clôture');
         TFQRS1(Ecran).FNomFiltre:= 'FERMETURE';
         Ecran.HelpContext := 7766000;
         if VH^.Precedent.Code <> '' then
           Exercice.Value  := VH^.Precedent.Code;
         end;
   'L' : begin      {Légale}
         Ecran.Caption     := TraduireMemoire('Journal des écritures');
         Ecran.HelpContext := 7712000;
         FTriParJnl        := True;
         end;
  end;
  UpdateCaption(Ecran);
  TitreEtat.Text := Ecran.Caption;

  if FTypeEdition = 'L' then begin
    Ecran.WindowState := wsMaximized;
    SetControlChecked('FAPERCU',   (ReadTokenSt(Argument) = 'X'));
    SetControlChecked('FCOULEUR',  (ReadTokenSt(Argument) = 'X'));
    end;
end ;

procedure TOF_CPJALOUVCLO.ChargementCritEdt;
begin
  if (TheData <> nil) and (TheData is (ClassCritEdt)) then begin
    Exercice.Value := ClassCritEdt(TheData).CritEdt.Exo.Code;
    // Date de Début de l'édition
    DateComptaDe.Text  := DateToStr(ClassCritEdt(TheData).CritEdt.Date1);
    // Date de Fin de l'édition
    DateComptaA.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date2);
    end;
  inherited;
end;

procedure TOF_CPJALOUVCLO.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPJALOUVCLO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPJALOUVCLO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPJALOUVCLO.ChoixMonnaieChange(Sender: TObject);
begin
  {Choix monnaie possible uniquement si une devise est sélectionnée}
  if (Devise.ItemIndex=0) or (Devise.Value=V_PGI.DevisePivot) then begin
    if ChoixMonnaie.ItemIndex <> 0 then
      HShowMessage('0;'+Ecran.Caption+';Vous devez d''abord sélectionner une devise particulière.;E;O;O;O;','','');
    ChoixMonnaie.ItemIndex := 0;
    end;
  AvecChoixMonnaie.Text := ChoixMonnaie.Value;
end;

procedure TOF_CPJALOUVCLO.ExoOnChange(Sender: TObject);
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

procedure TOF_CPJALOUVCLO.RecupCritEdt;
begin
  case FTypeEdition of
   'O' : begin      {Ouverture}
         JournalDe.DataType := 'TZJAN';
         JournalA.DataType  := 'TZJAN';
         PositionneFourchetteST(JournalDe, JournalA, tzJAN);
         end;
   'F' : begin      {Fermeture}
         JournalDe.DataType := 'TZJCLOTURE';
         JournalA.DataType  := 'TZJCLOTURE';
         PositionneFourchetteST(JournalDe, JournalA, tzJCloture);
         end;
   'L' : begin      {Légale}
         JournalDe.DataType := 'TZJOURNAL';
         JournalA.DataType  := 'TZJOURNAL';
         PositionneFourchetteST(JournalDe, JournalA, tzJournal);
         end;
  end;

  if (FTypeEdition in ['O', 'F']) then   {Ouverture/Fermeture}
    EdtObligatoire.Text := '-'
  else                                   {Légale}
    EdtObligatoire.Text := 'X';
  NumeroExercice.Text := Exercice.Value;

  If NumeroPieceDe.Text = '' then
    NumeroPieceDe.Text := '0';
  If NumeroPieceA.Text = '' then
    NumeroPieceA.Text := '999999999';
end;

Initialization
  registerclasses ( [ TOF_CPJALOUVCLO ] ) ;
end.

