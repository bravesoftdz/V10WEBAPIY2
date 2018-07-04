{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/12/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBROUILLARDBUD ()
Mots clefs ... : TOF;CPBROUILLARDBUD
    Version    |   Date   | Qui  |   Commentaires
--------------------------------------------------------------------------------------
 7.00.001.011    07/06/06    MD   FQ 18331 N° Aide incorrect
*****************************************************************}
Unit CPBROUILLARDBUD_TOF ;

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
     HEnt1,
     HMsgBox,
     UTOF,
     Ent1,
     TofMeth,
     uLibExercice,
     HCompte,
     SAISUTIL;

  procedure CPLanceFiche_CPBROUILLARDBUD(TypeEdition: Char);

Type
  TOF_CPBROUILLARDBUD = Class (TOF_Meth)
  private
    FTypeEdition: Char;

    {Critères}
    Axe:            THValComboBox;
    Budget:         THValComboBox;
    ExerciceDe:     THValComboBox;
    ExerciceA:      THValComboBox;
    PeriodeDe:      THValComboBox;
    PeriodeA:       THValComboBox;
    {Critères avancés}
    RefInterne:       THEdit;
    EcrValide:        TCheckBox;
    NatureBudget:     THValComboBox;
    NumeroPieceDe:    THEdit;
    NumeroPieceA:     THEdit;
    {Dev}
    DateComptaDe:     THEdit;
    DateComptaA:      THEdit;
    TitreEtat:        THEdit;

    procedure BudgetOnChange(Sender: TObject);
    procedure ExoOnChange(Sender: TObject) ;
    procedure PeriodeOnChange(Sender: TObject) ;

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

procedure CPLanceFiche_CPBROUILLARDBUD(TypeEdition: Char);
begin
  {G: Compte budgétaire
   S: Compte section}
  AGLLanceFiche('CP','CPBROUILBUD_QR1','','',TypeEdition);
end;

procedure TOF_CPBROUILLARDBUD.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDBUD.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDBUD.OnUpdate ;
var
  ClauseWhere: String;
begin
  Inherited ;

  {Clause Where}
  ClauseWhere := '';
  if RefInterne.Text<>'' then
    ClauseWhere := ClauseWhere + ' And Upper(BE_REFINTERNE) like "'+TraduitJoker(RefInterne.Text)+'"';

  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' ' + ClauseWhere;
end ;

procedure TOF_CPBROUILLARDBUD.OnLoad ;
begin
  Inherited ;

  RecupCritEdt;
end ;

procedure TOF_CPBROUILLARDBUD.OnArgument (S : String ) ;
begin
  Inherited ;

  if Length(S) = 0 then
    FTypeEdition := ' '
  else
    FTypeEdition := S[1];

  Pages          := TPageControl(Getcontrol('PAGES', true));
  {Critères}
  Axe            := THValComboBox(GetControl('BE_AXE', true));
  Budget         := THValComboBox(GetControl('BE_BUDJAL', true));
  ExerciceDe     := THValComboBox(GetControl('BE_EXERCICE', true));
  ExerciceA      := THValComboBox(GetControl('BE_EXERCICE_', true));
  PeriodeDe      := THValComboBox(GetControl('COMBOPERIODE', true));
  PeriodeA       := THValComboBox(GetControl('COMBOPERIODE_', true));
  {Critères avancés}
  RefInterne     := THEdit(Getcontrol('REFINTERNE', true));
  EcrValide      := TCheckBox(Getcontrol('BE_VALIDE', true));
  NatureBudget   := THValComboBox(Getcontrol('BE_NATUREBUD', true));
  NumeroPieceDe  := THEdit(GetControl('BE_NUMEROPIECE', true));
  NumeroPieceA   := THEdit(GetControl('BE_NUMEROPIECE_', true));
  {Dev}
  DateComptaDe   := THEdit(GetControl('BE_DATECOMPTABLE', true));
  DateComptaA    := THEdit(GetControl('BE_DATECOMPTABLE_', true));
  TitreEtat      := THEdit(Getcontrol('TITREETAT', true));


  Budget.OnChange     := BudgetOnChange;
  ExerciceDe.onChange := ExoOnChange;
  ExerciceA.onChange  := ExoOnChange;
  PeriodeDe.OnChange  := PeriodeOnChange;
  PeriodeA.OnChange   := PeriodeOnChange;

  Axe.ItemIndex := 0;
  if Budget.Items.Count > 0 then
    Budget.ItemIndex := 0;
  if (CtxPCl in V_PGI.PgiContexte) and (VH^.CPExoRef.Code <>'') then
    ExerciceDe.Value := VH^.CPExoRef.Code
  else
    ExerciceDe.Value := VH^.Entree.Code;
  ExerciceA.Value := ExerciceDe.Value;
  NatureBudget.ItemIndex  := 0;

  case FTypeEdition of
   'G' : begin      {Compte budgétaire}
         Ecran.Caption     := TraduireMemoire('Brouillard par comptes budgétaires');
         TFQRS1(Ecran).FNomFiltre :='BROUILCPT' ;
         TFQRS1(Ecran).CodeEtat   := 'CBG';
         Ecran.HelpContext        := 15282200 ;   // MD FQ 18331
         SetControlCaption('MVTANALYTIQUE', TraduireMemoire('Les &mouvements analytiques'));
         end;
   'S' : begin      {Section budgétaire}
         Ecran.Caption     := TraduireMemoire('Brouillard par sections budgétaires');
         TFQRS1(Ecran).FNomFiltre :='BROUILSEC' ;
         TFQRS1(Ecran).CodeEtat   := 'CBS';
         Ecran.HelpContext        := 15284200;  // MD FQ 18331
         SetControlCaption('MVTANALYTIQUE', TraduireMemoire('Les &mouvements des comptes'));
         end;
  end;
  UpdateCaption(Ecran);
  TitreEtat.Text := Ecran.Caption;
end ;

procedure TOF_CPBROUILLARDBUD.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDBUD.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDBUD.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPBROUILLARDBUD.ExoOnChange(Sender: TObject);
var
  Periode:   THValComboBox;
  DateCpt:   THedit;
  ItemIndex: Integer;
  Debut:     Boolean;
begin
  Debut := (Sender = ExerciceDe);
  if Debut and (ExerciceDe.Value > ExerciceA.Value) then
    ExerciceA.Value := ExerciceDe.Value
  else if (not Debut) and (ExerciceA.Value < ExerciceDe.Value) then
    ExerciceDe.Value := ExerciceA.Value;

  if Debut then begin
    Periode   := PeriodeDe;
    DateCpt   := DateComptaDe;
    ItemIndex := 0;
  end
  else begin
    Periode   := PeriodeA;
    DateCpt   := DateComptaA;
    ItemIndex := PeriodeA.Items.Count-1;
  end;

  ListePeriode((Sender as THValComboBox).Value, Periode.Items, Periode.Values,
    Debut) ;
  Periode.ItemIndex := ItemIndex;
  DateCpt.Text      := Periode.Value;
end;

procedure TOF_CPBROUILLARDBUD.PeriodeOnChange(Sender: TObject);
var
  Debut:     Boolean;
  d1, d2:    TDateTime;
  s1, s2:    string;
  Year, Month, Day : Word;
begin
  if Trim((Sender as THValComboBox).Value) = '' then
    Exit;

  Debut := (Sender = PeriodeDe);

  s1 := PeriodeDe.Value;
  s2 := PeriodeA.Value;

  d1 := StrToDate(PeriodeDe.Value);
  d2 := StrToDate(PeriodeA.Value);
  if Debut and (d1 > d2) then begin
    {Initialise la période fin au dernier jour du mois de la période début}
    DecodeDate(d1, Year, Month, Day);
    d2 := EncodeDate(Year, Month, DaysPerMonth(Year, Month));
    PeriodeA.Value := DateToStr(d2);
  end
  else if (not Debut) and (d2 < d1) then begin
    {Initialise la période début au premier jour du mois de la période fin}
    DecodeDate(d2, Year, Month, Day);
    d1 := EncodeDate(Year, Month, 1);
    PeriodeDe.Value := DateToStr(d1);
  end;

  DateComptaDe.Text := PeriodeDe.Value;
  DateComptaA.Text  := PeriodeA.Value;
end;

procedure TOF_CPBROUILLARDBUD.BudgetOnChange(Sender: TObject);
var QJal  : TQuery ;
begin
  QJal:=OpenSQL('Select BJ_EXODEB, BJ_EXOFIN, BJ_PERDEB, BJ_PERFIN from BUDJAL Where BJ_BUDJAL="'+Budget.Value+'"',True) ;
  if not QJal.EOF then
     begin
     ExerciceDe.Value:=QJal.FindField('BJ_EXODEB').AsString ;
     ExerciceA.Value:=QJal.FindField('BJ_EXOFIN').AsString ;
     PeriodeDe.Value:=QJal.FindField('BJ_PERDEB').AsString ;
     PeriodeA.Value:=QJal.FindField('BJ_PERFIN').AsString ;
     DateComptaDe.Text:=QJal.FindField('BJ_PERDEB').AsString ;
     DateComptaA.Text:=QJal.FindField('BJ_PERFIN').AsString ;
     end ;
  Ferme(QJal) ;
end ;

procedure TOF_CPBROUILLARDBUD.RecupCritEdt;
begin
  If NumeroPieceDe.Text = '' then
    NumeroPieceDe.Text := '0';
  If NumeroPieceA.Text = '' then
    NumeroPieceA.Text := '999999999';
end;

Initialization
  registerclasses ( [ TOF_CPBROUILLARDBUD ] ) ;
end.

