{***********UNITE*************************************************
Auteur  ...... :                 
Créé le ...... : 12/01/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPRESTANAECR ()
Mots clefs ... : TOF;CPRESTANAECR
*****************************************************************}
Unit CPRESTANAECR_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_main,
     EdtREtat,
     QRS1,
{$else}
     Maineagl,
     eMul,
     UtileAGL,
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
     uLibExercice ;

  procedure CPLanceFiche_CPRESTANAECR;

Type
  TOF_CPRESTANAECR = Class (TOF_Meth)
  private
    {Critères}
    Exercice:       THValComboBox;
    DateDe:         THEdit;
    DateA:          THEdit;
    JournalDe:      THEdit;
    JournalA:       THEdit;
    Axe:            THValComboBox;
    Modele:         THMultiValComboBox;
    CompteGeneDe:   THEdit;
    CompteGeneA:    THEdit;
    {Dev}
    NumExercice:        THEdit;

    procedure AxeChange(Sender : TObject);
    procedure ExoOnChange(Sender: TObject);
    procedure DateOnExit(Sender: TObject);

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
  uLibAnalytique, CPTPAYEURFACTURE_TOF;

procedure CPLanceFiche_CPRESTANAECR;
begin
  AGLLanceFiche('CP','CPRESTANAECR_QR1','','','');
end;

procedure TOF_CPRESTANAECR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANAECR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANAECR.OnUpdate ;
var
  SQL:         String;
  NumAxe:      String;
  RestrictAna: TRestrictionAnalytique;
begin
  Inherited ;
  NumAxe := Axe.Value[Length(Axe.Value)];

  RestrictAna := TRestrictionAnalytique.Create;
  try
    SQL :=
      'SELECT Y_DATECOMPTABLE, Y_GENERAL, Y_JOURNAL, Y_NUMEROPIECE,' +
           ' Y_NUMLIGNE, Y_NUMVENTIL, Y_REFEXTERNE, Y_DEBIT, Y_CREDIT,' +
           ' CRA_CODE,';
    if VH^.AnaCroisaxe then
      SQL := SQL + 'Y_SOUSPLAN'+NumAxe+' AS Y_SECTION, Y_AXE="A'+NumAxe+'"'
    else
      SQL := SQL + ' Y_SECTION, Y_AXE';
    SQL := SQL +
      ' FROM ANALYTIQ ' + RestrictAna.GetFromCompteInterdit(Axe.Value, 'Y_GENERAL', 'Y_AXE')+
      ' WHERE Y_AXE = "A'+IntToStr(GetFirstAxeAna(Axe.Value))+'"' +
      '   AND ' + TFQRS1(Ecran).WhereSQL +
      '   AND ' + RestrictAna.GetClauseCompteInterdit(Axe.Value, 'ANALYTIQ') +
      ' ORDER BY CRA_CODE, Y_GENERAL, Y_JOURNAL, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL';
    TFQRS1(Ecran).WhereSQL := SQL;
  finally
    RestrictAna.Free;
    end;
end ;

procedure TOF_CPRESTANAECR.OnLoad ;
begin
  Inherited ;
  RecupCritEdt;
end ;

procedure TOF_CPRESTANAECR.OnArgument (S : String ) ;
begin
  Inherited ;                
  Ecran.HelpContext := 1462000;// Lek 040706 1475000;  {FP 19/04/2006 FQ17838}

  Pages       := TPageControl(Getcontrol('PAGES', true));
  {Critères}
  Exercice    := THValComboBox(Getcontrol('EXERCICE', true));
  DateDe      := THEdit(Getcontrol('Y_DATECOMPTABLE', true));
  DateA       := THEdit(Getcontrol('Y_DATECOMPTABLE_', true));
  JournalDe   := THEdit(Getcontrol('Y_JOURNAL', true));
  JournalA    := THEdit(Getcontrol('Y_JOURNAL_', true));
  Axe         := THValComboBox(Getcontrol('AXE', true));
  Modele      := THMultiValComboBox(Getcontrol('CRA_CODE', true));
  CompteGeneDe:= THEdit(Getcontrol('Y_GENERAL', true));
  CompteGeneA := THEdit(Getcontrol('Y_GENERAL_', true));
  {Dev}
  NumExercice := THEdit(Getcontrol('Y_EXERCICE', true));

  Axe.OnChange  := AxeChange;
  DateDe.OnExit := DateOnExit;
  DateA.OnExit  := DateOnExit;
  Exercice.OnChange := ExoOnChange;

  CInitComboExercice(Exercice);       // Init de la combo Exercice en Relatif
  Axe.ItemIndex := 0;
  Axe.ItemIndex := GetFirstAxeAna(Axe.Value)-1;
  AxeChange(Axe);

  if (CtxPCl in V_PGI.PgiContexte) and (VH^.CPExoRef.Code <>'') then
    Exercice.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
  else
    Exercice.Value := CExerciceVersRelatif(VH^.Entree.Code) ;
end ;

procedure TOF_CPRESTANAECR.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANAECR.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANAECR.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANAECR.AxeChange(Sender: TObject);
begin
  if (not VH^.AnaCroisaxe) and (Axe.Value <> '') then
    begin
    Modele.Plus := 'CRA_AXE="'+Axe.Value+'"';
    CompteGeneDe.DataType := 'TZGVENTIL'+Axe.Value[Length(Axe.Value)];
    end
  else begin
    Modele.Plus := '';
    CompteGeneDe.DataType := 'TZGVENTIL'+IntToStr(GetFirstAxeAna(Axe.Value));
    end;
  Modele.Text           := '';
  CompteGeneDe.Text     := '';
  CompteGeneA.Text      := '';
  CompteGeneA.DataType  := CompteGeneDe.DataType;
end;

procedure TOF_CPRESTANAECR.ExoOnChange(Sender: TObject);
begin
  CExoRelatifToDates(Exercice.Value, DateDe, DateA);
  NumExercice.Text := CRelatifVersExercice(Exercice.Value);
end;

procedure TOF_CPRESTANAECR.DateOnExit(Sender: TObject);
var
  DateD: TDateTime;
  DateF: TDateTime;
begin
  DateD := StrToDate(DateDe.Text);
  DateF := StrToDate(DateA.Text);

  DoDateOnExit(THEdit(Sender), DateDe, DateA, DateD, DateF);
end;

procedure TOF_CPRESTANAECR.RecupCritEdt;
var
  i:     Integer;
  iAxe:  Integer;
begin
  iAxe := StrToInt(Axe.Value[Length(Axe.Value)]);
  i := Ord(tzGVentil1)+iAxe-1;
  PositionneFourchetteST(CompteGeneDe, CompteGeneA, TZoomTable(Chr(i)));
  PositionneFourchetteST(JournalDe, JournalA, tzJournal);

  if Trim(Modele.Text) = '' then
    Modele.Text := TraduireMemoire('<<Tous>>');
end;

Initialization
  registerclasses ( [ TOF_CPRESTANAECR ] ) ;
end.

