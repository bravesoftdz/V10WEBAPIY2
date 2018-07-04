{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 24/03/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BROBUDGEN ()
Mots clefs ... : TOF;BROBUDGEN
*****************************************************************}
Unit BROBUDGEN_TOF ;

Interface

Uses StdCtrls, Controls, Classes, 
{$IFDEF EAGLCLIENT}
  eMul, uTob, MaineAGL, eQRS1,
{$ELSE}
  db, dbtables, mul, FE_Main, QRS1,
{$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, Ent1;

procedure BrouillardBud(Code : Char);

Type
  TOF_BROBUDGEN = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    FJournal, FNatBud, FExercice, FExercice_, FPeriode1, FPeriode2 : THValComboBox;
    CodeNatureEcr : String;
    FDateCompta1, FDateCompta2 : THEdit;
    procedure FJournalChange(Sender : TObject);
    procedure FExerciceChange(Sender: TObject);
    procedure FExercice_Change(Sender: TObject);
    procedure FPeriode1Change(Sender: TObject);
    procedure FPeriode2Change(Sender: TObject);
    function GenereSQL : String;
    procedure DebToDate;
    procedure FinToDate;
  end ;

Implementation

procedure BrouillardBud(Code : Char);
begin
  AGLLanceFiche('CP', 'CPBBROBUDGEN', '', '', Code);
end;

procedure TOF_BROBUDGEN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BROBUDGEN.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BROBUDGEN.OnUpdate ;
begin
  Inherited ;

  TFQRS1(Ecran).WhereSQL := GenereSQL;
end ;

procedure TOF_BROBUDGEN.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BROBUDGEN.OnArgument (S : String ) ;
begin
  Inherited ;

  CodeNatureEcr := S;

  FJournal := THValComboBox(GetControl('FJOURNAL', True)); FJournal.OnChange := FJournalChange;
  FExercice := THValComboBox(GetControl('FEXERCICE', True)); FExercice.OnChange := FExerciceChange;
  FExercice_ := THValComboBox(GetControl('FEXERCICE_', True)); FExercice_.OnChange := FExercice_Change;
  FPeriode1 := THValComboBox(GetControl('FPERIODE1', True)); FPeriode1.OnChange := FPeriode1Change;
  FPeriode2 := THValComboBox(GetControl('FPERIODE2', True)); FPeriode2.OnChange := FPeriode2Change;
  FNatBud := THValComboBox(GetControl('FNATBUD', True));
  FDateCompta1 := THEdit(GetControl('FDATECOMPTA1', True));
  FDateCompta2 := THEdit(GetControl('FDATECOMPTA2', True));

  if FJournal.Values.Count>0 then FJournal.Value := FJournal.Values[0];
  SetControlVisible('FVALIDE', (CodeNatureEcr<>'S')and(CodeNatureEcr<>'R'));
  FNatBud.ItemIndex := 0;
  DebToDate;
  FinToDate;
end ;

procedure TOF_BROBUDGEN.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_BROBUDGEN.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BROBUDGEN.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BROBUDGEN.FJournalChange(Sender: TObject);
var
  QJal : TQuery;
begin
  QJal := OpenSQL('SELECT BJ_EXODEB, BJ_EXOFIN, BJ_PERDEB, BJ_PERFIN FROM BUDJAL WHERE BJ_BUDJAL="'+GetControlText('FJOURNAL')+'"', True);
  if Not QJal.EOF then begin
    FExercice.Value := QJal.FindField('BJ_EXODEB').AsString;
    FExercice_.Value := QJal.FindField('BJ_EXOFIN').AsString;
    FPeriode1.Value := QJal.FindField('BJ_PERDEB').AsString;
    FPeriode2.Value := QJal.FindField('BJ_PERFIN').AsString;
    FDateCompta1.Text := QJal.FindField('BJ_PERDEB').AsString;
    FDateCompta2.Text := QJal.FindField('BJ_PERFIN').AsString;
  end;
  Ferme(QJal);
end;

procedure TOF_BROBUDGEN.FExerciceChange(Sender: TObject);
begin
  if FExercice.Value > FExercice_.Value then FExercice_.Value := FExercice.Value;
  ListePeriode(FExercice.Value, FPeriode1.Items, FPeriode1.Values, True);
  FPeriode1.ItemIndex := 0;
  DebToDate;
end;

procedure TOF_BROBUDGEN.FExercice_Change(Sender: TObject);
begin
  if FExercice_.Value<FExercice.Value then FExercice.Value:=FExercice_.Value;
  ListePeriode(FExercice_.Value, FPeriode2.Items, FPeriode2.Values, False);
  FPeriode2.ItemIndex := FPeriode2.Items.Count-1;
  FinToDate;
end;

procedure TOF_BROBUDGEN.FPeriode1Change(Sender: TObject);
begin
  if StrToDate(FPeriode1.Value) > StrToDate(FPeriode2.Value) then FPeriode2.Value := FPeriode1.Value;
  FDateCompta1.Text := FPeriode1.Value;
end;

procedure TOF_BROBUDGEN.FPeriode2Change(Sender: TObject);
begin
  if StrToDate(FPeriode2.Value) < StrToDate(FPeriode1.Value) then FPeriode1.Value := FPeriode2.Value;
  FDateCompta2.Text := FPeriode2.Value;
end;

function TOF_BROBUDGEN.GenereSQL: String;
var
  sz : String;
  FValide : TCheckBox;
begin
  Result := 'SELECT BE_EXERCICE, BE_BUDJAL, BE_DATECOMPTABLE, BE_NUMEROPIECE, '+
            'BE_BUDGENE, BE_AXE, BE_QUALIFQTE2 , BE_REFINTERNE, BE_LIBELLE, BE_BUDSECT, '+
            'BE_QTE1, BE_QTE2, BE_QUALIFQTE1, BG_LIBELLE, BE_ETABLISSEMENT, '+
            'BE_NATUREBUD, BE_VALIDE, BE_TYPESAISIE, BS_LIBELLE, '+
            'BE_DEBIT DEBIT, BE_CREDIT CREDIT '+
            'FROM BUDECR '+
            'LEFT JOIN BUDGENE on BG_BUDGENE=BE_BUDGENE '+
            'LEFT JOIN BUDSECT On BS_AXE=BE_AXE and BS_BUDSECT=BE_BUDSECT '+
            'WHERE BE_DATECOMPTABLE>="'+USDateTime(StrToDate(FDateCompta1.Text))+'" And BE_DATECOMPTABLE<="'+USDateTime(StrToDate(FDateCompta2.Text))+'" '+
            'AND BE_QUALIFPIECE="'+CodeNatureEcr+'" ';
  sz := GetControlText('FNATURECPT');
  if sz<>'' then Result := Result + 'AND BE_AXE="'+sz+'" ';
  Result := Result + 'AND BE_BUDJAL="'+FJOURNAL.Value+'" '+
                     'AND BE_NUMEROPIECE>='+GetControlText('FNUMPIECE1')+' '+
                     'AND BE_NUMEROPIECE<='+GetControlText('FNUMPIECE2')+' '+
                     'AND BE_EXERCICE>="'+FExercice.Value+'" And BE_EXERCICE<="'+FExercice_.Value+'" ';
  sz := GetControlText('FREFINTERNE');
  if sz<>'' then Result := Result + 'AND UPPER(BE_REFINTERNE) LIKE "'+TraduitJoker(sz)+'" ';
  sz := GetControlText('FETAB');
  if sz<>'' then Result := Result + 'AND BE_ETABLISSEMENT="'+sz+'" ';
  if (V_PGI.Confidentiel<>'1') then Result := Result + 'AND BE_CONFIDENTIEL<>"1" ';

  FValide := TCheckBox(GetControl('FVALIDE', True));
  if FValide.Checked then Result := Result + 'AND BE_VALIDE="X" ' else
  if FValide.State = cbUnchecked then Result := Result + 'AND BE_VALIDE="-" ';
  if FNatBud.Value <>'' then Result := Result +'AND BE_NATUREBUD="'+FNatBud.Value+'" ';
  Result := Result + 'ORDER BY BE_NUMEROPIECE, BE_DATECOMPTABLE, BE_BUDGENE, BE_BUDJAL, BE_NATUREBUD, BE_QUALIFPIECE ';
end;

procedure TOF_BROBUDGEN.DebToDate;
begin
  FDateCompta1.Text := FPeriode1.Value;
end;

procedure TOF_BROBUDGEN.FinToDate;
begin
  FDateCompta2.Text := FPeriode2.Value;
end;

Initialization
  registerclasses ( [ TOF_BROBUDGEN ] ) ;
end.
