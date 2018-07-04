unit TofRbCyc;

interface
uses Classes, StdCtrls, SysUtils, dbtables,
     UTof, HCtrls, QRS1, Ent1, TofMeth, HTB97;

type
  TOF_GLCYCLE = class(TOF_Meth)
  private
    rub1, rub2: THEdit;
    Exo, Etab, Devise: THValComboBox;
    Date1, Date2 : THEdit;
    DevPivot: THEdit;
    Trier, Blocnote: TCheckbox;
    xxOrderBy: THEdit;
    BHelp: TToolbarButton97;
    procedure BHelpClick(Sender: TObject);
    procedure RubOnExit(Sender: TObject);
    procedure ExoOnChange(Sender: TObject);
    procedure OrderByChecked(Sender: TObject);
    procedure DateOnExit(Sender: TObject);
    procedure InitCritereAvance ;
    procedure ClearTable;
    procedure FillTable;
  public
    procedure OnArgument(Arguments : string); override ;
    procedure OnUpdate; override;
    procedure OnNew; override;
  end;


implementation
uses HEnt1;

{ TOF_GLCYCLE }

procedure TOF_GLCYCLE.OnArgument(Arguments: string);
begin
inherited ;
with TFQRS1(Ecran) do begin
  PageOrder(Pages);
  NatureEtat := Arguments;
  Rub1 := THEdit(GetControl('RC_RUBRIQUE'));
  Rub2 := THEdit(GetControl('RC_RUBRIQUE_'));
  Exo := THValComboBox(GetControl('E_EXERCICE'));
  Date1 := THEdit(GetControl('E_DATECOMPTABLE'));
  Date2 := THEdit(GetControl('E_DATECOMPTABLE_'));
  Etab := THValComboBox(GetControl('E_ETABLISSEMENT'));
  Devise := THValComboBox(GetControl('E_DEVISE'));
  DevPivot := THEdit(GetControl('DEVPIVOT'));
  Trier:=TCheckbox(GetControl('TRIER'));
  Blocnote:=TCheckbox(GetControl('BLOCNOTE'));
  xxOrderBy:=THEdit(GetControl('XX_ORDERBY'));
  BHelp := TToolbarButton97(GetControl('BAIDE'));
  Pages.ActivePage:=TFQRS1(Ecran).Pages.Pages[0];
end;
if (BHelp <> nil ) and (not Assigned(BHelp.OnClick)) then
  BHelp.OnClick := BHelpClick;
if Exo <> nil then begin
  Exo.OnChange := ExoOnChange;
  Exo.Value := VH^.Entree.Code;
end ;
if Rub1 <> nil then
  Rub1.OnExit := RubOnExit;
if Rub2 <> nil then
  Rub2.OnExit := RubOnExit;
if Date1 <> nil then
  Date1.OnExit := DateOnExit;
if Date2 <> nil then
  Date2.OnExit := DateOnExit;
if Trier <> nil then Trier.OnClick := OrderByChecked;
if Etab <> nil then Etab.ItemIndex := 0 ;
if Devise <> nil then Devise.ItemIndex := 0;
if (Devise <> nil) and (DevPivot<>nil) then
  DevPivot.Text := Devise.Items[Devise.Values.IndexOf(V_PGI.DevisePivot)];
InitCritereAvance;
end;

procedure TOF_GLCYCLE.OnNew;
begin

end;

procedure TOF_GLCYCLE.OnUpdate;
begin
  if (Rub1<>nil) and (Trim(Rub1.Text) = '') then Rub1.Text := '00000000';
  if (Rub2<>nil) and (Trim(Rub2.Text) = '') then Rub2.Text := 'ZZZZZZZZ';
  ClearTable;
  FillTable;
end;

procedure TOF_GLCYCLE.RubOnExit(Sender: TObject);
begin
  DoRubOnExit(THEdit(Sender), Rub1, Rub2);
end;

procedure TOF_GLCYCLE.BHelpClick(Sender: TObject);
begin

end;

procedure TOF_GLCYCLE.DateOnExit(Sender: TObject);
begin

end;

procedure TOF_GLCYCLE.ExoOnChange(Sender: TObject);
begin
  DoExoToDateOnChange(Exo, Date1, Date2);
end;


procedure TOF_GLCYCLE.ClearTable;
var s: string;
begin
  s := 'DELETE FROM RUBCOMPTE WHERE RC_UTILISATEUR=' + V_PGI.USER;
  ExecuteSQL(s);
end;


procedure TOF_GLCYCLE.FillTable;
var s, Inclu, Exclu: string;
    qrub, qCpt: TQuery;
begin
  s := 'SELECT * FROM RUBRIQUE'
     + ' WHERE'
     + ' RB_RUBRIQUE>=' + '"' + Rub1.Text + '"'
     + ' AND RB_RUBRIQUE<=' + '"' + Rub2.Text + '"'
     + ' AND RB_TYPERUB="CDR"'
     + ' ORDER BY RB_RUBRIQUE';
  qRub := OpenSql(s, True);
  try
    while not qRub.Eof do
    begin
      Inclu := AnalyseCompte(qRub.FieldByName('RB_COMPTE1').AsString, fbRubrique, FALSE, FALSE, FALSE);
      Exclu := AnalyseCompte(qRub.FieldByName('RB_EXCLUSION1').AsString, fbRubrique, true, FALSE, FALSE);
      Inclu := Uppercase(Inclu);
      Exclu := Uppercase(Exclu);
      Inclu := FindEtReplace (Inclu, 'LIKE', 'G_GENERAL like', true);
      Exclu := FindEtReplace (Exclu, 'NOT', 'G_GENERAL not', true);
      s := 'SELECT G_GENERAL, G_LIBELLE FROM GENERAUX';
      if trim(Inclu) <> '' then
        s := s + ' WHERE ' + Inclu;
      if trim(Exclu) <> '' then
        s := s + ' AND ' + Exclu;
      qCpt := OpenSql(s, True);
      try
        while not qCpt.Eof do
        begin
          s := 'INSERT INTO RUBCOMPTE '
             + '(RC_RUBRIQUE, RC_LIBELLE, RC_GENERAL, RC_INTITULE, RC_UTILISATEUR) '
             + 'VALUES ('
             + '"' + qRub.FieldByName('RB_RUBRIQUE').AsString + '"' + ','
             + '"' + qRub.FieldByName('RB_LIBELLE').AsString  + '"' + ','
             + '"' + qCpt.FieldByName('G_GENERAL').AsString   + '"' + ','
             + '"' + qCpt.FieldByName('G_LIBELLE').AsString   + '"' + ','
             + '"' + V_PGI.USER + '"'
             + ')';
          ExecuteSQL(s);
          qCpt.Next;
        end;
      finally
        Ferme(qCpt);
      end;
      qRub.Next;
    end;
  finally
    Ferme(qRub);
  end;
end;

procedure TOF_GLCYCLE.OrderByChecked(Sender: TObject);
begin
if Trier.Checked then
  xxOrderBy.Text:='RC_RUBRIQUE,RC_GENERAL,E_DATECOMPTABLE,E_JOURNAL,E_NUMEROPIECE,E_NUMLIGNE'
else
  xxOrderBy.Text:='RC_RUBRIQUE,RC_GENERAL,E_JOURNAL,E_NUMEROPIECE,E_NUMLIGNE';
end;

procedure TOF_GLCYCLE.InitCritereAvance;
begin
  Blocnote.Checked := false;
  Trier.Checked := true;
  OrderByChecked(self);
end;

initialization
RegisterClasses([TOF_GLCYCLE]) ;


end.
