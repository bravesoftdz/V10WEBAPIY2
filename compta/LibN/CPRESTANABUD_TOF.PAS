{***********UNITE*************************************************
Auteur  ...... :           
Créé le ...... : 12/01/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPRESTANABUD ()
Mots clefs ... : TOF;CPRESTANABUD
*****************************************************************}
Unit CPRESTANABUD_TOF ;

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
     UTOB,
     HTB97,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ELSE}
     tcalccum,   {TTypeCalc}
     {$ENDIF MODENT1}
     CALCOLE ;

  procedure CPLanceFiche_CPRESTANABUD;

Type
  TOF_CPRESTANABUD = Class (TOF_Meth)
  private
    FTobEdt: TOB;

    {Critères}
    Journal:      THValComboBox;
    CompteDe:     THEdit;
    CompteA:      THEdit;
    {Dev}
    Axe:          THEdit;

    procedure BValiderClick(Sender: TObject);

    procedure RecupCritEdt;
    procedure RemplirTobEdt;
    procedure RemplirCompteIntedit(SQL: String);
    function  AnalyseRubrique(RubName: String): String;
    function  SQLPrincipal(Axe, Jal, Compte, Section: String): String;
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

uses uLibAnalytique, CPTPAYEURFACTURE_TOF;

procedure CPLanceFiche_CPRESTANABUD;
begin
  AGLLanceFiche('CP','CPRESTANABUD_QR1','','','');
end;

procedure TOF_CPRESTANABUD.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANABUD.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANABUD.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANABUD.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANABUD.OnArgument (S : String ) ;
begin
  Inherited ;

  Pages       := TPageControl(Getcontrol('PAGES', true));
  {Critères}
  Journal     := THValComboBox(GetControl('BUDJAL', true));
  CompteDe    := THEdit(GetControl('BUDGENE', true));
  CompteA     := THEdit(GetControl('BUDGENE_', true));
  {Dev}
  Axe         := THEdit(GetControl('AXE', true));

  TToolbarButton97(GetControl('BVALIDER')).OnClick := BValiderClick;
end ;

procedure TOF_CPRESTANABUD.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANABUD.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANABUD.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPRESTANABUD.RecupCritEdt;
begin
  PositionneFourchetteST(CompteDe, CompteA, tzBudgen);
  if (Journal.ItemIndex=-1) and (Journal.Items.Count>0) then
    Journal.ItemIndex := 0;
end;

procedure TOF_CPRESTANABUD.BValiderClick(Sender: TObject);
var
  CritEdition: String;
begin
  RecupCritEdt;

  FTobEdt := TOB.Create('', nil, -1);
  try
    RemplirTobEdt;
    {Dans le paramétrage de l'état, il faut forcer 2 passages sinon la première ligne
     s'édite 2 fois}
    {Le paramètre Pages n'est pas utilisé en mode CWAS, il faut donc passer les
     paramètres dans la chaine CritEdition. Voir CPCONTROLERUB_TOF}
    CritEdition := ComposantToCritere(Pages);
    if CritEdition <> '' then
      CritEdition := Copy(CritEdition, 2, Length(CritEdition));
    LanceEtatTob('E', 'CRA', 'CRB', FTobEdt, TCheckBox(GetControl('FAPERCU')).Checked,
      TCheckBox(GetControl('FLISTE')).Checked, TCheckBox(GetControl('FREDUIRE')).Checked, {FP 19/04/2006 FQ17823}
      nil, '', Ecran.Caption, False,
      0, CritEdition);
  finally
    FreeAndNil(FTobEdt);
    end;
end;

function TOF_CPRESTANABUD.AnalyseRubrique(RubName: String): String;
var
  SQL:      String;
  QRub:     TQuery;

  JalBud:       String ;
  Laxe1,Laxe2:  String ;
  Lefb1,Lefb2: TFichierBase ;
  OnTableLibre: Boolean ;
  Where:        String;
  TypRub,LAxe,Cpt1,Cpt2,Cpt1Ex,Cpt2Ex: String ;
begin
  {Retourne la clause SQL pour filtrer les comptes comptables à partir des comptes budgétaires}
  Result := '';
  SQL := 'SELECT RB_TYPERUB, RB_AXE, RB_COMPTE1, RB_COMPTE2, RB_EXCLUSION1, RB_EXCLUSION2, RB_TABLELIBRE, RB_BUDJAL' +
         ' FROM RUBRIQUE' +
         ' WHERE RB_RUBRIQUE="'+RubName+'"';
  QRub:= OpenSQL(SQL, True);
  try
    if not QRub.Eof then
      begin
      TypRub := QRub.FindField('RB_TYPERUB').AsString ;
      Laxe   := QRub.FindField('RB_AXE').AsString ;
      Cpt1   := QRub.FindField('RB_COMPTE1').AsString ;
      Cpt2   := QRub.FindField('RB_COMPTE2').AsString ;
      Cpt1Ex := QRub.FindField('RB_EXCLUSION1').AsString ;
      Cpt2Ex := QRub.FindField('RB_EXCLUSION2').AsString ;
      JalBud := QRub.FindField('RB_BUDJAL').AsString ;
      OnTableLibre := QRub.FindField('RB_TABLELIBRE').AsString='X' ;
      QuelFBRub(Lefb1,Lefb2,TypRub,Laxe,Laxe1,Laxe2,UnBud,TRUE,False) ;
      Where:=AnalyseCompte(Cpt1,Lefb1,False,False,FALSE) ;
      if Where<>'' then Result:=Result+' And '+Where ;
      Where:=AnalyseCompte(Cpt1Ex,Lefb1,True,False,FALSE) ;
      if Where<>'' then Result:=Result+' And '+Where ;
      Where:=AnalyseCompte(Cpt2,Lefb2,False,OnTableLibre,FALSE) ;
      if Where<>'' then Result:=Result+' And '+Where ;
      Where:=AnalyseCompte(Cpt2Ex,Lefb2,True,OnTableLibre,FALSE) ;
      if Where<>'' then Result:=Result+' And '+Where ;
      end;
  finally
    FreeAndNil(QRub);
    end;
end;

procedure TOF_CPRESTANABUD.RemplirTobEdt;
var
  SQL:      String;
  QCrois:   TQuery;
begin
  {Recherche la liste des comptes budgétaires GENE/SECTION}
  SQL := 'SELECT BJ_AXE, BJ_BUDJAL, CX_COMPTE, CX_SECTION' +
         ' FROM BUDJAL join CROISCPT on (CX_JAL = BJ_BUDJAL)' +
         ' WHERE BJ_BUDJAL = "'+Journal.Value+'"' +
         '   AND CX_TYPE = "BUD"' +
         '   AND CX_COMPTE>="'+CompteDe.Text+'" AND CX_COMPTE<="'+CompteA.Text+'"' +
         '   AND CX_COMPTE <> BJ_GENEATTENTE' +
         '   AND CX_SECTION <> BJ_SECTATTENTE' +
         ' ORDER BY CX_COMPTE, CX_SECTION';
  QCrois := OpenSQL(SQL, True);
  try
    while not QCrois.Eof do
      begin
      Axe.Text := RechDom('TTAXE', QCrois.FindField('BJ_AXE').AsString, false);
      SQL := SQLPrincipal(QCrois.FindField('BJ_AXE').AsString,
               QCrois.FindField('BJ_BUDJAL').AsString,
               QCrois.FindField('CX_COMPTE').AsString,
               QCrois.FindField('CX_SECTION').AsString);
      RemplirCompteIntedit(SQL);
      QCrois.Next;
      end;
  finally
    FreeAndNil(QCrois);
    end;
end;

function TOF_CPRESTANABUD.SQLPrincipal(Axe, Jal, Compte, Section: String): String;
var
  RubName: String;
  NumAxe:  String;
  Clause:  String;
  RestrictAna: TRestrictionAnalytique;
begin
  RestrictAna := TRestrictionAnalytique.Create;
  try
    RestrictAna.AxeCroise := False;
    NumAxe  := Axe[Length(Axe)];
    RubName := 'G/S'+Jal+Compte+':'+Section;
    Clause  := AnalyseRubrique(RubName);
    if Trim(Clause) = '' then
      Result := ''
    else
      Result :=
      'SELECT "'+Jal+'" AS BJ_BUDJAL, "'+Compte+'" AS CX_COMPTE,'+
            ' "'+Section+'" AS CX_SECTION, CRA_CODE, G_GENERAL, S_AXE, S_SECTION'+
      ' FROM GENERAUX cross join SECTION'+
      ' '+RestrictAna.GetFromCompteInterdit(Axe, 'G_GENERAL', 'S_AXE') +
      ' WHERE S_AXE="'+Axe+'"'+
      '   '+Clause+
      '   AND '+RestrictAna.GetClauseCompteInterdit(Axe, 'SECTION') +
      ' ORDER BY G_GENERAL, S_SECTION';
  finally
    RestrictAna.Free;
    end;
end;

procedure TOF_CPRESTANABUD.RemplirCompteIntedit(SQL: String);
begin
  if SQL = '' then
    Exit;
  FTobEdt.LoadDetailFromSQL(Sql, True);
end;

Initialization
  registerclasses ( [ TOF_CPRESTANABUD ] ) ;
end.

