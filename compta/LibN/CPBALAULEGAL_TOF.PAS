{***********UNITE*************************************************
Auteur  ...... : Muriel Dumont
Créé le ...... : 03/01/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBALAULEGAL ()
Mots clefs ... : TOF;CPBALAULEGAL

Ce source permet d'éditer une balance auxiliaire mouvementée
Nature : CBA et Code : CBM

La première colonne de l'édition correspond à l'avant période
La deuxième colonne de l'édition correspond à la période sélectionnée par l'utilisateur ( = comparatif)
La troisième colonne de l'édition correspond au cumul des 2 premières colonnes
La quatrième colonne de l'édition correspond au solde du cumul
*****************************************************************}
Unit CPBALAULEGAL_TOF ;

Interface

Uses StdCtrls,
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
     sysutils,
     Ent1,
     HCtrls,
     HEnt1,
     UTOF,
     Forms,
     uLibExercice,  // CExerciceVersRelatif
     ParamSoc,      // GetParamSoc
     HTB97,
     AGLInit,       // TheData
     CritEdt,       // ClassCritEdt
     CPBALAUXI_TOF,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     UtilEdt;       // MajEditionLegal

  //procedure CPLanceFiche_CPBALAULEGAL(Argument: String='');
  procedure CPLanceFiche_CPBALAULEGAL(Apercu, Couleur: Boolean);

Type
  TOF_CPBALAULEGAL = Class (TOF_CPBALAUXI)
  private
    FExoDate        : TExoDate;
    FDateDeb        : TDateTime;
    FDateFin        : TDateTime;
    TitreEtat       : THEdit;
  public
    procedure InitZones                     ; override;
    procedure ParamTitresCols               ; override;
    procedure ParamComparatif               ; override;
    function  GetConditionSQLCpt:    String ; override;
    function  RecupWhereQualifPiece: String ; override;
    function  GetConditionSQLPeriode(vInPer: Integer): String; override;
    procedure ChargementCritEdt             ; override ;

    procedure OnNew                         ; override ;
    procedure OnDelete                      ; override ;
    procedure OnUpdate                      ; override ;
    procedure OnLoad                        ; override ;
    procedure OnArgument (S : String )      ; override ;
    procedure OnDisplay                     ; override ;
    procedure OnClose                       ; override ;
    procedure OnCancel                      ; override ;
  end ;

Implementation

procedure CPLanceFiche_CPBALAULEGAL(Apercu, Couleur: Boolean);
var Argument : String;
begin
  if Apercu then
    Argument := 'X;'
  else
    Argument := '-;';

  if Couleur then
    Argument := Argument + 'X;'
  else
    Argument := Argument + '-;';

  AGLLanceFiche('CP','CPBALAULEGAL_QR1','','',Argument);
end;

procedure TOF_CPBALAULEGAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBALAULEGAL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPBALAULEGAL.OnUpdate ;
var
  lStReq : String;
begin
  Inherited ;
  { Maj des soldes débiteurs }
  lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT1 = (CED_DEBIT1-CED_CREDIT1), CED_CREDIT1=0 '
          + ' WHERE ( CED_DEBIT1 >= CED_CREDIT1 ) AND CED_USER="' + V_PGI.User + '"' ;
  ExecuteSQL( lStReq ) ;
  { Maj des soldes créditeurs }
  lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT1 = (CED_CREDIT1-CED_DEBIT1), CED_DEBIT1 = 0 '
          + ' WHERE ( CED_DEBIT1 < CED_CREDIT1 ) AND CED_USER="' + V_PGI.User + '"' ;
  ExecuteSQL( lStReq ) ;
  { Maj des soldes débiteurs }
  lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT2 = (CED_DEBIT2-CED_CREDIT2), CED_CREDIT2=0 '
          + ' WHERE ( CED_DEBIT2 >= CED_CREDIT2 ) AND CED_USER="' + V_PGI.User + '"' ;
  ExecuteSQL( lStReq ) ;
  { Maj des soldes créditeurs }
  lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT2 = (CED_CREDIT2-CED_DEBIT2), CED_DEBIT2 = 0 '
          + ' WHERE ( CED_DEBIT2 < CED_CREDIT2 ) AND CED_USER="' + V_PGI.User + '"' ;
  ExecuteSQL( lStReq ) ;

  { On remet les valeurs initiales de l'état }
  DateComptaDe.Text := DateToStr(FDateDeb);
  DateComptaA.Text  := DateToStr(FDateFin);

  MajEditionLegal('BLT',CRelatifVersExercice(EXERCICE.Value),DateComptaDe.Text,DateComptaA.Text);
end ;

procedure TOF_CPBALAULEGAL.OnLoad ;
begin
  TFQRS1(Ecran).NatureEtat := 'CBA';
  TFQRS1(Ecran).CodeEtat   := 'CBM';
  Inherited ;
end ;

procedure TOF_CPBALAULEGAL.OnArgument (S : String ) ;
var
  Argument: String;
begin
  Argument        := S;
  SetControlChecked('FAPERCU',   ReadTokenSt(Argument) = 'X');
  SetControlChecked('FCOULEUR',  ReadTokenSt(Argument) = 'X');
  Inherited OnArgument(Argument);

  TitreEtat := THEdit(Getcontrol('TITREETAT', true));
  Ecran.Caption     := TraduireMemoire('Balance auxiliaire');
  TFQRS1(Ecran).FNomFiltre:= 'LEGALE';
  UpdateCaption(Ecran);
  TitreEtat.Text := Ecran.Caption;
  Ecran.WindowState := wsMaximized;
end ;

procedure TOF_CPBALAULEGAL.ChargementCritEdt;
begin
  if (TheData <> nil) and (TheData is (ClassCritEdt)) then begin
    Exercice.Value := CExerciceVersRelatif(ClassCritEdt(TheData).CritEdt.Exo.Code);
    // Date de Début de l'édition
    DateComptaDe.Text  := DateToStr(ClassCritEdt(TheData).CritEdt.Date1);
    // Date de Fin de l'édition
    DateComptaA.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date2);
     // Exercice et dates de sélection du comparatif
    ComparExo.Value   := Exercice.Value;
    AvecComparatif.checked := ClassCritEdt(TheData).CritEdt.AvecComparatif;
    end;
  { On sauvegarde les dates de la sélection }
  FDateDeb        := StrToDate(DateComptaDe.Text);
  FDateFin        := StrToDate(DateComptaA.Text);
  inherited;
  { La première colonne de l'état correspond à l'avant-période : on cumule donc les écritures
    antérieures à la date de début de sélection en débutant au premier jour de l'exercice;
    Remarques :
    - si la sélection débute au premier jour de l'exercice, la première colonne ne
      contiendra que les écritures d'A nouveaux --> date de fin = date de début d'exercice
    - attention ce test doit être fait après "inherited" sinon le contenu est écrasé }
  QuelDateDeExo(Exercice.Value,FExoDate);
  DateComptaDe.Text := DateToStr(FExodate.deb);
  if (FdateDeb > FExodate.deb) then
    DateComptaA.Text := DateToStr(FdateDeb - 1)
  else
    DateComptaA.Text := DateToStr(FExodate.deb);
end;

procedure TOF_CPBALAULEGAL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPBALAULEGAL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPBALAULEGAL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPBALAULEGAL.InitZones;
begin
  inherited;
end;

procedure TOF_CPBALAULEGAL.ParamComparatif;
begin
  ComparExo.Value   := Exercice.Value;
  ComparDateDe.text := DateToStr(FDateDeb);
  ComparDateA.text  := DateToStr(FDateFin);
  inherited;
end;

procedure TOF_CPBALAULEGAL.ParamTitresCols;
begin
  inherited;
  SetControlText('TitreCol1',TraduireMemoire('au') + ' ' + DateToStr(FdateDeb - 1)) ;
  SetControlText('TitreCol2',TraduireMemoire('du') + ' ' + DateToStr(FdateDeb) + ' ' +
                             TraduireMemoire('au') + ' ' + DateToStr(FDateFin)) ;
  SetControlText('TitreCol3',TraduireMemoire('au') + ' ' + DateToStr(FdateFin)) ;
  SetControlText('TitreCol4',TraduireMemoire('Solde au') + ' ' + DateToStr(FdateFin)) ;
end;

function TOF_CPBALAULEGAL.RecupWhereQualifPiece: String;
begin
  Result := ' AND E_QUALIFPIECE="N"';
end;

function TOF_CPBALAULEGAL.GetConditionSQLCpt: String;
begin
  Result := inherited GetConditionSQLCpt;
end;

function TOF_CPBALAULEGAL.GetConditionSQLPeriode(vInPer: Integer): String;
{ La première colonne de l'état correspond à l'avant-période : on cumule donc les écritures
  antérieures à la date de début de sélection en débutant au premier jour de l'exercice;
  La deuxième colonne de l'état correspond à la période sélectionnée

  Remarque :
    si la sélection débute au premier jour de l'exercice, la première colonne ne
    contiendra que les écritures d'A nouveaux --> on ne prend donc que les écritures
    qui sont sur le journal d'A nouveaux OAN }

begin
  Result := inherited GetConditionSQLPeriode(vInPer);
  if (vInPer = 1) and (FdateDeb = FExodate.Deb) then
    Result := Result + ' AND E_ECRANOUVEAU="OAN"';
  if vInPer = 2 then
    Result := Result + ' AND E_ECRANOUVEAU<>"OAN"';
end;

Initialization
  registerclasses ( [ TOF_CPBALAULEGAL ] ) ;
end.

