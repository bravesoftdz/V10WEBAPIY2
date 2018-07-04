{***********UNITE*************************************************
Auteur  ...... : Muriel Dumont
Créé le ...... : 03/01/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPBALGELEGAL ()
Mots clefs ... : TOF;CPBALGELEGAL

Ce source permet d'éditer une balance générale mouvementée
Nature : CLO et Code : CFS
Remarque : on fait appel à l'état de la balance de clôture dont la présentation correspond
           tout à fait à la balance mouvementée

La première colonne de l'édition correspond à l'avant période
La deuxième colonne de l'édition correspond à la période sélectionnée par l'utilisateur ( = comparatif)
La troisième colonne de l'édition correspond au cumul des 2 premières colonnes
La quatrième colonne de l'édition correspond au solde du cumul
*****************************************************************}
Unit CPBALGELEGAL_TOF ;

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
     uLibExercice,  // CExerciceVersRelatif
     ParamSoc,      // GetParamSoc
     HTB97,
     Forms,
     AGLInit,       // TheData
     CritEdt,       // ClassCritEdt
     CPBALGEN_TOF,
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     UtilEdt;       // MajEditionLegal

  //procedure CPLanceFiche_CPBALGELEGAL(Argument: String='');
  procedure CPLanceFiche_CPBALGELEGAL(Apercu, Couleur: Boolean);

Type
  TOF_CPBALGELEGAL = Class (TOF_CPBALGEN)
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
    procedure ChargementCritEdt             ; override ;
    function  GetConditionSQLPeriode(vInPer: Integer): String; override;

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

//procedure CPLanceFiche_CPBALGELEGAL(Argument: String='');
procedure CPLanceFiche_CPBALGELEGAL(Apercu, Couleur: Boolean);
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

  AGLLanceFiche('CP','CPBALGELEGAL_QR1','','',Argument);
end;

procedure TOF_CPBALGELEGAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBALGELEGAL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPBALGELEGAL.OnUpdate ;
var
  lStReq : String;
begin
  Inherited ;
  { Maj des soldes débiteurs 1 }
  lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT1 = (CED_DEBIT1-CED_CREDIT1), CED_CREDIT1=0 '
          + ' WHERE ( CED_DEBIT1 >= CED_CREDIT1 ) AND CED_USER="' + V_PGI.User + '"' ;
  ExecuteSQL( lStReq ) ;
  { Maj des soldes créditeurs 1 }
  lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT1 = (CED_CREDIT1-CED_DEBIT1), CED_DEBIT1 = 0 '
          + ' WHERE ( CED_DEBIT1 < CED_CREDIT1 ) AND CED_USER="' + V_PGI.User + '"' ;
  ExecuteSQL( lStReq ) ;

  { Maj des soldes débiteurs 2 }
  lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_DEBIT2 = (CED_DEBIT2-CED_CREDIT2), CED_CREDIT2=0 '
          + ' WHERE ( CED_DEBIT2 >= CED_CREDIT2 ) AND CED_USER="' + V_PGI.User + '"' ;
  ExecuteSQL( lStReq ) ;
  { Maj des soldes créditeurs 2 }
  lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_CREDIT2 = (CED_CREDIT2-CED_DEBIT2), CED_DEBIT2 = 0 '
          + ' WHERE ( CED_DEBIT2 < CED_CREDIT2 ) AND CED_USER="' + V_PGI.User + '"' ;
  ExecuteSQL( lStReq ) ;

  { On remet les valeurs initiales de l'état }
  DateComptaDe.Text := DateToStr(FDateDeb);
  DateComptaA.Text  := DateToStr(FDateFin);

  MajEditionLegal('BLG',CRelatifVersExercice(EXERCICE.Value),DateComptaDe.Text,DateComptaA.Text);
end ;

procedure TOF_CPBALGELEGAL.OnLoad ;
begin
  TFQRS1(Ecran).NatureEtat := 'CLO';
  TFQRS1(Ecran).CodeEtat   := 'CFS';
  Inherited ;
end ;

procedure TOF_CPBALGELEGAL.ChargementCritEdt;
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

procedure TOF_CPBALGELEGAL.OnArgument (S : String ) ;
var
  Argument: String;
begin
  Argument        := S;
  SetControlChecked('FAPERCU',   ReadTokenSt(Argument)= 'X');
  SetControlChecked('FCOULEUR',  ReadTokenSt(Argument)= 'X');
  Inherited OnArgument(Argument);

  TitreEtat := THEdit(Getcontrol('TITREETAT', true));
  Ecran.Caption := TraduireMemoire('Balance générale');
  TFQRS1(Ecran).FNomFiltre:= 'LEGALE';
  UpdateCaption(Ecran);
  TitreEtat.Text := Ecran.Caption;
  Ecran.WindowState := wsMaximized;
end ;

procedure TOF_CPBALGELEGAL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPBALGELEGAL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPBALGELEGAL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPBALGELEGAL.InitZones;
begin
  inherited;
end;

procedure TOF_CPBALGELEGAL.ParamComparatif;
begin
  { Mise à jour des dates de comparatif correspondant aux dates sélectionnées par l'utilisateur }
  ComparExo.Value   := Exercice.Value;
  ComparDateDe.text := DateToStr(FDateDeb);
  ComparDateA.text  := DateToStr(FDateFin);
  inherited;
end;

procedure TOF_CPBALGELEGAL.ParamTitresCols;
begin
  inherited;
  SetControlText('TitreCol1',TraduireMemoire('au') + ' ' + DateToStr(FdateDeb - 1)) ;
  SetControlText('TitreCol2',TraduireMemoire('du') + ' ' + DateToStr(FdateDeb) + ' ' +
                             TraduireMemoire('au') + ' ' + DateToStr(FDateFin)) ;
  SetControlText('TitreCol3',TraduireMemoire('au') + ' ' + DateToStr(FdateFin)) ;
  SetControlText('TitreCol4',TraduireMemoire('Solde au') + ' ' + DateToStr(FdateFin)) ;
end;

function TOF_CPBALGELEGAL.RecupWhereQualifPiece: String;
begin
  Result := ' AND E_QUALIFPIECE="N"';
end;

function TOF_CPBALGELEGAL.GetConditionSQLCpt: String;
begin
  Result := inherited GetConditionSQLCpt;
end;

function TOF_CPBALGELEGAL.GetConditionSQLPeriode(vInPer: Integer): String;
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
  registerclasses ( [ TOF_CPBALGELEGAL ] ) ;
end.

