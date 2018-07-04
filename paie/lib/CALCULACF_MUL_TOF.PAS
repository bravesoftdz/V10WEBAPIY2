{***********UNITE*************************************************
Auteur  ...... : MF
Créé le ...... : 25/07/2007
Modifié le ... : 25/07/2007
Description .. : Source TOF de la FICHE : CALCULACF_MUL ()
Suite ........ : Calcul des jours de fractionnement
Mots clefs ... : TOF;CALCULACF_MUL;FRACTIONNEMENT
*****************************************************************}
Unit CALCULACF_MUL_TOF ;

Interface

Uses
     AglInit,
     StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_Main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     MaineAgl,
{$ENDIF}
     forms,
     ComCtrls,
     HTB97,
     HCtrls,
     HEnt1,
     HMsgBox,
     HQry,
     UTOB,
     UTOF ;

Type
  TOF_CALCULACF_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    NomListe                : string;
    procedure  LanceCalculFractionnement(Sender: TObject);
  end ;

Implementation

uses
  PgCommun,
  PgOutils2;

procedure TOF_CALCULACF_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CALCULACF_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CALCULACF_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CALCULACF_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CALCULACF_MUL.OnArgument (S : String ) ;
var
  MoisE,AnneeE,ComboExer  : string;
  DebExer, FinExer        : TDateTime;
  BtnLance                : TToolBarButton97;
begin

  // Choix de la liste à utiliser
  NomListe := 'PGAFFECTSALPOPUL';

  if not ExisteSQL ('SELECT 1 FROM SALARIEPOPUL WHERE  PNA_TYPEPOP="ACF"') then
  // Aucune population paramétrée
  begin
    TFMul(Ecran).SetDBListe('PGSALARIES');
    NomListe := 'PGSALARIES';
  end
  else
  // Les populations ont été paramétrées
    SetControlText('PNA_TYPEPOP','ACF');

  // On ne traite que les salariés pour lesquels on gère les congés payés
  SetControlChecked ('PSA_CONGESPAYES',True);

  Inherited ;

  // Affichage date de la cernière paie
  RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer);
  SetControlText ('MOISANNEE',RechDom('JUMOISCLOTURE',MoisE,false)+ ' ' + AnneeE);

  BtnLance := TToolBarButton97(getcontrol('BLANCE'));
  if BtnLance <> nil then
    BtnLance.Onclick := LanceCalculFractionnement;

end ;

procedure TOF_CALCULACF_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CALCULACF_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CALCULACF_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure  TOF_CALCULACF_MUL.LanceCalculFractionnement(Sender: TObject);
var
  DebPer, FinPer, ExerPerEncours        : string;
begin
  if(TFmul(Ecran).Q.RecordCount <= 0) then
  begin
    PGIBox('Aucun salarié n''est concerné par le calcul.','Calcul des jours de fractionnement');
    exit;
  end;
  if not RendPeriodeEnCours(ExerPerEncours, DebPer, FinPer) = True then
  begin
    PGIBox('Aucun exercie actif en cours. Traitement impossible','Calcul des jours de fractionnement');
    exit;
  end;

{$IFDEF EAGLCLIENT}
  TheMulQ := TOB(Ecran.FindComponent('Q'));
{$ELSE}
  TheMulQ := THQuery(Ecran.FindComponent('Q'));
{$ENDIF}
  AglLanceFiche('PAY','CALCULACF', '', '', NomListe+';'+DebPer+';'+FinPer);
end;



Initialization
  registerclasses ( [ TOF_CALCULACF_MUL ] ) ;
end.
