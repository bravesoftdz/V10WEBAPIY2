{
PT1   : 31/10/2007 GGU V_80  Gestion du calcul des compteurs de présence
}
unit uDispatchPaieServeur;

interface

uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  classes
  , UTOB
  , HEnt1
  , Ent1
  ;

procedure InitApplication;
function myDispatch(Action, Param: string; RequestTOB: TOB): TOB;
function myProcCalcEdt(sf, sp: string): variant;

implementation

uses
  sysutils,
  UTofPG_PrepAuto,
  UTOFPG_SalDotCp_mul,
  entpaie,
  P5Def,
  UtofPG_PRESCALCCOMPTEURS,  //PT1
  PGDucsOutils,   //@@
//{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} HEnt1,HCtrls,
  eSession, eDispatch;

procedure InitApplication;
begin
  ProcDispatch := myDispatch;
  ProcCalcEdt := myProcCalcEdt;
end;

function myDispatch(Action, Param: string; RequestTOB: TOB): TOB;
var //s,
    D1,D2, SQL, Etab1, Etab2, LeMode : String;
    T1 : TOB;
    BoResult : Boolean;  //PT1
    LaDate : TDateTime;
    CompteursCalculator : TCompteursCalculator;  //PT1
begin
  {JP Pour le debugage : ligne de paramètre d'exécution à lancer une fois que les fichiers
        d'entrée et de sortie ont été "récupérés -i in.tmp -o out.tmp -a cloture -p aucun".
        C'est à quoi sert la boucle sans fin ci-dessous
   while True do begin end ;}
  // Init variable Halley
  InitAGL;
//  ChargeMagHalley ;
  InitLaVariablePaie;
  ChargeParamsPaie;
//  RequestTOB.savetofile('c:\temp\TOBProcess3.log', FALSE, TRUE, TRUE);
   T1 := RequestTOB.detail [0];
  // ----------------------------------------------
  // Fonction de préparation automatique des paies
  // ----------------------------------------------
  if UpperCase(Action) = 'PREPAUTO' then
  begin
    // Instanciation processus
    D1 := T1.GetValue('DD');
    D2 := T1.GetValue('DF');
    LeMode := T1.GetValue('LEMODE');
    SQL := T1.GetValue('LeSQL');
    result := Process_PREPAUTO(D1,D2,LeMode,SQL);
  end
  // ---------------------------------------------
  // Fonction de calcul des provisions CP et RTT
  // ---------------------------------------------
  else if UpperCase(Action) = 'PROVCP' then
  begin
    // Instanciation processus
    MessageCgi ('Avt Etab='+Etab1+';'+Etab2);
    LaDate := EncodeDate (StrToInt(T1.GetValue('YY')),StrToInt(T1.GetValue('MM')),StrToInt(T1.GetValue('DD')));
    Etab1 := T1.GetValue('ETAB1');
    Etab2 := T1.GetValue('ETAB2');
    result := Process_CalculProv(LaDate, Etab1, Etab2);
  end
  // ---------------------------------------
  // Fonction de génération des ods de paie
  // ---------------------------------------
  else if UpperCase(Action) = 'ODPAIES' then
  begin
    // Instanciation processus
  end
  //PT1
  // ---------------------------------------
  // Fonction de calcul des compteurs de présence
  // ---------------------------------------
  else if UpperCase(Action) = 'CALCULCOMPTEURSPRESENCE' then
  begin
    CompteursCalculator := TCompteursCalculator.Create;
    BoResult := CompteursCalculator.LancementDuCalcul(T1);
    FreeAndNil(CompteursCalculator);
    result := TOB.Create('CALCULCOMPTEURSPRESENCEPROCESS', nil, -1);
    result.AddChampSupValeur('OKTRAITE', BoResult);
  end
  // ----------------------------------------------
  // Fonction d'initialisation des DUCS
  // ----------------------------------------------
  else if UpperCase(Action) = 'INITDUCS' then
  begin
    // Instanciation processus
    result := Process_DucsInit(T1.GetValue('PERIODETAT'),
                       T1.GetValue('DATEDEB'),
                       T1.GetValue('DATFIN'),
                       T1.GetValue('NATDEB'),
                       T1.GetValue('NATFIN'),
                       T1.GetValue('ETABDEB'),
                       T1.GetValue('ETABFIN'));
  end
  else
    result := TOB.Create('', nil, -1);

  result.AddChampSupValeur('ERROR', ''); // créer au moins (au plus ?) ce champ
  // ne pas abuser des ERROR <> '' : elles sont transmises à l'administrateur...
end;

function myProcCalcEdt(sf, sp: string): variant;
begin
  // init, au cas où on n'a pas la fonction
  result := UnAssigned; // et pas autre chose
  // traitement ici :
  // ...
  // result := ''; si et seulement si le champ doit être blanc
end;


end.

