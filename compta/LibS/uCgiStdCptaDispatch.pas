unit uCgiStdCptaDispatch;

interface

uses
  classes,
  variants,
  HEnt1,      // DDWriteLN
  UTOB;

procedure InitApplication;
function  myDispatch(Action, Param : String ; RequestTOB : TOB) : TOB ;
function  myProcCalcEdt(sf, sp : string) : variant;
implementation

uses
  sysutils,
  uLibStdCpta, // ChargeToutLeStandardCompta
  Ent1,        // ChargeMagHalley
  AligneStd,   // TAligneStd
  eSession,
  eDispatch;

procedure InitApplication;
begin
  ProcDispatch := myDispatch;
  SetProcCalcEdt(myProcCalcEdt);
end;

procedure AffichageServer( st1, st2 : string );
begin
  DDWriteLN( St1 + ' - ' + St2 );
end;

function myDispatch(Action, Param : String ; RequestTOB : TOB) : TOB ;
var lInNumStd       : integer;
    lStLibelleStd   : string;
    lStAbregeStd    : string;
    lAligneStd      : TAligneStd;
    lstParamStd     : string;
begin
  //while True do begin end;

  InitAGL;

  lStLibelleStd := '';
  lStAbregeStd  := '';
  lstParamStd   := '';

  // Init variable Halley
  InitLaVariableHalley;
  ChargeMagHalley;

  DDWriteLN('Fonction MyDispatch');

  Result := TOB.Create('le result', nil, -1);
  // result doit être virtuelle, mais elle peut avoir des filles réelles
  // traitement ici :
  // ...

  // Chargement d'un standard
  if Action = 'CHARGESTD' then
  begin
    lInNumStd := StrToInt(Param);
    DDWriteLN( 'Entrée ChargeSTD');

    if not ChargeToutLeStandard( lInNumStd , True, False ) then
      Result.AddChampSupValeur('RESULT', 'PASOKCHARGE')
    else
      Result.AddChampSupValeur('RESULT', 'OKCHARGE');

    DDWriteLN( 'Sortie ChargeSTD');
  end
  else
  if Action = 'ENREGSTD' then
  begin
    lInNumStd := StrToInt(Param);

    if RequestTOB.FieldExists('LIBELLESTD') then
      lStLibelleStd := RequestTOB.GetValue('LIBELLESTD');
    if RequestTOB.FieldExists('ABREGESTD') then
      lStAbregeStd := RequestTOB.GetValue('ABREGESTD');

    if not EnregistreToutLeStandard( lInNumStd, lStLibelleStd, lStAbregeStd, False) then
      Result.AddChampSupValeur('RESULT', 'PASOKENREG')
    else
      Result.AddChampSupValeur('RESULT', 'OKENREG');
  end
  else
  if Action = 'ALIGNESTD' then
  begin
    lAligneStd := TAligneStd.Create;
    lAligneStd.OnInformation := nil;
    lAligneStd.ChargeAligneStdAvecTob( RequestTob );

    // GCO - 31/08/2007 - FQ 21289 Lecture du champ Sup. RAZFILTRE sur la
    // TOBMERE RequestTob
    if RequestTob.FieldExists('RAZFILTRES') then
      lALigneStd.FRazFiltres := RequestTob.GetString('RAZFILTRES') = 'X';
    // FIN GCO

    lAligneStd.Execute;
    lAligneStd.ChargeResultat( Result );
    if lAligneStd.LastError = 0 then
      Result.AddChampSupValeur('RESULT', 'OKALIGNE')
    else
      Result.AddChampSupValeur('RESULT', 'PASOKALIGNE');

    FreeAndNil(lAligneStd);
  end else
  if Action = 'SAVEDOSSIERASSTANDARD' then
  begin
    if RequestTOB.FieldExists('NUMSTD') then
    begin
      lInNumStd := RequestTOB.GetValue('NUMSTD');
      if RequestTOB.FieldExists('LIBELLESTD') then
        lstLibelleStd := RequestTOB.GetValue('LIBELLESTD');
      if RequestTOB.FieldExists('PARAMSTD') then
        lstParamStd := RequestTOB.GetValue('PARAMSTD');
      if SaveDossierAsStandard(lInNumStd, lstParamStd,lstLibelleStd) <> 0 then
        Result.AddChampSupValeur('RESULT', 'PASOK')
      else Result.AddChampSupValeur('RESULT', 'OK');
    end else Result.AddChampSupValeur('RESULT', 'PASOK');
  end;

  Result.AddChampSupValeur('ERROR', ''); // créer au moins (au plus ?) ce champ
  // ne pas abuser des ERROR <> '' : elles sont transmises à l'administrateur...
  DDWriteLN( 'Sortie MyDispatch');
end;

function myProcCalcEdt(sf, sp : string) : variant;
begin
  // init, au cas où on n'a pas la fonction
  result := UnAssigned; // et pas autre chose
  // traitement ici :
  // ...
  // result := ''; si et seulement si le champ doit être blanc
end;

end.

