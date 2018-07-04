{***********UNITE*************************************************
Auteur  ...... : PPZ
Créé le ...... : 11/04/2014
Modifié le ... :   /  /
Description .. : Traitement d'initialisation des affectations de la DSN
Mots clefs ... : PAIE;DSN
*****************************************************************}
{
}
unit PGInitRemDSN;

interface

uses
   sysutils,
   {$IFNDEF EAGLCLIENT}
   Fe_Main, uDbxDataSet,
   {$ELSE}
   MaineAGL,
   {$ENDIF}
   HCtrls,
   CbpDates;

procedure InitRemDSN;

implementation

procedure InitFromN4DS(NowSQL : string); forward;
function  GetCodeDSNprimeFromCodeN4DS(CodeN4DS : string) : string; forward;
function  GetCodeDSNautreFromCodeN4DS(CodeN4DS : string) : string; forward;

procedure InitFromAvantageNat(NowSQL : string); forward;
function  GetCodeDSNfromAvantageNat(AvantageNat : string) : string; forward;

procedure InitRemDSN;
var
  stNowSQL : string;
begin
  // On récupère la date de là tout de suite histoire de mettre tout le monde à la même date
  stNowSQL := UsDateTime(Trunc(Now));
  InitFromN4DS(stNowSQL);
  InitFromAvantageNat(stNowSQL);
end;

procedure InitFromN4DS(NowSQL : string);
var
  stSelect,
  stUpdatePrime,
  stUpdateAutre : string;
  qSelect  : TQuery;
  stPredefini,
  stNoDossier,
  stRubrique,
  stCodeN4DS,
  stCodeDSN : string;
begin
  // Construction de la requête de sélection des affectations N4DS
  stSelect := EmptyStr;
  stSelect := stSelect + 'SELECT PRM_PREDEFINI, ';
  stSelect := stSelect +        'PRM_NODOSSIER, ';
  stSelect := stSelect +        'PRM_RUBRIQUE, ';
  stSelect := stSelect +        'PUO_UTILSEGMENT ';
  stSelect := stSelect + 'FROM   REMUNERATION ';
  stSelect := stSelect + 'JOIN   PUBLICOTIS ';
  stSelect := stSelect + 'ON     PRM_RUBRIQUE = PUO_RUBRIQUE ';
  stSelect := stSelect + 'AND    PRM_PREDEFINI = PUO_PREDEFINI '; // On ne considère que les affectations du même niveau que les rémunérations
  stSelect := stSelect + 'WHERE  PRM_PREDEFINI <> "CEG" ';        // Les affectations de niveau CEGID ne sont pas traitées par cette outil d'initialisation
  stSelect := stSelect + 'AND    PUO_TYPEAFFECT = "4DS" ';
  stSelect := stSelect + 'AND    PUO_NATURERUB  = "REM" ';
  stSelect := stSelect + 'ORDER  BY PRM_RUBRIQUE';

  // Construction de la requête de mise à jour des rémunérations de type "Prime"
  stUpdatePrime := EmptyStr;
  stUpdatePrime := stUpdatePrime + 'UPDATE REMUNERATION ';
  stUpdatePrime := stUpdatePrime + 'SET    PRM_NATUREPRIME = "%s", ';
  stUpdatePrime := stUpdatePrime + Format('PRM_APPLICNATURE = "%s", ', [NowSQL]);
  stUpdatePrime := stUpdatePrime + Format('PRM_MODIFNATURE = "%s" ',   [NowSQL]);
  stUpdatePrime := stUpdatePrime + 'WHERE  PRM_PREDEFINI = "%s" ';
  stUpdatePrime := stUpdatePrime + 'AND    PRM_NODOSSIER = "%s" ';
  stUpdatePrime := stUpdatePrime + 'AND    PRM_RUBRIQUE = "%s"';

  // Construction de la requête de mise à jour des rémunérations de type "AutreElement"
  stUpdateAutre := EmptyStr;
  stUpdateAutre := stUpdateAutre + 'UPDATE REMUNERATION ';
  stUpdateAutre := stUpdateAutre + 'SET    PRM_AUTREELEMENT = "%s", ';
  stUpdateAutre := stUpdateAutre + Format('PRM_APPLICAUTREELE = "%s", ', [NowSQL]);
  stUpdateAutre := stUpdateAutre + Format('PRM_MODIFAUTREELEM = "%s" ',  [NowSQL]);
  stUpdateAutre := stUpdateAutre + 'WHERE  PRM_PREDEFINI = "%s" ';
  stUpdateAutre := stUpdateAutre + 'AND    PRM_NODOSSIER = "%s" ';
  stUpdateAutre := stUpdateAutre + 'AND    PRM_RUBRIQUE = "%s"';

  qSelect := OpenSQL(stSelect, True);
  try
    while not qSelect.eof do
    begin
      stPredefini := qSelect.FindField('PRM_PREDEFINI').AsString;
      stNoDossier := qSelect.FindField('PRM_NODOSSIER').AsString;
      stRubrique  := qSelect.FindField('PRM_RUBRIQUE').AsString;
      stCodeN4DS  := qSelect.FindField('PUO_UTILSEGMENT').AsString;

      // Récupération du code DSN de type "Prime" associé au code N4DS
      stCodeDSN := GetCodeDSNprimeFromCodeN4DS(stCodeN4DS);
      if stCodeDSN <> EmptyStr then
        // Exécution de l'update d'affectation DSN
        ExecuteSQL(Format(stUpdatePrime, [stCodeDSN, stPredefini, stNoDossier, stRubrique]));

      // Récupération du code DSN de type "AutreElement" associé au code N4DS
      stCodeDSN := GetCodeDSNautreFromCodeN4DS(stCodeN4DS);
      if stCodeDSN <> EmptyStr then
        // Exécution de l'update d'affectation DSN
        ExecuteSQL(Format(stUpdateAutre, [stCodeDSN, stPredefini, stNoDossier, stRubrique]));

      qSelect.Next;
    end;
  finally
    Ferme(qSelect);
  end;
end;

procedure InitFromAvantageNat(NowSQL : string);
var
  stSelect,
  stUpdate : string;
  qSelect  : TQuery;
  stPredefini,
  stNoDossier,
  stRubrique,
  stAvantageNat,
  stCodeDSN : string;
begin
  // Construction de la requête de sélection des avantages en nature
  stSelect := EmptyStr;
  stSelect := stSelect + 'SELECT PRM_PREDEFINI, ';
  stSelect := stSelect +        'PRM_NODOSSIER, ';
  stSelect := stSelect +        'PRM_RUBRIQUE, ';
  stSelect := stSelect +        'PRM_AVANTAGENAT ';
  stSelect := stSelect + 'FROM   REMUNERATION ';
  stSelect := stSelect + 'WHERE  PRM_PREDEFINI   <> "CEG" ';                              // Les affectations de niveau CEGID ne sont pas traitées par cette outil d'initialisation
  stSelect := stSelect + 'AND    PRM_AVANTAGENAT IN ("N", "L", "V", "T", "A")'; // Valeurs pour lesquelles on a une affectation

  // Construction de la requête de mise à jour des rémunérations de type "AutreElement"
  stUpdate := EmptyStr;
  stUpdate := stUpdate + 'UPDATE REMUNERATION ';
  stUpdate := stUpdate + 'SET    PRM_AUTREELEMENT = "%s", ';
  stUpdate := stUpdate + Format('PRM_APPLICAUTREELE = "%s", ', [NowSQL]);
  stUpdate := stUpdate + Format('PRM_MODIFAUTREELEM = "%s" ',  [NowSQL]);
  stUpdate := stUpdate + 'WHERE  PRM_PREDEFINI = "%s" ';
  stUpdate := stUpdate + 'AND    PRM_NODOSSIER = "%s" ';
  stUpdate := stUpdate + 'AND    PRM_RUBRIQUE = "%s"';

  qSelect := OpenSQL(stSelect, True);
  try
    while not qSelect.eof do
    begin
      stPredefini   := qSelect.FindField('PRM_PREDEFINI').AsString;
      stNoDossier   := qSelect.FindField('PRM_NODOSSIER').AsString;
      stRubrique    := qSelect.FindField('PRM_RUBRIQUE').AsString;
      stAvantageNat := qSelect.FindField('PRM_AVANTAGENAT').AsString;
      // Récupération du code DSN associé au code "Avantage en nature"
      stCodeDSN     := GetCodeDSNfromAvantageNat(stAvantageNat);
      // Exécution de l'update d'affectation DSN
      ExecuteSQL(Format(stUpdate, [stCodeDSN, stPredefini, stNoDossier, stRubrique]));

      qSelect.Next;
    end;
  finally
    Ferme(qSelect);
  end;
end;

function GetCodeDSNfromAvantageNat(AvantageNat : string) : string;
var
  codeDSN : string;
begin
  if  AvantageNat = 'N' then
    // Avantage en nature : repas
    codeDSN := '02'
  else if AvantageNat = 'L' then
    // Avantage en nature : logement
    codeDSN := '03'
  else if AvantageNat = 'V' then
    // Avantage en nature : vehicule
    codeDSN := '04'
  else if AvantageNat = 'T' then
    // Avantage en nature : NTIC
    codeDSN := '05'
  else if AvantageNat = 'A' then
    // Avantage en nature : autres
    codeDSN := '06'
  else
    codeDSN := EmptyStr;

  Result := codeDSN;
end;

function  GetCodeDSNprimeFromCodeN4DS(CodeN4DS : string) : string;
var
  iCodeN4DS : Integer;
  codeDSN : string;
begin
  iCodeN4DS := StrToInt(CodeN4DS);
  case iCodeN4DS of
    250 : codeDSN := '001';
    251 : codeDSN := '002';
    252 : codeDSN := '003';
    253 : codeDSN := '004';
    254 : codeDSN := '005';
    255 : codeDSN := '006';
    256 : codeDSN := '007';
    257 : codeDSN := '008';
    258 : codeDSN := '009';
    259 : codeDSN := '010';
    260 : codeDSN := '011';
    261 : codeDSN := '012';
    262 : codeDSN := '013';
    263 : codeDSN := '014';
    264 : codeDSN := '015';
    265 : codeDSN := '016';
    266 : codeDSN := '017';
    267 : codeDSN := '018';
    268 : codeDSN := '019';
    269 : codeDSN := '020';
    270 : codeDSN := '021';
    271 : codeDSN := '022';
    272 : codeDSN := '023';
    274 : codeDSN := '025';
    155 : codeDSN := '026';
    156 : codeDSN := '027';
    157 : codeDSN := '028';
    158 : codeDSN := '029';
    153 : codeDSN := '030';
    154 : codeDSN := '031';
    275 : codeDSN := '032';
    else
      codeDSN := EmptyStr;
  end;
  Result := codeDSN;
end;

function GetCodeDSNautreFromCodeN4DS(CodeN4DS : string) : string;
var
  iCodeN4DS : Integer;
  codeDSN : string;
begin
  iCodeN4DS := StrToInt(CodeN4DS);
  case iCodeN4DS of
    44  : codeDSN := '07';
    46  : codeDSN := '08';
    45  : codeDSN := '09';
    162 : codeDSN := '10';
    171 : codeDSN := '11';
    172 : codeDSN := '12';
    179 : codeDSN := '12';
    176 : codeDSN := '13';
    177 : codeDSN := '13';
    178 : codeDSN := '13';
    182 : codeDSN := '14';
    else
      codeDSN := EmptyStr;
  end;
  Result := codeDSN;  
end;

end.

