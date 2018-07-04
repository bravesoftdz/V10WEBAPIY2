{***********UNITE*************************************************
Auteur  ...... : PAIE PGI : GGU
Créé le ...... : 20/01/2007
Modifié le ... :   /  /
Description .. : Gestion des tables dynamiques
Mots clefs ... : PAIE;tables dynamiqes;
*****************************************************************}
{
PT1    : 05/06/2007 FC V_72 Mémorisation des codes libres dont la valeur est manquante pour les proposer à la saisie
PT2    : 27/11/2007 GGU V_80 Gestion du log d'un code type
PT3 13/12/2007 V8 GGU FQ 14620 Ajout d'une 4e nature de table : Elément national
PT4 13/12/2007 V8 GGU FQ 15058 Gestion des sens pour les critères des tables dynamiques
PT5 16/01/2008 V8 GGU Les tables dynamiques variables renvoyant des réels sont mal gérées
PT7 21/04/2008 GGU V81 FQ 15361 Gestion uniformisée des zones libres - tables dynamiques
PT8 07/05/2008 GGU V81 FQ 15361 Si je sélectionne une date qui ne correspond à aucune valeur message de violation d'accès s'affiche.
PT9 16/05/2008 GGU V81 FQ 15437 table dynamique avec elt nationaux , resultat VALEUR : on a le code au lieu  de la valeur de l'élément national
PT15 10/09/2008 FC Tester la longueur des chaines comparées quand on récupère le résultat d'une table dynamique
}
unit PGTablesDyna;

interface
{$IFDEF PAIEGRH}
{$DEFINE TABLESDYNA}
{$ENDIF}
{$IFDEF CGIPAIE}
{$DEFINE TABLESDYNA}
{$ENDIF}

Uses
  {$IFNDEF EAGLCLIENT}
    db,
    {$IFNDEF EAGLSERVER} Fiche, FichList, {$ENDIF}
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$else}
    eFiche, eFichList,
  {$ENDIF}
    Controls, HEnt1, HCtrls, UTob , sysutils, StdCtrls, Classes;

  Type
    TTablesDynamiques = class (TObject)
  public
    constructor Create;
    destructor Destroy; override;
    { Cherche le libellé d'une table dynamique de type Code Libre en fonction du salarié
      l'établissement et la convention du salarié sont de paramètres optionnels qui
      permettent d'accélérer les temps de traitement }
    function GetValueFromDynaTableCOD(Salarie, CodeTable: String;
      Date: TDate; Code : String; SalEtab : String = ''; SalConv : String = ''; GetLibelle : Boolean = True): String;
{$IFDEF TABLESDYNA}
    { Appel la fonction appropriée DSA ou VAR }
    function GetValueFromDynaTable(Salarie, CodeTable: String; CleSynElt : String; var ValidOK : Boolean;  //PT1
      const DateDeb, DateFin: TDateTime; TOB_Rub: TOB;
       SalEtab : String = ''; SalConv : String = ''; DateReferenceAge : TDateTime = 0; DateReferenceAnciennete : TDateTime = 0 ; Diag: TObject = nil): Double;
    { Cherche la valeur d'une table dynamique de type Donnée salarié
      Le parametre DateReference permet de fixer une date de référence pour le calcul
      de l'age et de l'ancienneté. Si ce paramettre n'est pas passé ou est à 0 ou idate1900
      on utilise le parametre date en tant que date de référence
      l'établissement et la convention du salarié sont de paramètres optionnels qui
      permettent d'accélérer les temps de traitement  }
    function GetValueFromDynaTableDSA(Salarie, CodeTable: String;
      Date : TDateTime ; SalEtab : String = ''; SalConv : String = '' ; DateReferenceAge : TDateTime = 0; DateReferenceAnciennete : TDateTime = 0; Entete : TOB = nil): Double;
    { Cherche la valeur d'une table dynamique de type variable
      l'établissement et la convention du salarié sont de paramètres optionnels qui
      permettent d'accélérer les temps de traitement
      DateDeb, DateFin et TOB_Rub servent au calcul de la variable }
    function GetValueFromDynaTableVAR(Salarie, CodeTable: String;
      const DateDeb, DateFin: TDateTime; TOB_Rub: TOB;
      SalEtab : String = ''; SalConv : String = ''; Diag: TObject = nil; Entete : TOB = nil): Double;
    { PT3
      Cherche la valeur d'une table dynamique de type élément national
      l'établissement et la convention du salarié sont de paramètres optionnels qui
      permettent d'accélérer les temps de traitement
      DateDeb sert au calcul de la variable }
    function GetValueFromDynaTableELT(Salarie, CodeTable: String;
      const DateDeb : TDateTime;
      SalEtab : String = ''; SalConv : String = ''; Diag: TObject = nil; Entete : TOB = nil): Double;  //PT3
{$ENDIF}
  private
    TobEntetesTable, TobEnteteTableGen, TobEnteteTableTTCon, TobEnteteTableTTEtb : TOB;
    LastCodeTable : String;
    TobDetailsTable : TOB;
    { Gestion d'une liste de sauvegarde des valeurs du salarié courrant }
    LastSal : String;
    LastSalEtab, LastSalConv : String;
    SavLastSalResultCodeTable : Array of String;
    SavLastSalResultValue : Array of Double;
{$IFDEF TABLESDYNA}
    ClePGSynEltNAt : String;
    ValidationOK:Boolean;
    function GetDonneeSalarie(Salarie: String; Critere: Integer; TypeCritere: String; DateReferenceAge, DateReferenceAnciennete: TDateTime): String;
{$ENDIF}
    function GetEnteteFromDynaTable(Salarie, CodeTable: String; Date : TDateTime; SalEtab : String = ''; SalConv : String = '') : TOB;
    procedure GetEtabConv(Salarie: String; var SalEtab, SalConv: String);
    function TraiteSensTOB(VALCRIT, Donnee, sens: String; var LastVALCRIT: String): Boolean;
    { Gestion d'une liste de sauvegarde des valeurs du salarié courrant }
{$IFDEF TABLESDYNA}
    procedure EvtOnChangeSal(NewSal : String);
    function SearchResultOnSaveSalResult(var Valeur : Double) : Boolean;
    procedure AddResultOnSaveSalResult(Valeur : Double);
    procedure MemElt(Salarie,CodeTable,Libelle,NivSais,ValNiv,Predefini:STring;DatVal:TDateTime);  //PT1
{$ENDIF}
  end;

  function lpad(Str: String; Size : Integer; Pad : String = '0') : String;

  function GetPlusPGCOMBOZONELIBRE(DateApplication : TDateTime; CodeTable, Salarie : String) : String;
  function GetLibelleZLTableDyna(DateApplication : TDateTime; CodeTable, Salarie, Valeur : String) : String; //PT6

const
  { Nombre de caractères et Format des codes des critères 1 et 2 des tables dynamiques }
  NbrC_DSA : integer = 4; FormatC_DSA : String = '0000;0;0';
  NbrC_COD : integer = 3; FormatC_COD : String = '000;0;0';
  NbrC_ELT : integer = 17; FormatC_ELT : String = '00000000000000000;0;0'; //PT3
  NbrC_VAR : integer = 17; FormatC_VAR : String = '00000000000000000;0;0';
implementation

uses
{$IFNDEF EAGLSERVER}
   uTobdebug,
{$ENDIF}
  {$IFDEF TABLESDYNA} P5Util, {$ENDIF}
   Variants, StrUtils
  , ParamSoc  //PT4
  ;

function lpad(Str: String; Size : Integer; Pad : String = '0') : String;
begin
  result := Str;
  While length(result) < Size do
  begin
    result := Pad + result;
  end;
end;

Procedure GetEtablissementConventionSalarie(Salarie : String; var SalEtab, SalConv : String);
var
  QEtabConv : TQuery;
  stQEtabConv  : String;
begin
  SalEtab := '...'; //PT7
  SalConv := '000'; //PT7
  if Salarie <> '' then //PT7
  begin
    stQEtabConv   := 'SELECT PSA_ETABLISSEMENT, PSA_CONVENTION '
                   + ' FROM SALARIES '
                   + 'WHERE PSA_SALARIE   = "'+Salarie+'" ';
    QEtabConv := OpenSQL(stQEtabConv,True,1);
    if not QEtabConv.Eof then
    begin
      SalEtab := QEtabConv.FindField('PSA_ETABLISSEMENT').AsString;
      SalConv := QEtabConv.FindField('PSA_CONVENTION').AsString;
    end;
    Ferme(QEtabConv);
  end;
end;

function GetPlusPGCOMBOZONELIBRE(DateApplication : TDateTime; CodeTable, Salarie : String) : String;
var
{
  _ConvSal, _Etab : String;
  _Predefini, _NivSaisie, _ValSaisie : String;
  _DateValid : TDateTime;
  _St : String;
  _Q : TQuery; //PT2
}
  ObjTableDyn : TTablesDynamiques;
  TobEnt : Tob;  
  Function WherePredefini(NiveauPredefini : String; Prefixe : String = 'PTE') : String;
  begin
    if NiveauPredefini = 'CEG' then
      result := ' AND '+Prefixe+'_PREDEFINI="CEG" '
    else if NiveauPredefini = 'STD' then
      result := ' AND '+Prefixe+'_PREDEFINI="STD" '
    else if NiveauPredefini = 'DOS' then
      result := ' AND '+Prefixe+'_PREDEFINI="DOS" AND '+Prefixe+'_NODOSSIER="' + V_PGI.NoDossier + '" ';
  end;
  Function WhereSaisie(NiveauSaisie, EtabSalarie, ConvSalarie : String; Prefixe : String = 'PTE') : String;
  begin
    if NiveauSaisie = 'GEN' then
      result := ' AND '+Prefixe+'_NIVSAIS="GEN" '
    else if NiveauSaisie = 'ETB' then
      result := ' AND '+Prefixe+'_NIVSAIS="ETB" AND '+Prefixe+'_VALNIV="' + EtabSalarie + '" '
    else if NiveauSaisie = 'CON' then
      result := ' AND '+Prefixe+'_NIVSAIS="CON" AND '+Prefixe+'_VALNIV="' + ConvSalarie + '" ';
  end;
{
  Function WhereComplet(NiveauPredefini, NiveauSaisie, EtabSalarie, ConvSalarie : String; Prefixe : String = 'PTE') : String;
  begin
    result := ' WHERE '+Prefixe+'_CODTABL="'+CodeTable+'" AND '+Prefixe+'_DTVALID<="'+UsDateTime(DateApplication)+'" '
            + WherePredefini(NiveauPredefini, Prefixe)
            + WhereSaisie(NiveauSaisie, EtabSalarie, ConvSalarie, Prefixe);
  end;
  Function RequeteExist(NiveauPredefini, NiveauSaisie, EtabSalarie, ConvSalarie : String; Prefixe : String = 'PTE') : String;
  begin
    result := 'SELECT PTE_CODTABL FROM TABLEDIMENT'
            + WhereComplet(NiveauPredefini, NiveauSaisie, EtabSalarie, ConvSalarie, Prefixe);
  end;
  Procedure ExisteTableDyna(Etablissement, Convention : String; var NiveauPredefini, NiveauSaisie, ValeurSaisie : String; var DateValidite : TDateTime); //PT6
  begin
    NiveauPredefini := 'DOS';
    NiveauSaisie := 'ETB';
    ValeurSaisie := Etablissement; //PT6
    //Vérifier s'il n'existe pas des valeurs pour l'établissement du salarié
    if not ExisteSQL(RequeteExist(NiveauPredefini, NiveauSaisie, ValeurSaisie, ValeurSaisie)) then
    begin
      NiveauSaisie := 'GEN';
      //Vérifier s'il n'existe pas des valeurs pour le dossier en général
      if not ExisteSQL(RequeteExist(NiveauPredefini, NiveauSaisie, ValeurSaisie, ValeurSaisie)) then
      begin
        NiveauPredefini := 'STD';
        NiveauSaisie := 'CON';
        ValeurSaisie := Convention; //PT6
        //Vérifier s'il n'existe pas des valeurs pour STD + convention
        if not ExisteSQL(RequeteExist(NiveauPredefini, NiveauSaisie, ValeurSaisie, ValeurSaisie)) then
        begin
          ValeurSaisie := '000';
          //Vérifier s'il n'existe pas des valeurs pour STD + Convention 000
          if not ExisteSQL(RequeteExist(NiveauPredefini, NiveauSaisie, ValeurSaisie, ValeurSaisie)) then
          begin
            NiveauSaisie := 'GEN';
            //Vérifier s'il n'existe pas des valeurs pour STD + GEN
            if not ExisteSQL(RequeteExist(NiveauPredefini, NiveauSaisie, ValeurSaisie, ValeurSaisie)) then
            begin
              NiveauPredefini := 'CEG';
            end;
          end;
        end;
      end;
    end;
    _St := 'SELECT MAX(PTE_DTVALID) AS DTVALID FROM TABLEDIMENT'+ WhereComplet(NiveauPredefini, NiveauSaisie, ValeurSaisie, ValeurSaisie);
    _Q := OpenSQL(_St,True,1);
    DateValidite := iDate1900;
    if not _Q.Eof then
      DateValidite := _Q.FindField('DTVALID').AsDateTime;
    Ferme(_Q);
  end;
}
begin
  if DateApplication = iDate1900 then
    DateApplication := Date;
  ObjTableDyn := TTablesDynamiques.Create; // PT175
  TobEnt := ObjTableDyn.GetEnteteFromDynaTable(Salarie, CodeTable, DateApplication);
  if Assigned(TobEnt) then     //PT8
    result := ' WHERE PTD_DTVALID = "' + USDATETIME(TobEnt.GetDateTime('PTE_DTVALID')) + '" AND PTD_CODTABL="'+CodeTable+'"'
            + WherePredefini(TobEnt.GetString('PTE_PREDEFINI'), 'PTD')
            + WhereSaisie(TobEnt.GetString('PTE_NIVSAIS'), TobEnt.GetString('PTE_VALNIV'), TobEnt.GetString('PTE_VALNIV'), 'PTD')
  else
    result := ' WHERE (1 = 2)'; //PT8
  FreeAndNil(ObjTableDyn);
{
  GetEtablissementConventionSalarie(Salarie, _Etab, _ConvSal);
  ExisteTableDyna(_Etab, _ConvSal, _Predefini, _NivSaisie, _ValSaisie, _DateValid); //PT6
  result := ' WHERE PTD_DTVALID = "' + USDATETIME(_DateValid) + '" AND PTD_CODTABL="'+CodeTable+'"'
          + WherePredefini(_Predefini, 'PTD')
          + WhereSaisie(_NivSaisie, _ValSaisie, _ValSaisie, 'PTD');
}
end;

function GetLibelleZLTableDyna(DateApplication : TDateTime; CodeTable, Salarie, Valeur : String) : String; //PT6
var
  Q : TQuery;
  SQL : String;
begin
  SQL := ' SELECT ##TOP 1## PTD_LIBELLECODE FROM TABLEDIMDET '
       + GetPlusPGCOMBOZONELIBRE(DateApplication, CodeTable, Salarie)
       + ' AND PTD_VALCRIT1="' + Valeur + '"';
  Q := OpenSQL(SQL,True);
  If Not Q.Eof then Result := Q.FindField('PTD_LIBELLECODE').AsString;
  Ferme(Q);
end;

{ TTablesDynamiques }

constructor TTablesDynamiques.Create;
begin
  inherited;
  LastCodeTable := '';
  LastSal:= '';
  LastSalEtab := '';
  LastSalConv := '';
  TobEntetesTable := TOB.Create('Tob entetes tables dyna',nil,-1);
  TobDetailsTable := TOB.Create('Tob details tables dyna',nil,-1);
  SetLength(SavLastSalResultCodeTable,0);
  SetLength(SavLastSalResultValue,0);
end;

destructor TTablesDynamiques.Destroy;
begin
  LastCodeTable := '';
  LastSal:= '';
  LastSalEtab := '';
  LastSalConv := '';
  if assigned(TobEntetesTable) then FreeAndNil(TobEntetesTable);
  if assigned(TobDetailsTable) then FreeAndNil(TobDetailsTable);
  SetLength(SavLastSalResultCodeTable,0);
  SetLength(SavLastSalResultValue,0);
  inherited;
end;

{$IFDEF TABLESDYNA}
Function TTablesDynamiques.GetDonneeSalarie(Salarie : String; Critere : Integer; TypeCritere : String; DateReferenceAge,  DateReferenceAnciennete : TDateTime) : String;
var
  Prefixe, Suffixe, StTypeDate : String;
  DateReference : TDateTime;
  Procedure GetPrefixSuffix(Critere : integer; var Prefix, Suffix : String);
  var
    QChamps : TQuery;
    stQChamps : String;
  begin
    stQChamps  := 'SELECT PAI_PREFIX, PAI_SUFFIX '
                 + ' FROM PAIEPARIM '
                 + 'WHERE ##PAI_PREDEFINI## AND PAI_IDENT = '+intToStr(Critere);
    QChamps := OpenSQL(stQChamps,True,1);
    if not QChamps.Eof then
    begin
      Prefix := QChamps.FindField('PAI_PREFIX').AsString;
      Suffix := QChamps.FindField('PAI_SUFFIX').AsString;
    end else begin
      Prefix := '';
      Suffix := '';
    end;
    Ferme(QChamps);
  end;
  Function GetDonnee(Prefix, Suffix : String) : Variant;
  var
    QDonnee : TQuery;
    stQDonnee : String;
    FlagAnciennete : Boolean;
  begin
    { Si on cherche l'ancienneté, il faut ramener l'ancienneté et la date d'entrée
      au cas ou l'ancienneté ne serait pas renseignée }
    FlagAnciennete := False;
    If Suffix = 'DATEANCIENNETE' then
    begin
      FlagAnciennete := True;
      Suffix := 'DATEANCIENNETE, PSA_DATEENTREE';
    end;
    stQDonnee  := 'SELECT '+Prefix+'_'+Suffix
                + ' FROM '+PrefixeToTable(Prefix)+' '
                + 'WHERE '+Prefix+'_SALARIE = "'+Salarie+'" ';
    QDonnee := OpenSQL(stQDonnee, True, 1);
    if not QDonnee.Eof then
    begin
      if FlagAnciennete then
      begin
        Result :=  QDonnee.FindField('PSA_DATEANCIENNETE').AsDateTime;
        if result = iDate1900 then
          Result :=  QDonnee.FindField('PSA_DATEENTREE').AsDateTime;
      end else
        Result :=  QDonnee.FindField(Prefix+'_'+Suffix).AsVariant;
    end;
    Ferme(QDonnee);
  end;
begin
  Result := '';
  Prefixe := '';
  if (Critere = 1500) or (Critere = 1501) then
  begin
    if Critere = 1500 then
    begin
      StTypeDate := 'DATENAISSANCE';
      DateReference := DateReferenceAge;
    end else begin
      StTypeDate := 'DATEANCIENNETE';
      DateReference := DateReferenceAnciennete;
    end;
    if      TypeCritere = 'ANN' then Result := IntToStr(Round( AncienneteAnnee(VarToDateTime(GetDonnee('PSA',StTypeDate)), DateReference)))
    else if TypeCritere = 'MOI' then Result := IntToStr(Round( AncienneteMois(VarToDateTime(GetDonnee('PSA',StTypeDate)), DateReference)))
    else if TypeCritere = 'JOU' then Result := IntToStr(Round( DateReference - VarToDateTime(GetDonnee('PSA',StTypeDate))));
    Result := lpad(Result, NbrC_DSA);
  end else begin
    GetPrefixSuffix(Critere,Prefixe, Suffixe);
    if Prefixe <> '' then Result := GetDonnee(Prefixe, Suffixe);
  end;
end;
{$ENDIF}

Procedure TTablesDynamiques.GetEtabConv(Salarie : String; var SalEtab, SalConv : String);
begin
  GetEtablissementConventionSalarie(Salarie, SalEtab, SalConv);
end;

Function TTablesDynamiques.GetEnteteFromDynaTable(Salarie, CodeTable: String; Date : TDateTime; SalEtab : String = ''; SalConv : String = '') : TOB;// ; Date: TDate
var
  stQCriteres: String; // , StKeyNiveau
  TobEnteteTableCon, TobEnteteTableEtb : TOB;
begin
  if SalEtab <> '' then  LastSalEtab := SalEtab;
  if SalConv <> '' then  LastSalConv := SalConv;
  if (LastSalEtab = '') or (LastSalConv = '') then
    GetEtabConv(Salarie,LastSalEtab,LastSalConv);
  if LastCodeTable <> CodeTable then
  begin
    stQCriteres := 'SELECT * FROM TABLEDIMENT '
                 + 'WHERE ##PTE_PREDEFINI## '
                 + '  AND PTE_CODTABL   = "'+CodeTable+'" '
                 + '  AND PTE_DTVALID <= "'+USDATETIME(Date)+'"'
                 + ' ORDER BY PTE_PRIORITENIV DESC, PTE_DTVALID DESC';
    TobEntetesTable.LoadDetailFromSQL(stQCriteres);
    LastCodeTable := CodeTable;
    TobEnteteTableGen := TobEntetesTable.FindFirst(['PTE_NIVSAIS'],['GEN'],False);
    TobEnteteTableTTCon := TobEntetesTable.FindFirst(['PTE_NIVSAIS','PTE_VALNIV'],['CON','000'],False);
    TobEnteteTableTTEtb := TobEntetesTable.FindFirst(['PTE_NIVSAIS','PTE_VALNIV'],['ETB','...'],False);
  end;
  TobEnteteTableCon := TobEntetesTable.FindFirst(['PTE_NIVSAIS','PTE_VALNIV'],['CON',LastSalConv],False);
  TobEnteteTableEtb := TobEntetesTable.FindFirst(['PTE_NIVSAIS','PTE_VALNIV'],['ETB',LastSalEtab],False);

  if Assigned(TobEnteteTableEtb) then
    Result := TobEnteteTableEtb
  else if Assigned(TobEnteteTableCon) then
    Result := TobEnteteTableCon
  else if Assigned(TobEnteteTableTTEtb) then
    Result := TobEnteteTableTTEtb
  else if Assigned(TobEnteteTableTTCon) then
    Result := TobEnteteTableTTCon
  else
    Result := TobEnteteTableGen;
end;

Function TTablesDynamiques.TraiteSensTOB(VALCRIT, Donnee, sens : String; var LastVALCRIT : String) : Boolean;
var
  ValNewSens : Boolean; //PT4
begin
  result := False;
  if sens = '' then sens := '=';
  { On regarde le paramsoc pour savoir si le changement à eu lieu }
  ValNewSens := GetParamSocSecur('SO_TDYCHGTSENS', False);
  if (ValNewSens and (sens = '<=')) or ((not ValNewSens) and (sens = '>=')) then  //PT4   sens = '>='
    if (VALCRIT >= Donnee) and ((VALCRIT <= LastVALCRIT) or (LastVALCRIT = '')) then
      result := True;
  if (ValNewSens and (sens = '<' )) or ((not ValNewSens) and (sens = '>' )) then  //PT4   sens = '>'
    if (VALCRIT >  Donnee) and ((VALCRIT <= LastVALCRIT) or (LastVALCRIT = '')) then
      result := True;
  if (ValNewSens and (sens = '>=')) or ((not ValNewSens) and (sens = '<=')) then  //PT4   sens = '<='
    if (VALCRIT <= Donnee) and ((VALCRIT >= LastVALCRIT) or (LastVALCRIT = '')) then
      result := True;
  if (ValNewSens and (sens = '>' )) or ((not ValNewSens) and (sens = '<' )) then  //PT4   sens = '<'
    if (VALCRIT <  Donnee) and ((VALCRIT >= LastVALCRIT) or (LastVALCRIT = '')) then
      result := True;
  if sens = '=' then
    if VALCRIT = Donnee then
      result := True;
end;

{$IFDEF TABLESDYNA}
function TTablesDynamiques.GetValueFromDynaTableVAR(Salarie, CodeTable: String; const DateDeb, DateFin: TDateTime; TOB_Rub: TOB; SalEtab : String = ''; SalConv : String = '' ; Diag: TObject = nil; Entete : TOB = nil): Double;
var
  ResultatVariable : String;
begin
  if Entete = nil then
    Entete := GetEnteteFromDynaTable(Salarie, CodeTable, DateDeb , SalEtab, SalConv);
  ResultatVariable := lpad(floatToStr(Trunc(ValVariable(lpad(Entete.GetString('PTE_CRITERE1'), 4 ), DateDeb, DateFin, TOB_Rub, Diag))), NbrC_VAR ); // PT5 : Ajout de Trunc()
  result := StrToFloat(GetValueFromDynaTableCOD(Salarie, CodeTable, Date, ResultatVariable, SalEtab, SalConv, False));
//Pour la nature Variable on n'a rien à demander
//  if result = -123456 then
//  begin
//    MemElt(Salarie,CodeTable,Entete.GetString('PTE_LIBELLE'),Entete.GetString('PTE_NIVSAIS'),Entete.GetString('PTE_VALNIV'),Entete.GetString('PTE_PREDEFINI'),DateDeb);
//    result := 0;
//  end;
end;

//début PT3
function TTablesDynamiques.GetValueFromDynaTableELT(Salarie, CodeTable: String; const DateDeb: TDateTime; SalEtab : String = ''; SalConv : String = '' ; Diag: TObject = nil; Entete : TOB = nil): Double;
var
  ResultatElementNational : String;
begin
  if Entete = nil then
    Entete := GetEnteteFromDynaTable(Salarie, CodeTable, DateDeb , SalEtab, SalConv);
  ResultatElementNational := lpad(floatToStr(Trunc(ValEltNat(lpad(Entete.GetString('PTE_CRITERE1'), 4 ), DateDeb, Diag))), NbrC_ELT );    // PT5 : Ajout de Trunc()
  result := StrToFloat(GetValueFromDynaTableCOD(Salarie, CodeTable, Date, ResultatElementNational, SalEtab, SalConv, False));
  if result = -123456 then
  begin
    MemElt(Salarie,CodeTable,Entete.GetString('PTE_LIBELLE'),Entete.GetString('PTE_NIVSAIS'),Entete.GetString('PTE_VALNIV'),Entete.GetString('PTE_PREDEFINI'),DateDeb);
    result := 0;
  end;
end;
//Fin PT3

function TTablesDynamiques.GetValueFromDynaTable(Salarie,
  CodeTable :String; CleSynElt : String; var ValidOK : Boolean; const DateDeb, DateFin: TDateTime; TOB_Rub: TOB;    //PT1
  SalEtab, SalConv: String; DateReferenceAge,
  DateReferenceAnciennete: TDateTime; Diag: TObject): Double;
var
  Entete : TOB;
  Code : String;
  QCode : TQuery;
  stQCode : String;
  NewDate : TDateTime;
begin
  ClePGSynEltNAt := CleSynElt;   //PT1
  ValidationOK := ValidOK;
  result := 0;
  Entete := GetEnteteFromDynaTable(Salarie, CodeTable, DateDeb , SalEtab, SalConv);
  if Entete <> nil then
  begin
    if Entete.GetString('PTE_NATURETABLE') = 'DSA' then
    begin
      result := GetValueFromDynaTableDSA(Salarie,CodeTable,DateDeb,SalEtab,SalConv,DateReferenceAge, DateReferenceAnciennete, Entete);
//Debut PT194
//      if Diag <> nil then Diag.Items.add('   '+TraduireMemoire('avec pour résultat : ') + FloatToStr(result));
      LogMessage(Diag, '   '+TraduireMemoire('avec pour résultat : ') + FloatToStr(result), 'TDY', CodeTable, '', False, -1, RechDom('PGNATURETABLE','DSA', False));
//Fin PT194
    end else if Entete.GetString('PTE_NATURETABLE') = 'VAR' then
    begin
      result := GetValueFromDynaTableVAR(Salarie, CodeTable, DateDeb, DateFin, TOB_Rub, SalEtab, SalConv, Diag, Entete);
//Debut PT194
//      if Diag <> nil then Diag.Items.add('   '+TraduireMemoire('avec pour résultat : ') + FloatToStr(result));
      LogMessage(Diag, '   '+TraduireMemoire('avec pour résultat : ') + FloatToStr(result), 'TDY', CodeTable, '', False, -1, RechDom('PGNATURETABLE','VAR', False));
//Fin PT194
    end else if Entete.GetString('PTE_NATURETABLE') = 'ELT' then //PT3
    begin
      result := GetValueFromDynaTableELT(Salarie, CodeTable, DateDeb, SalEtab, SalConv, Diag, Entete);
      LogMessage(Diag, '   '+TraduireMemoire('avec pour résultat : ') + FloatToStr(result), 'TDY', CodeTable, '', False, -1, RechDom('PGNATURETABLE','ELT', False));
    end else if Entete.GetString('PTE_NATURETABLE') = 'COD' then
    begin
      Code := '';
      NewDate := DateDeb;
      stQCode := ' SELECT PHD_NEWVALEUR, PHD_DATEAPPLIC FROM PGHISTODETAIL '
                      + ' WHERE PHD_PGTYPEINFOLS = "ZLS" AND PHD_CODTABL = "'+CodeTable+'" AND PHD_SALARIE = "'+Salarie+'" '
                      + ' AND PHD_DATEAPPLIC = (SELECT MAX(PHD_DATEAPPLIC) '
                      +                       ' FROM PGHISTODETAIL '
                      +                       ' WHERE PHD_PGTYPEINFOLS = "ZLS" '
                      +                       ' AND PHD_CODTABL = "'+CodeTable+'" '
                      +                       ' AND PHD_SALARIE = "'+Salarie+'" '
                      +                       ' AND PHD_DATEAPPLIC <= "'+ USDATETIME(DateDeb) +'") '
                      + ' ORDER BY PHD_DATECREATION DESC ';
      QCode := OpenSQL(stQCode, True, 1);
      if not QCode.Eof then
      begin
        Code := QCode.Fields[0].AsString;
        NewDate := QCode.Fields[1].AsDateTime;
      end;
      Ferme(QCode);
      result := StrToFloat(GetValueFromDynaTableCOD(Salarie, CodeTable, NewDate, Code, SalEtab, SalConv, False));
      //DEB PT1
      if result = -123456 then
      begin
        MemElt(Salarie,CodeTable,Entete.GetString('PTE_LIBELLE'),Entete.GetString('PTE_NIVSAIS'),Entete.GetString('PTE_VALNIV'),Entete.GetString('PTE_PREDEFINI'),DateDeb);
        result := 0;
      end;
      //FIN PT1
//Debut PT194
//      if Diag <> nil then Diag.Items.add('   '+TraduireMemoire('avec pour valeur d''entrée "') + Code + TraduireMemoire('" et pour résultat ') + FloatToStr(result));
      LogMessage(Diag, '   '+TraduireMemoire('avec pour valeur d''entrée "') + Code + TraduireMemoire('" et pour résultat ') + FloatToStr(result), 'TDY', CodeTable, '', False, -1, RechDom('PGNATURETABLE','COD', False));
//Fin PT194
    end;// else begin Raise Exception.Create('Aucune table ne correspond à ces critères');
  end else begin
      LogMessage(Diag, '   Table inconnue pour ce salarié', 'TDY', CodeTable);
  end;
  ValidOK := ValidationOK;
end;

//DEB PT1
procedure TTablesDynamiques.MemElt(Salarie,CodeTable,Libelle,NivSais,ValNiv,Predefini:STring;DatVal:TDateTime);
var
  TobCodeLibre,TC:Tob;
  sDate : String;
  St:String;
begin
  ValidationOK := False;

  if TypeTraitement = 'PREPA' then
  begin
    { Si on est en préparation automatique, ajouter un message d'erreur dans la liste }
    St := 'Le code libre ' + CodeTable + ' n''a pas de valeur au niveau salarié ' + Salarie + '. Le bulletin n''a pas pu être calculé juste pour le salarié ' + Salarie;
    TraceErr.Items.Add(St);
  end;

  { Récupérer les valeurs par défaut dans le cas où un niveau plus générique existe (DOS ou STD) }
  TobCodeLibre := TOB.Create('Lescodeslibres', Nil, -1);
  TC := TOB.Create('PGSYNELTNAT', TobCodeLibre, -1);

  TC.PutValue('PEY_LIBELLE', Libelle);
  TC.PutValue('PEY_THEMEELT', ' ');
  TC.PutValue('PEY_MONETAIRE', 'X');
  TC.PutValue('PEY_ABREGE', ' ');
  TC.PutValue('PEY_BLOCNOTE', ' ');
  TC.PutValue('PEY_DECALMOIS', 'X');
  TC.PutValue('PEY_REGIMEALSACE', '-');
  TC.PutValue('PEY_CONVENTION', ' ');
  TC.PutValue('PEY_TYPUTI','3');
  TC.PutValue('PEY_TYPNIV', 'SAL');
  TC.PutValue('PEY_VALNIV', Salarie);
  TC.PutValue('PEY_LIBVALNIV', Salarie);
  TC.PutValue('PEY_ETABLISSEMENT', ' ');
  TC.PutValue('PEY_CODEPOP', ' ');
  TC.PutValue('PEY_SALARIE', Salarie);
  TC.PutValue('PEY_SALARIE', Salarie);
  TC.PutValue('PEY_CODEELT', CodeTable);

  sDate := DateTimeToStr(DatVal);
  if MidStr(sDate,1,2) <> '01' then
    sDate := '01' + MidStr(sDate,3,8);
  TC.PutValue('PEY_DATVAL', StrToDateTime(sDate));

  TC.PutValue('PEY_MONTANT', 0);
  TC.PutValue('PEY_MONTANTEURO', 0);
  TC.PutValue('PEY_PREDEFINI', Predefini);

{$IFNDEF EABSENCES}
  TC.PutValue('PEY_NODOSSIER', ClePGSynEltNAt); // Sert d'identifiant, contient un n° de salarié :
                                                // - soit celui en cours pour le traitement individuel
                                                // - soit le 1er de la liste si préparation automatique
{$ENDIF EABSENCES}

  TC.InsertOrUpdateDB;
  TobCodeLibre.Free;
end;
//FIN PT1

function TTablesDynamiques.GetValueFromDynaTableDSA(Salarie, CodeTable: String; Date : TDateTime ; SalEtab : String = ''; SalConv : String = '' ; DateReferenceAge : TDateTime = 0; DateReferenceAnciennete : TDateTime = 0; Entete : TOB = nil): Double;
var
  stQResultat: String;
  Donnee1, Donnee2 : String;
  TempTobDetail : TOB;
  LastVALCRIT1, LastVALCRIT2, VALCRIT1, VALCRIT2 : String;
  stTrav1, stTrav2 : string; //PT15
begin
  if DateReferenceAge < 10 then DateReferenceAge := Date;
  if DateReferenceAnciennete < 10 then DateReferenceAnciennete := Date;
  Result := 0;
  if Salarie <> LastSal then
    EvtOnChangeSal(Salarie)
  else
    if SearchResultOnSaveSalResult(Result) then exit; // Si on a déjà la valeur en mémoire, on répond tout de suite
  if Entete = nil then
    Entete := GetEnteteFromDynaTable(Salarie, CodeTable, Date , SalEtab, SalConv);
  if Assigned(Entete) then
  begin
    if Entete.GetInteger('PTE_CRITERE1') <> 0 then
      Donnee1 := GetDonneeSalarie(Salarie, Entete.GetInteger('PTE_CRITERE1'), Entete.GetString('PTE_TYPECRITERE1'), DateReferenceAge, DateReferenceAnciennete)
    else
      Donnee1 := '';
    if Entete.GetInteger('PTE_CRITERE2') <> 0 then
      Donnee2 := GetDonneeSalarie(Salarie, Entete.GetInteger('PTE_CRITERE2'), Entete.GetString('PTE_TYPECRITERE2'), DateReferenceAge, DateReferenceAnciennete)
    else
      Donnee2 := '';
    TempTobDetail :=  TobDetailsTable.FindFirst(['PTD_CODTABL', 'PTD_DTVALID', 'PTD_NIVSAIS', 'PTD_VALNIV', 'PTD_PREDEFINI'],
                                                [CodeTable, Entete.GetDateTime('PTE_DTVALID'), Entete.GetString('PTE_NIVSAIS'), Entete.GetString('PTE_VALNIV'), Entete.GetString('PTE_PREDEFINI')],
                                                False);
    if not Assigned(TempTobDetail) then
    begin
      stQResultat := 'SELECT * '
                   + ' FROM TABLEDIMDET '
                   + 'WHERE PTD_CODTABL = "'+CodeTable+'" '
                   + '  AND PTD_DTVALID = "'+USDATETIME(Entete.GetDateTime('PTE_DTVALID'))+'" '
                   + '  AND PTD_NIVSAIS = "'+Entete.GetString('PTE_NIVSAIS')+'" '
                   + '  AND PTD_VALNIV  = "'+Entete.GetString('PTE_VALNIV')+'" '
                   + '  AND PTD_PREDEFINI = "'+Entete.GetString('PTE_PREDEFINI')+'" ';
      TobDetailsTable.LoadDetailFromSQL(stQResultat,True);
      TempTobDetail :=  TobDetailsTable.FindFirst(['PTD_CODTABL', 'PTD_DTVALID', 'PTD_NIVSAIS', 'PTD_VALNIV', 'PTD_PREDEFINI'],
                                                  [CodeTable, Entete.GetDateTime('PTE_DTVALID'), Entete.GetString('PTE_NIVSAIS'), Entete.GetString('PTE_VALNIV'), Entete.GetString('PTE_PREDEFINI')],
                                                  False);
    end;
    LastVALCRIT1 := '';
    LastVALCRIT2 := '';
    While Assigned(TempTobDetail) do
    begin
      VALCRIT1 := TempTobDetail.GetString('PTD_VALCRIT1');
      VALCRIT2 := TempTobDetail.GetString('PTD_VALCRIT2');
      //DEB PT15
      stTrav1 := Donnee1;
      while Length(VALCRIT1) > Length(stTrav1) do
        stTrav1 := '0' + stTrav1;
      stTrav2 := Donnee2;
      while Length(VALCRIT2) > Length(stTrav2) do
        stTrav2 := '0' + stTrav2;
      //FIN PT15
      if (    TraiteSensTOB(VALCRIT1, stTrav1, Entete.GetString('PTE_SENS1'), LastVALCRIT1)  //PT15
          and TraiteSensTOB(VALCRIT2, stTrav2, Entete.GetString('PTE_SENS2'), LastVALCRIT2)) then  //PT15
      begin
        if Entete.GetString('PTE_TYPERESULTAT') = 'VAL' then
          result := TempTobDetail.GetValue('PTD_RESULTAT')
        else if Entete.GetString('PTE_TYPERESULTAT') = 'ELT' then begin
          result := ValEltNat(TempTobDetail.GetValue('PTD_RESULTAT'), Date);
        end else if Entete.GetString('PTE_TYPERESULTAT') = 'VAR' then begin
        { Les variables de paies ne sont pas traités. Si elles l'étaient, il faudrait
          appeler la fonction
        //  result := ValVariable(TempTobDetail.GetValue('PTD_RESULTAT'), Date, Date, TOB_Rub); }
        end;
        LastVALCRIT1 := TempTobDetail.GetString('PTD_VALCRIT1');
        LastVALCRIT2 := TempTobDetail.GetString('PTD_VALCRIT2');
      end;
      TempTobDetail :=  TobDetailsTable.FindNext (['PTD_CODTABL', 'PTD_DTVALID', 'PTD_NIVSAIS', 'PTD_VALNIV', 'PTD_PREDEFINI'],
                                                  [CodeTable, Entete.GetDateTime('PTE_DTVALID'), Entete.GetString('PTE_NIVSAIS'), Entete.GetString('PTE_VALNIV'), Entete.GetString('PTE_PREDEFINI')],
                                                  False);
    end;
    AddResultOnSaveSalResult(result);
  end;
end;
{$ENDIF}

function TTablesDynamiques.GetValueFromDynaTableCOD(Salarie, CodeTable: String;
  Date: TDate; Code : String; SalEtab : String = ''; SalConv : String = ''; GetLibelle : Boolean = True): String;
var
  stQResultat : String;
  LastVALCRIT, VALCRIT : String;
  Entete, TempTobDetail : TOB;
begin
//  if GetLibelle then Result := '' else result := '0';
  Result := '-123456';    //PT1
  Entete := GetEnteteFromDynaTable(Salarie, CodeTable, Date , SalEtab, SalConv);
  if length(Code) < NbrC_COD then Code := lpad(Code, NbrC_COD);
  if Assigned(Entete) then
  begin
    TempTobDetail :=  TobDetailsTable.FindFirst(['PTD_CODTABL', 'PTD_DTVALID', 'PTD_NIVSAIS', 'PTD_PREDEFINI'],
                                                [CodeTable, Entete.GetDateTime('PTE_DTVALID'), Entete.GetString('PTE_NIVSAIS'), Entete.GetString('PTE_PREDEFINI')],
                                                False);
    if not Assigned(TempTobDetail) then
    begin
      stQResultat := 'SELECT * '
                   + ' FROM TABLEDIMDET '
                   + 'WHERE PTD_CODTABL = "'+CodeTable+'" '
                   + '  AND PTD_DTVALID = "'+USDATETIME(Entete.GetDateTime('PTE_DTVALID'))+'" '
                   + '  AND PTD_NIVSAIS = "'+Entete.GetString('PTE_NIVSAIS')+'" '
                   + '  AND PTD_PREDEFINI = "'+Entete.GetString('PTE_PREDEFINI')+'" ';
      TobDetailsTable.LoadDetailFromSQL(stQResultat,true);
      TempTobDetail :=  TobDetailsTable.FindFirst(['PTD_CODTABL', 'PTD_DTVALID', 'PTD_NIVSAIS', 'PTD_PREDEFINI'],
                                                  [CodeTable, Entete.GetDateTime('PTE_DTVALID'), Entete.GetString('PTE_NIVSAIS'), Entete.GetString('PTE_PREDEFINI')],
                                                  False);
    end;
    LastVALCRIT := '';
    While Assigned(TempTobDetail) do
    begin
      VALCRIT := TempTobDetail.GetString('PTD_VALCRIT1');
      VALCRIT := lpad(VALCRIT, length(Code));
      if TraiteSensTOB(VALCRIT, Code, Entete.GetString('PTE_SENSINT'), LastVALCRIT) then
      begin
        if GetLibelle then
          Result := TempTobDetail.GetValue('PTD_LIBELLECODE')
        else
{$IFDEF TABLESDYNA}
        begin
          if Entete.GetString('PTE_TYPERESULTAT') = 'VAL' then
            result := TempTobDetail.GetValue('PTD_RESULTAT')
          else if Entete.GetString('PTE_TYPERESULTAT') = 'ELT' then begin
            result :=  FloatToStr(ValEltNat(TempTobDetail.GetValue('PTD_RESULTAT'), Date));
          end else if Entete.GetString('PTE_TYPERESULTAT') = 'VAR' then begin
          { Les variables de paies ne sont pas traités. Si elles l'étaient, il faudrait
            appeler la fonction
          //  result := ValVariable(TempTobDetail.GetValue('PTD_RESULTAT'), Date, Date, TOB_Rub); }
          end;
        end;
{$ELSE !TABLESDYNA}
         result := TempTobDetail.GetValue('PTD_RESULTAT');
{$ENDIF TABLESDYNA}
//Fin PT9
        LastVALCRIT := VALCRIT;//TempTobDetail.GetValue('PTD_VALCRIT1');
      end;
      TempTobDetail := TobDetailsTable.FindNext (['PTD_CODTABL', 'PTD_DTVALID', 'PTD_NIVSAIS', 'PTD_VALNIV', 'PTD_PREDEFINI'],
                                                 [CodeTable, Entete.GetDateTime('PTE_DTVALID'), Entete.GetString('PTE_NIVSAIS'), Entete.GetString('PTE_VALNIV'), Entete.GetString('PTE_PREDEFINI')],
                                                 False);
    end;
  end;
end;

{$IFDEF TABLESDYNA}
procedure TTablesDynamiques.EvtOnChangeSal(NewSal : String);
begin
  SetLength(SavLastSalResultCodeTable,0);
  SetLength(SavLastSalResultValue,0);
  LastSal := NewSal;
  LastSalEtab := '';
  LastSalConv := '';
end;

procedure TTablesDynamiques.AddResultOnSaveSalResult(Valeur: Double);
begin
  SetLength(SavLastSalResultCodeTable,Length(SavLastSalResultCodeTable)+1);
  SetLength(SavLastSalResultValue,Length(SavLastSalResultValue)+1);
  SavLastSalResultCodeTable[Length(SavLastSalResultCodeTable)-1] := LastCodeTable;
  SavLastSalResultValue[Length(SavLastSalResultValue)-1] := Valeur;
end;

function TTablesDynamiques.SearchResultOnSaveSalResult(var Valeur : Double) : Boolean;
var
  index : Integer;
begin
  result := False;
  for index := 0 to Length(SavLastSalResultCodeTable)-1 do
  begin
    if LastCodeTable = SavLastSalResultCodeTable[index] then
    begin
      Valeur := SavLastSalResultValue[index];
      result := True;
      break;
    end;
  end;
end;
{$ENDIF !TABLESDYNA}

end.
