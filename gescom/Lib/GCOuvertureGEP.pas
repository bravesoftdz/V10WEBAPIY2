unit GCOuvertureGEP;

interface

uses
  UTob,
  M3FP,
  Forms,
  Controls,
  hmsgbox,
  MenuSpec,
  TiersUtil,
  UtilArticle,
  CreePieceOle,
  {$IFDEF SAV}
    CreeParcOle,
  {$ENDIF SAV}
  {$IFDEF AFFAIRE}
//  AffaireGEP,
  {$ENDIF AFFAIRE}
  FactGrp,
//  FactPieceContainer,
  StockUtil,
  BTPUTIL,
  EntGC,
  factcomm,uEntCommun;

procedure OnAfterTransformePiece (Parms : array of Variant; nb: integer );

// Validation de pièce
(*
function OnBeforeValidePiece (LaForme : TForm; PieceContainer : TPieceContainer) : boolean; overload;
function OnBeforeValidePiece (LaForme : TForm; PieceContainer : TPieceContainer ; var TexteMessageGep : string) : boolean; overload;
function OnAfterValidePiece (LaForme : TForm; PieceContainer : TPieceContainer) : boolean;
{ Suppression d'une pièce }
function OnBeforeDeletePiece (LaForme : TForm; PieceContainer : TPieceContainer) : boolean;
//procedure OnAfterDeletePiece (LaForme : TForm; PieceContainer : TPieceContainer);
{ Ouverture d'une pièce }
procedure OnLoadPiece  (LaForme : TForm; PieceContainer : TPieceContainer);
{ Abandon en saisie de pièce }
procedure OnAfterCancelPiece (LaForme : TForm; PieceContainer : TPieceContainer);

// en sortie de ligne de pièce
function OnExitLigne (LaForme : TForm; PieceContainer : TPieceContainer; TobL : Tob) : boolean;
// GC_20080115_DBR_GC15712_DEBUT
// en sortie de ligne de pièce, donner la possibilité d'insérer d'autres lignes
procedure OnInsertLigne (LaForme : TForm; PieceContainer : TPieceContainer; TobL, TobInsert : Tob);
// GC_20080115_DBR_GC15712_FIN
{ After Calcul Prix de revient }
procedure PdR_AfterCalcul(Const iIdentifiant: integer);

{ Evenement pièce standard }
function OnEvenementPiece (ProcName : string ; LaForme : TForm; PieceContainer : TPieceContainer; TobL : Tob = nil) : boolean; overload;
function OnEvenementPiece (ProcName : string ; LaForme : TForm; PieceContainer : TPieceContainer; var sFocus : string ; TobL : Tob = nil) : boolean; overload;
function OnEvenementPiece (ProcName : string ; LaForme : TForm; PieceContainer : TPieceContainer; var dPrix, dRemise, dQte : double ; var sFocus : string ; TobL : Tob = nil) : boolean; overload;

function GCExitQuantite (LaForme : TForm; PieceContainer : TPieceContainer; var dPrix, dRemise : double ; TobL : Tob) : boolean;
function GCExitTiersPiece (LaForme : TForm; PieceContainer : TPieceContainer) : boolean;
function GCOuvertureTaxe (LaForme : TForm; PieceContainer : TPieceContainer) : boolean;
*)
{ Evénements tables }
procedure OnNewRecord(Const TableName: string; TobData: Tob);
procedure OnDeleteRecord(Const TableName: string; TobData: Tob);
procedure OnAfterDeleteRecord(Const TableName: string; TobData: Tob);
procedure OnUpdateRecord(Const TableName, Ikc: string; TobData: Tob);
procedure OnAfterUpdateRecord(Const TableName, Ikc: string; TobData: Tob);

//GP_20080429_TS_GP14877 >>>
{ Evénements métiers - STOCK }
{$IFDEF STK}
function G_GetDatePeremption(const Context: String; const DefaultValue: TDateTime): TDateTime;
function G_GetDateDispo(const Context: String; const DefaultValue: TDateTime): TDateTime;
//GP_20080429_TS_GP14877 <<<
//GP_20080430_TS_GP14873
function G_GetEvaluationGSF(const Context: String): String;
{$ENDIF STK}

{ Evénements Calcul CBN }
procedure CBN_OnBeforeCalcul(Const Origine, Ctx: string);
procedure CBN_OnAfterCalcul(Const Origine, Ctx: string);

{ Evénements validation proposition CBN }
{ Fabrication }
procedure CBN_BeforeGenereFabrication(Const Ctx, TypeProposition, Nature, User: string; TobWEV: Tob);
procedure CBN_AfterGenereFabrication(Const Ctx, TypeProposition, Nature, User: string; TobWEV: Tob);
procedure CBN_OnNewWOT(Const Ctx, TypeProposition, Nature, User: string; TobWEV, TobData: Tob);
procedure CBN_OnNewWOL(Const Ctx, TypeProposition, Nature, User: string; TobWEV, TobData: Tob);
procedure CBN_BeforeGenerePiece(Const Ctx, TypeProposition, Nature, User: string; TobWEV: Tob);
procedure CBN_AfterGenerePiece(Const Ctx, TypeProposition, Nature, User: string; TobWEV: Tob);
procedure CBN_OnNewPiece(Const Ctx, TypeProposition, Nature, User: string; TobWEV, TobData: Tob);
procedure CBN_OnNewLigne(Const Ctx, TypeProposition, Nature, User: string; TobWEV, TobData: Tob);

{ Evénements Calcul SCM }
procedure SCM_OnBeforeSession(Const Origine, Ctx: string);
procedure SCM_OnAfterSession(Const Origine, Ctx: string);
procedure SCM_OnBeforeCalculPropSCM(Const Origine, Ctx: string);
procedure SCM_OnAfterCalculPropSCM(Const Origine, Ctx: string);

{ Evénements Calcul Plan de charge }
procedure PDC_OnBeforeCalcul(Const Ctx: string);
procedure PDC_OnAfterCalcul(Const Ctx: string);

{ tiers 360° }
//function OnEvenementTiers360 (ProcName : string ; LaForme : TForm; CodeCBS, CodeTiers, CodeAuxiliaire, stArg, Valeur : string) : string;

implementation

uses
  {$IFNDEF EAGLCLIENT}
   uDbxDataSet,
  {$ENDIF !EAGLCLIENT}
  HCtrls, UtilPgi,
  hEnt1,
  wCommuns,
  wActionsTables,
//GP_20080429_TS_GP14877 >>>
  Facture,
  Variants
//GP_20080429_TS_GP14877 <<<
{ GC_20080923_PFA_010;15690}
   ,GerePiece
  ;

var TextMsg : string;

{ Tiers }
Function AGLCreateTiers (parms: array of variant; nb: integer ): variant;
var theTob : Tob;
TobR : Tob;
begin
    TheTob := Tob(LongInt(Parms[1]));
    if nb > 2 then
      TobR := Tob(LongInt(Parms [2]))
    else TobR := nil;
    
    if theTob = nil then Result := ''
    else Result:= CreateTiersFromTob (string(Parms[0]), TheTob) ;
    
    if (TheTob <> nil) and (TobR <> nil) then
      if TheTob.FieldExists ('Error') then
        TobR.AddChampSupValeur ('MESSAGE', TheTob.GetString ('Error'));
End;

{ Article }
Function AGLCreateArticle (parms: array of variant; nb: integer ): variant;
var theTob : Tob;
begin
    TheTob := Tob(LongInt(Parms[1]));
    if theTob = nil then Result := ''
    else Result:= wCreateArticleFromTob (string(Parms[0]), TheTob) ;
End;

{ Dépot }
Function AGLCreateDepot (parms: array of variant; nb: integer ): variant;
var
  TobGDE, theTob : Tob;
begin
  TheTob := Tob(LongInt(Parms[1]));
  TobGDE := Tob.Create('GDE', nil, -1);
  try
    if theTob = nil then
      Result := ''
    else
    begin
      TobRealToVirtual(TheTob, TobGDE);
      TobGDE.AddChampSupValeur('GDE_DEPOT', string(Parms[0]));
      Result := CreateTable('DEPOTS', TobGDE);
    end;
  finally
    TobGDE.free;
  end;
End;
{ Commercial }
Function AGLCreateCial (parms: array of variant; nb: integer ): variant;
var
  TobGCL, theTob : Tob;
begin
  theTob := Tob(LongInt(Parms[1]));
  TobGCL := Tob.Create('GCL', nil, -1);
  try
    if theTob = nil then
      Result := ''
    else
    begin
      TobRealToVirtual(theTob, TobGCL);
      TobGCL.AddChampSupValeur('GCL_COMMERCIAL', string(Parms[0]));
      Result := CreateTable('COMMERCIAL', TobGCL);
    end;
  finally
    TobGCL.Free;
  end;
End;

{ mouvement }
function AglCreateMouvement (Parms : array of Variant; nb : integer) : Variant;
begin
{$IFDEF STK}
  Result := CreationMouvement (String (Parms [0]), String (Parms [1]), String (Parms [2]), String (Parms [3]), Double (Parms [4]));
{$ENDIF STK}
end;

{ Pièces }
function AglCreatePiece (Parms : array of Variant; nb: integer ): variant;
begin
  if nb < 1 then Result := ''
  else Result := CreeLaPieceDepuisFichier (String(Parms [0]));
end;

{$IFDEF SAV}
{ Parc }
function AglCreateParc (Parms : array of Variant; nb: integer ): variant;
begin
  if nb < 1 then Result := ''
  else Result := CreeLeParcDepuisFichier (String(Parms [0]));
end;
{$ENDIF SAV}

function AglCreatePieceTob (Parms : array of Variant; nb: integer ): variant;
{$IFNDEF BTP}
var
{ GC_BBI_GC_Pied & Port dans création de pièce tob_DEBUT }
  TobD, TobR, TobP : Tob;
  stError : string;
{$ENDIF}
begin
{$IFNDEF BTP}
  TobD := Tob(LongInt(Parms [0]));
  TobR := Tob(LongInt(Parms [1]));
  if Assigned(Tob(LongInt(Parms [2]))) then
    TobP := Tob(LongInt(Parms [2]))
  else
    TobP := nil;

  stError := CreerLaPieceTob (TobD, TobR, TobP);
{ GC_BBI_GC_Pied & Port dans création de pièce tob_FIN }

  if stError <> '' then
  begin
    if not TobR.FieldExists ('MESSAGE') then TobR.AddChampSupValeur ('MESSAGE', stError)
    else TobR.SetString ('MESSAGE', stError);
  end;
{$ENDIF}
end;

{ GC_20080916_PFA_010;15690}
function AglSupprimePiece (Parms : array of Variant; nb: integer ): variant;
var
{$IFNDEF BTP}
  TobR : Tob;
//  stError : string;
{$ENDIF}
  StA : String;
	cledoc : r_cledoc;
begin
  StA := String(Parms [0]);
{$IFNDEF BTP}
  TobR := Tob(LongInt(Parms [1]));
  if not Assigned(TobR) then
    TobR := Tob.Create('TobR',nil,-1);
  Result := SupprimePiecesFromClesDoc(StA,TobR);
{$ELSE}
 	DecodeRefPiece (Sta,cledoc);
	BTPSupprimePiece (cledoc);
{$ENDIF}
end;

function AglTransformePiece (Parms : array of Variant; nb: integer ): variant;
begin
  if nb < 6 then Result := ''
  else if nb=7 then
  begin
    Result := TransfoBatchPiece (String(Parms [0]), String(Parms [1]), String(Parms [2]), Integer(Parms [3]), Integer(Parms [4]),
                                 TDateTime(Parms [5]));//, String(Parms [6]));
  end else
  begin
    Result := TransfoBatchPiece (String(Parms [0]), String(Parms [1]), String(Parms [2]), Integer(Parms [3]), Integer(Parms [4]),
                                 TDateTime(Parms [5]));
  end;
end;

function AglEncoursTiers (Parms : array of Variant; nb : integer) : variant;
var
  sType, Societe, Tiers : String;
  TobTiers : Tob;
  QTiers : TQuery;

  function GetBase (UneBase, LaTable : string) : String;
  begin
    if isMssql then result := UneBase + '.DBO.' + LaTable
    else result := UneBase + '.' + LaTable;
  end;

begin
  Result := -1;
  if nb < 2 then exit;
  Tiers := String(Parms [0]);
  Societe := String (Parms [1]);
  if nb > 2 then
    sType := String (Parms [2])
  else sType := '';

  if Tiers = '' then exit;

  Result := 0;

  TobTiers := Tob.Create ('', nil, -1);
  QTiers := OpenSql ('SELECT T_AUXILIAIRE, T_TIERS, T_TOTALDEBIT, T_TOTALCREDIT FROM ' + GetBase (Societe, 'TIERS') + ' WHERE T_TIERS="' + Tiers + '"',  true);
  try
    if not QTiers.Eof then
    begin
      TobTiers.SelectDB ('', QTiers);

      if (sType = '') or (sType = 'COM') then
        Result := RisqueTiersGC(TOBTiers, Societe);
      if (sType = '') or (sType = 'CPT') then
        Result := Result + RisqueTiersCPTA(TOBTiers, V_PGI.DateEntree, Societe);
    end else
      Result := 'Le Tiers n''existe pas';
  finally
    Ferme (QTiers);
    TobTiers.Free;
  end;

end;

{ traitement sur pièce }
procedure OnAfterTransformePiece (Parms : array of Variant; nb: integer );
begin
  vmDisp_CallXX ('GC_AfterTransformePiece', -1);
end;

{ Saisie de pièce }
function OnMultiSelectionArticle (Parms : array of Variant; nb: integer ) : Variant;
begin
  vmDisp_CallXX ('GC_MultiSelectionArticle', -1);
end;

{ A faire dans la pièce :
  on initialise la ligne par le source métier et on donne la main
  (sortie de référence article, sortie de ligne, sortie du complément de ligne)
  on récupère les valeurs des champs autorisés
  on recalcul la pièce
  ex : saisie article donne la main. calcul d'une remise spécifique -> calcul de la pièce }

function OnInitialisationQte (Parms : array of Variant; nb: integer ) : Variant;
begin
  Result := 1;
  vmDisp_CallXX ('GC_InitialisationQte', -1);
end;

{ doit relancer en sortie le recalcul de la pièce après avoir récupérer les nouvelles valeurs des champs autorisés }
function OnExitReferenceArticle (Parms : array of Variant; nb: integer ) : variant;
begin
end;

// en sortie de ligne de pièce
(*
function OnExitLigne (LaForme : TForm; PieceContainer : TPieceContainer; TobL : Tob) : boolean;
var
  TobData, TobResultat, TobTemp : TOB;
begin
  Result := true;

  TobData := Tob.Create ('', nil, -1);
  TobResultat := Tob.Create ('', nil, -1);
  try
    TobTemp := Tob.Create ('', TobData, -1);
    TobTemp.Dupliquer(PieceContainer.TCPiece, false, true);
    TobTemp := Tob.Create ('', TobData, -1);
    TobTemp.Dupliquer (TobL, false, true);

    TobResultat.AddChampSupValeur ('VALIDE', 0);
    TobResultat.AddChampSupValeur ('MESSAGE', '');

    vmDisp_CallProc ([LongInt(LaForme), 'GC_ExitLigne', LongInt(TobData), LongInt(TobResultat)], 4);

    if TobResultat.GetInteger ('VALIDE') = 1 then
    begin
      if TobResultat.GetString ('MESSAGE') <> '' then
        if PgiAsk (TobResultat.GetString ('MESSAGE')) = mrno then Result := false;
    end else if TobResultat.GetInteger ('VALIDE') = 2 then
    begin
      PgiBox (TobResultat.GetString ('MESSAGE'));
      Result := false;
    end;
  finally
    TobData.Free;
    TobResultat.Free;
  end;
end;

function OnEvenementPiece (ProcName : string ; LaForme : TForm; PieceContainer : TPieceContainer; TobL : Tob = nil) : boolean; overload;
var dPx, dRem, dQte : double;
    sFocus : string;
begin
  result := OnEvenementPiece(ProcName, LaForme, PieceContainer, dPx, dRem, dQte, sFocus, TobL);
end;

function OnEvenementPiece (ProcName : string ; LaForme : TForm; PieceContainer : TPieceContainer ; var sFocus : string ; TobL : Tob = nil) : boolean; overload;
var dPx, dRem, dQte : double;
begin
  result := OnEvenementPiece(ProcName, LaForme, PieceContainer, dPx, dRem, dQte, sFocus, TobL);
end;

function OnEvenementPiece (ProcName : string ; LaForme : TForm; PieceContainer : TPieceContainer; var dPrix, dRemise, dQte : double ; var sFocus : string ; TobL : Tob = nil) : boolean;
var
  TobData, TobResultat, TobTemp, TobTempL : TOB;
begin
  Result := true;

  { GC_DBR_20081015_FQ;010;16683 }
  if PieceContainer.LeCodeTiers <> '' then exit;

  TobData := Tob.Create ('', nil, -1);
  try
    TobResultat := Tob.Create ('', nil, -1);
    try
      TobTemp := Tob.Create ('', TobData, -1);
      TobTemp.Dupliquer (PieceContainer.TCPiece, false, true);
      if assigned (TobL) then
      begin
        TobTempL := Tob.Create ('', TobData, -1);
        TobTempL.Dupliquer (TobL, false, true);
      end;

      TobResultat.AddChampSupValeur ('VALIDE', 0);
      TobResultat.AddChampSupValeur ('MESSAGE', '');
      TobResultat.AddChampSupValeur ('FOCUS', '', false);
      TobResultat.AddChampSupValeur ('PRIX', dPrix, false);
      TobResultat.AddChampSupValeur ('REMISE', dRemise, false);
      TobResultat.AddChampSupValeur ('QUANTITE', dQte, false);

      vmDisp_CallProc ([LongInt(LaForme), ProcName, LongInt(TobData), LongInt(TobResultat)], 4);

      case TobResultat.GetInteger ('VALIDE') of
        1 : begin
              result := true;
              if ProcName = 'GC_SetInfoLigne' then
              begin
                dPrix   := TobResultat.GetDouble ('PRIX');
                dRemise := TobResultat.GetDouble ('REMISE');
                dQte    := TobResultat.GetDouble ('QUANTITE');
              end else
              begin
                if TobResultat.GetString ('MESSAGE') <> '' then
                  if PgiAsk (TobResultat.GetString ('MESSAGE')) = mrno then
                  begin
                    sFocus := TobResultat.GetString('FOCUS');
                    result := false;
                  end;
              end;
            end;
        2 : begin
              if ProcName = 'GC_SetInfoLigne' then
              else
              begin
                PgiBox (TobResultat.GetString ('MESSAGE'));
                sFocus := TobResultat.GetString('FOCUS');
                result := false;
              end;
            end;
      end;
    finally
      TobResultat.Free;
    end;
  finally
    TobData.Free;
  end;
end;

function GCOuvertureTaxe (LaForme : TForm; PieceContainer : TPieceContainer) : boolean;
var
  TobData, TobResultat, TobTemp : TOB;
begin
  TobData := Tob.Create ('', nil, -1);
  try
    TobResultat := Tob.Create ('', nil, -1);
    try
      TobTemp := Tob.Create ('', TobData, -1);
      TobTemp.Dupliquer (PieceContainer.TCPiece, false, true);

      TobResultat.AddChampSupValeur ('MESSAGE', '');
      TobResultat.AddChampSupValeur ('OUVERTURETAXE', false, false);

      vmDisp_CallProc ([LongInt(LaForme), 'GC_OUVERTURETAXE', LongInt(TobData), LongInt(TobResultat)], 4);

      if TobResultat.GetString ('MESSAGE') <> '' then
        PgiBox (TobResultat.GetString ('MESSAGE'));

      result := TobResultat.GetBoolean('OUVERTURETAXE');
      
    finally
      TobResultat.Free;
    end;
  finally
    TobData.Free;
  end;
end;

function GCExitTiersPiece (LaForme : TForm; PieceContainer : TPieceContainer) : boolean;
var
  TobData, TobResultat, TobTemp : TOB;
begin
  Result := true;
  TobData := Tob.Create ('', nil, -1);
  try
    TobResultat := Tob.Create ('', nil, -1);
    try
      TobTemp := Tob.Create ('', TobData, -1);
      TobTemp.Dupliquer (PieceContainer.TCPiece, false, true);

      TobResultat.AddChampSupValeur ('VALIDE', 0);
      TobResultat.AddChampSupValeur ('MESSAGE', '');
      TobResultat.AddChampSupValeur ('OUVERTURECPLTPIECE', false, false);

      vmDisp_CallProc ([LongInt(LaForme), 'GC_ExitTiersPiece', LongInt(TobData), LongInt(TobResultat)], 4);

      case TobResultat.GetInteger ('VALIDE') of
        1 : begin
              result := true;
              if TobResultat.GetString ('MESSAGE') <> '' then
                if PgiAsk (TobResultat.GetString ('MESSAGE')) = mrno then
                  result := false;
            end;
        2 : begin
              PgiBox (TobResultat.GetString ('MESSAGE'));
              result := false;
            end;
      end;

      if TobResultat.GetBoolean ('OUVERTURECPLTPIECE') then
        TFFacture(LaForme).CpltEnteteClick(nil);
    finally
      TobResultat.Free;
    end;
  finally
    TobData.Free;
  end;
end;

function GCExitQuantite (LaForme : TForm; PieceContainer : TPieceContainer; var dPrix, dRemise : double ; TobL : Tob) : boolean;
var
  TobData, TobResultat, TobTemp, TobTempL : TOB;
begin
  Result := true;

  TobData := Tob.Create ('', nil, -1);
  try
    TobResultat := Tob.Create ('', nil, -1);
    try
      TobTemp := Tob.Create ('', TobData, -1);
      TobTemp.Dupliquer (PieceContainer.TCPiece, false, true);
      if assigned (TobL) then
      begin
        TobTempL := Tob.Create ('', TobData, -1);
        TobTempL.Dupliquer (TobL, false, true);
      end;

      TobResultat.AddChampSupValeur ('MESSAGE', '');
      TobResultat.AddChampSupValeur ('PRIX', dPrix, false);
      TobResultat.AddChampSupValeur ('REMISE', dRemise, false);

      vmDisp_CallProc ([LongInt(LaForme), 'GC_ExitQuantite', LongInt(TobData), LongInt(TobResultat)], 4);

      if TobResultat.GetString ('MESSAGE') <> '' then
        PgiBox (TobResultat.GetString ('MESSAGE'));

      dPrix   := TobResultat.GetDouble ('PRIX');
      dRemise := TobResultat.GetDouble ('REMISE');

    finally
      TobResultat.Free;
    end;
  finally
    TobData.Free;
  end;
end;
*)
{-------------------------------------------------------------------------------
  Suite au calcul de prix de revient
  Point d'entrée pour lancer une consolidation par exemple
--------------------------------------------------------------------------------}
procedure PdR_AfterCalcul(Const iIdentifiant: integer);
begin {OnAfterCalculPdr}
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'W_AFTERCALCULPDR', iIDentifiant], 3);
end; {OnAfterCalculPdr}

{ Avant calcul CBN }
procedure CBN_OnBeforeCalcul(Const Origine, Ctx: string);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'CBN_OnBeforeCalcul', Origine, Ctx], 4);
end;

{ Après calcul CBN }
procedure CBN_OnAfterCalcul(Const Origine, Ctx: string);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'CBN_OnAfterCalcul', Origine, Ctx], 4);
end;

{ Avant création des propositions d'ordre de fabrication issues CBN }
procedure CBN_BeforeGenereFabrication(Const Ctx, TypeProposition, Nature, User: string; TobWEV: Tob);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'CBN_BeforeGenereFabrication', Ctx, TypeProposition, Nature, User, Integer(TobWEV)], 7);
end;

{ Après création des propositions d'ordre de fabrication issues CBN }
procedure CBN_AfterGenereFabrication(Const Ctx, TypeProposition, Nature, User: string; TobWEV: Tob);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'CBN_AfterGenereFabrication', Ctx, TypeProposition, Nature, User, Integer(TobWEV)], 7);
end;

{ Création WOT depuis validation proposition d'ordre CBN }
procedure CBN_OnNewWOT(Const Ctx, TypeProposition, Nature, User: string; TobWEV, TobData: Tob);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'CBN_OnNewWOT', Ctx, TypeProposition, Nature, User, Integer(TobWEV), Integer(TobData)], 8);
end;

{ Création WOL depuis validation proposition d'ordre CBN }
procedure CBN_OnNewWOL(Const Ctx, TypeProposition, Nature, User: string; TobWEV, TobData: Tob);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'CBN_OnNewWOL', Ctx, TypeProposition, Nature, User, Integer(TobWEV), Integer(TobData)], 8);
end;

{ Avant création pièces Achat, Transfert }
procedure CBN_BeforeGenerePiece(Const Ctx, TypeProposition, Nature, User: string; TobWEV: Tob);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'CBN_BeforeGenerePiece', Ctx, TypeProposition, Nature, User, Integer(TobWEV)], 7);
end;

{ Après création pièces Achat, Transfert }
procedure CBN_AfterGenerePiece(Const Ctx, TypeProposition, Nature, User: string; TobWEV: Tob);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'CBN_AfterGenerePiece', Ctx, TypeProposition, Nature, User, Integer(TobWEV)], 7);
end;

{ Création pièce propositions achat, sous-traitance d'achat, transfert issues du CBN }
procedure CBN_OnNewPiece(Const Ctx, TypeProposition, Nature, User: string; TobWEV, TobData: Tob);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'CBN_OnNewPiece', Ctx, TypeProposition, Nature, User, Integer(TobWEV), Integer(TobData)], 8);
end;

{ Création ligne propositions achat, sous-traitance d'achat, transfert issues du CBN }
procedure CBN_OnNewLigne(Const Ctx, TypeProposition, Nature, User: string; TobWEV, TobData: Tob);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'CBN_OnNewLigne', Ctx, TypeProposition, Nature, User, Integer(TobWEV), Integer(TobData)], 8);
end;

{ Avant ouverture session SCM }
procedure SCM_OnBeforeSession(Const Origine, Ctx: string);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'SCM_OnBeforeSession', Origine, Ctx], 4);
end;

{ Après ouverture session SCM }
procedure SCM_OnAfterSession(Const Origine, Ctx: string);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'SCM_OnAfterSession', Origine, Ctx], 4);
end;

{ Avant calcul CBN dernier niveau suite à la validation d'un planning SCM}
procedure SCM_OnBeforeCalculPropSCM(Const Origine, Ctx: string);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'SCM_OnBeforeCalculPropSCM', Origine, Ctx], 4);
end;

{ Après calcul CBN dernier niveau suite à la validation d'un planning SCM}
procedure SCM_OnAfterCalculPropSCM(Const Origine, Ctx: string);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'SCM_OnAfterCalculPropSCM', Origine, Ctx], 4);
end;


{ Avant calcul Plan de charge }
procedure PDC_OnBeforeCalcul(Const Ctx: string);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'PDC_OnBeforeCalcul', Ctx], 3);
end;

{ Après calcul Plan de charge }
procedure PDC_OnAfterCalcul(Const Ctx: string);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'PDC_OnAfterCalcul', Ctx], 3);
end;

procedure OnNewRecord(Const TableName: string; TobData: Tob);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'AGLOnNewRecord', TableName, Integer(TobData)], 4);
end;

procedure OnDeleteRecord(Const TableName: string; TobData: Tob);
var
  TobError: Tob;
begin
  if V_Pgi.CGEPDisabled then exit;

  TobError := Tob.Create('', nil, -1);
  try
    TobError.AddChampSupValeur('Error', '', false);
    vmDisp_CallProc([integer(nil), 'AGLOnDeleteRecord', TableName, Integer(TobData), Integer(TobError)], 5);

    if TobError.GetString('Error') <> '' then
      TobData.AddChampSupvaleur('Error',TobError.GetString('Error'));
  finally
    TobError.free;
  end;
end;

procedure OnAfterDeleteRecord(Const TableName: string; TobData: Tob);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'AGLOnAfterDeleteRecord', TableName, Integer(TobData)], 4);
end;

procedure OnUpdateRecord(Const TableName, Ikc: string; TobData: Tob);
var
  TobError: Tob;
begin
  if V_Pgi.CGEPDisabled then exit;

  TobError := Tob.Create('', nil, -1);
  try
    TobError.AddChampSupValeur('Error', '', false);
    vmDisp_CallProc([integer(nil), 'AGLOnUpdateRecord', TableName, Ikc, Integer(TobData), Integer(TobError)], 6);

    if TobError.GetString('Error') <> '' then
      TobData.AddChampSupvaleur('Error',TobError.GetString('Error'));
  finally
    TobError.free;
  end;
end;

procedure OnAfterUpdateRecord(Const TableName, Ikc: string; TobData: Tob);
begin
  if V_Pgi.CGEPDisabled then exit;

  vmDisp_CallProc([integer(nil), 'AGLOnAfterUpdateRecord', TableName, Ikc, Integer(TobData)], 5);
end;
(*
function OnBeforeValidePiece (LaForme : TForm; PieceContainer : TPieceContainer) : boolean;
var
  TobData, TobResultat, TobTemp : TOB;
begin
  Result := true;
  TextMsg := '';
  TobData := Tob.Create ('', nil, -1);
  try
    TobTemp := Tob.Create ('', TobData, -1);
    TobTemp.Dupliquer (PieceContainer.TCPiece, true, true);
    TobTemp := Tob.Create ('', TobData, -1);
    TobTemp.Dupliquer (PieceContainer.TCArticles, true, true);

    TobResultat := Tob.Create ('', nil, -1);
    try
      TobResultat.AddChampSupValeur ('VALIDE', 0);
      TobResultat.AddChampSupValeur ('MESSAGE', '');

      vmDisp_CallProc ([LongInt(LaForme), 'GC_BeforeValidePiece', LongInt(TobData), LongInt(TobResultat)], 4);

      if TobResultat.GetInteger ('VALIDE') = 1 then
      begin
        if (TobResultat.GetString ('MESSAGE') <> '') and PieceContainer.WithInterface then
          if PgiAsk (TobResultat.GetString ('MESSAGE')) = mrno then Result := false;
      end else if TobResultat.GetInteger ('VALIDE') = 2 then
      begin
        if PieceContainer.WithInterface then PgiBox (TobResultat.GetString ('MESSAGE'));
        TextMsg := TobResultat.GetString ('MESSAGE');
        PieceContainer.TCPiece.AddChampSupValeur ('ErrorGEP', TextMsg);
        Result := false;
      end;

      if result then
      begin
        {$IFNDEF ERADIO}
        { GC_FQ 12257_JSI_DEB }
        if assigned(LaForme) then
        begin
          if GCOuvertureTaxe (LaForme, PieceContainer) then
            TFFacture(LaForme).ChangeRegime;
        end;
        { GC_FQ 12257_JSI_FIN }
        {$ENDIF !ERADIO}
      end;
    finally
      TobResultat.Free;
    end;
  finally
    TobData.Free;
  end;
end;

function OnBeforeValidePiece (LaForme : TForm; PieceContainer : TPieceContainer ; var TexteMessageGep : string) : boolean;
begin
end;

function OnAfterValidePiece (LaForme : TForm; PieceContainer : TPieceContainer) : boolean;
var
  TobData, TobResultat, TobTemp : TOB;
begin
  Result := true;

  TobData := Tob.Create ('', nil, -1);
  TobResultat := Tob.Create ('', nil, -1);
  try
    TobTemp := Tob.Create ('', TobData, -1);
    TobTemp.Dupliquer (PieceContainer.TCPiece, true, true);

    vmDisp_CallProc ([LongInt(LaForme), 'GC_AfterValidePiece', LongInt(TobData), LongInt(TobResultat)], 4);
  finally
    TobResultat.Free;
    TobData.Free;
  end;
end;

function OnBeforeDeletePiece (LaForme : TForm; PieceContainer : TPieceContainer) : boolean;
var
  TobData, TobResultat, TobPiece, TobArticle : tob;
begin
  Result := true;
  TextMsg := '';
  TobData := Tob.Create ('', nil, -1);
  try
    TobPiece := Tob.Create ('', TobData, -1);
    TobPiece.Dupliquer (PieceContainer.TCPiece, true, true);
    TobArticle := Tob.Create ('', TobData, -1);
    TobArticle.Dupliquer (PieceContainer.TCArticles, true, true);

    TobResultat := Tob.Create ('', nil, -1);
    try
      TobResultat.AddChampSupValeur ('VALIDE', 0);
      TobResultat.AddChampSupValeur ('MESSAGE', '');

      vmDisp_CallProc ([LongInt(LaForme), 'GC_BeforeDeletePiece', LongInt(TobData), LongInt(TobResultat)], 4);

      if TobResultat.GetInteger ('VALIDE') = 1 then
      begin
        if TobResultat.GetString ('MESSAGE') <> '' then
          if PgiAsk (TobResultat.GetString ('MESSAGE')) = mrno then Result := false;
      end else if TobResultat.GetInteger ('VALIDE') = 2 then
      begin
        if PieceContainer.WithInterface then PgiBox (TobResultat.GetString ('MESSAGE'));
        TextMsg := TobResultat.GetString ('MESSAGE');
        Result := false;
      end;
    finally
      TobResultat.Free;
    end;
  finally
    TobData.Free;
  end;
end;

procedure OnAfterDeletePiece (LaForme : TForm; PieceContainer : TPieceContainer);
var
  TobData, TobPiece : tob;
begin
  TobData := Tob.Create ('', nil, -1);
  try
    TobPiece := Tob.Create ('', TobData, -1);
    TobPiece.Dupliquer (PieceContainer.TCPiece, true, true);

    vmDisp_CallProc ([LongInt(LaForme), 'GC_AfterDeletePiece', LongInt(TobData)], 3);
  finally
    TobData.Free;
  end;
end;

procedure OnLoadPiece (LaForme : TForm; PieceContainer : TPieceContainer);
var
  TobData, TobAction : tob;
begin
  TobData := Tob.Create ('', nil, -1);
  try
    TobAction := Tob.Create ('', TobData, -1);
    TobAction.AddChampSupValeur ('ACTION', PieceContainer.ActionCurr);

    vmDisp_CallProc ([LongInt(LaForme), 'GC_LoadPiece', LongInt(TobData)], 3);
  finally
    TobData.Free;
  end;
end;

procedure OnAfterCancelPiece (LaForme : TForm; PieceContainer : TPieceContainer);
begin
  vmDisp_CallProc ([LongInt(LaForme), 'GC_AfterCancelPiece'], 2);
end;

function OnEvenementTiers360 (ProcName : string ; LaForme : TForm; CodeCBS, CodeTiers, CodeAuxiliaire, stArg, Valeur : string) : string;
begin
// result:=OnEvenementTiers360(ecran,'GC_RecupSQL360',CodeCBS, CodeTiers, CodeAuxiliaire);
  Result := vmDisp_CallFonc ([LongInt(LaForme), ProcName, CodeCBS,CodeTiers, CodeAuxiliaire,stArg, Valeur], 7);
end;

// GC_20080115_DBR_GC15712_DEBUT
// en sortie de ligne de pièce, donner la possibilité d'insérer d'autres lignes
procedure OnInsertLigne (LaForme : TForm; PieceContainer : TPieceContainer; TobL, TobInsert : Tob);
var
  TobData, TobTemp, TobRecup : TOB;
begin
  TobData := Tob.Create ('', nil, -1);
  try
    TobTemp := Tob.Create ('', TobData, -1);
    TobTemp.Dupliquer(PieceContainer.TCPiece, false, true);
    TobTemp := Tob.Create ('', TobData, -1);
    TobTemp.Dupliquer (TobL, false, true);
// GC_20080115_JPL_GC15712_DEBUT
    if assigned (PieceContainer.TCLigFormule) then
    begin
      TobRecup := PieceContainer.TCLigFormule.FindFirst(['GLF_INDICEG','GLF_NUMLIGNE'],[TobL.GetInteger('GL_INDICEG'),TobL.GetInteger('GL_NUMLIGNE')],True);
      if TobRecup <> Nil then
      begin
        TobTemp := Tob.Create ('', TobData, -1);
        TobTemp.Dupliquer (TobRecup, true, true);
        // GC_20080516_JPL_GC15712: on ne doit pas libérer la tob origine de LIGNEFORMULE
        //TobRecup.Free;
      end;
    end;
// GC_20080115_JPL_GC15712_FIN

    vmDisp_CallProc ([LongInt(LaForme), 'G_InsertLigne', LongInt(TobData), LongInt(TobInsert)], 4);
  finally
    TobData.Free;
  end;
end;
// GC_20080115_DBR_GC15712_FIN

//GP_20080429_TS_GP14877 >>>
{ Evénements métiers - STOCK --------------------------------------------------------------------- }

{ appel d'une fonction inexistante afin de récupérer la valeur par défaut
  si rien n'a été fait dans l'événement ou si l'événement n'existe pas }
function fxDefaultValue: Variant;
begin
  Result := vmDisp_CallFonc([LongInt(nil), 'G_xxx', ''], 1)
end;
*)
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 28/04/2008
Modifié le ... :   /  /
Description .. : Evénement CBS : permet d'avoir un point d'entrée
Suite ........ : CBS afin de calculer une date de péremption d'un
Suite ........ : article en entrée de stock.
Paramètres ... : Context : STRING :
Suite ........ : ce contexte contient des données tokenisées :
Suite ........ : ARTICLE="l'article entrée en stock"
Suite ........ : DATEENTREELOT="la date d'entrée du lot"
Suite ........ : DUREEVIE="La durée de vie de l'article en nombre de jours"(> WPF)
Suite ........ : PREFIXEORI="Origine du mouvement : Préfixe de la table d'origine"
Suite ........ : NATUREORI="Nature de l'origine"
Suite ........ : SOUCHEORI="Souche d'origine"
Suite ........ : NUMEROORI="Numéro d'origine"
Suite ........ : INDICEORI="Indice d'origine"
Suite ........ : OPECIRCORI="Phase d'origine"
Suite ........ : NUMLIGNEORI="Ligne d'origine"
*****************************************************************}
{$IFDEF STK}
function G_GetDatePeremption(const Context: String; const DefaultValue: TDateTime): TDateTime;
begin
  { appel de la fonction métier }
  Result := VarAsType(vmDisp_CallFonc([LongInt(nil), 'G_GetDatePeremption', Context], 1), varDate);

  { Valeur par défaut }
  if (Result = VarAsType(fxDefaultValue(), varDate)) or (Result = iDate1900) then
    Result := DefaultValue
end;
{$ENDIF STK}

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 28/04/2008
Modifié le ... :   /  /
Description .. : Evénement CBS : permet d'avoir un point d'entrée
Suite ........ : CBS afin de calculer une date de mise à dispo d'un
Suite ........ : article en entrée de stock.
Paramètres ... : Context : STRING :
Suite ........ : ce contexte contient des données tokenisées :
Suite ........ : ARTICLE="l'article entrée en stock"
Suite ........ : DATEENTREELOT="la date d'entrée du lot"
Suite ........ : DELGARDE="La durée de vie de l'article en nombre de jours"(> WPF)
Suite ........ : PREFIXEORI="Origine du mouvement : Préfixe de la table d'origine"
Suite ........ : NATUREORI="Nature de l'origine"
Suite ........ : SOUCHEORI="Souche d'origine"
Suite ........ : NUMEROORI="Numéro d'origine"
Suite ........ : INDICEORI="Indice d'origine"
Suite ........ : OPECIRCORI="Phase d'origine"
Suite ........ : NUMLIGNEORI="Ligne d'origine"
*****************************************************************}
{$IFDEF STK}
function G_GetDateDispo(const Context: String; const DefaultValue: TDateTime): TDateTime;
begin
  { appel de la fonction métier }
  Result := VarAsType(vmDisp_CallFonc([LongInt(nil), 'G_GetDateDispo', Context], 1), varDate);

  { Valeur par défaut }
  if (Result = VarAsType(fxDefaultValue(), varDate)) or (Result = iDate1900) then
    Result := DefaultValue
end;
{$ENDIF STK}
//GP_20080429_TS_GP14877 <<<

//GP_20080430_TS_GP14873 >>>
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut Sublet
Créé le ...... : 29/04/2008
Modifié le ... :   /  /
Description .. : Evénement CBS : permet d'avoir un point d'entrée CBS afin
Suite ........ : de calculer un élément d'une formule de lot/série
Paramètres ... : Context : STRING :
Suite ........ : ce contexte contient des données tokenisées
Suite ........ :
*****************************************************************}
{$IFDEF STK}
function G_GetEvaluationGSF(const Context: String): String;
begin
  { appel de la fonction métier }
  Result := VarToStr(vmDisp_CallFonc([LongInt(nil), 'G_GetEvaluationGSF', Context], 1));

  { Valeur par défaut }
  if (Result = VarToStr(fxDefaultValue())) and (VarType(fxDefaultValue()) <> VarType(Result)) then
    Result := ''
end;
{$ENDIF STK}
//GP_20080430_TS_GP14873 <<<

procedure InitAglFunc;
begin
  RegisterAglFunc ('GC_CreateTiers', False, 3, AglCreateTiers);
  RegisterAglFunc ('GC_CreateArticle', False, 2, AglCreateArticle);
  RegisterAglFunc ('GC_CreatePiece', False , 1, AGLCreatePiece);
  RegisterAglFunc ('GC_CreateDepot', False , 1, AGLCreateDepot);
  RegisterAglFunc ('GC_CreateCial', False , 1, AGLCreateCial);
  {$IFDEF SAV}
    RegisterAglFunc ('GC_CreateParc', False , 1, AGLCreateParc);
  {$ENDIF SAV}
  { GC_BBI_GC_Pied & Port dans création de pièce tob }
  RegisterAglFunc ('GC_CreatePieceTob', False , 3, AGLCreatePieceTob);
  RegisterAglFunc ('GC_CreateMouvement', false, 5, AglCreateMouvement);
  RegisterAglFunc ('GC_TransformePiece', false, 7, AglTransformePiece);
  {$IFDEF AFFAIRE}
//    RegisterAglFunc ('G_CreateAffaire', false, 5, AGLCreeAffaire);
  {$ENDIF AFFAIRE}
  RegisterAglFunc ('G_EncoursTiers', false, 1, AglEncoursTiers);
  { GC_20080916_PFA_010;15690}
  RegisterAglFunc ('GC_SupprimePiece', False , 3, AGLSupprimePiece);
end;

Initialization
  InitAglFunc();
end.
