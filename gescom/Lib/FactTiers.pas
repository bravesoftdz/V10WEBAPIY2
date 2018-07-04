unit FactTiers ;

interface

uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, UTOB, HCtrls, LookUp, SysUtils, FactUtil,
{$IFDEF EAGLCLIENT}
     Utileagl, MaineAgl,
{$ELSE}
     Fe_Main, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     AGLInitGC, EntGC, TiersUtil, UtilPGI,Paramsoc ;

type T_RechTiers = (trtOk, trtIncorrect, trtFerme);

function  GetTiers(TH: THCritMaskEdit): boolean;
procedure TiersVersPiece(TOBTiers, TOBPiece: TOB);
procedure ChangeTzTiersSaisie(NaturePieceG: String3);
function  RemplirTOBTiers(TOBTiers: TOB; CodeTiers, NaturePiece: string; ControleFerme: boolean): T_RechTiers;
function ReferenceIntExtIsOk(From: T_CtrlRefIntExt; TobPiece: TOB; Reference: string): Boolean; 


implementation

uses
  HMsgBox,FactTOB
  ;

Function GetTiersMul_FO ( TH : THCritMaskEdit ) : boolean ;
Var St, sTiers, sWhere : String ;
    Ind : integer ;
    Ok  : boolean ;
BEGIN
// Constitution de la clause where pour la recherche des tiers
Result := False ;
sWhere := '' ;
St := 'GCTIERSSAISIE' ;
Ind := TTToNum(St) ;
if Ind > 0 then sWhere := V_PGI.DECombos[Ind].Where ;
{$IFDEF EAGLCLIENT}
Ok := False ;
{$ELSE}
Ok := (Pos('CLI', sWhere) > 0) ;
{$ENDIF}
if Ok then
   BEGIN
   if Ind > 0 then
      BEGIN
      sWhere := V_PGI.DECombos[Ind].Where ;
      sWhere := FindEtReplace(sWhere,'"','',True) ;
      sWhere := FindEtReplace(sWhere,'(','',True) ;
      sWhere := FindEtReplace(sWhere,')','',True) ;
      END ;
   sWhere := 'T_TIERS='+ TH.Text +';'+ sWhere ;
   ////if pos > 0 then sWhere := 'XX_WHERE='+FindEtReplace(V_PGI.DECombos[pos].Where,'=','##',True) ;
   sTiers := AGLLanceFiche('MFO', 'FOCLIMUL_MODE', sWhere, '', 'SELECTION') ;
   if sTiers <> '' then
      BEGIN
      TH.Text := sTiers ;
      Result := True ;
      END ;
   END else
   BEGIN
   Result := (DispatchRecherche(TH, 2, sWhere, '', '') <> '') ;
   END ;
END ;

function GetTiersMul(TH: THCritMaskEdit): boolean;
var St, sw: string;
  ii: integer;
begin
  if ctxFO in V_PGI.PGIContexte then Result := GetTiersMul_FO(TH) else
  begin
    St := 'GCTIERSSAISIE';
    if (ctxAffaire in V_PGI.PGIContexte) and (TH.DataType <> '') then st := TH.DataType;
    ii := TTToNum(St);
    if ii > 0 then sw := V_PGI.DECombos[ii].Where;
    Result := (DispatchRecherche(TH, 2, sw, '', '') <> '');
  end;
end;

function GetSqlTiersLike(LeDataType, Letiers : string; Exact : boolean=false) : string;
var St, sw  : string;
    ii : integer;
begin
  Result := '';
  St := 'GCTIERSSAISIE';
  if (ctxAffaire in V_PGI.PGIContexte) and (LeDataType <> '') then
    st := LeDataType;
  ii := TTToNum(St);
  if ii > 0 then
    sw := V_PGI.DECombos[ii].Where;
  if not Exact then
    Result := ' FROM TIERS WHERE T_TIERS LIKE "' + Letiers + '%" '
    else
    Result := ' FROM TIERS WHERE T_TIERS = "' + Letiers + '" ';
  if sw <> '' then
    Result := Result + ' AND ' + sw;
end;

function GetTiers(TH: THCritMaskEdit): boolean;
var TQ : TQuery;
    Sql : string;
begin
  Result := false;
  Sql := GetSqlTiersLike(TH.DataType, TH.Text);
  TQ := OpenSQL ('SELECT ##TOP 2## T_TIERS' + Sql, true);
  try
    if TQ.RecordCount <> 1 then
    begin
    if GetParamSocSecur('SO_GCRECHTIERSAV', False) then
        Result := GetTiersMul(TH)
    else
        Result := LookUpCombo(TH);
    end else if TQ.RecordCount = 1 then
    begin
      Ferme (TQ);
      TQ := OpenSQL ('SELECT T_TIERS' + Sql, true,-1, '', True);
      TH.Text := TQ.Fields [0].AsString;
      Result := true;
    end;
  finally
    Ferme (TQ);
  end;
end;

procedure TiersVersPiece(TOBTiers, TOBPiece: TOB);
var CodeLivr      : String;
    CodeFact      : String;
    CodePay       : String;
    AuxiLivr      : String;
    AuxiFact      : String;
    AuxiPay       : String;
    NaturePieceG  : string;
    ModeReglement : String;
    TiersLivre    : string;
    i             : integer;
begin
  if not TobTiers.FieldExists('YTC_TIERSLIVRE') then
    TiersLivre := ''
  else
    TiersLivre := NullToVide(Tobtiers.GetValue('YTC_TIERSLIVRE'));
  //if (Tobtiers.GetValue('YTC_TIERSLIVRE') <> '') then
  //  CodeLivr := TiersAuxiliaire(TOBTiers.GetValue('YTC_TIERSLIVRE'), True)
  if TiersLivre <> '' then
    CodeLivr := TiersAuxiliaire(TiersLivre, True)
  else
    CodeLivr := TOBTiers.GetValue('T_TIERS');

  AuxiLivr := TOBTiers.GetValue('T_AUXILIAIRE');
  AuxiFact := TOBTiers.GetValue('T_FACTURE');
  AuxiPay := TOBTiers.GetValue('T_PAYEUR');

  if ((AuxiFact = '') or (AuxiFact = AuxiLivr)) then
  begin
    AuxiFact := AuxiLivr;
    CodeFact := CodeLivr;
  end
  else
  begin
    CodeFact := TiersAuxiliaire(AuxiFact, True);
  end;

  if ((AuxiPay = '') or (AuxiPay = AuxiFact)) then
  begin
    AuxiPay := AuxiFact;
    CodePay := CodeFact;
  end
  else
  begin
    CodePay := TiersAuxiliaire(AuxiPay, True);
  end;
  
  TOBPiece.PutValue('GP_TIERSLIVRE', CodeLivr);
  //  TOBPiece.PutValue('GP_NADRESSELIV', TOBTiers.GetValue('YTC_NADRESSELIV'));
  TOBPiece.PutValue('GP_TIERSFACTURE', CodeFact);
  //  TOBPiece.PutValue('GP_NADRESSEFAC', TOBTiers.GetValue('YTC_NADRESSEFAC'));
  TOBPiece.PutValue('GP_TIERSPAYEUR', CodePay);
  TOBPiece.PutValue('GP_RIB', TOBTiers.GetValue('RIB'));
  TOBPiece.PutValue('GP_FACTUREHT', TOBTiers.GetValue('T_FACTUREHT'));
  TOBPiece.PutValue('GP_ESCOMPTE', TOBTiers.GetValue('T_ESCOMPTE'));

  //recherche du mode de régelement du client payers !!!
  ModeReglement := RechZoneTiers ('T_MODEREGLE', 'T_AUXILIAIRE', CodePay, 'CLI');
  if ModeReglement <> '' then
    TOBPiece.PutValue('GP_MODEREGLE',ModeReglement)
  else
    TOBPiece.PutValue('GP_MODEREGLE', TOBTiers.GetValue('T_MODEREGLE'));

  TOBPiece.PutValue('GP_REGIMETAXE', TOBTiers.GetValue('T_REGIMETVA'));
  TOBPiece.PutValue('GP_QUALIFESCOMPTE', TOBTiers.GetValue('T_QUALIFESCOMPTE'));
  if TOBTiers.FieldExists('YTC_TARIFSPECIAL') then // JT - Ajout test car si n'existe pas, msg erreur en vision SAV
    TOBPiece.PutValue('GP_TARIFSPECIAL', TOBTiers.GetValue('YTC_TARIFSPECIAL'));
//  TOBPiece.PutValue('GP_COMMSPECIAL', TOBTiers.GetValue('YTC_COMMSPECIAL'));
  TOBPiece.PutValue('GP_TARIFTIERS', TOBTiers.GetValue('T_TARIFTIERS'));
  TOBPiece.PutValue('GP_PRESCRIPTEUR', TOBTiers.GetValue('T_PRESCRIPTEUR'));
  NaturePieceG := TOBPiece.GetValue('GP_NATUREPIECEG');
  if GetInfoParPiece(NaturePieceG, 'GPP_BLOBTIERS') = 'X' then TOBPiece.PutValue('GP_BLOCNOTE', TOBTiers.GetValue('T_BLOCNOTE'));
  if TOBTiers.FieldExists('YTC_TABLELIBRETIERS1') then
  begin
    if (TOBTiers.GetValue('T_NATUREAUXI') = 'CLI') or (TOBTiers.GetValue('T_NATUREAUXI') = 'PRO') then
    begin
      for i := 1 to 9 do TOBPiece.PutValue('GP_LIBRETIERS' + IntToStr(i), TOBTiers.GetValue('YTC_TABLELIBRETIERS' + IntToStr(i)));
      TOBPiece.PutValue('GP_LIBRETIERSA', TOBTiers.GetValue('YTC_TABLELIBRETIERSA'));
    end else
    begin
      if TOBTiers.GetValue('T_NATUREAUXI') = 'FOU' then
        for i := 1 to 3 do TOBPiece.PutValue('GP_LIBRETIERS' + IntToStr(i), TOBTiers.GetValue('YTC_TABLELIBREFOU' + IntToStr(i)))
    end;
  end;
  if (ctxAffaire in V_PGI.PGIContexte) and (TOBTiers.FieldExists('YTC_RESSOURCE1')) then
  begin
    for i := 1 to 3 do TOBPiece.PutValue('GP_TIERSSAL' + IntToStr(i), TOBTiers.GetValue('YTC_RESSOURCE' + IntToStr(i)));
  end;
  PutValueDetail(TOBPiece, 'GP_TIERS', TOBTiers.getValue('T_TIERS'));
end;

procedure ChangeTzTiersSaisie(NaturePieceG: String3);
var St, StNat, sWhere, VenteAchat: string;
  ii: integer;
begin
  StNat := GetInfoParPiece(NaturePieceG, 'GPP_NATURETIERS');
  VenteAchat := GetInfoParPiece(NaturePieceG, 'GPP_VENTEACHAT');
  sWhere := FabricWhereToken(stNat, 'T_NATUREAUXI');
  St := 'GCTIERSSAISIE';
  ii := TTToNum(St);
  if ii > 0 then
  begin
    V_PGI.DECombos[ii].Where := sWhere;
    if VenteAchat = 'VEN' then V_PGI.DECombos[ii].Libelle := 'Clients' else
      if VenteAchat = 'ACH' then V_PGI.DECombos[ii].Libelle := 'Fournisseurs' else
      V_PGI.DECombos[ii].Libelle := 'Tiers';
  end;
end;

function RemplirTOBTiers(TOBTiers: TOB; CodeTiers, NaturePiece: string; ControleFerme: boolean): T_RechTiers;
var Q: TQuery;
  isFerme, Auto: boolean;
  TOBT: TOB;
  Auxi, sRib: string;
  TiersComplOk: Boolean; 
begin
  Result := trtOk;
  if CodeTiers = '' then
  begin
    TOBTiers.ClearDetail;
    TOBTiers.InitValeurs(False);
  end else if CodeTiers <> TOBTiers.GetValue('T_TIERS') then
  begin
    { pour éviter les nulls dans TobTiers si pas d'enreg. dans TIERSCOMPL }
    TiersComplOk := False;
    while (not TiersComplOk) and (Result = trtOk) do
    begin
      Q := OpenSQL('SELECT * FROM TIERS LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS WHERE T_TIERS="' + CodeTiers + '"', True,-1, '', True);
      if not Q.EOF then
        if not TOBTiers.SelectDB('', Q) then
          Result := trtIncorrect;
      Ferme(Q);
      if Result = trtOk then
      begin
        { Pas d'enreg dans TiersCompl => On le crée }
        if (TobTiers.G('YTC_AUXILIAIRE') = Null) or (TobTiers.G('YTC_TIERS') = Null) then
          { GPAO Annecy: Si la ligne ci-dessous provoque un bugg, veuillez prévenir le DEV à ANNECY }
          { avant de la mettre en remarque car sans elle cela ne fonctionne pas terrible. Merci }
          InitTiersCompl(TobTiers.G('T_AUXILIAIRE'), TobTiers.G('T_TIERS'))
        else
          TiersComplOk := True; { Tout est Ok => on sort de la boucle }
      end;
    end;
    if Result = trtOk then
    begin
      {RIB}
      Auxi := TOBTiers.GetValue('T_AUXILIAIRE');
      sRib := '';
      Q := OpenSQL('SELECT * FROM RIB WHERE R_AUXILIAIRE="' + Auxi + '" AND R_PRINCIPAL="X"', True,-1, '', True);
      if not Q.EOF then sRib := EncodeRIB(Q.FindField('R_ETABBQ').AsString, Q.FindField('R_GUICHET').AsString,
          Q.FindField('R_NUMEROCOMPTE').AsString, Q.FindField('R_CLERIB').AsString,
          Q.FindField('R_DOMICILIATION').AsString);
      Ferme(Q);
      TOBTiers.PutValue('RIB', sRIB);
      {Exceptions Tiers}
      TOBTiers.ClearDetail;
      if NaturePiece <> '' then
      begin
        Q := OpenSQL('SELECT * FROM TIERSPIECE WHERE GTP_TIERS="' + CodeTiers + '" AND GTP_NATUREPIECEG="' + NaturePiece + '"', True,-1, '', True);
        if not Q.EOF then TOBTiers.LoadDetailDB('TIERSPIECE', '', '', Q, False);
        Ferme(Q);
      end;
    end;
  end;
  isFerme := (TOBTiers.GetValue('T_FERME') = 'X');
  Auto := True;
  if ((isFerme) and (ControleFerme)) then
  begin
    TOBT := TOBTiers.FindFirst(['GTP_NATUREPIECEG'], [NaturePiece], False);
    if TOBT <> nil then Auto := (TOBT.GetValue('GTP_SUSPENSION') = 'X') else Auto := False;
  end;
  if not Auto then
  begin
    Result := trtFerme;
    TOBTiers.ClearDetail;
    TOBTiers.InitValeurs(False);
  end;
end;

function ReferenceIntExtIsOk(From: T_CtrlRefIntExt; TobPiece: TOB; Reference: string): Boolean; 
{
   Contrôle de la référence externe
   Renvoi FALSE si une pièce de même nature contient déjà cette référence externce, pour le TIERS
}
var
  Msg, AlerteMode, FieldName, Sql, DatePiece: string;
  Numero: Integer;
  Q: TQuery;

  function GetAlerteMode: string;
  var
    ParPieceFieldName: string;
  begin
    Result := '';
    case From of
      crInterne: ParPieceFieldName := 'GPP_REFINTCTRL';
      crExterne: ParPieceFieldName := 'GPP_REFEXTCTRL';
    else EXIT;
    end;
    Result := GetInfoParPiece(TobPiece.GetValue('GP_NATUREPIECEG'), ParpieceFieldName);
  end;

begin
  Result := True; { par défaut : Tout va bien }
  if Reference <> '' then
  begin
    AlerteMode := GetAlerteMode;
    case From of
      crInterne:
        begin
          FieldName := 'GP_REFINTERNE';
          Msg := 'la référence interne';
        end;
      crExterne:
        begin
          FieldName := 'GP_REFEXTERNE';
          Msg := 'la référence externe';
        end;
    end;
    if (AlerteMode <> '') and (AlerteMode <> '000') then { 000 = Pas de contrôle }
    begin
      Sql := 'SELECT GP_NUMERO, GP_DATEPIECE '
        + 'FROM PIECE '
        + 'WHERE GP_NATUREPIECEG="' + TobPiece.GetValue('GP_NATUREPIECEG') + '" '
        + 'AND GP_TIERS="' + TobPiece.GetValue('GP_TIERS') + '" '
        + 'AND ' + FieldName + '="' + Reference + '"' { = Référence INTERNE / EXTERNE }
      + 'AND GP_NUMERO <> ' + IntToStr(TobPiece.GetValue('GP_NUMERO')); { Modif = Pas la même pièce }

      try
        Q := OpenSQL(Sql, True,-1, '', True);
        if not Q.EOF then
        begin
          if (AlerteMode = '001') or (AlerteMode = '002') then
          begin
            Numero := Q.FindField('GP_NUMERO').AsInteger;
            DatePiece := DateToStr(Q.FindField('GP_DATEPIECE').AsDateTime);
            PgiInfo(TraduireMemoire(Msg) + ' : ' + Reference + ' ' + TraduireMemoire('existe déjà') + ' '
              + '(' + TraduireMemoire('Pièce N°') + ' : ' + IntToStr(Numero)
              + ' ' + TraduireMemoire('datée du') + ' : ' + DatePiece + ')', TraduireMemoire('Contrôle de') + ' ' + Msg);
          end;
          if AlerteMode = '002' then Result := False;
        end;
      finally
        Ferme(Q);
      end;
    end;
  end;
end;

end.
