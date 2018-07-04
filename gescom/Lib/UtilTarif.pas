unit UtilTarif;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  SysUtils, Classes, Hctrls, SaisUtil,
  {$IFDEF EAGLCLIENT}
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,
  {$ENDIF}
  {$IFNDEF EAGLSERVER}
  UtilArticle,
  {$ENDIF}
  UTOB, HEnt1, ParamSoc
  ;

{$IFNDEF EAGLSERVER}
function ChercheTarif(OArt, OTiers, OLigne, OPiece, OTarif: TOB; QueQte, EnHT: boolean): Boolean;
function ChercheTarifTTCMode(OTarif: Tob; CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat: string; LaDate: TDateTime; EnHT:
  Boolean): Boolean;
{$IFDEF BTP}
procedure GetTarifGlobal(CodeArticle, TarifArticle, SousTarifArticle, VenteAchat: string; TOBArt, TOBTiers, TOBTarif: TOB; EnHt: boolean; QteRef : double = 1; DateRef : TdateTime=0);
{$ENDIF}
{$IFDEF CHR}
function ChercheTarifCHR(OArt, OTiers, OLigne, OPiece, OTarif: TOB; QueQte, EnHT: boolean): Boolean;
function  SQLTarifCHR (CodeArt,TarifArticle,CodeTiers,TarifTiers,CodeDevise,CodeDepot : String ; LaDate : TDateTime ; Qte : double;CodeTypres:string) : String ;
{$ENDIF}
{$ENDIF}
procedure CalcPriorite(Q: TQuery);
function ChercheWEBtarif(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, LaDate, Qte, RemiseTiers: OLEVariant; var LePrix, LaRemise:
  OLEVariant): OLEVariant;
function SQLTarif(CodeArt, TarifArticle, SousFamilleTarif ,CodeTiers, TarifTiers, CodeDevise, CodeDepot: string; LaDate: TDateTime; Qte: double): string;
function CompleteSQLTarif(VenteAchat: string; QueQte, EnHT: Boolean): string;
implementation

uses TarifUtil;

procedure CalcPriorite(Q: TQuery);
var Priorite: Integer;
begin
  Priorite := 0;
  if Q.FindField('GF_ARTICLE').AsString <> '' then Priorite := Priorite + 40 else
    if Q.FindField('GF_TARIFARTICLE').AsString <> '' then Priorite := Priorite + 10;
  if Q.FindField('GF_SOUSFAMTARART').AsString  <> '' then Priorite := Priorite + 10;
  if Q.FindField('GF_TIERS').AsString <> '' then Priorite := Priorite + 40;
  if Q.FindField('GF_TARIFTIERS').AsString <> '' then Priorite := Priorite + 5;
  if Q.FindField('GF_DEVISE').AsString <> '' then Priorite := Priorite + 2;
  if Q.FindField('GF_DEPOT').AsString <> '' then Priorite := Priorite + 1;
  if Q.FindField('GF_CASCADEREMISE').AsString = 'FOR' then Priorite := Priorite + 200;
  if Q.FindField('GF_CASCADEREMISE').AsString = 'CAS' then Priorite := Priorite + 100;
  if Q.FindField('GF_FERME').AsString = 'X' then Priorite := 0;
  Q.FindField('GF_PRIORITE').AsInteger := Priorite;
end;

function VerifConditions(Q: TQuery; OArt, OTiers, OLigne, OPiece: TOB): Boolean;
var Conditions: TStrings;
  i: Integer;
  s: string;
  sTable: string;
  sChamp: string;
  sOperateur: string;
  sValeur: string;
  vChamp: Variant;
  vValeur: Variant;
  ssChamp: string;
  ssValeur: string;
begin
  Result := True;
  Conditions := TStringList.Create;
  {$IFDEF EAGLCLIENT}
  Conditions.Text := Q.FindField('GF_CONDAPPLIC').AsString;
  {$ELSE}
  Conditions.Assign(TMemoField(Q.FindField('GF_CONDAPPLIC')));
  {$ENDIF}
  if Conditions = nil then
  begin
    Conditions.Free;
    Exit;
  end;
  for i := 0 to Conditions.Count - 1 do
  begin
    s := Conditions.Strings[i];
    sTable := ReadTokenSt(s);
    sChamp := ReadTokenSt(s);
    sOperateur := ReadTokenSt(s);
    sValeur := ReadTokenSt(s);
    {Champ de test}
    if sTable = 'TIERS' then vChamp := OTiers.GetValue(sChamp);
    if sTable = 'ARTICLE' then vChamp := OArt.GetValue(sChamp);
    if sTable = 'LIGNE' then vChamp := OLigne.GetValue(sChamp);
    if sTable = 'PIECE' then vChamp := OPiece.GetValue(sChamp);
    ssChamp := Uppercase(VarAsType(vChamp, varString));
    {Valeur de test}
    try // Fiche DBR 10163
      vValeur := VarAsType(sValeur, VarType(vChamp));
      ssValeur := Uppercase(sValeur);

      {Test de la condition}
      if sOperateur = '=' then Result := (vChamp = vValeur);
      if sOperateur = '<>' then Result := (vChamp <> vValeur);
      if sOperateur = '>' then Result := (vChamp > vValeur);
      if sOperateur = '<' then Result := (vChamp < vValeur);
      if sOperateur = '>=' then Result := (vChamp >= vValeur);
      if sOperateur = '<=' then Result := (vChamp <= vValeur);
      if sOperateur = 'C' then Result := (Copy(ssChamp, 1, Length(ssValeur)) = ssValeur);
      if sOperateur = 'D' then Result := (Copy(ssChamp, 1, Length(ssValeur)) <> ssValeur);
      if sOperateur = 'L' then Result := (Pos(ssValeur, ssChamp) > 0);
      if sOperateur = 'M' then Result := (Pos(ssValeur, ssChamp) = 0);
    except
    end;
    if not Result then
    begin
      Conditions.Free;
      Exit;
    end;
  end;
  Conditions.Free;
end;

function SQLTarif(CodeArt, TarifArticle, SousFamilleTarif ,CodeTiers, TarifTiers, CodeDevise, CodeDepot: string; LaDate: TDateTime; Qte: double): string;
var SQL: string;
begin
	if (VarIsNull (SousFamilleTarif)) or (VarAsType (SousFamilleTarif,varString)=#0) then SousFamilleTarif := '';
  SQL := 'Select * From TARIF Where ' +
    '((GF_ARTICLE="' + CodeArt + '" AND GF_TIERS="' + CodeTiers + '" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE="" AND GF_SOUSFAMTARART="") OR' +
    ' (GF_ARTICLE="' + CodeArt + '" AND GF_TIERS="" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE="" AND GF_SOUSFAMTARART="") OR';
  SQL := SQL +
    ' (GF_ARTICLE="' + CodeArt + '" AND GF_TIERS="" AND GF_TARIFTIERS="' + TarifTiers + '") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="' + CodeTiers + '" AND GF_TARIFARTICLE="") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="' + CodeTiers + '" AND GF_TARIFARTICLE="' + TarifArticle + '" AND GF_SOUSFAMTARART="") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="' + CodeTiers + '" AND GF_TARIFARTICLE="' + TarifArticle + '" AND GF_SOUSFAMTARART="'+SousFamilleTarif+'") OR';
  SQL := SQL +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="' + TarifTiers + '" AND GF_TARIFARTICLE="' + TarifArticle + '") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE="' + TarifArticle + '") OR';
  SQL := SQL +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="' + TarifTiers + '" AND GF_TARIFARTICLE="") OR' +
    ' (GF_ARTICLE="" AND GF_TIERS="" AND GF_TARIFTIERS="" AND GF_TARIFARTICLE=""))';
  SQL := SQL + ' AND GF_DATEDEBUT<="' + USDatetime(LaDate) + '" AND GF_DATEFIN>="' + USDatetime(LaDate) + '"' +
    ' AND GF_BORNEINF<=' + StrFPoint(Qte) + ' AND GF_BORNESUP>=' + StrFPoint(Qte) +
    ' AND (GF_DEPOT="' + CodeDepot + '" OR GF_DEPOT="")' +
    ' AND (GF_DEVISE="' + CodeDevise + '" OR GF_DEVISE="")' +
    ' AND GF_FERME="-"';
  Result := SQL;
end;

function CompleteSQLTarif(VenteAchat: string; QueQte, EnHT: Boolean): string;
var SQL: string;
begin
  SQL := '';
  if QueQte then SQL := SQL + ' AND GF_QUANTITATIF="X" ';
  if VenteAchat <> 'VEN' then SQL := SQL + ' AND GF_NATUREAUXI="FOU" ' else SQL := SQL + ' AND GF_NATUREAUXI<>"FOU" ';
  if EnHT then SQL := SQL + ' AND GF_REGIMEPRIX<>"TTC" ' else SQL := SQL + ' AND GF_REGIMEPRIX<>"HT" ';
  SQL:=SQL+' ORDER BY GF_PRIORITE DESC, GF_DATEDEBUT DESC' ;
  Result := SQL;
end;

function ChercheWEBtarif(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, LaDate, Qte, RemiseTiers: OLEVariant; var LePrix, LaRemise:
  OLEVariant): OLEVariant;
var SQL, StCascade, QuelleCascade: string;
  Q: TQuery;
  OTarif: TOB;
  Remise: double;
begin
  Result := False;
  LePrix := 0;
  LaRemise := 0;
  Remise := 0;
  OTarif := TOB.Create('TARIF', nil, -1);

  SQL := SQLTarif(CodeArt, TarifArticle,'', CodeTiers, TarifTiers, CodeDevise, CodeDepot, LaDate, Qte);
  SQL := SQL + ' AND GF_NATUREAUXI<>"FOU" ';
  SQL := SQL + ' AND GF_REGIMEPRIX<>"TTC" ';
  SQL := SQL + ' ORDER BY GF_PRIORITE DESC';

  Q := OpenSQL(SQL, True);
  if not Q.EOF then
  begin
    Result := True;
    OTarif.SelectDB('', Q);
    StCascade := StCascade + Q.FindField('GF_CALCULREMISE').AsString;
    QuelleCascade := Q.FindField('GF_CASCADEREMISE').AsString;
    Remise := Q.FindField('GF_REMISE').AsFloat;
    if Q.FindField('GF_CASCADEREMISE').AsString = 'CAS' then
    begin
      repeat
        Q.Next;
        if Q.EOF then Break;
        QuelleCascade := Q.FindField('GF_CASCADEREMISE').AsString;
        if Q.FindField('GF_PRIXUNITAIRE').AsFloat <> 0 then OTarif.PutValue('GF_PRIXUNITAIRE', Q.FindField('GF_PRIXUNITAIRE').AsFloat);
        Remise := 100.0 * (1.0 - (1.0 - Remise / 100.0) * (1.0 - Q.FindField('GF_REMISE').AsFloat / 100));
        StCascade := StCascade + Q.FindField('GF_CALCULREMISE').AsString;
      until ((Q.EOF) or (Q.FindField('GF_CASCADEREMISE').AsString <> 'CAS'));
      OTarif.PutValue('GF_REMISE', Remise);
    end;
  end;
  Ferme(Q);
  if QuelleCascade = 'FOR' then else {rien à faire, le tarif a été trouvé}
    if QuelleCascade = 'CAS' then else {rien à faire, la remise cascade a été caclulée}
    if QuelleCascade = 'MIE' then
  begin
    {prendre la meilleure entre remise issue du tarif et celle du tiers}
    if RemiseTiers > Remise then OTarif.PutValue('GF_REMISE', RemiseTiers);
  end else if QuelleCascade = 'CUM' then
  begin
    {cumuler la remise du tarif et celle du tiers}
    Remise := 100.0 * (1.0 - (1.0 - Remise / 100.0) * (1.0 - RemiseTiers / 100));
    OTarif.PutValue('GF_REMISE', Remise);
  end;
  OTarif.PutValue('GF_CALCULREMISE', StCascade);
  if Result then
  begin
    LaRemise := OTarif.GetValue('GF_REMISE');
    LePrix := OTarif.GetValue('GF_PRIXUNITAIRE');
  end;
  OTarif.Free;
end;

{$IFNDEF EAGLSERVER}

function ChercheTarif(OArt, OTiers, OLigne, OPiece, OTarif: TOB; QueQte, EnHT: boolean): Boolean;
var Q: TQuery;
  Find,TarifFFO: boolean;
  SQL, VenteAchat: AnsiString;
  QuelleCascade, CodeArt, CodeTiers, CodeDepot, CodeDevise, TarifArticle, TarifTiers, StCascade, SousFamilleTarif: string;
  LaDate: TDateTime;
  Qte, RemiseTiers, Remise: Double;
  {$IFDEF MODE} TarifNegoce : boolean ; {$ENDIF}
begin
  {$IFDEF CHR}
  result := ChercheTarifCHR(OArt, OTiers, Oligne, OPiece, OTarif, False, EnHT);
  {$ELSE} // Fin CHR
  Find := FALSE;
  TarifFFO := False;
  StCascade := '';
  Remise := 0;
  QuelleCascade := 'MIE';
  CodeArt := OLigne.GetValue('GL_ARTICLE');
  CodeTiers := OTiers.GetValue('T_TIERS');
  CodeDepot := OLigne.GetValue('GL_DEPOT');
  CodeDevise := OPiece.GetValue('GP_DEVISE');
  Qte := OLigne.GetValue('GL_QTEFACT');
  TarifArticle := OLigne.GetValue('GL_TARIFARTICLE');
  LaDate := OPiece.GetValue('GP_DATEPIECE');
  TarifTiers := OPiece.GetValue('GP_TARIFTIERS');
  RemiseTiers := OTiers.GetValue('T_REMISE');
  VenteAchat := OPiece.GetValue('GP_VENTEACHAT');
  if (VarIsNull(OArt.GetValue('GA2_SOUSFAMTARART'))) or
  	 (VarAsType(OArt.GetValue('GA2_SOUSFAMTARART'),VarString)=#0) then
  begin
  	SousFamilleTarif := '';
  end else
  begin
  	SousFamilleTarif := Oart.GEtVAlue('GA2_SOUSFAMTARART');
  end;

  // Debut Modif AC1 13/05/2003 (On cherche le tarif d'un établissement)
  if (CtxMode in V_PGI.PGIContexte) then CodeDepot := OLigne.GetValue('GL_ETABLISSEMENT');
  // Fin Modif AC1 13/05/2003

  {$IFDEF MODE}
  if (VenteAchat='VEN') and EnHT then TarifNegoce := True else TarifNegoce := False ;
  // Recherche des tarifs détail
  if ((CtxMode in V_PGI.PGIContexte) and (OArt.GetValue('GA_STATUTART')='UNI') and not TarifNegoce)
  or (CtxFo in V_PGI.PGIContexte)  then
  begin
    Find := ChercheTarifTTCMode(OTarif, CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat, LaDate, EnHT);
    TarifFFO := True ;
  end ;
  {$ENDIF}

  if not TarifFFO then
  begin
    {$IFDEF MODE} // Debut Modif Mode AC 3/06/2003  optimisation chargement tarif
    if not (CtxFO in V_PGI.PGIContexte) and (OArt.GetValue('GA_STATUTART') = 'GEN') then
    begin
      Result := True;
      exit;
    end;
    CodeArt := CodeArticleUnique(Copy(CodeArt, 1, 18), '', '', '', '', '');
    SQL := SQLTarif(CodeArt, TarifArticle,SousFamilleTarif, CodeTiers, TarifTiers, CodeDevise, CodeDepot, LaDate, Qte);
    SQL := SQl + CompleteSQLTarif(VenteAchat, QueQte, EnHT);
    Q := OpenSQL(SQL, True);
    {$ELSE}

    SQL := SQLTarif(CodeArt, TarifArticle,'', CodeTiers, TarifTiers, CodeDevise, CodeDepot, LaDate, Qte);
    SQL := SQl + CompleteSQLTarif(VenteAchat, QueQte, EnHT);

    Q := OpenSQL(SQL, True);
    {$ENDIF} // fin Modif Mode

    if not Q.EOF then
    begin
      {Vérification des conditions d'applications}
      while not Q.EOF do
      begin
        if not VerifConditions(Q, OArt, OTiers, OLigne, OPiece) then Q.Next else Break;
      end;
      if Q.EOF then
      begin
        Ferme(Q);
        Result := False;
        Exit;
      end;
      OTarif.SelectDB('', Q);
      StCascade := StCascade + Q.FindField('GF_CALCULREMISE').AsString;
      QuelleCascade := Q.FindField('GF_CASCADEREMISE').AsString;
      Remise := Q.FindField('GF_REMISE').AsFloat;
      Find := TRUE;
      if ((not QueQte) and (Q.FindField('GF_CASCADEREMISE').AsString = 'CAS')) then
        // Si premier = cascade --> dérouler cascade, y compris l'éventuel dernier non cascade
      begin
        repeat
          Q.Next;
          if Q.EOF then Break;
          QuelleCascade := Q.FindField('GF_CASCADEREMISE').AsString;
          if Q.FindField('GF_PRIXUNITAIRE').AsFloat <> 0 then OTarif.PutValue('GF_PRIXUNITAIRE', Q.FindField('GF_PRIXUNITAIRE').AsFloat);
          Remise := 100.0 * (1.0 - (1.0 - Remise / 100.0) * (1.0 - Q.FindField('GF_REMISE').AsFloat / 100));
          StCascade := StCascade + Q.FindField('GF_CALCULREMISE').AsString;
        until ((Q.EOF) or (Q.FindField('GF_CASCADEREMISE').AsString <> 'CAS'));
        OTarif.PutValue('GF_REMISE', Remise);
      end;
    end else
    begin
    	OTarif.InitValeurs;
    end;
    Ferme(Q);
    if QuelleCascade = 'FOR' then else {rien à faire, le tarif a été trouvé}
      if QuelleCascade = 'CAS' then else {rien à faire, la remise cascade a été caclulée}
      if ((not QueQte) and (QuelleCascade = 'MIE')) then
    begin
      {prendre la meilleure entre remise issue du tarif et celle du tiers}
      if RemiseTiers > Remise then OTarif.PutValue('GF_REMISE', RemiseTiers);
    end else if ((not QueQte) and (QuelleCascade = 'CUM')) then
    begin
      {cumuler la remise du tarif et celle du tiers}
      Remise := 100.0 * (1.0 - (1.0 - Remise / 100.0) * (1.0 - RemiseTiers / 100));
      OTarif.PutValue('GF_REMISE', Remise);
    end;
    OTarif.PutValue('GF_CALCULREMISE', StCascade);
   end ;
  Result := Find;
  {$ENDIF} // Fin CHR
end;
{$IFDEF BTP}

procedure GetTarifGlobal(CodeArticle, TarifArticle, SousTarifArticle, VenteAchat: string; TOBArt, TOBTiers, TOBTarif: TOB; EnHt: boolean; QteRef : double = 1; DateRef : TdateTime=0);
var Q: TQuery;
  Find: boolean;
  SQL: AnsiString;
  QuelleCascade, CodeTiers, CodeDepot, CodeDevise, TarifTiers, StCascade: string;
  LaDate: TDateTime;
  Qte, RemiseTiers, Remise, Prix: Double;
begin
  Find := FALSE;
  StCascade := '';
  Remise := 0;
  QuelleCascade := 'MIE';
  CodeTiers := TobTiers.GetValue('T_TIERS');
  CodeDepot := '';
  CodeDevise := GetParamSoc('SO_DEVISEPRINC');
  TarifTiers := TOBTIers.GetValue('T_TARIFTIERS');
//  Qte := 1;
	Qte := QteRef;
//  LaDate := DateRef;
  if DateRef = 0 then LaDate := V_PGI.DateEntree else LaDate := DateRef;
  TarifTiers := '';
  RemiseTiers := 0;

  SQL := SQLTarif(CodeArticle, TarifArticle, SousTarifArticle ,CodeTiers, TarifTiers, CodeDevise, CodeDepot, LaDate, Qte);
  SQL := SQl + CompleteSQLTarif(VenteAchat, false, EnHT);

  Q := OpenSQL(SQL, True);

  if not Q.EOF then
  begin
    {Vérification des conditions d'applications}
    while not Q.EOF do
    begin
      if not VerifConditions(Q, TOBArt, TOBTiers, nil, nil) then Q.Next else Break;
    end;
    if Q.EOF then
    begin
      Ferme(Q);
      Exit;
    end;
    TOBTarif.SelectDB('', Q);
    TOBTarif.PutValue('GF_PRIXANCIEN', TobTarif.GetValue('GF_PRIXUNITAIRE'));
    if Q.FindField('GF_CALCULREMISE').AsString <> '' then
    begin
      if stcascade <> '' then StCascade := Q.FindField('GF_CALCULREMISE').AsString + '%' + '+' + stcascade
      else stCascade := Q.FindField('GF_CALCULREMISE').AsString + '%';
    end;
    QuelleCascade := Q.FindField('GF_CASCADEREMISE').AsString;
    Remise := Q.FindField('GF_REMISE').AsFloat;
    Find := TRUE;
    if ((Q.FindField('GF_CASCADEREMISE').AsString = 'CAS')) then
      // Si premier = cascade --> dérouler cascade, y compris l'éventuel dernier non cascade
    begin
      repeat
        Q.Next;
        if Q.EOF then Break;
        QuelleCascade := Q.FindField('GF_CASCADEREMISE').AsString;
        if Q.FindField('GF_PRIXUNITAIRE').AsFloat <> 0 then TOBTarif.PutValue('GF_PRIXUNITAIRE', Q.FindField('GF_PRIXUNITAIRE').AsFloat);
        Remise := 100.0 * (1.0 - (1.0 - Remise / 100.0) * (1.0 - Q.FindField('GF_REMISE').AsFloat / 100));
        if Q.FindField('GF_CALCULREMISE').AsString <> '' then
        begin
          if stcascade <> '' then StCascade := Q.FindField('GF_CALCULREMISE').AsString + '%' + '+' + stcascade
          else stCascade := Q.FindField('GF_CALCULREMISE').AsString + '%';
        end;
      until ((Q.EOF) or (Q.FindField('GF_CASCADEREMISE').AsString <> 'CAS'));
      TOBTarif.PutValue('GF_REMISE', Remise);
    end;
  end;
  Ferme(Q);
  if QuelleCascade = 'FOR' then else {rien à faire, le tarif a été trouvé}
    if QuelleCascade = 'CAS' then else {rien à faire, la remise cascade a été caclulée}
    if ((QuelleCascade = 'MIE')) then
  begin
    {prendre la meilleure entre remise issue du tarif et celle du tiers}
    if RemiseTiers > Remise then TOBTarif.PutValue('GF_REMISE', RemiseTiers);
  end else if ((QuelleCascade = 'CUM')) then
  begin
    {cumuler la remise du tarif et celle du tiers}
    Remise := 100.0 * (1.0 - (1.0 - Remise / 100.0) * (1.0 - RemiseTiers / 100));
    TOBTarif.PutValue('GF_REMISE', Remise);
  end;
  TOBTarif.PutValue('GF_CALCULREMISE', StCascade);
end;
{$ENDIF}

function ChercheTarifTTCMode(OTarif: Tob; CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat: string; LaDate: TDateTime; EnHT:
  Boolean): Boolean;
var TOBTarif, TOBMode: Tob;
  QMode, QQ: TQuery;
  TypeTarfEtab, CodeDevisePiece: string;
  Arron, InfoRem, Demarque, NatureType, NomChamp, NomChampDev: string;
  Remise, PrixBase, Prix, PrixCon, PrixConAr: Double;
begin
  Result := False;
  CodeDevisePiece := CodeDevise;
  Remise := 0;
  if EnHT then
  begin
    NatureType := 'ACH';
    NomChamp := 'ET_TYPETARIFACH';
    NomChampDev := 'ET_DEVISEACH';
  end else
  begin
    NatureType := 'VTE';
    NomChamp := 'ET_TYPETARIF';
    NomChampDev := 'ET_DEVISE';
  end;
  // Recherche prix de base
  TOBMode := TOB.Create('', nil, -1);
  QQ := OpenSQL('Select ' + NomChamp + ',' + NomChampDev + ' from ETABLISS Where ET_ETABLISSEMENT="' + CodeDepot + '"', True);
  if not QQ.EOF then
  begin
    TypeTarfEtab := QQ.FindField(NomChamp).AsString;
    CodeDevise := QQ.FindField(NomChampDev).AsString
  end;
  if TypeTarfEtab = '' then TypeTarfEtab := '...';
  Ferme(QQ);
  QMode := OpenSQL('Select GFM_TARFMODE,GFM_TYPETARIF,GFM_NATURETYPE,GFM_DATEDEBUT,GFM_PROMO from TARIFMODE where gfm_typetarif in ("' + TypeTarfEtab +
    '","...") order by GFM_TYPETARIF DESC,GFM_DATEDEBUT DESC', True);
  TOBMode.LoadDetailDB('TARIFMODE', '', '', QMode, False);
  Ferme(QMode);
  TOBTarif := CreerTobTarifArtDim(CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDevise, CodeDepot, VenteAchat, LaDate, EnHT);
  InfoRem := RechTarifSpec(TOBTarif, TOBMode, CodeArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType);
  Prix := Valeur(ReadTokenSt(InfoRem));
  Demarque := ReadTokenSt(InfoRem);
  PrixBase := RechPrixTarifBase(TOBTarif, TOBMode, CodeArt, CodeDepot, CodeDevise, TypeTarfEtab, NatureType);
  if Prix = 0 then
  begin
    if PrixBase <> 0 then
    begin
      if CodeDevise <> CodeDevisePiece then ToxConvToDev2(PrixBase, PrixCon, PrixConAr, CodeDevise, CodeDevisePiece, Date, nil);
      if PrixConAr = 0 then PrixConAr := PrixBase;
      OTarif.PutValue('GF_PRIXUNITAIRE', PrixConAr);
      Result := True;
    end;
    InfoRem := '';
    InfoRem := ChercheMieRem(TOBTarif, TobMode, CodeArt, TarifArticle, CodeTiers, TarifTiers, CodeDepot, CodeDevise, TypeTarfEtab, NatureType);
    Remise := Valeur(ReadTokenSt(InfoRem));
    Arron := ReadTokenSt(InfoRem);
    Demarque := ReadTokenSt(InfoRem);
    if Remise <> 0 then
    begin
      OTarif.PutValue('GF_REMISE', Remise);
      Result := True;
    end;
    if Arron <> '' then OTarif.PutValue('GF_ARRONDI', Arron);
    OTarif.PutValue('GF_DEMARQUE', Demarque);
  end else
  begin
    if CodeDevise <> CodeDevisePiece then ToxConvToDev2(Prix, PrixCon, PrixConAr, CodeDevise, CodeDevisePiece, Date, nil);
    if PrixConAr = 0 then PrixConAr := Prix;
    OTarif.PutValue('GF_PRIXUNITAIRE', PrixConAr);
    OTarif.PutValue('GF_DEMARQUE', Demarque);
    OTarif.PutValue('GF_REMISE', Remise);
    Result := True;
  end;
  OTarif.AddChampSupValeur('PRIXBASE', PrixBase);
  TOBTarif.Free;
  TOBMode.Free;
end;

{$IFDEF CHR}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : GF
Créé le ...... : 05/11/2002
Modifié le ... :   /  /
Description .. : Idem fonction Cherchetarif
Suite ........ : sauf ajout code ressource -> type de ressource
Mots clefs ... :
*****************************************************************}
function  ChercheTarifCHR (OArt,OTiers,OLigne,OPiece,OTarif : TOB ; QueQte,EnHT : boolean ) : Boolean ;
Var Q:TQuery;
    Find : boolean ;
    SQL: AnsiString ;
    QuelleCascade,CodeArt,CodeTiers,CodeDepot,CodeDevise,TarifArticle,TarifTiers,CodeRessource: String ;
    LaDate : TDateTime ;
    Qte : Double ;
    i:integer;
begin
  Find:=FALSE ;
  QuelleCascade:='MIE' ;
  CodeArt:=OLigne.GetValue('GL_ARTICLE')   ;
  CodeTiers:=OTiers.GetValue('T_TIERS') ;
  CodeDepot:=OLigne.GetValue('GL_DEPOT')   ;
  CodeDevise:=OPiece.GetValue('GP_DEVISE') ;

  Qte:=OLigne.GetValue('GL_QTESIT');
  CodeRessource := OLigne.GetValue('GL_RESSOURCE');
  if Qte = 0 then Qte := 1;
  TarifArticle:=OLigne.GetValue('GL_TARIFARTICLE') ;
  LaDate:=OLigne.GetValue('GL_DATEPRODUCTION');
  if (LaDate = iDate1900) then LaDate := V_PGI.DateEntree;
  TarifTiers:=OPiece.GetValue('GP_TARIFTIERS');

  //RemiseTiers:=OTiers.GetValue('T_REMISE') ; VenteAchat:=OPiece.GetValue('GP_VENTEACHAT') ;

  //SQL:=SQLTarif(CodeArt,TarifArticle,CodeTiers,TarifTiers,CodeDevise,CodeDepot,LaDate,Qte) ;
  SQL:=SQLTarifCHR(CodeArt,TarifArticle,'',CodeTiers,TarifTiers,CodeDevise,CodeDepot,LaDate,Qte,CodeRessource) ;

  Q:=OpenSQL(SQL,True) ;

  if (OPiece.detail.count = 0) then // on vient de la préfacturation
  begin
    if Not Q.EOF then
    begin
     OTarif.LoadDetailDB('TARIF','','',Q,False);
     Find:=TRUE ;
    end ;
    Ferme(Q) ;
    if (OTarif.detail.count > 1) then
    begin
      i:=0;
      while i < OTarif.detail.count do
      begin
       if (CodeRessource <>'') and (OTarif.detail[i].getvalue('GF_TYPRES') <> CodeRessource) then
           OTarif.detail[i].Free
       else i:=i+1;
      end;
    end;
    if (OTarif.detail.count = 0) then Result:=False else
    Result:=Find ;
  end else
  begin
    if Not Q.EOF then
    begin
     OTarif.SelectDB('',Q);
     Find:=TRUE ;
    end ;
    Ferme(Q) ;
    Result:=Find;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : GF
Créé le ...... : 08/11/2002
Modifié le ... :   /  /
Description .. : Complète la requete SQL de la recherche de tarif avec le
Suite ........ : type de ressource
Mots clefs ... :
*****************************************************************}
function  SQLTarifCHR (CodeArt,TarifArticle,CodeTiers,TarifTiers,CodeDevise,CodeDepot : String ; LaDate : TDateTime ; Qte : double ; CodeTypres:string) : String ;
Var SQL : String ;
begin
  SQL:='' ;
  SQL:='Select * From TARIF Where '+
     '(GF_ARTICLE="'+CodeArt+'" AND GF_TARIFTIERS="'+TarifTiers+'") ';
  SQL:=SQL+' AND GF_DATEDEBUT<="'+USDatetime(LaDate)+'" AND GF_DATEFIN>="'+USDatetime(LaDate)+'"'+
         ' AND GF_BORNEINF<='+StrFPoint(Qte)+' AND GF_BORNESUP>='+StrFPoint(Qte)+
         ' AND (GF_DEPOT="'+CodeDepot+'" OR GF_DEPOT="")'+
         ' AND (GF_DEVISE="'+CodeDevise+'" OR GF_DEVISE="")'+
         ' AND GF_FERME="-"' ;
  if existeSQL('select HTR_TYPRES from HRTYPRES where HTR_RESSOURCE = "'
                     +CodeTypres+'" and HTR_NATURERES="TRE"') then
  begin
     SQL:=SQL+' AND (GF_TYPRES="'+Codetypres+'" OR GF_TYPRES="")';
  end;
  Result:=SQL ;
end;

{$ENDIF} // CHR 16/05/02
{$ENDIF}

end.
