unit FactTimbres;

interface
Uses HEnt1, UTOB, Ent1,AGLInit,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, DB,
{$ENDIF}
     SysUtils,
     Math, EntGC, Classes,HCtrls,
     Paramsoc,
     uEntCommun,HmsgBox
     ;


function FindTimbre(Code : string) : TOB;
procedure InitParamTimbres;
procedure LibereParamTimbres;
procedure CalculeTimbres (TOBTimbres,TOBPiece,TOBEches : TOB);
procedure ChargeLesTimbres (TOBPiece,TOBTimbres : TOB);
procedure LoadlesTimbres (cledoc : R_CLEDOC; TOBTimbres : TOB);
function  GetTotalTimbres (TOBTimbres,TOBPiece,TOBEches : TOB) : Double;
procedure ValideLesTimbres(TOBPiece, TOBTimbres : TOB);


implementation

uses UtilTOBPiece,FactUtil;
var TOBTIMBRESP : TOB;

function FindTimbre(Code : string) : TOB;
begin
  if TOBTIMBRESP = nil then InitParamTimbres;
  Result := TOBTIMBRESP.findFirst(['BTP_CODE'],[Code],True);
end;

function CalculeMontantTimbre(TOBParam,TOBEche : TOB) : double;
var Mt : Double;
		Vdiv,Vmod : integer;
begin
  Result := 0;
	if TOBParam.getvalue('BTP_TYPETIMBRE')='001' then
  begin
    // % de l'echeance
    Result := TOBEche.GetValue('GPE_MONTANTDEV') * TOBParam.getvalue('BTP_POURCENT')/100;
  end else if TOBParam.getvalue('BTP_TYPETIMBRE')='002' then
  begin
    // montant fixe
    Result := TOBParam.getvalue('BTP_MONTANT');
  end else if TOBParam.getvalue('BTP_TYPETIMBRE')='003' then
  begin
    vDiv := 0;
    vMod := 0;
    // gestion tranches
    if TOBEche.GetValue('GPE_MONTANTDEV') < TOBParam.getvalue('BTP_SEUILAPPLIC') then Exit;
    Vdiv := (TOBEche.GetValue('GPE_MONTANTDEV') div TOBParam.getvalue('BTP_TRANCHE')) * TOBParam.getvalue('BTP_MONTANT');
    if TOBParam.GetValue('BTP_FRACTION')='X' then
    begin
    	vMod := (TOBEche.GetValue('GPE_MONTANTDEV') mod TOBParam.getvalue('BTP_TRANCHE'));
      if Vmod > 0 then Vmod := TOBParam.getvalue('BTP_MONTANT');
    end;
    Mt := VDiv + Vmod;
    if Mt < TOBParam.GetValue('BTP_MINI') then Mt := TOBParam.GetValue('BTP_MINI');
    if Mt > TOBParam.GetValue('BTP_MAXI') then Mt := TOBParam.GetValue('BTP_MAXI');
    Result := Mt;
  end;
end;

procedure CalculeUnTimbre (TOBPiece,TOBTimbre,TOBEches: TOB);
var TOBP,TOBE : TOB;
		indice : Integer;
    MtTImbre : Double;
begin
  if TOBTimbre = nil then Exit;
  if TOBTimbre.GetValue('BT0_CODE')='' then Exit;
  TOBP :=  FindTimbre(TOBTimbre.GetValue('BT0_CODE'));
  if TOBP = nil then exit;
  TOBTimbre.PutValue('BT0_BASETIMBRE',0);
  TOBTimbre.PutValue('BT0_VALEUR',0);
  // La c'est bon on a retrouve le paramétrage --> on peut docn calculer le montant de timbre
  for Indice := 0 to TOBEches.detail.count -1 do
  begin
    TOBE := TOBEches.detail[Indice];
    if Pos(TOBE.GetValue('GPE_MODEPAIE'),TOBP.getValue('BTP_LISTEMODPAIE'))=0 then Continue;
    //
    TOBTimbre.PutValue('BT0_BASETIMBRE',TOBTimbre.GetValue('BT0_BASETIMBRE') + TOBE.GetValue('GPE_MONTANTDEV'));
    MtTImbre := CalculeMontantTimbre(TOBP,TOBE);
    TOBTimbre.PutValue('BT0_VALEUR',TOBTimbre.GetValue('BT0_VALEUR') + MtTImbre);
    TOBE.PutValue('GPE_TIMBRE', MtTImbre);
    TOBE.PutValue('GPE_CODE', TOBTimbre.GetValue('BT0_CODE'));
  end;
end;

procedure CalculeTimbres (TOBTimbres,TOBPiece,TOBEches : TOB);
var Indice : Integer;
begin
	for Indice := 0 to TOBTimbres.detail.count - 1 do
  begin
    CalculeUnTimbre (TOBPiece,TOBTimbres.detail[Indice],TOBEches);
  end;
end;

procedure InitParamTimbres;
var QQ : TQuery;
begin
  TOBTIMBRESP := TOB.Create('LES TIMBRESP',nil,-1);
  TRY
    QQ := OpenSQL('SELECT * FROM TIMBRESPARAM',True,-1,'',true);
    if not QQ.eof then
    begin
      TOBTIMBRESP.LoadDetailDB('TIMBRESPARAM','','',QQ,false);
    end;
  FINALLY
  	Ferme(QQ);
  end;
end;

procedure LibereParamTimbres;
begin
  if TOBTIMBRESP <> nil then FreeAndNil(TOBTIMBRESP);
end;

procedure LoadlesTimbres (cledoc : R_CLEDOC; TOBTimbres : TOB);
var QQ : TQuery;
begin
  QQ := OpenSql ('SELECT * FROM TIMBRESPIECE WHERE '+WherePiece(cledoc,ttdTimbres ,false),True,-1,'',True);
  if not QQ.Eof then
  begin
    TOBTimbres.LoadDetailDB('TIMBRESPIECE','','',QQ,false);
  end;
  ferme (QQ);
end;

procedure ChargeLesTimbres (TOBPiece,TOBTimbres : TOB);
var Indice : Integer;
		TOBT : TOB;
begin
  TOBTimbres.ClearDetail;
	if GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT') <> 'VEN' then Exit;
  //
  for Indice := 0 to TOBTIMBRESP.detail.count -1 do
  begin
    TOBT := TOB.Create('TIMBRESPIECE',TOBTimbres,-1);
    TOBT.PutValue('BT0_CODE',TOBTIMBRESP.detail[Indice].GetValue('BTP_CODE'));
    TOBT.PutValue('BT0_INDICE',indice);
  end;
end;

function  GetTotalTimbres (TOBTimbres,TOBPiece,TOBEChes : TOB) : Double;
var indice : Integer;
begin
  Result := 0;
  for Indice := 0 to TOBTimbres.detail.Count -1 do
  begin
    Result := Result + TOBTimbres.detail[Indice].GetValue('BT0_VALEUR');
  end;
end;

procedure ValideLesTimbres(TOBPiece, TOBTimbres : TOB);
var Indice : Integer;
		Nature,Souche : string;
    Numero,IndiceG : Integer;
    okok : boolean;
begin
  if TOBTimbres.detail.Count =0 then Exit;

  Nature := TOBPiece.GetValue('GP_NATUREPIECEG');
  Souche := TOBPiece.GetValue('GP_SOUCHE');
  Numero := TOBPiece.GetValue('GP_NUMERO');
  IndiceG := TOBPiece.GetValue('GP_INDICEG');
  for Indice := 0 to TOBTimbres.detail.Count -1 do
  begin
    TOBTimbres.detail[Indice].PutValue('BT0_NATUREPIECEG', Nature);
    TOBTimbres.detail[Indice].PutValue('BT0_SOUCHE', Souche);
    TOBTimbres.detail[Indice].PutValue('BT0_NUMERO', Numero);
    TOBTimbres.detail[Indice].PutValue('BT0_INDICEG', IndiceG);
  end;
  okok := false;
  TRY
    okok :=  TOBTimbres.InsertDB(nil);
  EXCEPT
    on E: Exception do
    begin
      PgiError('Erreur SQL : ' + E.Message, 'Erreur mise à jour des Timbres');
    end;
  END;
  if not okok then 
  begin
    V_PGI.IoError := oeUnknown;
  end;

end;

end.
