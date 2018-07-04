unit ULiquidTva2014;

interface
uses
  SysUtils, Classes,
	UTOB,
  HEnt1,
  EntGC,
  Aglinit,
  HMsgBox,
  Facture
  ;

//
//
function TraitementAutoliquidTva (TOBL , TOBOuvrages : TOB) : boolean;
function IncidenceDateLiquidTva (TOBPiece,TOBOuvrages : TOB) : boolean;
function IsAutoLiquidationTva (TOBPiece : TOB) : boolean;
function IsAutoLiquidationTvaST (Sstraitant : string) : boolean;
function GetTvaST (Sstraitant : string) : string;

implementation
uses AffaireUtil,FactOuvrage;

function GetTvaST (Sstraitant : string) : string;
var TOBI : TOB;
begin
	Result := '';
  if TheParamTVA = nil then
  begin
  	PgiInfo ('ERREUR : Document Non affecté pour calcul ST');
    Exit;
  end;
	if (Sstraitant <> '') then
  begin
    if TheParamTva.TOBSSTrait <> nil then
    begin
      TOBI := TheParamTva.TOBSSTrait.FindFirst(['BPI_TIERSFOU'],[Sstraitant],true);
      result := TOBI.GetString('BPI_FAMILLETAXE');
    end;
  end;
end;

function IsAutoLiquidationTvaST (Sstraitant : string) : boolean;
var TOBI : TOB;
begin
	Result := false;
  if TheParamTVA = nil then
  begin
  	PgiInfo ('ERREUR : Document Non affecté pour calcul ST');
    Exit;
  end;
	if (Sstraitant <> '') then
  begin
    if TheParamTva.TOBSSTrait <> nil then
    begin
      TOBI := TheParamTva.TOBSSTrait.FindFirst(['BPI_TIERSFOU'],[Sstraitant],true);
      (* ----------
      if TOBI <> nil then
      begin
        if TOBI.GetValue('BPI_DATECONTRAT') >= StrToDate('01/01/2014') then Result := True;
      end;
      ----------- *)
      result := TOBI.GetBoolean('BPI_AUTOLIQUID');
    end;
  end;

end;


function IsAutoLiquidationTva (TOBPiece : TOB) : boolean;
var TOBA : TOB;
begin
	Result := false;

  // ---- Ooooops la boulette -- l'automatisme sur l'autoliquidation de tva ne s'effectue que sur des pièces de vente ---
  if GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT') <> 'VEN' then exit;

	if (TOBPiece.GetValue('GP_DATEPIECE')>= StrToDate('01/01/2014')) and (TOBPiece.GetString('GP_AFFAIRE')<>'') then
  begin
		TOBA := FindCetteAffaire (TOBPiece.GetString('GP_AFFAIRE'));
    if TOBA = nil then Exit;
  	if (TOBA.GeTValue('AFF_SSTRAITANCE')='X') and (TOBA.GeTValue('AFF_DATESSTRAIT')>=StrToDate('01/01/2014')) then
    	Result := true;
  end;
end;

function TraitementAutoliquidTva (TOBL , TOBOuvrages : TOB) : boolean;
var TOBA : TOB;
		TOBPiece : TOB;
    IndiceL : Integer;
begin
  Result := false;
  TOBPiece := TOBL.Parent;
  IndiceL := TOBL.GetIndex;
  // ----
  // ---- Ooooops la boulette -- l'automatisme sur l'autoliquidation de tva ne s'effectue que sur des pièces de vente ---
  if GetInfoParPiece(TOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT') <> 'VEN' then exit;
  // ----------------------------
	if (TOBPIECE.GetValue('GP_DATEPIECE')>= StrToDate('01/01/2014')) and (TOBL.GetString('GL_AFFAIRE')<>'') then
  begin
		TOBA := FindCetteAffaire (TOBL.GetString('GL_AFFAIRE'));
    if TOBA = nil then Exit;
  	if (TOBA.GeTValue('AFF_SSTRAITANCE')='X') and (TOBA.GeTValue('AFF_DATESSTRAIT')>=StrToDate('01/01/2014')) then
    begin
      {
			TOBL.SetString('GL_FAMILLETAXE1',VH_GC.AutoLiquiTVAST);
			TOBL.SetString('GL_FAMILLETAXE3','');
			TOBL.SetString('GL_FAMILLETAXE4','');
			TOBL.SetString('GL_FAMILLETAXE5','');
      if (TOBL.Getvalue('GL_TYPEARTICLE') = 'OUV') or (TOBL.GetValue('GL_TYPEARTICLE') = 'ARP') then
      begin
      	AppliqueChangeTaxeOuv(TOBPiece, TOBOuvrages, IndiceL + 1, TOBL.GetValue('GL_FAMILLETAXE1'),
        											TOBL.GetValue('GL_FAMILLETAXE2'),TOBL.GetValue('GL_FAMILLETAXE3'),
                              TOBL.GetValue('GL_FAMILLETAXE4'),TOBL.GetValue('GL_FAMILLETAXE5'));
      end;
      }
      TOBPIECE.setBoolean('GP_AUTOLIQUID',true);
      Result := True;
    end;
  end;
end;

function IncidenceDateLiquidTva (TOBPiece,TOBOuvrages : TOB) : boolean;
var i : Integer;
begin
  Result := false;
	for i := 0 to TOBPiece.Detail.count -1 do
  begin
    if TOBPiece.detail[I].GetString('GL_TYPELIGNE')<>'ART' Then continue;
    if TraitementAutoliquidTva (TOBPiece.detail[I],TOBOuvrages) then Result := true;
  end;
end;

end.
