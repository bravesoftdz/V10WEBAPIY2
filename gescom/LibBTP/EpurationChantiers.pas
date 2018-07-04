unit EpurationChantiers;

interface
Uses Classes,
     Windows,
     sysutils,
     ComCtrls,
     StdCtrls,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fe_Main,
{$ELSE}
     MainEagl,
{$ENDIF}
     HCtrls,
     Hpanel,
     HEnt1,
     HMsgBox,
     Vierge,
     EntGc,
     BTPUtil,
     ExtCtrls,
     SaisUtil,
     UtilPgi,
     UtilEpurePiece,
     Forms,
     UTOB;

type
	TEpurationChantier = class (TObject)
  	private
    	fTOBChantier : TOB;
      fchantier : string;
      fresult : integer;
      fAvecControle : boolean;
      ffiche : TForm;
      function ControleChantierSupp : boolean;
      procedure TraiteChantier;
    	procedure SetChantier(const Value: TOB);
    public
      property Chantier : TOB read fTOBChantier write SetChantier;
      property result : integer read fresult;
      property Fiche : TForm read ffiche write ffiche;
      property AvecControle : Boolean read fAvecControle write fAvecControle;
      procedure lancetraitement;
  end;

function TraitementEpurationChantier (TOBChantier: TOB;AvecControle : boolean; Ecran : Tform) : integer;

implementation
uses CalcOLEGenericBTP;

function TraitementEpurationChantier (TOBChantier: TOB;AvecControle : boolean; Ecran : Tform) : integer;
var EpurationChantier : TEpurationChantier;
begin
   EpurationChantier := TEpurationChantier.create;
   EpurationChantier.Chantier := TOBChantier;
   EpurationChantier.AvecControle := AvecControle;
   EpurationChantier.Fiche := Ecran;
   TRY
     EpurationChantier.lanceTraitement;
   FINALLY
     result := EpurationChantier.Result;
   	 EpurationChantier.free;
   END;
end;

{ TEpurationChantier }

function TEpurationChantier.ControleChantierSupp: boolean;
begin
  if Not AvecControle then
  begin
    result := true;
    exit;
  end;
  // Verification que les clotures de facturation ont bien été effectuées
  result := not (ExisteSql ('SELECT AFF_ETATAFFAIRE FROM PIECE LEFT JOIN AFFAIRE ON AFF_AFFAIRE=GP_AFFAIREDEVIS '+
                 'WHERE GP_NATUREPIECEG="DBT" AND GP_AFFAIRE="'+fchantier+'" AND '+
                 'AFF_ETATAFFAIRE <> "TER"'));
end;


procedure TEpurationChantier.lancetraitement;
begin
	if ControleChantierSupp then
  begin
    TraiteChantier;
  end else
  begin
  	fresult := 1;
  end;

end;

procedure TEpurationChantier.SetChantier(const Value: TOB);
begin
  fTOBChantier := Value;
  fChantier := fTOBChantier.getValue('AFF_AFFAIRE');
end;

procedure TEpurationChantier.TraiteChantier;
var TOBPiece,TOBP : TOB;
		Indice : integer;
    LblEncours : TLabel;
begin
  LblEncours := nil;
  if ffiche <> nil then
  begin
    if ffiche.FindComponent('LblEncours')<> nil then
    begin
     	LblEncours := TLabel(ffiche.FindComponent('LblEncours'));
      LblEncours.Visible := True;
      TAnimate(FFiche.FindComponent('Animate1')).Visible := false;
    end;
  end;
	TOBPiece := TOB.Create ('LES PIECES',nil,-1);
  TOBPiece.LoadDetailDBFromSQL('PIECE','SELECT GP_NATUREPIECEG,GP_SOUCHE,GP_NUMERO,GP_INDICEG,GP_AFFAIRE,GP_AFFAIREDEVIS FROM PIECE WHERE GP_AFFAIRE="'+fCHantier+'"',false,true);
  For Indice := 0 to TOBPiece.detail.count-1 do
  begin
    TOBP := TOBPiece.detail[indice];
    if LblEncours <> nil then
    begin
      LblEncours.Caption := 'Chantier '+ BTPCodeAffaireAffiche(fchantier)+' '+RechDom('GCNATUREPIECEG',TOBP.GetString('GP_NATUREPIECEG'),false)+' N° : '+InttOStr(TOBP.GetInteger('GP_NUMERO'));
      ffiche.Refresh;
    end;
    EpureLaPiece (TOBP);
    if LblEncours <> nil then
    begin
      LblEncours.Caption := '';
      ffiche.Refresh;
    end;
    if V_PGI.Ioerror <> OeOk then break;
  end;
  if V_PGI.ioError = OeOk then
  begin
  	ExecuteSql ('DELETE FROM BDETETUDE WHERE BDE_AFFAIRE="'+fchantier+'"');
  	ExecuteSql ('DELETE FROM BSITUATIONS WHERE BST_AFFAIRE="'+fchantier+'"');
  	ExecuteSql ('DELETE FROM BSTRDOC WHERE BSD_AFFAIRE="'+fchantier+'"');
  	ExecuteSql ('DELETE FROM BTMEMOFACTURE WHERE BMF_AFFAIRE="'+fchantier+'"');
  	ExecuteSql ('DELETE FROM CONSOMMATIONS WHERE BCO_AFFAIRE="'+fchantier+'"');
  	ExecuteSql ('DELETE FROM PHASESCHANTIER WHERE BPC_AFFAIRE="'+fchantier+'"');
  	ExecuteSql ('DELETE FROM AFFAIREFRSGEST WHERE BAF_AFFAIRE="'+fchantier+'"');
  	ExecuteSql ('DELETE FROM AFFAIREINTERV WHERE BAI_AFFAIRE="'+fchantier+'"');
  	ExecuteSql ('DELETE FROM BTRESTEADEP WHERE RAD_AFFAIRE="'+fchantier+'"');
  	ExecuteSql ('DELETE FROM AFFAIRE WHERE AFF_AFFAIRE="'+fchantier+'"');
  end;
end;

end.
