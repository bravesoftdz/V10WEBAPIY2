unit ChangeVersions;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
	ComCtrls, StdCtrls, Hctrls, ExtCtrls, Hgauge,UTob,EntGc, UtilPGI,MajTable,uEntCommun,
  TntStdCtrls;

type

	TMajParam = class
  public
    LanceMajRecalcPrio    : boolean;
		LanceMajPieceLiv 			: boolean;
		LanceMajTvaMillieme 	: boolean;
    LanceMajParamSoc 			: boolean;
		LanceMajReconduction 	: boolean;
	  LanceMajCleTelephone 	: boolean;
		OKConsommations 			: boolean;
    MajLigneCOmplOk 			: boolean;
    FamilleOuvOk 					: boolean;
    Fv12120               : Boolean;
    LanceSupEtats					: Boolean;
    LanceMasquerNatures		: Boolean;
    GestionEvtminutes		  : Boolean;
    DateFACTAFFok   		  : Boolean;
    OkMAJCodeBarre        : Boolean;
    LanceMajConsosDepuisAchat     : boolean;
    LanceMajLienExcel     : boolean;
    UpdateEdition7CEGID : Boolean;
    ReajusteUA : Boolean;
    MajCodetarifs : Boolean;
    PartageFamillesTarifs : Boolean;
    Situationreajuste : Boolean;
    constructor create;
  end;

  TFMajVerBtp = class(TForm)
    Titre: THLabel;
    PB: TEnhancedGauge;
  private
  	param : TMajParam;
    procedure MajVerBTP664;
    procedure MajVerBTP667;
    procedure TraitementMajPiece;
  	procedure TraitementTvaMillieme;
  	procedure TraitementReconduction;
    procedure MajCleTelephone;
    Procedure MajCleTelAdresse;
    procedure MajCleTelContact;
    procedure TraitementParamSoc;
    procedure TraitementMajTel;
    function TraiteChangeGamme : integer;
    procedure OoopsConsommations;
    procedure MajLigneCompl;
    procedure MajFamilleOuv;
    procedure CorrectionFV12120;
    procedure SuppressionEtats;
    procedure MasquerNatures;
    procedure LanceRecalcPrio;
    procedure GestionEvtMinutes;
    function TraiteDateFactAff : integer;
    procedure LanceMajConsosDepuisAchat;
    procedure lanceMajLienXls;
    procedure lanceUpdateEdition7CEGID;
    procedure reajusteDocAchat;
    procedure reajusteTarifs;
    procedure AjoutPartageFamillesTarifs;
    procedure ReajusteSituationAnterieure;
    procedure MAJTableCodeBarre;
  end;

{$IFNDEF EAGLCLIENT}
function TraiteChangementVersions (VersionBase : String) : integer;
function TraiteChangementGamme : integer;
function RecupVersionbase : String;
{$ENDIF}
Procedure ChangeTabletteCompta;

    function VerificationCodeBarre: Boolean;
implementation
{$R *.DFM}

uses HEnt1,
{$IFNDEF EAGLCLIENT}
	DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} MajStruc, HQry,DBCtrls,
{$ENDIF}
  ParamSoc,
  Ent1, HmsgBox,UtilSoc,uBTPVerrouilleDossier,TarifUtil,UtilTOBPiece,BTRechargeFamTarif; (*,STKMoulinette, yTarifsBascule,Wjetons  ;*)

Function RecupVersionBase : String;
var QQ : Tquery;
begin
  result := '';
  QQ := OpenSql ('SELECT BTV_VERSIONBASEB FROM BTMAJSTRUCTURES WHERE BTV_IDBTP=0',True,-1,'',true);
  if not QQ.eof then
  begin
    result := QQ.Fields[0].AsString;
  end;
  ferme (QQ);
end;

function TraitementAlancer (P : TMajparam) : boolean;
begin

	result :=(P.LanceMajRecalcPrio) or
           (P.LanceMajPieceLiv) or
  				 (P.LanceMajTvaMillieme) or
           (P.LanceMajReconduction)  or
           (P.LanceMajParamSoc) or
           (not P.LanceMajCleTelephone) or
           (not P.MajLigneCOmplOk) or
           (not P.OKConsommations) or
           (not P.FamilleOuvOk) or
           (P.LanceSupEtats) Or
           (P.LanceMasquerNatures) Or
           (P.Fv12120) or
           (not P.GestionEvtminutes) or
           (P.LanceMajConsosDepuisAchat) or
           (P.LanceMajLienExcel) or
           (not P.DateFACTAFFok) or
           (P.UpdateEdition7CEGID) or
           (P.ReajusteUA) or
           (P.PartageFamillesTarifs) or
           (P.Situationreajuste);
end;


function TestChangeVersion ( P : TMajParam; VersionBase : String): boolean;
var QQ : Tquery;
		Nbr : integer;
begin
  QQ := OpenSql ('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_RECALCPRIO" AND SOC_DATA="X"',true);
  P.LanceMajRecalcPrio := QQ.eof;
  Ferme (QQ);
  
  (* ---- ohh putain --
  QQ := OpenSql ('SELECT GP_NATUREPIECEG FROM PIECE WHERE GP_NATUREPIECEG="BLC"',true,-1,'',true);
  P.LanceMajPieceLiv := not QQ.eof;
  ferme(QQ);
  -------- *)

  QQ := OpenSql ('SELECT BPT_CATEGORIETAXE FROM PARAMTAXE',true,-1,'',true);
  P.LanceMajTvaMillieme := QQ.eof;
  Ferme (QQ);

  QQ := OpenSql ('SELECT BRE_CODE FROM BRECONDUCTION',true,-1,'',true);
  P.LanceMajReconduction := QQ.eof;
  Ferme (QQ);

  QQ := OpenSql ('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_METIER" AND SOC_DATA="PGE"',true,-1,'',true);
  P.LanceMajParamSoc := QQ.eof;
  Ferme (QQ);

	P.OKConsommations := true;
  QQ := OpenSql ('SELECT COUNT(*) AS NBR FROM CONSOMMATIONS WHERE BCO_AFFAIRE0=""',True,-1,'',true);
  if not QQ.eof then
  begin
    Nbr := QQ.findField('NBR').AsInteger;
    if Nbr > 0 then P.OKConsommations := false;
  end;
  ferme (QQ);

  P.MajLigneCOmplOk  := GetparamSoc('SO_MAJLIGCOMPL');
  P.LanceMajCleTelephone := GetparamSoc('SO_FORMATTEL');
  //
  P.FamilleOuvOk := false;
  QQ := OpenSql ('SELECT COUNT(*) AS NBR FROM CHOIXCOD WHERE CC_TYPE="BO1"',True,-1,'',true);
  if not QQ.eof then
  begin
    Nbr := QQ.findField('NBR').AsInteger;
    if Nbr > 0 then P.FamilleOuvOk := true;
  end;
  ferme (QQ);
  //
  P.FV12120 := false;
  QQ := OpenSql ('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_FV12120" AND SOC_DATA="OUI"', True);
  P.FV12120 := QQ.eof;
  ferme (QQ);
  //
  QQ := OpenSql ('SELECT MO_CODE FROM MODELES WHERE MO_TYPE="E" AND MO_NATURE="GPJ" AND MO_CODE="AAC"',true);
  P.LanceSupEtats := not QQ.eof;
  ferme(QQ);
  //
  QQ := OpenSql ('SELECT GPP_NATUREPIECEG FROM PARPIECE WHERE GPP_NATUREPIECEG="ACV" AND GPP_MASQUERNATURE="-"',true);
  P.LanceMasquerNatures := not QQ.eof;
  ferme(QQ);
  //
  QQ := OpenSql ('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_GESTIONEVTMIN" AND SOC_DATA="X"', True);
  P.GestionEvtminutes := not QQ.eof;
  ferme (QQ);

  QQ := OpenSql ('SELECT AFA_NUMECHE FROM FACTAFF WHERE AFA_DATEDEBUTFAC='+DateToStr(Idate1900),true);
  P.DateFACTAFFok := QQ.eof;
  ferme(QQ);

  P.LanceMajLienExcel  := false;
  QQ := OpenSql ('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_BTMAJLIENXLS" AND SOC_DATA="X"', True);
  P.LanceMajLienExcel := QQ.eof;
  ferme (QQ);


  P.LanceMajConsosDepuisAchat := false;
  if (VersionBase > '850.Z5') and (VersionBase < '850.ZK') then P.LanceMajConsosDepuisAchat := true;


  P.UpdateEdition7CEGID := false;
  QQ := OpenSql ('SELECT DH_NOMCHAMP FROM DECHAMPS WHERE DH_NOMCHAMP="PSP_ORDRE" AND DH_CONTROLE=""', True);
  P.UpdateEdition7CEGID := QQ.eof;
  ferme (QQ);

  if (VersionBase < '998.ZR') then P.ReajusteUA := true;

  P.MajCodetarifs := false;
  if ExisteSQL('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="TAR"') then P.MajCodetarifs := True;;

  P.PartageFamillesTarifs := false;
  if ExisteSQL('SELECT * FROM DESHARE') then P.PartageFamillesTarifs := True;

  P.Situationreajuste := False;
  if ExisteSQL('SELECT 1 FROM LIGNE WHERE GL_NATUREPIECEG="FBT" and SUBSTRING(GL_PIECEORIGINE,10,3)="ETU"') then P.Situationreajuste := true;

  P.OkMAJCodeBarre := VerificationCodeBarre;

end;

Procedure TFMajVerBtp.TraitementMajPiece;
Begin
  if not param.LanceMajPieceLiv then exit;
	Titre.Caption := 'Mise à jour sur pièces';
  Titre.visible := true;
  Titre.Refresh;
  PB.Visible := true;
  PB.MinValue := 0;
  PB.MaxValue  := 16;
  self.Refresh;
  ExecuteSQL('UPDATE PIECE SET GP_NATUREPIECEG = "LBT", GP_SOUCHE = "LBT" WHERE GP_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE PIEDBASE SET GPB_NATUREPIECEG = "LBT", GPB_SOUCHE = "LBT" WHERE GPB_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE PIECEADRESSE SET GPA_NATUREPIECEG = "LBT", GPA_SOUCHE = "LBT" WHERE GPA_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE PIECERG SET PRG_NATUREPIECEG = "LBT", PRG_SOUCHE = "LBT" WHERE PRG_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE PIEDBASERG SET PBR_NATUREPIECEG = "LBT", PBR_SOUCHE = "LBT" WHERE PBR_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE PIEDECHE SET GPE_NATUREPIECEG = "LBT", GPE_SOUCHE = "LBT" WHERE GPE_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE PIEDPORT SET GPT_NATUREPIECEG = "LBT", GPT_SOUCHE = "LBT" WHERE GPT_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE LIGNE SET GL_NATUREPIECEG = "LBT", GL_SOUCHE = "LBT" WHERE GL_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE LIGNECOMPL SET GLC_NATUREPIECEG = "LBT", GLC_SOUCHE = "LBT" WHERE GLC_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE LIGNEFORMULE SET GLF_NATUREPIECEG = "LBT", GLF_SOUCHE = "LBT" WHERE GLF_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE LIGNELOT SET GLL_NATUREPIECEG = "LBT", GLL_SOUCHE = "LBT" WHERE GLL_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE LIGNENOMEN SET GLN_NATUREPIECEG = "LBT", GLN_SOUCHE = "LBT" WHERE GLN_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE LIGNEOUV SET BLO_NATUREPIECEG = "LBT", BLO_SOUCHE = "LBT" WHERE BLO_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE LIGNESERIE SET GLS_NATUREPIECEG = "LBT", GLS_SOUCHE = "LBT" WHERE GLS_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE ACOMPTES SET GAC_NATUREPIECEG = "LBT", GAC_SOUCHE = "LBT" WHERE GAC_NATUREPIECEG="BLC"');
  PB.Progress := PB.Progress + 1;
  ExecuteSQL('UPDATE SOUCHE SET SH_NUMDEPART = (SELECT SH_NUMDEPART FROM SOUCHE WHERE SH_TYPE="GES" AND SH_SOUCHE="GLC") WHERE SH_TYPE="GES" AND SH_SOUCHE="LBT"');
  PB.Progress := PB.Progress + 1;
  PB.visible := false;
  Titre.visible := false;
end;

Procedure TFMajVerBtp.MajVerBTP664;
Begin
  TraitementMajPiece;
  TraitementTvaMillieme;
  TraitementReconduction;
  TraitementParamSoc;
End;

Procedure TFMajVerBtp.MajVerBTP667;
Begin
  MajCleTelephone;
  MajCleTelAdresse;
  MajCleTelContact;
  TraitementMajTel;
End;


function VerifChangeGame : boolean;
var Gamme : string;
begin
  Result := false;
//	Gamme := GetparamSocSecur('SO_BTGAMME','');

//uniquement en line
//  Result := (Gamme <> 'S1');
//
//  Result := (Gamme <> 'S3');
//
end;

{$IFNDEF EAGLCLIENT}
function TraiteChangementGamme : integer;
var XX : TFMajVerBtp;
begin
	result := 0;

  if VerifChangeGame then
  begin
    XX := TFMajVerBtp.create (application);
    result := XX.TraiteChangeGamme;
    XX.free;
  end;
end;
{$ENDIF}

Procedure ChangeTabletteCompta;
Var StReq : Integer;
Begin
//uniquement en line
{*
//Modification de la Table DESCOMBO pour gestion des modifications
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=0 WHERE DO_COMBO="TZGCOLLCLIENT"     AND DO_TAGMODIF<>0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=0 WHERE DO_COMBO="TZGCOLLFOURN"      AND DO_TAGMODIF<>0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=0 WHERE DO_COMBO="TZGCOLLTOUTDEBIT"  AND DO_TAGMODIF<>0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=0 WHERE DO_COMBO="TZGCOLLTOUTCREDIT" AND DO_TAGMODIF<>0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=0 WHERE DO_COMBO="TZGPRODUIT" 				AND DO_TAGMODIF<>0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=0 WHERE DO_COMBO="TZGCHARGE" 				AND DO_TAGMODIF<>0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=0 WHERE DO_COMBO="TZGENERAL" 				AND DO_TAGMODIF<>0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=1 WHERE DO_COMBO="TZGCOLLCLIENT"     AND DO_TAGMODIF=0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=1 WHERE DO_COMBO="TZGCOLLFOURN"      AND DO_TAGMODIF=0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=1 WHERE DO_COMBO="TZGCOLLTOUTDEBIT"  AND DO_TAGMODIF=0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=1 WHERE DO_COMBO="TZGCOLLTOUTCREDIT" AND DO_TAGMODIF=0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=1 WHERE DO_COMBO="TZGPRODUIT" 				AND DO_TAGMODIF=0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=1 WHERE DO_COMBO="TZGCHARGE" 				AND DO_TAGMODIF=0');
StReq := ExecuteSql('UPDATE DECOMBOS SET DO_TAGMODIF=1 WHERE DO_COMBO="TZGENERAL" 				AND DO_TAGMODIF=0');
//
AvertirTable('TZGCOLLCLIENT');
AvertirTable('TZGCOLLFOURN');
AvertirTable('TZGCOLLTOUTDEBIT');
AvertirTable('TZGCOLLTOUTCREDIT');
AvertirTable('TZGPRODUIT');
AvertirTable('TZGCHARGE');
AvertirTable('TZGENERAL');
	*}

//
end;

{$IFNDEF EAGLCLIENT}
function TraiteChangementVersions (VersionBase : String) : integer;
var XX : TFMajVerBtp;
		P : TMajParam;
begin
	result := 0;
	P := TMajParam.Create;
  if ISVerrouille Then
  begin
    PGIInfo ('Un traitement de mise à jour est en cours..#13#10Merci de vous reconnecter ultérieurement');
    result := -1 ;
    exit;
  end;
  TRY
    TestChangeVersion (P, VersionBase);
    if TraitementAlancer (P)then
    begin
      XX := TFMajVerBtp.create (application);
      XX.param := P;
      XX.Show;
      TRY
      XX.LanceRecalcPrio;
      XX.MajVerBTP664;
      XX.MajVerBTP667;
      XX.OoopsConsommations;
      XX.MajLigneCompl;
      XX.MajFamilleOuv;
      XX.CorrectionFV12120;
      XX.MasquerNatures;
      XX.SuppressionEtats; // a faire impérativement après MasquerNatures
      XX.GestionEvtMinutes;
      result := XX.TraiteDateFactAff;
      XX.LanceMajConsosDepuisAchat;
      XX.lanceMajLienXls;
      XX.lanceUpdateEdition7CEGID;
      XX.reajusteDocAchat;
      XX.reajusteTarifs;
      XX.AjoutPartageFamillesTarifs;
      XX.ReajusteSituationAnterieure;
      XX.MAJTableCodeBarre;
      FINALLY
        XX.free;
      END;
    end;
  FINALLY
    P.free;
    Deverrouille;
  END;
end;
{$ENDIF}

procedure TFMajVerBtp.ReajusteSituationAnterieure;
begin
  if Not param.Situationreajuste then Exit;
  (*
  Titre.Caption := 'Réajustement situations antérieure';
  Titre.visible := true;
  Titre.Refresh;
  PB.Visible := true;
  PB.MinValue := 0;
  PB.MaxValue  := 1;
  self.Refresh;
	ExecuteSQL('UPDATE LIGNE SET GL_PIECEORIGINE=GL_PIECEPRECEDENTE WHERE GL_NATUREPIECEG="FBT" AND SUBSTRING(GL_PIECEORIGINE,10,3)="ETU"');
  PB.Progress := PB.Progress + 1;
  PB.visible := false;
  Titre.visible := false;
  *)
end;

procedure TFMajVerBtp.OoopsConsommations;
begin
  if param.OKConsommations then exit;
  Titre.Caption := 'Mise à jour des consommations';
  Titre.visible := true;
  Titre.Refresh;
  PB.Visible := true;
  PB.MinValue := 0;
  PB.MaxValue  := 1;
  self.Refresh;
	ExecuteSQL('UPDATE CONSOMMATIONS SET BCO_AFFAIRE0=SUBSTRING(BCO_AFFAIRE,1,1) WHERE BCO_AFFAIRE0=""');
  PB.Progress := PB.Progress + 1;
  PB.visible := false;
  Titre.visible := false;
end;

procedure TFMajVerBtp.TraitementTvaMillieme;
var LesParamTaxes,Latablette,UneTaxe,Unparam : TOB;
		SQL : string;
    QQ : TQuery;
    Indice : integer;
begin
  if not param.LanceMajTvaMillieme then exit;
  lesParamTaxes := TOB.Create ('LES PARAM TAXES',nil,-1);
  laTablette := TOB.Create ('LES COMMUNS',nil,-1);
  Titre.Caption := 'Mise à jour du paramétrage des Taxes';
  Titre.visible := true;
  Titre.Refresh;
  PB.Visible := true;
  PB.MinValue := 0;
  PB.MaxValue  := 1;
  self.Refresh;
  TRY
    Sql := 'SELECT * FROM CHOIXCOD WHERE CC_TYPE="GCX"';
    QQ := OpenSql (Sql,True,-1,'',true);
    LaTablette.LoadDetailDB ('COMMUN','','',QQ,false);
    ferme (QQ);
    if Latablette.detail.count > 0 then
    begin
      PB.Progress := PB.Progress + 1;
    	for Indice := 0 to Latablette.detail.count -1 do
      begin
        UneTaxe := LaTablette.detail[Indice];
        UnParam := TOB.create ('PARAMTAXE',LesParamtaxes,-1);
        UnParam.putvalue('BPT_CATEGORIETAXE',UneTaxe.getValue('CC_CODE'));
        if UneTaxe.getValue('CC_CODE') = 'TX1' then UnParam.putValue('BPT_TYPETAXE','TVA') else
        if UneTaxe.getValue('CC_CODE') = 'TX2' then UnParam.putValue('BPT_TYPETAXE','TPF') else
        																						UnParam.putValue('BPT_TYPETAXE','');
        UnParam.putvalue('BPT_LIBELLE',UneTaxe.getValue('CC_LIBELLE'));
      end;
      lesParamTaxes.InsertDB (nil);
    end;
  FINALLY
    PB.visible := false;
    Titre.visible := false;
  	lesParamTaxes.free;
    laTablette.free;
    GetParamTaxe;
  END;
end;

procedure TFMajVerBtp.TraitementReconduction;
var LesParamTaxes,Latablette,UneTaxe,Unparam : TOB;
	SQL : string;
    QQ : TQuery;
    Indice : integer;
begin
  if not param.LanceMajReconduction then exit;
  lesParamTaxes := TOB.Create ('LES RECONDUCTIONS',nil,-1);
  laTablette := TOB.Create ('LES COMMUNS',nil,-1);
  Titre.Caption := 'Mise à jour du paramétrage des Types de Reconduction';
  Titre.visible := true;
  Titre.Refresh;
  PB.Visible := true;
  PB.MinValue := 0;
  PB.MaxValue  := 1;
  self.Refresh;
  TRY
    Sql := 'SELECT * FROM COMMUN WHERE CO_TYPE="RCA"';
    QQ := OpenSql (Sql,True,-1,'',true);
    LaTablette.LoadDetailDB ('COMMUN','','',QQ,false);
    ferme (QQ);
    if Latablette.detail.count > 0 then
    begin
      PB.Progress := PB.Progress + 1;
    	self.Refresh;
    	for Indice := 0 to Latablette.detail.count -1 do
      begin
        UneTaxe := LaTablette.detail[Indice];
        UnParam := TOB.create ('BRECONDUCTION',LesParamtaxes,-1);
        UnParam.putvalue('BRE_CODE',UneTaxe.getValue('CO_CODE'));
        UnParam.putValue('BPT_TYPETAXE','');
        UnParam.putvalue('BRE_LIBELLE',UneTaxe.getValue('CO_LIBELLE'));
        UnParam.putvalue('BRE_TYPEACTION',UneTaxe.getValue(''));
				UnParam.putvalue('BRE_VAL',UneTaxe.getValue('X'));
      end;
      lesParamTaxes.InsertDB (nil);
    end;
  FINALLY
    PB.visible := false;
    Titre.visible := false;
  	lesParamTaxes.free;
    laTablette.free;
    GetParamTaxe;
  END;
end;

//Maj de la zone T_CleTELEPHONE dans la table Tiers. Concaténation de T_TELEPHONE
Procedure TFMajVerBtp.MajCleTelephone;
Var TobCli : Tob;
		QQ     : TQuery;
		Indice : integer;
    i      : Integer;
    NoTel  : String;
    Car    : String;
    Cletel : String;
Begin

	if param.LanceMajCleTelephone then exit;

  QQ := OpenSQL('SELECT * FROM TIERS WHERE T_NATUREAUXI="CLI" AND T_TELEPHONE <> ""', false,-1,'',true);

  TobCli := TOB.Create('lesTiers', nil, -1);
  Tobcli.LoadDetailDB('TIERS', '', '', QQ,True);

  Ferme(QQ);

  Titre.Caption := 'Mise à Jour du Téléphone des Tiers';
  Titre.visible := true;
  Titre.Refresh;
  PB.Visible := true;
  PB.MinValue := 0;
  PB.MaxValue  := TOBCli.detail.count;
  self.Refresh;
	TRY
    if TobCli.detail.count > 0 then
    Begin
      for Indice := 0 to TobCli.detail.count -1 do
      Begin
        PB.Progress := PB.Progress + 1;
  			self.Refresh;
        //Chargement du N° de téléphone sans Car. Spéciaux
        if TobCli.Detail[Indice].GetValue('T_TELEPHONE') <> '' then
        Begin
          Notel := TobCli.Detail[Indice].GetValue('T_TELEPHONE');
          TobCli.Detail[Indice].PutValue('T_CLETELEPHONE',CleTelephone(Notel));
        End;
        if TobCli.Detail[Indice].GetValue('T_FAX') <> '' then
        Begin
          Notel := TobCli.Detail[Indice].GetValue('T_FAX');
          TobCli.Detail[Indice].PutValue('T_CLEFAX',CleTelephone(Notel));
        End;
        if TobCli.Detail[Indice].GetValue('T_TELEPHONE2') <> '' then
        Begin
          Notel := TobCli.Detail[Indice].GetValue('T_TELEPHONE2');
          TobCli.Detail[Indice].PutValue('T_CLETELEPHONE2',CleTelephone(Notel));
        End;
        if TobCli.Detail[Indice].GetValue('T_TELEX') <> '' then
        Begin
          Notel := TobCli.Detail[Indice].GetValue('T_TELEX');
          TobCli.Detail[Indice].PutValue('T_CLETELEPHONE3',CleTelephone(Notel));
        End;
      end;
      TobCli.SetAllModifie(true);
      TobCli.InsertOrUpdateDB(False);
    end;
  FINALLY
    TobCli.Free;
    PB.visible := false;
    Titre.visible := false;
  END;

end;

Procedure TFMajVerBtp.MajCleTelAdresse;
Var TobAdr,TOBADRB : Tob;
		QQ,QQ1     : TQuery;
		Indice : integer;
    i      : Integer;
    NoTel  : String;
    Car    : String;
    Cletel : String;
Begin

	if param.LanceMajCleTelephone then exit;

  QQ := OpenSQL('SELECT * FROM ADRESSES WHERE ADR_TELEPHONE <> ""', false,-1,'',true);

  TobAdr := TOB.Create('lesAdresses', nil, -1);
  TOBADRB := TOB.Create('BADRESSES', nil, -1);
  TobAdr.LoadDetailDB('ADRESSES', '', '', QQ,True);

  Ferme(QQ);

  Titre.Caption := 'Mise à Jour du Téléphone des Adresses';
  Titre.visible := true;
  Titre.Refresh;
  PB.Visible := true;
  PB.MinValue := 0;
  PB.MaxValue  := TobAdr.detail.count;
  self.Refresh;
	TRY
    if TobAdr.detail.count > 0 then
    Begin
      for Indice := 0 to TobAdr.detail.count -1 do
      Begin
        PB.Progress := PB.Progress + 1;
        //Chargement du N° de téléphone sans Car. Spéciaux
        if TobAdr.Detail[Indice].GetValue('ADR_TELEPHONE') <> '' then
        begin
          Notel := TobAdr.Detail[Indice].GetValue('ADr_TELEPHONE');
          QQ := OpenSql ('SELECT * FROM BADRESSES WHERE BA0_TYPEADRESSE="'+TobAdr.Detail[Indice].getString('ADR_TYPEADRESSE')+'" AND BA0_NUMEROADRESSE='+TobAdr.Detail[Indice].getString('ADR_NUMEROADRESSE'),true,1,'',true);
          if not QQ.eof then
          begin
            TOBADRB.SelectDB('',QQ);
          end else
          begin
            TOBADRB.InitValeurs(false);
            TOBADRB.SetString('BA0_TYPEADRESSE',TOBADR.detail[Indice].getString('ADR_TYPEADRESSE'));
            TOBADRB.SetInteger('BA0_NUMEROADRESSE',TOBADR.detail[Indice].getInteger('ADR_NUMEROADRESSE'));
          end;
          ferme (QQ);
          //
          TobAdrB.SetAllModifie(true);
          TobAdrB.PutValue('BA0_CLETELEPHONE',CleTelephone(Notel));
          TOBADRB.InsertOrUpdateDB(false); 
        end;
      end;
    end;
  FINALLY
    Titre.visible := false;
    PB.Visible := false;
    TobAdr.Free;
    TobAdrB.Free;
  END;


end;

Procedure TFMajVerBtp.MajCleTelContact;
Var TobContact : Tob;
		QQ    		 : TQuery;
		Indice 		 : integer;
    i      		 : Integer;
    NoTel  		 : String;
    Car    		 : String;
    Cletel 		 : String;
    req				 : String;
Begin

	if param.LanceMajCleTelephone then exit;

  QQ := OpenSQL('SELECT * FROM CONTACT WHERE C_TELEPHONE <> ""', false,-1,'',true);

  TobContact := TOB.Create('lesContacts', nil, -1);
  TobContact.LoadDetailDB('CONTACT', '', '', QQ,True);

  Ferme(QQ);

  Titre.Caption := 'Mise à Jour du Téléphone des Contacts';
  Titre.visible := true;
  Titre.Refresh;
  PB.Visible := true;
  PB.MinValue := 0;
  PB.MaxValue  := TobContact.detail.count;
  self.Refresh;
	TRY
    if TobContact.detail.count > 0 then
       Begin
     	 for Indice := 0 to TobContact.detail.count -1 do
           Begin
           PB.Progress := PB.Progress + 1;
           //Chargement du N° de téléphone sans Car. Spéciaux
		       if TobContact.Detail[Indice].GetValue('C_TELEPHONE') <> '' then
    	        begin
              Notel := TobContact.Detail[Indice].GetValue('C_TELEPHONE');
	            TobContact.Detail[Indice].PutValue('C_CLETELEPHONE',CleTelephone(Notel));
       		    end;
		       if TobContact.Detail[Indice].GetValue('C_FAX') <> '' then
    	        begin
              Notel := TobContact.Detail[Indice].GetValue('C_FAX');
	            TobContact.Detail[Indice].PutValue('C_CLEFAX',CleTelephone(Notel));
       		    end;
		       if TobContact.Detail[Indice].GetValue('C_TELEX') <> '' then
    	        begin
              Notel := TobContact.Detail[Indice].GetValue('C_TELEX');
	            TobContact.Detail[Indice].PutValue('C_CLETELEX',CleTelephone(Notel));
       		    end;
           end;
       TobContact.SetAllModifie(true);
  		 TobContact.InsertOrUpdateDB(False);
       end;
  FINALLY
    TobContact.Free;
    PB.visible := false;
    Titre.visible := false;
  END;

end;

{
procedure TFMajVerBtp.SetSoucheLignePhase;
begin
  ExecuteSQL('UPDATE LIGNEPHASES SET BLP_SOUCHE = (SELECT GPP_SOUCHE FROM PARPIECE WHERE GPP_NATUREPIECEG=LIGNEPHASES.BLP_NATUREPIECEG) WHERE BLP_SOUCHE IS NULL');
end;
}
procedure TFMajVerBtp.TraitementParamSoc;
begin
	ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="PGE" WHERE SOC_NOM="SO_METIER"');
	ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_GCTOXCONFIRM"');
end;

procedure TFMajVerBtp.TraitementMajTel;
begin
	ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_FORMATTEL"');
end;

function TFMajVerBtp.TraiteChangeGamme : integer;
begin
	result := 0;
  Titre.Caption := 'Mise à niveau en gamme de produit';
  Titre.visible := true;
  Titre.Refresh;
  PB.Visible := true;
  PB.MinValue := 0;
  PB.MaxValue  := 1;
  self.Refresh;
  ChangeTabletteCompta;
//uniquement en line
{*
  ExecuteSQL ('UPDATE MENU SET MN_ACCESGRP="----------------------------------------------------------------------------------------------------" where mn_1=26 and mn_2=1 and mn_4=0 and (mn_3=1 or mn_3=3 or mn_3=6 or mn_3=8)');
  ExecuteSQL ('UPDATE PARPIECE SET GPP_REFINTEXT="INT" WHERE GPP_NATUREPIECEG IN ("DBT","FBT","ABT")');
  //
  ExecuteSQL ('UPDATE MENU SET MN_LIBELLE="Devis - Factures" WHERE MN_1=0 AND MN_2=145');
  ExecuteSQL ('UPDATE MENU SET MN_LIBELLE="Fiches" WHERE MN_1=145 AND MN_2=1 AND MN_3=0');
  ExecuteSQL ('UPDATE MENU SET MN_LIBELLE="Chantier" WHERE MN_1=149 AND MN_2=1 AND MN_3=0');
  SetparamSoc('SO_BTGAMME','S1');
  result := 1;
  ExecuteSQL ('UPDATE MENU SET MN_LIBELLE="Etudes et situations" WHERE MN_1=0 AND MN_2=145');
  ExecuteSQL ('UPDATE MENU SET MN_LIBELLE="Fiche Affaire" WHERE MN_1=145 AND MN_2=1 AND MN_3=0');
  ExecuteSQL ('UPDATE MENU SET MN_LIBELLE="Affaire" WHERE MN_1=149 AND MN_2=1 AND MN_3=0');
  SetparamSoc('SO_BTGAMME','S3');
  result := 1;
*}

  Titre.visible := false;
  PB.Visible := false;
  Titre.Refresh;
end;

procedure TFMajVerBtp.MajLigneCompl;
var TheChaine : string;
begin
	if param.MajLigneCOmplOk then exit;
  Titre.Caption := 'Traitement en cours ...';
  Titre.visible := true;
  Titre.Refresh;
  PB.Visible := true;
  PB.MinValue := 0;
  PB.MaxValue  := 2;
  self.Refresh;
  TheChaine := 'DELETE FROM LIGNECOMPL WHERE GLC_NATUREPIECEG="DBT" AND GLC_NUMERO IN '+
  						 '(SELECT GP_NUMERO from piece where GP_NATUREPIECEG="DBT" and '+
               '(SELECT count (DISTINCT GL_NUMORDRE) FROM LIGNE WHERE GL_NATUREPIECEG="DBT" AND GL_NUMERO=GP_NUMERO) <> '+
               '(SELECT count (DISTINCT GL_NUMLIGNE) FROM LIGNE WHERE GL_NATUREPIECEG="DBT" AND GL_NUMERO=GP_NUMERO))';

  ExecuteSql (TheChaine);
  PB.Progress := PB.Progress + 1;
  self.refresh;

	TheChaine := 'UPDATE LIGNE SET GL_NUMORDRE=GL_NUMLIGNE WHERE GL_NATUREPIECEG="DBT" AND GL_NUMERO IN '+
  						 '(SELECT GP_NUMERO from piece where GP_NATUREPIECEG="DBT" and '+
               '(SELECT count (DISTINCT GL_NUMORDRE) FROM LIGNE WHERE GL_NATUREPIECEG="DBT" AND GL_NUMERO=GP_NUMERO) <> '+
               '(SELECT count (DISTINCT GL_NUMLIGNE) FROM LIGNE WHERE GL_NATUREPIECEG="DBT" AND GL_NUMERO=GP_NUMERO))';

  ExecuteSql (TheChaine);
  PB.Progress := PB.Progress + 1;
  self.refresh;

	ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_MAJLIGCOMPL"');
  Titre.visible := false;
  PB.Visible := false;
  Titre.Refresh;

end;

{ TMajParam }

constructor TMajParam.create;
begin
  LanceMajPieceLiv := false;
  LanceMajTvaMillieme := false;
  LanceMajParamSoc := false;
  LanceMajReconduction := false;
  LanceMajCleTelephone := true;
  OKConsommations := true;
  MajLigneCOmplOk := true;
  FamilleOuvOk := true;
  Fv12120 := False;
  LanceSupEtats := false;
  LanceMasquerNatures := false;
	DateFACTAFFok := false;
end;

procedure TFMajVerBtp.MajFamilleOuv;
var QQ : TQuery;
    TOBFamilles : TOB;
    indice : integer;
begin
  // Correction 032;12594
  if not IsMasterOnShare('GCFAMILLENIV1') then exit;
  // ---
  if param.FamilleOuvOk then exit;

  TOBFamilles := TOB.Create ('LES FAMILLES',nil,-1);
  QQ := OpenSql ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="FN1"',true,-1,'',true);
  TOBFamilles.LoadDetailDB ('CHOIXCOD','','',QQ,false);
  ferme (QQ);
  for Indice := 0 to TOBfamilles.detail.count -1 do
  begin
    TOBFamilles.detail[Indice].putvalue('CC_TYPE','BO1');
  end;
  if TOBFamilles.detail.count > 0 then TOBFamilles.InsertDB (nil,true);
  TOBfamilles.ClearDetail;

  QQ := OpenSql ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="FN2"',true,-1,'',true);
  TOBFamilles.LoadDetailDB ('CHOIXCOD','','',QQ,false);
  ferme (QQ);
  for Indice := 0 to TOBfamilles.detail.count -1 do
  begin
    TOBFamilles.detail[Indice].putvalue('CC_TYPE','BO2');
  end;
  if TOBFamilles.detail.count > 0 then TOBFamilles.InsertDB (nil,true);
  TOBfamilles.ClearDetail;
  QQ := OpenSql ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="FN3"',true,-1,'',true);
  TOBFamilles.LoadDetailDB ('CHOIXCOD','','',QQ,true);
  ferme (QQ);
  for Indice := 0 to TOBfamilles.detail.count -1 do
  begin
    TOBFamilles.detail[Indice].putvalue('CC_TYPE','BO3');
  end;
  if TOBFamilles.detail.count > 0 then TOBFamilles.InsertDB (nil,true);
  TOBFamilles.Free;
end;

procedure TFMajVerBtp.CorrectionFV12120;
Var TobAppels : Tob;
    TobPiece	: Tob;
    TobLigne	: Tob;
    FactHT		: String;
    Affaire		: String;
    Tiers			: String;
    NatPiece	: String;
    Souche		: String;
    Numero		: String;
    Req				: String;
    QQ				: TQuery;
    Indice 		: Integer;
    Ind				: Integer;
    I					: Integer;
begin

	if Not param.Fv12120 then exit;

  TobAppels := Tob.Create('LESAFFAIRES', Nil, -1);
  TobPiece  := Tob.Create('LESPIECES', Nil, -1);
  TobLigne  := Tob.Create('LESLIGNES', Nil, -1);

  //Lecture des affaires de type Appels
  QQ := OpenSQL ('Select * From AFFAIRE Where AFF_AFFAIRE0="W"', True);
  TobAppels.LoadDetailDB('AFFAIRE', '', '', QQ, True);
  Ferme (QQ);

  For Indice := 0 to TobAppels.Detail.Count -1 do
  		Begin
      Affaire := TobAppels.Detail[Indice].GetValue('AFF_AFFAIRE');
      Tiers		:= TobAppels.Detail[Indice].GetValue('AFF_TIERS');
  	  //Lecture du tiers associé à l'appel
      QQ := OpenSQL ('Select T_FACTUREHT FROM TIERS WHERE T_TIERS="' + Tiers + '"', True);
      If Not QQ.Eof Then FactHT := QQ.FindField('T_FACTUREHT').AsString;
      Ferme(QQ);
  	  //Lecture des Entêtes de pièces
      QQ := OpenSQL ('SELECT GP_FACTUREHT, GP_NATUREPIECEG, GP_SOUCHE, GP_NUMERO ' +
                                            'FROM PIECE WHERE ' +
                                            'GP_AFFAIRE="' + Affaire + '" AND GP_TIERS="' + Tiers + '"', True);
      TobPiece.LoadDetailDB('PIECE', '','', QQ, True);
      Ferme(QQ);
      For Ind := 0 To TobPiece.Detail.Count -1 Do
      		Begin
          NatPiece := TobPiece.Detail[Ind].GetValue('GP_NATUREPIECEG');
          Souche   := TobPiece.Detail[Ind].GetValue('GP_SOUCHE');
          Numero	 := TobPiece.Detail[Ind].GetValue('GP_NUMERO');
          //lecture des lignes de Pièces
          if FactHT <> TobPiece.Detail[Ind].GetValue('GP_FACTUREHT') then
             Begin
             QQ := OpenSQL ('SELECT GL_FACTUREHT FROM LIGNE WHERE GL_NATUREPIECEG="' + NatPiece + '"' +
                            ' AND GL_SOUCHE="' + Souche +  '"' +
                            ' AND GL_NUMERO="' + Numero +  '"' +
                            ' AND GL_AFFAIRE="' + Affaire + '"' +
                            ' AND GL_TIERS="' + Tiers + '"', True);
             TobLigne.LoadDetailDB('LIGNE', '','', QQ, True);
             Ferme(QQ);
             For I := 0 To TobLigne.Detail.Count -1 Do
                 Begin
                 TobLigne.detail[I].PutValue('GL_FACTUREHT', FactHT);
                 End;
             TobPiece.detail[Ind].PutValue('GP_FACTUREHT', FactHT);
             End;
          End;
      //Mise à jour de l'Appel
			TobAppels.Detail[Indice].PutValue('AFF_AFFAIREHT', FactHT);
	    end;

   TobAppels.UpdateDB(true);
   TobPiece.UpdateDB(true);
   TobLigne.UpdateDB(true);

  //Mise à jour du paramètre
	ExecuteSQL ('UPDATE PARAMSOC SET SOC_DATA="OUI" WHERE SOC_NOM="SO_FV12120"');
  Titre.visible := false;
  PB.Visible := false;

  TobAppels.Free;
  TobPiece.Free;
  TobLigne.Free;

  Titre.Refresh;

end;

procedure TFMajVerBtp.MasquerNatures;
var nb : Integer;
begin                                                         
	if Not param.LanceMasquerNatures then exit;

	Nb:=ExecuteSQL ('UPDATE PARPIECE SET GPP_MASQUERNATURE="X" WHERE GPP_NATUREPIECEG IN'+
                  '("ACV","AFP","ALF","ALV","APR","APV","AVS","BRC","BSA","BSP","CCE","CCH","CCR","CMC","CMF","CSA",'+
                  '"CSP","FCF","FFA","FFO","FRA","FRF","FSI","LCE","LCR","OTR","PAF","PRE","PRF","PRO","RCC","RCF","RPR","SSE","TRV")');
end;

procedure TFMajVerBtp.SuppressionEtats;
var nb : Integer;
begin
	if Not param.LanceSupEtats then exit;

	Nb:=ExecuteSQL ('delete from modeles where mo_type="E" and mo_nature="GPJ" '+
                  'and mo_code not like "Z%" '+
                  'and mo_code not in ("B01","B02","B03","BAT","BCF","BCI","BDP","BLB","BMN","BON","CFB","DBM",'+
                  '"DBT","DDT","DGD","DGR","RFB","SI0","SIA","SIR","SIT","SMR","SMT","SR0","ST0","STM") '+
                  'and mo_code not in (select gpp_impetat from parpiece where gpp_impetat=mo_code and gpp_masquernature="-") '+
                  'and mo_code not in (select gpp_impmodele from parpiece where gpp_impmodele=mo_code and gpp_masquernature="-")');
end;

procedure TFMajVerBtp.LanceRecalcPrio;
var lastKey : integer;
    TOBDet : TOB;
    Sql : string;
    QQ : TQUery;
    Indice : integer;
begin
  if not param.LanceMajRecalcPrio then exit;
  Titre.Caption := 'Traitement en cours ...';
  Titre.visible := true;
  Titre.Refresh;
  lastkey := 0;
  TOBDet := TOB.Create ('LES TARIFS',nil,-1);
  repeat
    TOBDet.clearDetail;
    Sql := 'SELECT TOP 1000 * FROM TARIF ';
    if lastKey <>0 THEN Sql := Sql + ' WHERE GF_TARIF > '+IntToStr(lastkey);
    Sql := Sql + ' ORDER BY GF_TARIF';
    QQ := OpenSql (Sql,True,-1,'',true);
    if not QQ.eof then TOBDet.LoadDetailDB ('TARIF','','',QQ,false);
    if TOBDet.detail.count > 0 then
    begin
      PB.Visible := true;
      PB.MinValue := 0;
      PB.MaxValue  := 1000;
      PB.Progress := 0;
      for Indice := 0 to TOBDet.detail.count -1 do
      begin
        CalcPriorite (TOBDet.detail[Indice]);
        lastKey := TOBDet.detail[Indice].GetValue('GF_TARIF');
        TOBDet.detail[Indice].SetAllModifie (true);
        TOBDet.detail[Indice].UpdateDB ;
        PB.Progress := PB.Progress + 1;
        self.refresh;
      end;
    end;
  until TOBDet.detail.count = 0;
  ExecuteSql ('UPDATE PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_RECALCPRIO"');
  Titre.visible := false;
  PB.Visible := false;
  Titre.Refresh;
end;

procedure TFMajVerBtp.GestionEvtMinutes;
begin
  if param.GestionEvtminutes then exit; // deja traité
  Titre.Caption := 'Traitement en cours ...';
  Titre.visible := true;
  Titre.Refresh;
  ExecuteSql ('UPDATE BTEVENPLAN SET BEP_DUREE = ROUND(BEP_DUREE * 60,0)');
  ExecuteSql ('UPDATE PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_GESTIONEVTMIN"');
  Titre.visible := false;
  PB.Visible := false;
  Titre.Refresh;
end;

function TFMajVerBtp.TraiteDateFactAff : integer;
begin
	result := 0;
  if param.DateFACTAFFok then exit; // deja traité
  PgiInfo ('Merci de relancer l''application#13#10'+
  				 'Afin d''effectuer la "Réactualisation des échéances de contrats" dans le menu "Traitements"');
  result := -1;
end;

// traitement de maj des consos depuis achats depuis le 1er Janvier 2011
// suite bug FQ 15594 : maj des LBT non remisées
procedure TFMajVerBtp.LanceMajConsosDepuisAchat;
var lastKey : integer;
    TOBDet : TOB;
    Sql : string;
    QQ : TQUery;
    Indice : integer;
begin
  if not param.LanceMajConsosDepuisAchat then exit;
  Titre.Caption := 'Traitement en cours, phase 1 ...';
  Titre.visible := true;
  Titre.Refresh;

  // 1-	Requête maj depuis BLF :
  Sql :=  'UPDATE C1 SET C1.BCO_DPA = C2.BCO_DPA, C1.BCO_DPR = C2.BCO_DPR, C1.BCO_PUHT = C2.BCO_PUHT,'+
          'C1.BCO_MONTANTACH = C2.BCO_MONTANTACH, C1.BCO_MONTANTPR= C2.BCO_MONTANTPR, C1.BCO_MONTANTHT = C2.BCO_MONTANTHT '+
          'FROM CONSOMMATIONS AS C1 JOIN CONSOMMATIONS AS C2 ON C1.BCO_NUMMOUV = C2.BCO_LIENVENTE '+
          'WHERE C1.BCO_DATEMOUV>="01/01/2011" AND C1.BCO_NATUREPIECEG="LBT" AND C2.BCO_NATUREPIECEG="BLF"';
  ExecuteSql (Sql);

  Titre.Caption := 'Traitement en cours, phase 2 ...';
  Titre.Refresh;
  // 2-	Requête maj lignes FF d'indice 1 à partir de la ligne d'indice 0 :
  Sql :=  'UPDATE C1 SET C1.BCO_DPA = C2.BCO_DPA, C1.BCO_DPR = C2.BCO_DPR, C1.BCO_PUHT = C2.BCO_PUHT,'+
          'C1.BCO_MONTANTACH = C2.BCO_MONTANTACH, C1.BCO_MONTANTPR= C2.BCO_MONTANTPR, C1.BCO_MONTANTHT = C2.BCO_MONTANTHT '+
          'FROM CONSOMMATIONS AS C1 JOIN CONSOMMATIONS AS C2 ON C1.BCO_NUMMOUV = C2.BCO_NUMMOUV '+
          'WHERE C1.BCO_DATEMOUV>="01/01/2011" AND C1.BCO_NATUREPIECEG="FF" AND C1.BCO_INDICE=1 AND C2.BCO_INDICE=0';
  ExecuteSql (Sql);

  Titre.Caption := 'Traitement en cours, phase 3 ...';
  Titre.Refresh;
  // 3-	Requête maj depuis FF :
  Sql :=  'UPDATE C1 SET C1.BCO_DPA = C2.BCO_DPA, C1.BCO_DPR = C2.BCO_DPR, C1.BCO_PUHT = C2.BCO_PUHT,'+
          'C1.BCO_MONTANTACH = C2.BCO_MONTANTACH, C1.BCO_MONTANTPR= C2.BCO_MONTANTPR, C1.BCO_MONTANTHT = C2.BCO_MONTANTHT '+
          'FROM CONSOMMATIONS AS C1 JOIN CONSOMMATIONS AS C2 ON C1.BCO_NUMMOUV = C2.BCO_LIENVENTE '+
          'WHERE C1.BCO_DATEMOUV>="01/01/2011" AND C1.BCO_NATUREPIECEG="LBT" AND C2.BCO_NATUREPIECEG="FF"';
  ExecuteSql (Sql);

  Titre.visible := false;
  PB.Visible := false;
  Titre.Refresh;
end;

procedure TFMajVerBtp.lanceMajLienXls;
var QQ: TQuery;
		TOBDet : TOB;
    cledoc : r_cledoc;
    Sql : string;
    Indice : integer;
begin
  if not param.LanceMajLienExcel then exit;
  Titre.Caption := 'Mise à jour des liens Excel (Appels d''offres) ...';
  Titre.visible := true;
  Titre.Refresh;
  TOBDet := TOB.Create ('LES LIEN EXCELS',nil,-1);
  QQ := OpenSql ('SELECT * FROM BLIENEXCEL',true,-1,'',true);
  TOBDet.LoadDetailDB('BLIENEXCEL','','',QQ,false);
  ferme (QQ);
  TRY
    PB.Visible := true;
    PB.MinValue := 0;
    PB.MaxValue  := TOBdet.detail.count -1;
    for Indice := 0 to TOBdet.detail.count -1 do
    begin
      DecodeRefPiece(TOBDet.detail[Indice].getValue('BLX_PIECEASSOCIEE'),cledoc);
      SQl := 'UPDATE BLIENEXCEL SET BLX_NUMLIGNE=(SELECT GL_NUMORDRE FROM LIGNE WHERE '+
             WherePiece(cledoc,ttdligne,false) + ' AND GL_NUMLIGNE=BLX_NUMLIGNE) '+
             'WHERE BLX_PIECEASSOCIEE="'+
             TOBDet.detail[Indice].getValue('BLX_PIECEASSOCIEE')+'" AND BLX_NUMLIGNE='+
             inttostr(TOBDet.detail[Indice].getValue('BLX_NUMLIGNE'));
      ExecuteSql (SQL);
  		PB.Progress := PB.Progress + 1;
      PB.Update;
    end;
  FINALLY
    ExecuteSql ('UPDATE PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM="SO_BTMAJLIENXLS"');
  END;
end;

procedure TFMajVerBtp.lanceUpdateEdition7CEGID;
var QQ: TQuery;
		TOBDet : TOB;
    cledoc : r_cledoc;
    Sql : string;
    Indice : integer;
begin
  if not param.LanceMajLienExcel then exit;
  Titre.Caption := 'Mise à jour Edition 7 CEGID PAIE-COMPTA ...';
  Titre.visible := true;
  Titre.Refresh;
  TRY
    PB.Visible := true;
    PB.MinValue := 0;
    PB.MaxValue  := 1;
    //
    ExecuteSql ('UPDATE DECHAMPS SET DH_CONTROLE="L" WHERE DH_PREFIXE="PSP" AND DH_NOMCHAMP="PSP_ORDRE" AND DH_CONTROLE=""');
    PB.Progress := PB.Progress + 1;
    PB.Update;
    //
  FINALLY
    // Rien a faire
  END;
end;

procedure TFMajVerBtp.reajusteDocAchat;
begin
  if not param.ReajusteUA then Exit;
  Titre.Caption := 'Réajustements des documents d''achats ...';
  Titre.visible := true;
  Titre.Refresh;
  PB.Visible := true;
  PB.MinValue := 0;
  PB.MaxValue  := 2;

	ExecuteSQL('UPDATE LIGNE set gl_qualifqteach=gl_qualifqtevte where gl_naturepieceg IN '+
  					 '("CF","BLF","FF","CFR","LFR","BFA","AFS") and gl_typeligne ="ART"');
  PB.Progress := PB.Progress + 1;
  PB.Update;
	ExecuteSQL('update ligne set gl_qualifqteVTE=('+
  					 'select GA_QUALIFUNITEVTE FROM ARTICLE WHERE GA_ARTICLE=GL_ARTICLE) where '+
             'gl_naturEpieceg IN ("CF","BLF","FF","CFR","LFR","BFA","AFS") and gl_typeligne ="ART"');
  PB.Progress := PB.Progress + 1;
  PB.Update;
end;

procedure TFMajVerBtp.reajusteTarifs;
begin
	InitialisationTARIF;
end;

procedure TFMajVerBtp.AjoutPartageFamillesTarifs;
Var StSQL       : String;

  function getBasePrincipale() : string;
  var Qbase : tQuery;
  begin
    result := V_PGI.SchemaName;
    Qbase := openSql('SELECT ##TOP 1## DS_NOMBASE FROM DESHARE',true);
    try
      if not Qbase.Eof then
        result := Qbase.FindField('DS_NOMBASE').asString;
    finally
      ferme(Qbase);
    end;
  end;
begin
  // Traitement si partage de données
  if ExisteSQL('SELECT * FROM DESHARE') then
  begin
    if Not ExisteSQL('SELECT * FROM DESHARE WHERE DS_NOMTABLE="BTFAMILLETARIF"') then
    begin
      StSQL := 'INSERT INTO DESHARE (DS_NOMTABLE, DS_MODEFONC, DS_NOMBASE, DS_TYPTABLE, DS_VUE) VALUES ("BTFAMILLETARIF", "LIB", "'+ getBasePrincipale +'","TAB", "")';
      ExecuteSQL(StSql);
    end;
    if Not ExisteSQL('SELECT * FROM DESHARE WHERE DS_NOMTABLE="BTSOUSFAMILLETARIF"') then
    begin
      StSQL := 'INSERT INTO DESHARE (DS_NOMTABLE, DS_MODEFONC, DS_NOMBASE, DS_TYPTABLE, DS_VUE) VALUES ("BTSOUSFAMILLETARIF", "LIB", "'+ getBasePrincipale +'","TAB", "")';
      ExecuteSQL(StSql);
    end;
  end;

end;

Function VerificationCodeBarre : Boolean;
Var StSQL : string;
    QQ    : TQuery;
begin

  Result := False;

  QQ := OpenSQL('SELECT * from BTCODEBARRE WHERE BCB_NATURECAB IN ("MAR","ARP")',False);
  if QQ.Eof then
  begin
    StSQL := 'SELECT Count(*) AS NBART FROM ARTICLE WHERE GA_CODEBARRE <> ""';
    QQ := OpenSQL(StSQL,False);
    if QQ.EOF then
      Result := False
    else
    begin
      if StrToInt(QQ.FindField('NBArt').AsString) > 0 then
        Result := True
      else
        Result := False;
    end;
    Ferme (QQ);
  end
  else
    Result := False;

  Ferme(QQ);

end;


procedure TFMajVerBtp.MAJTableCodeBarre;
var SQl           : String;
begin

  if Not param.OkMAJCodeBarre then Exit;
  Sql := 'INSERT INTO BTCODEBARRE (BCB_NATURECAB,BCB_IDENTIFCAB,BCB_CABPRINCIPAL,BCB_QUALIFCODEBARRE,BCB_CODEBARRE) '+
         'SELECT GA_TYPEARTICLE,GA_ARTICLE,"X",GA_QUALIFCODEBARRE,GA_CODEBARRE FROM ARTICLE WHERE GA_CODEBARRE <> ""';
  ExecuteSql (Sql);
end;



end.
