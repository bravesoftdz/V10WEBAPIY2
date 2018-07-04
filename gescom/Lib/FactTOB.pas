unit FactTOB ;

interface

uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, UTOB, UtilPGI, EntGC, SysUtils, Math, HCtrls,uEntCommun,ParamSoc
{$IFDEF EAGLCLIENT}
{$ELSE}
  ,Db,{$IFNDEF DBXPRESS} dbTables {$ELSE} uDbxDataSet {$ENDIF}
{$ENDIF}
{$IFDEF BTP}
     ,factTOBBTP,BTStructChampSup
{$ENDIF}
 ;


function  TOB2CleDoc(TOBG: TOB): R_CleDoc;
function  GetTOBLigne(TOBPiece: TOB; ARow: integer): TOB;
procedure InsertTOBLigne(TOBPiece: TOB; ARow: integer);
procedure AddLesSupLigne(TOBL: TOB; Soeur: Boolean; WithLigneFac : boolean=false);
procedure AddLesSupEntete(TOBE: TOB);
procedure AddLesSupEnteteLigneTarif(TobE: Tob);
procedure PutLesSupEnteteLigneTarif(TobPiece, TobLigneTarif, TobLigneTarif_O: Tob);
function  GetTobLigneTarifInTobPiece(TobPiece: Tob): Tob;
procedure PutTobPieceInTobLigneTarif(TobPiece, TobLigneTarif: Tob);
function  GetTobPieceInTobLigneTarif(TobLigneTarif: Tob): Tob;
procedure InitLesSupMontants(TOBL: TOB);
procedure InitLesSupLigne(TOBL: TOB; InitLigneCompl : boolean=true);
function NewTOBLigne(TOBPiece: TOB; ARow: integer; ForFacture : boolean=false): TOB;
procedure NewTOBLigneFille(TOBL: TOB);
procedure InitTobLigne(TobLigne: Tob);
function  GetChampLigne(TOBPiece: TOB; Champ: string; ARow: integer): Variant;
procedure InitTOBPiece(TOBP: TOB);
procedure InitLesCols;
function  FabricWhereNatArt(NaturePieceG, DomainePiece, SelectFourniss: string): string;
function  EstRempliGC(GS: THGrid; Lig: integer): boolean;
procedure DepileTOBLignes(GS: THGrid; TOBPiece: TOB; ARow, NewRow: integer);
procedure VideCodesLigne(TOBPiece: TOB; ARow: integer);
procedure PieceAjouteSousDetail(FTOBPiece: TOB; AvecInit : boolean=true; AvecInitLigneCompl : boolean= true; FromLoad : boolean=false);
procedure PutValueDetail(TOBPiece: TOB; ChampMaj: string; VV: Variant; TOBArticles:TOB=nil);
procedure AddLesSupLigneCompl(Tobl: Tob; Soeur: Boolean);
procedure InitLesSupLigneCompl(Tobl: Tob; Soeur: Boolean);
function  TypeDeChamp(NomChamp: String) : String;
procedure CopieChampsSup (TOBDest,TOBProv : TOB);
procedure CopieChamps (TOBDest,TOBProv : TOB);
procedure AddLesSupLigneFac(TOBL : TOB;Soeur : boolean);
procedure ProtectionCoef (TOBL : TOB);
procedure ReinitLigneFac (TOBL : TOB);


const
  SG_NL           : integer = 0  ;
  SG_ContreM      : integer = -1 ;
  SG_RefArt       : integer = -1 ;
  SG_Catalogue    : integer = -1 ;
  SG_Lib          : integer = 2  ;
  SG_QF           : integer = 3  ;
  SG_QS           : integer = -1 ;
  SG_Px           : integer = 5  ;
  SG_Rem          : integer = 6  ;
  SG_Aff          : integer = 7  ;
  SG_Rep          : integer = 8  ;
  SG_Dep          : integer = 9  ;
  SG_RV           : integer = 10 ;
  SG_RL           : integer = 11 ;
  SG_Montant      : integer = 12 ;
  SG_DateLiv      : integer = -1 ;
  SG_Total        : integer = -1 ;
  SG_QR           : integer = -1 ;
  SG_QReste       : integer = -1 ;
  SG_MtReste      : Integer = -1 ;
  SG_Motif        : integer = -1 ;
  SG_QA           : integer = -1 ;
  SG_Pct          : integer = -1 ;
  SG_Unit         : integer = -1 ;
  SG_TypA         : integer = -1 ;
  SG_Folio        : integer = -1 ;
  SG_DateProd     : integer = -1 ;
  SG_Regrpe       : integer = -1 ;
  SG_LibRegrpe    : integer = -1 ;
  SG_PxNet        : integer = -1 ;
  SG_TypeRem      : integer = -1 ;
  SG_Circuit      : integer = -1 ;
  SG_PXAch        : integer = -1 ;
  SG_MontantAch   : integer = -1 ;
  SG_REFTiers     : integer = -1 ;
  SG_DETAILBORD   : Integer = -1 ;
  SG_DEJAFACT     : integer = -1 ;
  SG_montantSit   : integer = -1 ;
  SG_MontantMarche: integer = -1 ;
  SG_QTEPREVUE    : integer = -1 ;
  SG_POURCENTAVANC: integer = -1 ;
  SG_QTESITUATION : integer = -1 ;
  SG_MTDEJAFACT   : integer = -1 ;
  SG_TEMPS        : integer = -1 ;
  SG_TEMPSTOT     : integer = -1;
  SG_QUALIFTEMPS  : integer = -1 ;
  SG_CODTVA1      : integer = -1 ;
  SG_NATURETRAVAIL: integer = -1 ;
  SG_FOURNISSEUR  : integer = -1 ;
  SG_VOIRDETAIL   : integer = -1 ;
  SG_REFCATALOGUE : integer = -1 ;
  SG_MTPRODUCTION : integer = -1 ;
  SG_MATERIEL     : integer = -1 ;
  SG_TYPEACTION   : integer = -1 ;
  SG_QTEAPLANIF   : Integer = -1;
  SG_COEFMARG : integer = -1 ;
  SG_POURMARG : integer = -1;
  SG_POURMARQ : integer = -1;
  SG_NUMAUTO : integer = -1;
  SG_QTESAIS : integer = -1;
  SG_RENDEMENT : integer = -1;
  SG_PERTE : integer = -1;
  SG_QUALIFHEURE : integer = -1;
  SG_QTECOND : integer = -1;
  SG_COEFCOND : integer = -1;
  SG_CODECOND : integer = -1;
  SG_CODEMARCHE : integer = -1;
  SG_TOTALTS : integer = -1;

implementation
uses
  FactTarifs,FactComm,FactDomaines,UspecifPoc,saisutil
  ,CbpMCD
  ,CbpEnumerator
  ;

procedure ReinitLigneFac (TOBL : TOB);
begin
  TOBL.SetString('GL_CODEARTICLE','');
  TOBL.SetString('GL_TYPENOMENC','');
  TOBL.SetString('GL_QUALIFQTEACH','');
  TOBL.SetString('GL_QUALIFQTEVTE','');
  TOBL.SetString('GL_QUALIFQTESTO','');
  TOBL.SetString('GL_TYPEARTICLE','');
  TOBL.SetString('GL_FAMILLENIV1','');
  TOBL.SetString('GL_FAMILLENIV2','');
  TOBL.SetString('GL_FAMILLENIV3','');
  TOBL.SetString('GL_FAMILLETAXE1','');
  TOBL.SetString('GL_FAMILLETAXE2','');
  TOBL.SetString('GL_FAMILLETAXE3','');
  TOBL.SetString('GL_FAMILLETAXE4','');
  TOBL.SetString('GL_FAMILLETAXE5','');
  TOBL.SetDouble('GL_DPR',0);
  TOBL.SetDouble('GL_DPA',0);
  TOBL.SetDouble('GL_PMAP',0);
  TOBL.SetDouble('GL_PMRP',0);
  TOBL.SetDouble('GL_PUHTDEV',0);
  //
  TOBL.SetDouble('GL_PUHT',0);
  TOBL.SetDouble('GL_PUHTNET',0);
  TOBL.SetDouble('GL_PUHTNETDEV',0);
  TOBL.SetDouble('GL_PUHTDEV',0);
  //
  TOBL.SetDouble('GL_PUTTC',0);
  TOBL.SetDouble('GL_PUTTCNET',0);
  TOBL.SetDouble('GL_PUTTCNETDEV',0);
  //
  TOBL.SetDouble('GL_PUHTBASE',0);
  TOBL.SetDouble('GL_PUTTCBASE',0);
  TOBL.SetDouble('GL_PUHTORIGINE',0);
  TOBL.SetDouble('GL_PUHTORIGINEDEV',0);
  //
  TOBL.SetDouble('GL_QTEFACT',0);
  TOBL.SetDouble('GL_QTESTOCK',0);
  TOBL.SetDouble('GL_QTERESTE',0);
  //
  TOBL.SetDouble('GL_MTRESTE',0);

end;

function TOB2CleDoc(TOBG: TOB): R_CleDoc;
var CleDoc: R_CleDoc;
  Pref: string;
begin
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  Result := CleDoc;
  if TOBG = nil then Exit;
  if TOBG.NomTable = 'PIECE' then Pref := 'GP' else
    if TOBG.NomTable = 'FACTURE' then Pref := 'GP' else
    if TOBG.NomTable = 'LIGNE' then Pref := 'GL' else
    if TOBG.NomTable = 'PIEDBASE' then Pref := 'GPB' else
    if TOBG.NomTable = 'PIEDECHE' then Pref := 'GPE' else
    if TOBG.NomTable = 'PIECERG' then Pref := 'PRG' else
    if TOBG.NomTable = 'PIEDBASERG' then Pref := 'PBR' else
    if TOBG.NomTable = 'LIGNEOUV' then Pref := 'BLO' else
    if TOBG.NomTable = 'LIGNEOUVPLAT' then Pref := 'BOP' else
    Exit;
  CleDoc.NaturePiece := TOBG.GetValue(Pref + '_NATUREPIECEG');
  CleDoc.Souche := TOBG.GetValue(Pref + '_SOUCHE');
  CleDoc.DatePiece := TOBG.GetValue(Pref + '_DATEPIECE');
  CleDoc.NumeroPiece := TOBG.GetValue(Pref + '_NUMERO');
  CleDoc.Indice := TOBG.GetValue(Pref + '_INDICEG');
  if Pref = 'GL' then
  begin
  	CleDoc.NumLigne  := TOBG.GetValue(Pref + '_NUMLIGNE');
  	CleDoc.NumOrdre  := TOBG.GetValue(Pref + '_NUMORDRE');
  end;
  Result := CleDoc;
end;

function GetTOBLigne(TOBPiece: TOB; ARow: integer): TOB;
begin
  Result := nil;
  if ((ARow <= 0) or (ARow > TOBPiece.Detail.Count)) then Exit;
  
  Result := TOBPiece.Detail[ARow - 1];
end;

procedure InsertTOBLigne(TOBPiece: TOB; ARow: integer);
begin
  NewTOBLigne(TOBPiece, ARow);
end;

procedure AddLesSupLigne(TOBL: TOB; Soeur: Boolean; WithLigneFac : boolean=false);
var prefixe : string;
begin
  prefixe := Getprefixetable(TOBL);
  // Nouvelle gestion des champm sup -- voir ADO
  AddlesChampsSupLigneEtl (TOBL);
  //
  TOBL.AddChampSup('RECALCULTARIF', Soeur);
  TOBL.AddChampSup('RECALCULCOMM', Soeur);
  TOBL.AddChampSup('TOTREMLIGNETTCDEV', Soeur);
  TOBL.AddChampSup('TOTREMLIGNETTC', Soeur);
  TOBL.AddChampSup('TOTREMPIEDTTCDEV', Soeur);
  TOBL.AddChampSup('TOTREMPIEDTTC', Soeur);
  TOBL.AddChampSup('TOTESCLIGNETTCDEV', Soeur);
  TOBL.AddChampSup('TOTESCLIGNETTC', Soeur);
  TOBL.AddChampSup('COMPTAAFFAIRE', Soeur);
  TOBL.AddChampSup('OUVREANAL', Soeur);
  TOBL.AddChampSup('QTEORIG', Soeur);
  TOBL.AddChampSup('MARGE', Soeur);
  TOBL.AddChampSup('ARTSLIES', Soeur);
  TOBL.AddChampSup('QTECHANGE', Soeur);
  TOBL.AddChampSup('_CONTREMARTREF', Soeur);
  TOBL.AddChampSup('SUPPRIME', Soeur);
  TOBL.AddChampSup('NUMBON', Soeur);
  // Code Barre
  if not TOBL.FieldExists('REGROUPE_CB') then TOBL.AddChampSupValeur ('REGROUPE_CB', '-');
  if not TOBL.FieldExists('UNI_OU_DIM') then TOBL.AddChampSupValeur ('UNI_OU_DIM', '');
  // -- Modif BTP
  TOBL.AddChampsup(prefixe+'_RECALCULER', Soeur);
  TOBL.AddChampsupValeur('INDICERG', 0, Soeur);
//  TOBL.AddChampSupValeur('TOTALACHAT', 0,Soeur);
//  TOBL.AddChampSupValeur('TOTALACHATNET', 0,Soeur);
//  TOBL.AddChampSupValeur('TOTALREVIENT',0, Soeur);
  TOBL.AddChampSup('FROMTARIF', Soeur);
  {$IFDEF BTP}
  AddLesSupLignesBtp (TOBL,Soeur);
  {$ENDIF}
end;

procedure AddLesSupEntete(TOBE: TOB);
Var Str : string;
begin
  if (not TOBE.FieldExists('AFF_GENERAUTO')) then TOBE.AddChampsupValeur('AFF_GENERAUTO', 'DIR', False);
  if GetParamSocSecur('SO_OKSIZERO',False) then
    Str := 'X'
  else
    Str := '-';
  if (not TOBE.FieldExists('AFF_OKSIZERO')) then TOBE.AddChampsupValeur('AFF_OKSIZERO',  Str, False);
  if (not TOBE.FieldExists('ETATDOC')) then TOBE.AddChampsupValeur('ETATDOC', 'ENC', False);
  if (not TOBE.FieldExists('_GERE_EN_STOCK')) then TOBE.AddChampsupValeur('_GERE_EN_STOCK', 'X', False);
  if (not TOBE.FieldExists('_BLOQUETARIF')) then TOBE.AddChampsupValeur('_BLOQUETARIF', '-', False);
  if (not TOBE.FieldExists('_NEWFRAIS_')) then TOBE.AddChampsupValeur('_NEWFRAIS_', '-', False);
  if (not TOBE.FieldExists('FAMILLETAXE1')) then
  begin
    TOBE.AddchampSupValeur('FAMILLETAXE1', '', False);
    TOBE.AddchampSupValeur('FAMILLETAXE2', '', False);
    TOBE.AddchampSupValeur('FAMILLETAXE3', '', False);
    TOBE.AddchampSupValeur('FAMILLETAXE4', '', False);
    TOBE.AddchampSupValeur('FAMILLETAXE5', '', False);
    TOBE.AddchampSupValeur('GP_RECALCULER', '-', False);
    TOBE.AddchampSupValeur('GP_TARIFSGROUPES', '-', False);
    TOBE.AddChampSUpValeur('AVANCSAISIE', 0, false);
    TOBE.AddChampsupValeur('AVANCPREC', 0, false);
  end;
  TOBE.AddChampsupValeur('RUPTMILLIEME', '', False);
  TOBE.AddChampsupValeur('COEFFGFORCE',0,false);
  AddLesSupEnteteLigneTarif(TobE);
// MODIF LS Pour affichage montant de situation
	TOBE.AddChampSupValeur('MONTANTSIT',0,false);
	TOBE.AddChampSupValeur('GP_RECALCACHAT','-',false);
	TOBE.AddChampSupValeur('ISDEJAFACT','-',false);
	TOBE.AddChampSupValeur('ESCREMMULTIPLE','-',false);
//
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 17/09/2003
Description .. : Crée le champ sup contenant le pointeur vers la
Suite ........ : TobLigneTarif
*****************************************************************}
procedure AddLesSupEnteteLigneTarif(TobE: Tob); 
begin
  if Assigned(TobE) then
  begin
    if not TobE.FieldExists('LIGNETARIF') then
      TobE.AddChampSupValeur('LIGNETARIF', null);
    if not TobE.FieldExists('LIGNETARIF_O') then
      TobE.AddChampSupValeur('LIGNETARIF_O', null);
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 17/09/2003
Description .. : Dépose les pointeurs sur les TobligneTarif et
Suite ........ : TobLigneTarif_O dans la TobPiece
*****************************************************************}
procedure PutLesSupEnteteLigneTarif(TobPiece, TobLigneTarif, TobLigneTarif_O: Tob);
begin
(*
  if Assigned(TobPiece) then
  begin
    AddLesSupEnteteLigneTarif(TobPiece);
    if Assigned(TobLigneTarif) then
      TobPiece.PutValue('LIGNETARIF', Longint(TobLigneTarif));
    if Assigned(TobLigneTarif_O) then
      TobPiece.PutValue('LIGNETARIF_O', Longint(TobLigneTarif_O));
  end;
*)
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 25/09/2003
Description .. : Récupère la TobLigneTarif de la TobPiece
*****************************************************************}
function GetTobLigneTarifInTobPiece(TobPiece: Tob): Tob;
begin
  if Assigned(TobPiece) and TobPiece.FieldExists('LIGNETARIF') and (NullToVide(TobPiece.G('LIGNETARIF')) <> '') then
    Result := Tob(LongInt(TobPiece.G('LIGNETARIF')))
  else
    Result := nil;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 30/09/2003
Description .. : Crée le champ sup contenant la référence à la tobPiece
Suite ........ : dans la TobLigneTarif
*****************************************************************}
procedure PutTobPieceInTobLigneTarif(TobPiece, TobLigneTarif: Tob);
begin
  if Assigned(TobPiece) and Assigned(TobLigneTarif) then
  begin
    if not TobLigneTarif.FieldExists('PIECE') then
      TobLigneTarif.AddChampSupValeur('PIECE', Longint(TobPiece))
    else
      TobLigneTarif.P('PIECE', Longint(TobPiece));
  end;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 30/09/2003
Description .. : Récupère la TobPiece dans la TobLigneTarif
*****************************************************************}
function GetTobPieceInTobLigneTarif(TobLigneTarif: Tob): Tob;
begin
  if Assigned(TobLigneTarif) and TobLigneTarif.FieldExists('PIECE') then
    Result := Tob(LongInt(TobLigneTarif.G('PIECE')))
  else
    Result := nil;
end;

function ConvertDateString (TOBDest,TOBProv : TOB ; NomChamp : String) : boolean;
var YY,MM,JJ : Word;
		DateStr : string;
begin
	result := true;
  DateStr := TOBProv.getValue(NomChamp);
  if Length (DateStr) <> 8 then BEGIN result := false; exit; END;
  TRY
    YY := StrtoInt(copy(DateStr,1,4));
    MM := StrToInt(copy(DateStr,5,2));
    JJ := StrToInt(copy(DateStr,7,2));
    TOBDest.putValue(nomchamp,EncodeDate (YY,MM,JJ));
  EXCEPT
  	on EConvertError do result := false;
  END;
end;

procedure ConvertTypeChamps (TOBDest,TOBProv : TOB ; NomChamp,TypeChamps : string);
begin
	TRY
    if Pos(TypeChamps,'DATE')>0 then
    begin
      if not IsValidDate(TOBProv.GetValue(NomChamp)) then
      begin
      	// le format est AAAAMMJJ ??
        if not ConvertDateString (TOBDest,TOBProv,NomChamp) then
        begin
        	TOBDest.putValue(NomChamp,IDate1900);
        end;
      end else
      begin
      	TOBDest.putValue(NomChamp,StrToDate(TOBProv.GetValue(NomChamp)));
      end;
    end else if Pos(TypeChamps,'DOUBLE,RATE,EXTENDED')>0 then
    begin
			TRY
        TOBDest.putValue(NomChamp,valeur(TOBProv.GetValue(NomChamp)));
      except
        TOBDest.putValue(NomChamp,0);
      end;
    end else if Pos(TypeChamps,'INTEGER,SMALLINT')>0 then
    begin
			TRY
      	TOBDest.putValue(NomChamp,StrToInt(TOBProv.GetValue(NomChamp)));
      except
        TOBDest.putValue(NomChamp,0);
      end;
    end else
    begin
      TOBDest.putValue(NomChamp,TOBProv.GetValue(NomChamp));
    end;
  EXCEPT
  	raise  Exception.create ('La conversion de type pour le champs '+NomChamp+' de type '+TypeChamps+' a échoué');
  END;
end;

procedure CopieChamps (TOBDest,TOBProv : TOB);
var
  TobLigne: Tob;
  i: Integer;
  NomChamp: String;
begin
  if Assigned(TobProv) then
  begin
    for i := 1 to TobProv.NbChamps do
    begin
      NomChamp := TobProv.GetNomChamp(i);
      if not TOBDest.fieldExists (NomChamp) then
      begin
        TOBDest.AddChampSup(NomChamp,false);
      end;
      TOBDest.putValue(NomChamp,TOBProv.GetValue(NomChamp));
    end;
  end;
end;

procedure CopieChampsSup (TOBDest,TOBProv : TOB);
var i,Itable,Indice : integer;
		NomChamp : string;
    prefixe : string;
    ConvertOk : boolean;
    Mcd : IMCDServiceCOM;
    Field     : IFieldCOM ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
	ConvertOk := false;
  i := 1000;
  NomChamp := TOBProv.GetNomChamp(i);
  while NomChamp <> '' do
  begin
  	if TOBDest.FieldExists (NomChamp) then
    begin
    	if (not VarIsNull(TOBProv.GetValue(NomChamp)) and (VarAsType(TOBProv.GetValue(nomChamp), varString) <> #0 )) then
      begin
        Field := Mcd.GetField(NomChamp);
        if Assigned(Field) then
        begin
          ConvertTypeChamps (TOBDest,TOBProv,NomChamp,Field.tipe);
          ConvertOk := true;
        end;
        if not ConvertOk then
        begin
    			TOBDest.putValue(NomChamp,TOBProv.GetValue(Nomchamp));
        end;
      end;
    end;
    inc(i);
  	NomChamp := TOBProv.GetNomChamp(i);
  end;
end;

procedure AddLesSupLigneFac(TOBL : TOB;Soeur : boolean);
var
  TobLigneFac: Tob;
  i: Integer;
  NomChamp: String;
  Value: Variant;
begin
  if Assigned(Tobl) then
  begin
    TobLigneFac := Tob.Create('LIGNEFAC', nil, -1);
    try
      for i := 1 to TobLigneFac.NbChamps do
      begin
        NomChamp := TobLigneFac.GetNomChamp(i);
        if (TypeDeChamp(NomChamp) = 'I') or (TypeDeChamp(NomChamp) = 'N') then
          Value := 0
        else if (TypeDeChamp(NomChamp) = 'D') then
          Value := iDate1900
        else if (TypeDeChamp(NomChamp) = 'B') then
          Value := '-'
        else
          Value := '';
        if not TOBL.fieldExists(TobLigneFac.GetNomChamp(i)) then
        begin
        	Tobl.AddChampSupValeur(TobLigneFac.GetNomChamp(i), Value, Soeur);
        end;
      end;
    finally
      TobLigneFac.Free;
    end;
  end;
end;

procedure AddLesSupLigneCompl(Tobl: Tob; Soeur: Boolean);
var
  TobLigneCompl: Tob;
  i: Integer;
  NomChamp: String;
  Value: Variant;
begin
  if Assigned(Tobl) then
  begin
    TobLigneCompl := Tob.Create('LIGNECOMPL', nil, -1);
    try
      for i := 1 to TobLigneCompl.NbChamps do
      begin
        NomChamp := TobLigneCompl.GetNomChamp(i);
        if (TypeDeChamp(NomChamp) = 'I') or (TypeDeChamp(NomChamp) = 'N') then
          Value := 0
        else if (TypeDeChamp(NomChamp) = 'D') then
          Value := iDate1900
        else if (TypeDeChamp(NomChamp) = 'B') then
          Value := '-'
        else
          Value := '';
        if not TOBL.fieldExists(TobLigneCompl.GetNomChamp(i)) then
        begin
        	Tobl.AddChampSupValeur(TobLigneCompl.GetNomChamp(i), Value, Soeur);
        end;
      end;
    finally
      TobLigneCompl.Free;
    end;
  end;
end;

procedure InitLesSupLigneCompl(Tobl: Tob; Soeur: Boolean);
var
  TobLigneCompl: Tob;
  i: Integer;
  NomChamp: String;
  Value: Variant;
begin
  if Assigned(Tobl) then
  begin
    TobLigneCompl := Tob.Create('LIGNECOMPL', nil, -1);
    try
      for i := 1 to TobLigneCompl.NbChamps do
      begin
        NomChamp := TobLigneCompl.GetNomChamp(i);
        if (Tobl.FieldExists(NomChamp)) then
        begin
          if (TypeDeChamp(NomChamp) = 'I') or (TypeDeChamp(NomChamp) = 'N') then
            Value := 0
          else if (TypeDeChamp(NomChamp) = 'D') then
            Value := iDate1900
          else if (TypeDeChamp(NomChamp) = 'B') then
            Value := '-'
          else
            Value := '';
          Tobl.PutValue(TobLigneCompl.GetNomChamp(i), Value);
        end;
      end;
    finally
      TobLigneCompl.Free;
    end;
  end;
end;

procedure InitLesSupMontants(TOBL: TOB);
begin
  TOBL.PutValue('TOTREMLIGNETTCDEV', 0);
  TOBL.PutValue('TOTREMLIGNETTC', 0);
  TOBL.PutValue('TOTREMPIEDTTCDEV', 0);
  TOBL.PutValue('TOTREMPIEDTTC', 0);
  TOBL.PutValue('TOTESCLIGNETTCDEV', 0);
  TOBL.PutValue('TOTESCLIGNETTC', 0);
end;

procedure InitLesSupLigne(TOBL: TOB; InitLigneCompl : boolean=true);
var prefixe : string;
begin
  prefixe := Getprefixetable(TOBL);
  TOBL.PutValue('RECALCULTARIF', '-');
  TOBL.PutValue('RECALCULCOMM' , '-');
  TOBL.PutValue('COMPTAAFFAIRE', '');
  TOBL.PutValue('OUVREANAL', '-');
  TOBL.PutValue('QTEORIG', 0);
  TOBL.PutValue('MARGE', 0);
  TOBL.PutValue('ARTSLIES', ''); // Ce champs boolean se doit d'être initialisé à vide
  TOBL.PutValue('QTECHANGE', '-');
  TOBL.PutValue('_CONTREMARTREF', '');
  TOBL.putvalue('SUPPRIME', '-');
  // -- Modif BTP
  TOBL.Putvalue(prefixe+'_RECALCULER', '-');
  TOBL.Putvalue('INDICERG', 0);
  TOBL.PutValue('FROMTARIF', '-');
  TOBL.SetString('NUMBON','');
  // ---
  {$IFDEF BTP}
  InitLesSupLigneBtp (TOBL);
  {$ENDIF}
  InitLesSupMontants(TOBL);
  if InitLigneCompl then InitLesSupLigneCompl(Tobl, False);
end;

function NewTOBLigne(TOBPiece: TOB; ARow: integer; ForFacture : boolean=false): TOB;
var TOBL: TOB;
    MyTobLigneTarif: Tob;
    MyCleDoc: R_CleDoc;
begin
  TOBL := TOB.Create('LIGNE', TOBPiece, ARow - 1);
  NewTOBLigneFille(TOBL);
  AddLesSupLigne(TOBL, False,ForFacture);
  InitLesSupLigne(TOBL);
  Result := TOBL;
end;

procedure NewTOBLigneFille(TOBL: TOB);
begin
	if TOBL.detail.count < 2 then TOB.Create('', TOBL, -1); {Analytique vente/achat}
  if TOBL.detail.count < 2 then TOB.Create('', TOBL, -1); {Analytique stock}
end;

procedure InitTobLigne(TobLigne: Tob);
begin
  if Assigned(TobLigne) then
  begin
    NewTOBLigneFille(TobLigne);
    AddLesSupLigne(TobLigne, False);
    InitLesSupLigne(TobLigne);
  end;
end;

function GetChampLigne(TOBPiece: TOB; Champ: string; ARow: integer): Variant;
var TOBL: TOB;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  Result := TOBL.GetValue(Champ);
end;

procedure InitTOBPiece(TOBP: TOB);
begin
  if TOBP = nil then Exit;
  TOBP.PutValue('GP_INDICEG', 0);
  TOBP.PutValue('GP_ARRONDILIGNE', CheckToString(VH_GC.GCArrondiLigne));
  TOBP.PutValue('GP_UTILISATEUR', V_PGI.User);
  TOBP.PutValue('GP_CREATEUR', V_PGI.User);
  TOBP.PutValue('GP_VIVANTE', 'X');
  TOBP.PutValue('GP_SAISIECONTRE', '-');
  TOBP.PutValue('GP_ETATVISA', 'NON');
  TOBP.PutValue('GP_DATECREATION', Date);
  TOBP.PutValue('GP_DATELIVRAISON', Date);
  TOBP.PutValue('GP_HEURECREATION', NowH);
  TOBP.PutValue('GP_DATEMODIF', NowH);
  TOBP.PutValue('GP_CREEPAR', 'SAI');
  TOBP.PutValue('GP_ETATEXPORT', 'ATT');
  TOBP.PutValue('GP_ETATCOMPTA', 'ATT');
  {$IFNDEF GPAO}
  // Modif BTP
  (*
  TOBP.PutValue('PMAP', 0);
  TOBP.PutValue('PMRP', 0);
  *)
  TOBP.PutValue('GP_MONTANTPA', 0);
  TOBP.PutValue('GP_MONTANTPR', 0);
  TOBP.PutValue('GP_MONTANTPAFG', 0); // Prix achat pour calcule du coef de FG

  {$ENDIF}
  // Modif Mode
  TOBP.PutValue('GP_TICKETANNULE', '-');

end;

procedure InitLesCols;
begin
  SG_NL := -1;
  SG_RefArt := -1;
  SG_CATALOGUE := -1;
  SG_Lib := -1;
  SG_QF := -1;
  SG_QS := -1;
  SG_Px := -1;
  SG_Rem := -1;
  SG_Aff := -1;
  SG_Rep := -1;
  SG_Dep := -1;
  SG_RV := -1;
  SG_RL := -1;
  SG_ContreM := -1;
  SG_Montant := -1;
  SG_DateLiv := -1;
  SG_Total := -1;
  SG_QR := -1;
  SG_Motif := -1;
  SG_QReste := -1; { NEWPIECE }
  // MODIF BTP
  SG_QA := -1;
  SG_Pct := -1;
  SG_Unit := -1;
  SG_TypA := -1;
  SG_PxAch := -1;
  SG_MTPRODUCTION := -1;
  SG_MontantSit := -1;
  SG_MontantAch := -1;
  SG_MontantMarche := -1;
  // ----
  // DEBUT AJOUT CHR
  SG_Folio := -1;
  SG_DateProd := -1;
  SG_Regrpe := -1;
  SG_LibRegrpe := -1;
  // FIN AJOUT CHR
  // DEBUT Modif MODE 31/07/2002
  SG_PxNet := -1;
  SG_TypeRem := -1;
  // FIN Modif MODE 31/07/2002

  SG_Circuit := -1;
{$IFDEF BTP}
	SG_REFTiers := -1;
  SG_DETAILBORD := -1;
  SG_DEJAFACT := -1;
  SG_QTEPREVUE := -1;
  SG_QTESITUATION:= -1;
  SG_POURCENTAVANC:= -1;
  SG_MontantSit:= -1;
  SG_MTDEJAFACT := -1;
  SG_TEMPS := -1;  SG_TEMPSTOT := -1;
  SG_QUALIFTEMPS := -1;
  SG_CODTVA1 := -1;
  SG_NATURETRAVAIL := -1;
  SG_VOIRDETAIL := -1;
  SG_FOURNISSEUR := -1;
  SG_COEFMARG := -1;
  SG_POURMARG := -1;
  SG_POURMARQ := -1;
  SG_NUMAUTO := -1;
  SG_QTESAIS := -1;
  SG_RENDEMENT := -1;
  SG_PERTE := -1;
  SG_QUALIFHEURE := -1;
  SG_QTECOND := -1;
  SG_COEFCOND := -1;
  SG_CODECOND := -1;
  SG_CODEMARCHE := -1;
  SG_TOTALTS := -1;
{$ENDIF}
end;

function FabricWhereNatArt(NaturePieceG, DomainePiece, SelectFourniss: string): string;
var sNatArt, StF, VenteAchat, ChampFiltre, ValsFiltre, ChampArt: string;
  ArtStock: boolean;
begin
  Result := '';
  if NaturePieceG = '' then Exit;
  {Sélection des types articles autorisés}
  sNatArt := GetInfoParPiece(NaturePieceG, 'GPP_TYPEARTICLE');
  {$IFNDEF BTP}
  {Ajout automatique pour la gestion d'affaire des prestations + frais}
  if ctxAffaire in V_PGI.PGIContexte then
  begin
    VenteAchat := GetInfoParPiece(NaturePieceG, 'GPP_VENTEACHAT');
    if VenteAchat = 'VEN' then
    begin
      if Pos('PRE', sNatArt) = 0 then sNatArt := sNatArt + 'PRE;';
      if Pos('FRA', sNatArt) = 0 then sNatArt := sNatArt + 'FRA;';
    end;
  end;
  {$ENDIF}
  Result := FabricWhereToken(sNatArt, 'GA_TYPEARTICLE');
  {Sélection éventuelles des filtres articles}
  ChampFiltre := GetInfoParPiece(NaturePieceG, 'GPP_FILTREARTCH');
  ValsFiltre := GetInfoParPiece(NaturePieceG, 'GPP_FILTREARTVAL');
  ArtStock := (GetInfoParPiece(NaturePieceG, 'GPP_ARTSTOCK') = 'X');
  if ((ChampFiltre <> '') and (ValsFiltre <> '')) then
  begin
    ChampArt := 'GA_' + RechDom('GCFILTREARTPIECE', ChampFiltre, True);
    if ChampToNum(ChampArt) > 0 then
    begin
      StF := FabricWhereToken(ValsFiltre, ChampArt);
      if StF <> '' then
      begin
        if Result = '' then Result := StF else Result := Result + 'AND ' + StF;
      end;
    end;
  end;

  {Sélection éventuelle des articles uniquement tenue en stock}
  if ArtStock then
  begin
    StF := '(GA_TENUESTOCK="X")';
    if Result = '' then Result := StF else Result := Result + 'AND ' + StF;
  end;

  {Sélection éventuelle des domaines articles}
  if DomainePiece <> '' then
  begin
    StF := '(GA_DOMAINE="' + DomainePiece + '" OR GA_DOMAINE="")';
    if Result = '' then Result := StF else Result := Result + 'AND ' + StF;
  end;

  {Sélection éventuelle du fournisseur défini en entête de document d'achat}
  if SelectFourniss <> '' then
  begin
    StF := 'GA_FOURNPRINC="' + SelectFourniss + '"';
    if Result = '' then Result := StF else Result := Result + ' AND ' + StF;
  end;
end;

function EstRempliGC(GS: THGrid; Lig: integer): boolean;
var LaCol: integer;
begin
  Result := False;
  for LaCol := GS.FixedCols to GS.ColCount - 1 do
  begin
    if LaCol = SG_AFF then continue; // si seule l'affaire est renseignée =ligne vide
    if LaCol = SG_ContreM then continue; // si seule la colonne contremarque est renseignée =ligne vide
    if LaCol = SG_CODEMARCHE then Continue;
    if LaCol = SG_NATURETRAVAIL then Continue;
    if LaCol = SG_VOIRDETAIL then Continue;
    if GS.Colwidths[LaCol] = -1 then continue; // non visible;
    Result := ((Result) or (GS.Cells[LaCol, Lig] <> ''));
    if Result then Break;
  end;
end;

procedure DepileTOBLignes(GS: THGrid; TOBPiece: TOB; ARow, NewRow: integer);
var i, Limite, MaxL,MinL: integer;
begin
  if NewRow > ARow then Exit;
  Limite := TOBPiece.Detail.Count + 1;
  MaxL := Max(Limite, ARow);
  MinL := Min(Limite, ARow);
//  for i := MaxL downto NewRow + 1 do
  for i := MaxL downto NewRow + 1  do
  begin
    if not EstRempliGC(GS, i) then Limite := i else Break;
  end;
  for i := TOBPiece.Detail.Count - 1 downto Limite - 1 do
  begin
    TOBPiece.Detail[i].Free;
//    DeleteTobLigneTarif(TobPiece, i);
    GS.Cells[SG_NL, i + 1] := '';
    if (SG_AFF <> -1) then GS.Cells[SG_AFF, i + 1] := '';
    if (SG_ContreM <> -1) then GS.Cells[SG_ContreM, i + 1] := '';
  end;
end;

procedure VideCodesLigne(TOBPiece: TOB; ARow: integer);
var TOBL: TOB;
begin
  TOBL := GetTOBLigne(TOBPiece, ARow);
  if TOBL = nil then Exit;
  TOBL.InitValeurs;
  InitLesSupLigne(TOBL);
end;

procedure PieceAjouteSousDetail(FTOBPiece: TOB; AvecInit : boolean=true; AvecInitLigneCompl : boolean= true; FromLoad : boolean=false);
var i: integer;
  TOBL: TOB;
  SavType : string;
begin
  {$IFDEF BTP}
  InitLesSupEnteteBtp (FTobPiece);
  {$ENDIF}

  if FTOBPiece.Detail.Count > 0 then
  begin
  // Oulla ca pue le bug a plein pif
  //  AddLesSupLigne(FTOBPiece.Detail[0], True);

//    InitLesSupLigne(FTOBPiece.Detail[0]);
    for i := 0 to FTOBPiece.Detail.Count - 1 do
    begin
      TOBL := FTOBPiece.Detail[i];
      (* correction LS *)
      // ppfffff
      Savtype := TOBL.GetValue('BNP_TYPERESSOURCE');
			AddLesSupLigne(TOBL, false);
      (* --- *)
      if AvecInit then InitLesSupLigne(TOBL,AvecInitLigneCompl);
      if FromLoad then TOBL.PutValue('BNP_TYPERESSOURCE',Savtype);
      if (TOBL.getValue('GL_TYPELIGNE')='COM') and (TOBL.getValue('GL_TYPEARTICLE')='RRR') then
      begin
      	TOBL.putValue('MODIFIABLE','-');
        FTOBPiece.PutValue('ESCREMMULTIPLE','X');
      end;
      TOB.Create('', TOBL, -1);
      TOB.Create('', TOBL, -1);
    end;
  end;
end;

procedure PutValueDetail(TOBPiece: TOB; ChampMaj: string; VV: Variant; TOBArticles:TOB=nil);
var il, i: integer;
  NomLig: string;
  TOBLig,TOBART: TOB;
  QQ : TQuery;
  PrixPasModif : boolean;
  DEV : Rdevise;
begin
  prixpasModif := false;
  if TOBPiece = nil then Exit;
  TOBPiece.PutValue(ChampMaj, VV);
  if TOBPiece.Detail.Count <= 0 then Exit;
  NomLig := ChampMaj;
  NomLig[2] := 'L';
  { TP - Bug 10853 - Les champs sups des TobLigne n'ont pas forcément toujours le même index .... }
(*
  il := 0;
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
    TOBLig := TOBPiece.Detail[i];
    if i = 0 then il := TOBLig.GetNumChamp(NomLig);
    if il > 0 then TOBLig.PutValeur(il, VV) else Break;
  end;
*)
  for i := 0 to TOBPiece.Detail.Count - 1 do
  begin
  	TOBArt := nil;
    TOBLig := TOBPiece.Detail[i];
    il := TOBLig.GetNumChamp(NomLig);
    if il > 0 then
    begin
      TOBLig.PutValeur(il, VV);
      if (ChampMaj = 'GP_DOMAINE') and (TOBLIG.GetValue('GL_ARTICLE')<>'') then
      begin
        if TOBArticles <> nil then
        begin
			    TOBArt :=TOBArticles.findFirst(['GA_ARTICLE'],[TOBLig.getValue('GL_ARTICLE')],true);
          prixpasModif :=  (TOBART.getValue('GA_PRIXPASMODIF')='X');
        end;
        if (TOBART = nil) then
        begin
          TOBART := TOB.Create('ARTICLE',nil,-1);
          QQ := OPenSql('SELECT GA_PRIXPASMODIF FROM ARTICLE WHERE GA_ARTICLE="'+TOBLig.getValue('GL_ARTICLE')+'"',true,1,'',true);
          if not QQ.eof then
          begin
            TOBART.selectDB('',QQ);
            prixpasModif :=  (TOBART.getValue('GA_PRIXPASMODIF')='X');
          end;
          ferme(QQ);
          TOBARt.free;
        end;
        if not PrixPasModif then
        begin
      	  AppliqueCoefDomaineLig (TOBLIG);
          if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBLIG,DEV);
        end;
      end;
    end;
  end;
end;

function TypeDeChamp(NomChamp: String) : String;
var
  iTable, iChamp: Integer;
  s: String;
	Mcd : IMCDServiceCOM;
	Field     : IFieldCOM ;
begin
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();

  Result := '';
  Field := Mcd.GetField(NomChamp);
  if Assigned(Field) then
  begin
    s := Field.tipe;
    if (s = 'INTEGER') or (s = 'SMALLINT') then Result := 'I'
    else if (s = 'DOUBLE') or (s = 'EXTENDED') or (s = 'DOUBLE') or (s = 'RATE') then Result := 'N'
    else if (s = 'DATE') then Result := 'D'
    else if (s = 'BLOB') or (s = 'DATA') then Result := 'M'
    else if (s = 'BOOLEAN') then Result := 'B'
    else Result := 'C';
  end;
end;

procedure ProtectionCoef (TOBL : TOB);
begin
  if TOBL.GetValue('GL_COEFMARG')>99999 then TOBL.putValue('GL_COEFMARG',0); // protection du coef de marge
  if TOBL.GetValue('GL_COEFFG')>99999 then TOBL.putValue('GL_COEFG',0); // protection du coef de frais généraux
  if TOBL.GetValue('GL_COEFFC')>99999 then TOBL.putValue('GL_COEFC',0); // protection du coef de frais généraux
  if TOBL.GetValue('GL_COEFFR')>99999 then TOBL.putValue('GL_COEFR',0); // protection du coef de frais répartis
end;

end.
