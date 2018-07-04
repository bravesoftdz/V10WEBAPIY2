{***********UNITE*************************************************
Auteur  ...... : factTvaMilliem
Créé le ...... : 24/01/2005
Modifié le ... :   /  /
Description .. : Fiche de répartition de tva au millieme  ()
Mots clefs ... : 
*****************************************************************}
Unit FactTvaMilliem ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
     uTob,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HTB97,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOB,
     EntGc,
     UTOF,uEntCommun,UtilTOBPiece ;

Type

	TREPARTTVAMILL = class
    private
    	fusable : boolean;
      fTOBRepart,fTOBPiece,fTOBBases,fTOBOuvrages : TOB;
  		procedure AppliqueDetailOuvrage(TOBOuvrage : TOB; Taxe,Famille : string; Millieme : double);
  	public
    	constructor create (FF : TForm=nil);
    	destructor  destroy ; override;
      procedure   Applique;
			procedure   AppliqueFromPiece;
      procedure   AppliqueLig (TOBL : TOB);
      procedure   Charge;
      procedure   Ecrit;
      procedure   DefiniRepartTva (TheRepart : string);
      procedure   Delete;
      procedure   SetDocument;
      procedure   AppliqueOuvrage (TOBL : TOB; Taxe,Famille : string; millieme : double);
      procedure   InitAppliquePiece;
      procedure   InitApplique (TOBL : TOB);
    	procedure   InitAppliqueOuv(TOBL: TOB);
      procedure   Initialise;

      property    IsExists : boolean read fUsable;
      property    Tobrepart : TOB read fTOBRepart;
      Property    TOBpiece : TOB read fTOBpiece write fTOBpiece;
      Property    TOBBases : TOB read fTOBBases write fTOBBases;
      Property    TOBOuvrages : TOB read fTOBOuvrages write fTOBOuvrages;
  end;

procedure AppliqueMillieme (Tobrepart,TobPiece,TobBases,TobOuvrage : TOB );
procedure AppliqueMilliemeLig (Tobrepart,TobPiece,TobBases,TobOuvrage,TOBL : TOB );
procedure AppliqueMilliemeOuvrage(TobOuvrage,TobRepart,TOBL: TOB);
procedure AppliqueMilliemeDetailOuvrage(TOBOuvrage,TOBrepart: TOB);
procedure DecodeRepartTva (Therepart : string;TOBMilliemes : TOB);
Function EncodeRepartTva (TOBMilliemes : TOB) : string;
function GetMillieme (CodeTaxe : string; TOBRepart : TOB) : double;
procedure InitAppliqueMillieme  (TOBL : TOB);
Function READTOKENINT (var TheChaine : string; Separateur : string) : string;
procedure ValideMillieme (TOBMilliemes : TOB);

implementation
uses Facture,FactTOB,FactUtil;

{ TREPARTTVAMILL }


procedure InitAppliqueMillieme  (TOBL : TOB);
var Indice : integer;
		uneTaxe : TOB;
    prefixe : string;
begin
  prefixe := TableToPrefixe (TOBL.NomTable);
	for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
    Unetaxe  := VH_GC.TOBParamTaxe.detail[Indice];
    if UneTaxe.getValue('BPT_TYPETAXE') <> '' then
    begin
      if UneTaxe.getValue('BPT_TYPETAXE') = 'TVA' Then
      begin
      	// on ne reinitialise que les TVA
        TOBL.PutValue(prefixe+'_FAMILLETAXE'+IntToStr(Indice+1),'');
        if TOBL.fieldExists ('MILLIEME'+IntToStr(Indice+1)) then TOBL.PutValue('MILLIEME'+IntToStr(Indice+1),0)
        																 										else TOBL.AddChampSupValeur ('MILLIEME'+IntToStr(Indice+1),0);
      end;
    end;
  end;
end;

procedure AppliqueMillieme (Tobrepart,TobPiece,TobBases,TobOuvrage : TOB );
var Indice : integer;
		TOBL : TOB;
begin
  if (TOBpiece = nil) or (TobBases = nil) then exit; // on peut poooooo
  if (not TOBPiece.fieldExists ('RUPTMILLIEME')) then TOBPiece.AddChampSupValeur ('RUPTMILLIEME','');
  TOBPiece.putValue('RUPTMILLIEME',EncodeRepartTva (TOBRepart));
  for Indice := 0 to TOBpiece.detail.count -1 do
  begin
  	TOBL := TOBPiece.detail[Indice];
    if not IsArticle (TOBL) then continue;
    AppliqueMilliemeLig (Tobrepart,TobPiece,TobBases,TobOuvrage,TOBL);
  end;
end;

procedure AppliqueMilliemeLig (Tobrepart,TobPiece,TobBases,TobOuvrage,TOBL : TOB );
var IndTva : integer;
		TOBT,TOBTYPT : TOB;
    TheTAxe : string;
begin
  if TOBRepart.detail.count = 0 then exit;
	//
  InitAppliqueMillieme (TOBL);
  for IndTva := 0 to TOBRepart.detail.count -1 do
  begin
    TOBT := TOBRepart.detail[IndTva];
    TOBTYPT := VH_GC.TOBParamTaxe.findFirst (['BPT_CATEGORIETAXE'],[TOBT.GetValue('BPM_CATEGORIETAXE')],true);
    if TOBTYPT <> nil then
    begin
      if TOBTYPT.GetValue('BPT_TYPETAXE') = 'TVA' then
      begin
        TheTaxe := copy (TOBT.GetValue('BPM_CATEGORIETAXE'),3,1);
        TOBL.PutValue('GL_FAMILLETAXE'+TheTaxe,TOBT.GetValue('BPM_FAMILLETAXE'));
        TOBL.PutValue('MILLIEME'+TheTaxe,TOBT.GetValue('BPM_MILLIEME'));
        if (TOBL.GetValue('GL_TYPEARTICLE')='OUV') or (TOBL.GetValue('GL_TYPEARTICLE')='ARP') Then
        begin
          AppliqueMilliemeOuvrage (TobOuvrage,Tobrepart,TOBL);
        end;
      end;
    end;
  end;
  TOBL.PutValue('GL_RECALCULER','X');
end;


procedure AppliqueMilliemeOuvrage(TobOuvrage,TobRepart,TOBL: TOB);
var Num : integer;
		TOBO : TOB;
begin
	if TOBOuvrage = nil then exit;
  Num := TOBL.GetValue('GL_INDICENOMEN');
  if Num = 0 then exit;
  TOBO := TOBOUvrage.detail[Num-1];
  if TOBO = nil then exit;
  AppliqueMilliemeDetailOuvrage(TOBO,TOBrepart);
end;

procedure AppliqueMilliemeDetailOuvrage(TOBOuvrage,TOBrepart: TOB);
var Indice,IndTva : integer;
		TOBO,TOBT,TOBTYPT : TOB;
    TheTaxe : string;
begin
  for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
    TOBO := TOBOuvrage.detail[Indice];
  	InitAppliqueMillieme (TOBO);
    for IndTva := 0 to TOBRepart.detail.count -1 do
    begin
      TOBT := TOBRepart.detail[IndTva];
      TOBTYPT := VH_GC.TOBParamTaxe.findFirst (['BPT_CATEGORIETAXE'],[TOBT.GetValue('BPM_CATEGORIETAXE')],true);
      if TOBTYPT <> nil then
      begin
        if TOBTYPT.GetValue('BPT_TYPETAXE') = 'TVA' then
        begin
          TheTaxe := copy (TOBT.GetValue('BPM_CATEGORIETAXE'),3,1);
        	TOBO.PutValue('BLO_FAMILLETAXE'+TheTaxe,TOBT.GetValue('BPM_FAMILLETAXE'));
        	TOBO.PutValue('MILLIEME'+TheTaxe,TOBT.GetValue('BPM_MILLIEME'));
        end;
      end;
    end;
  end;
end;

function READTOKENINT(var TheChaine: string; Separateur: string): string;
var Position : integer;
begin
  Result := '';
  Position := Pos (Separateur,TheChaine);
  if Position < 0 then exit;
  result := Copy (TheChaine,1,Position-1);
  TheChaine := Copy (TheChaine,Position+1,Length(TheChaine));
end;

Function EncodeRepartTva (TOBMilliemes : TOB) : string;
var Indice : integer;
begin
	result := '';
  if TOBMilliemes = nil then exit;
  if TOBMilliemes.detail.count = 0 then exit;
  for Indice := 0 to TOBMilliemes.detail.count -1 do
  begin
    if Indice > 0 then
    begin
      result := result +';'+
      					TOBMilliemes.detail[Indice].getValue('BPM_CATEGORIETAXE')+'+'+
                TOBMilliemes.detail[Indice].getValue('BPM_FAMILLETAXE')+'='+
                FloatToStr (TOBMilliemes.detail[Indice].getValue('BPM_MILLIEME'));
    end else
    begin
      result := TOBMilliemes.detail[Indice].getValue('BPM_CATEGORIETAXE')+'+'+
      					TOBMilliemes.detail[Indice].getValue('BPM_FAMILLETAXE')+'='+
                FloatToStr (TOBMilliemes.detail[Indice].getValue('BPM_MILLIEME'));
    end;
  end;
end;

procedure DecodeRepartTva (Therepart : string;TOBMilliemes : TOB);
var TheCategorieTva,TheFamilleTva,Millieme,LaRepart,UneRepart : string;
    TOBRepart : TOB;
begin
	if TOBMilliemes = nil then exit;
  if TOBMilliemes.detail.count > 0 then TOBMilliemes.ClearDetail;
  LaRepart := TheRepart;
  repeat
  	UneRepart := READTOKENST (LaRepart);

    // on recupere une chaine sous une forme "Categorie Taxe + Famille Taxe = Millieme"
    if UneRepart = '' then break;
    // recup Categorie Taxe
    TheCategorieTva := READTOKENINT (UneRepart,'+');
    if TheCategorieTva = '' then exit;
    // recup Famille Taxe
    TheFamilleTva := READTOKENINT (UneRepart,'=');
    if TheFamilleTva = '' then exit;
    // Recup millimeme
    Millieme := UneRepart;
    if Millieme = '' then exit;
    // On peut donc maintenant creer un ligne de Repartition

    TOBRepart := TOB.Create ('BTPIECEMILIEME',TOBMilliemes,-1);
    TOBRepart.PutValue ('BPM_CATEGORIETAXE',TheCategorieTva);
    TOBRepart.PutValue ('BPM_FAMILLETAXE',TheFamilleTva);
    TOBRepart.PutValue ('BPM_MILLIEME',Valeur(Millieme));
  until TheRepart = '';
end;

function GetMillieme (CodeTaxe : string; TOBRepart : TOB) : double;
var UneTOB : TOB;
begin
	result := 0;
  if TOBREpart = nil then exit;
  UneTOB := TOBRepart.findFirst (['BPM_CATEGORIETAXE'],[CodeTaxe],true);
  if UneTOB <> nil then result := UneTOB.GetValue('BPM_MILLIEME');
end;

procedure ValideMillieme (TOBMilliemes : TOB);
begin
	if not TOBMilliemes.InsertDB (nil,true) then V_PGI.IOerror := OeUnknown;
end;

procedure TREPARTTVAMILL.Applique;
var Indice : integer;
		TOBL : TOB;
begin
  if (fTOBpiece = nil) or (fTobBases = nil) then exit; // on peut poooooo
  if (not fTOBPiece.fieldExists ('RUPTMILLIEME')) then fTOBPiece.AddChampSupValeur ('RUPTMILLIEME','');
  fTOBPiece.putValue('RUPTMILLIEME',EncodeRepartTva (fTOBRepart));
  for Indice := 0 to fTOBpiece.detail.count -1 do
  begin
  	TOBL := fTOBPiece.detail[Indice];
    if not IsArticle (TOBL) then continue;
    AppliqueLig (TOBL);
  end;
end;

procedure TREPARTTVAMILL.AppliqueDetailOuvrage(TOBOuvrage : TOB; Taxe,Famille : string; Millieme : double);
var Indice : integer;
		TOBO: TOB;
begin
  for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
    TOBO := TOBOuvrage.detail[Indice];
    TOBO.PutValue('BLO_FAMILLETAXE'+Taxe,Famille);
    TOBO.PutValue('MILLIEME'+Taxe,Millieme);
    if TOBO.detail.count > 0 then AppliqueDetailOuvrage (TOBO,Taxe,Famille,Millieme);
  end;
end;

procedure TREPARTTVAMILL.AppliqueFromPiece;
var TheRepart : String;
begin
  if fTOBPiece.FieldExists ('RUPTMILLIEME') then
  begin
  	TheRepart := fTOBPiece.GetValue('RUPTMILLIEME');
    DefiniRepartTva (TheRepart);
    if fTOBRepart.detail.count > 0 then
    begin
      fusable := true;
			Applique;
    end;
  end;
end;

procedure TREPARTTVAMILL.AppliqueLig(TOBL: TOB);
var IndTva : integer;
		TOBT,TOBTYPT : TOB;
    TheTAxe,Famille : string;
    Millieme : double;
begin
	if (not fusable) and (fTOBRepart.detail.count = 0) then exit;

  if fTOBRepart.detail.count = 0 then
  begin
  	exit;
  end else
  begin
    if (not fusable) then fusable := true;
  end;
	//
  InitApplique (TOBL);
  if (TOBL.GetValue('GL_TYPEARTICLE')='OUV') or (TOBL.GetValue('GL_TYPEARTICLE')='ARP') Then
  begin
  	InitAppliqueOuv (TOBL);
  end;


  for IndTva := 0 to fTOBRepart.detail.count -1 do
  begin
    TOBT := fTOBRepart.detail[IndTva];
    TOBTYPT := VH_GC.TOBParamTaxe.findFirst (['BPT_CATEGORIETAXE'],[TOBT.GetValue('BPM_CATEGORIETAXE')],true);
    if TOBTYPT <> nil then
    begin
      if TOBTYPT.GetValue('BPT_TYPETAXE') = 'TVA' then
      begin
        TheTaxe := copy (TOBT.GetValue('BPM_CATEGORIETAXE'),3,1);
        Famille := TOBT.GetValue('BPM_FAMILLETAXE');
        Millieme := TOBT.GetValue('BPM_MILLIEME');
        TOBL.PutValue('GL_FAMILLETAXE'+TheTaxe,Famille);
        TOBL.PutValue('MILLIEME'+TheTaxe,Millieme);
        if (TOBL.GetValue('GL_TYPEARTICLE')='OUV') or (TOBL.GetValue('GL_TYPEARTICLE')='ARP') Then
        begin
          AppliqueOuvrage (TOBL,TheTaxe,Famille,Millieme);
        end;
      end;
    end;
  end;
  TOBL.PutValue('GL_RECALCULER','X');
end;

procedure TREPARTTVAMILL.AppliqueOuvrage (TOBL : TOB; Taxe,Famille : string; millieme : double);
var Num : integer;
		TOBO : TOB;
begin
	if fTOBOuvrages = nil then exit;
  Num := TOBL.GetValue('GL_INDICENOMEN');
  if Num = 0 then exit;
  TOBO := fTOBOUvrages.detail[Num-1];
  if TOBO = nil then exit;
  AppliqueDetailOuvrage(TOBO,Taxe,Famille,Millieme);
end;

procedure TREPARTTVAMILL.Charge;
var cledoc : r_cledoc;
		Sql: String;
    QQ : TQuery;
begin
  if (fTOBpiece = nil) or (fTobBases = nil) then exit; // on peut poooooo
  fusable := true;
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := fTOBPiece.getValue('GP_NATUREPIECEG');
  CleDoc.DatePiece := fTOBPiece.getValue('GP_DATEPIECE');
  CleDoc.Souche := fTOBPiece.getValue('GP_SOUCHE');
  CleDoc.NumeroPiece := fTOBPiece.getValue('GP_NUMERO');
  CleDoc.Indice := fTOBPiece.getValue('GP_INDICEG');
  SQl := 'SELECT * FROM BTPIECEMILIEME WHERE '+WherePiece(Cledoc,TTdRepartmill,false);
  QQ := OpenSql (Sql,True,-1, '', True);
  if not QQ.eof then
  begin
  	fTOBRepart.LoadDetailDB ('BTPIECEMILIEME','','',QQ,False);
  end;
  Applique;
  ferme (QQ);
end;

constructor TREPARTTVAMILL.create(FF: TForm=nil);
begin
	fTOBRepart := TOB.Create ('LA REPARTITION AU 1000',nil,-1);
  fusable := false;
end;

procedure TREPARTTVAMILL.DefiniRepartTva(TheRepart: string);
begin
	DecodeRepartTva (TheRepart,fTOBRepart);
end;

procedure TREPARTTVAMILL.Delete;
var cledoc : r_cledoc;
begin
  if not fUsable then exit;
  FillChar(CleDoc, Sizeof(CleDoc), #0);
  CleDoc.NaturePiece := fTOBPiece.getValue('GP_NATUREPIECEG');
  CleDoc.DatePiece := fTOBPiece.getValue('GP_DATEPIECE');
  CleDoc.Souche := fTOBPiece.getValue('GP_SOUCHE');
  CleDoc.NumeroPiece := fTOBPiece.getValue('GP_NUMERO');
  CleDoc.Indice := fTOBPiece.getValue('GP_INDICEG');
  ExecuteSql('DELETE BTPIECEMILIEME WHERE '+WherePiece(Cledoc,TTdRepartmill,false));
end;

destructor TREPARTTVAMILL.destroy;
begin
	fTOBRepart.free;
  inherited;
end;

procedure TREPARTTVAMILL.Ecrit;
begin
  if not fUsable then exit;
  Delete;
	SetDocument;
	if not fTOBRepart.InsertDB (nil,true) then
  begin
    MessageValid := 'Erreur mise à jour Répartition millième';
    V_PGI.IOerror := OeUnknown;
  end;
end;


procedure TREPARTTVAMILL.InitApplique  (TOBL : TOB);
var Indice : integer;
		uneTaxe : TOB;
    prefixe : string;
begin
  prefixe := TableToPrefixe (TOBL.NomTable);
	for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
  begin
    Unetaxe  := VH_GC.TOBParamTaxe.detail[Indice];
    if UneTaxe.getValue('BPT_TYPETAXE') <> '' then
    begin
      if UneTaxe.getValue('BPT_TYPETAXE') = 'TVA' Then
      begin
      	// on ne reinitialise que les TVA
        TOBL.PutValue(prefixe+'_FAMILLETAXE'+IntToStr(Indice+1),'');
        if TOBL.fieldExists ('MILLIEME'+IntToStr(Indice+1)) then TOBL.PutValue('MILLIEME'+IntToStr(Indice+1),0)
        																 										else TOBL.AddChampSupValeur ('MILLIEME'+IntToStr(Indice+1),0);
      end;
    end;
  end;
end;

procedure TREPARTTVAMILL.InitAppliqueOuv  (TOBL : TOB);

		procedure InitAppliqueOuvDet (TOBOuv : TOB);
    var Indice,IIndice : integer;
        uneTaxe,TOBO : TOB;
        prefixe : string;
    begin
    	for IIndice := 0 to TOBOUV.detail.count -1 do
      begin
      	TOBO := TOBOUV.detail[Iindice];
        prefixe := TableToPrefixe (TOBO.NomTable);
        for Indice := 0 to VH_GC.TOBParamTaxe.detail.count -1 do
        begin
          Unetaxe  := VH_GC.TOBParamTaxe.detail[Indice];
          if UneTaxe.getValue('BPT_TYPETAXE') <> '' then
          begin
            if UneTaxe.getValue('BPT_TYPETAXE') = 'TVA' Then
            begin
              // on ne reinitialise que les TVA
              TOBO.PutValue(prefixe+'_FAMILLETAXE'+IntToStr(Indice+1),'');
              if TOBO.fieldExists ('MILLIEME'+IntToStr(Indice+1)) then TOBO.PutValue('MILLIEME'+IntToStr(Indice+1),0)
                                                                  else TOBO.AddChampSupValeur ('MILLIEME'+IntToStr(Indice+1),0);
            end;
          end;
        end;
        if TOBO.detail.count > 0 then InitAppliqueOuvDet(TOBO);
      end;
    end;

var Num : integer;
		TOBO : TOB;
begin
	if fTOBOuvrages = nil then exit;
  Num := TOBL.GetValue('GL_INDICENOMEN');
  if Num = 0 then exit;
  TOBO := fTOBOUvrages.detail[Num-1];
  if TOBO = nil then exit;
  InitAppliqueOuvDet(TOBO);
end;

procedure TREPARTTVAMILL.InitAppliquePiece;
var Indice : integer;
		TOBL : TOB;
begin
  if (fTOBpiece = nil) then exit; // on peut poooooo
  for Indice := 0 to fTOBpiece.detail.count -1 do
  begin
  	TOBL := fTOBPiece.detail[Indice];
    if not IsArticle (TOBL) then continue;
    InitApplique (TOBL);
    if (TOBL.GetValue('GL_TYPEARTICLE')='OUV') or (TOBL.GetValue('GL_TYPEARTICLE')='ARP') Then
    begin
      InitAppliqueOuv (TOBL);
    end;
  end;

end;

procedure TREPARTTVAMILL.Initialise;
begin
  fTOBRepart.ClearDetail;
	fTOBPiece := nil;
  fTOBBases := nil;
  fTOBOuvrages := nil;
end;

procedure TREPARTTVAMILL.SetDocument;
var Indice : integer;
		TOBR : TOB;
begin
	if fTobrepart.detail.count = 0 then exit;
	Indice := 0;
  // nettoyage
	repeat
    TOBR := fTOBRepart.detail[Indice];
    if (TOBR.getValue('BPM_FAMILLETAXE')='') OR (TOBR.getValue('BPM_MILLIEME') = 0) then TOBR.free else Inc(Indice);
  until indice >= fTOBRepart.detail.count;

	for Indice := 0 to fTOBRepart.detail.count -1  do
  begin
    TOBR := fTOBRepart.detail[Indice];
    TOBR.putValue('BPM_NATUREPIECEG',fTOBPiece.getValue('GP_NATUREPIECEG'));
    TOBR.putValue('BPM_SOUCHE',fTOBPiece.getValue('GP_SOUCHE'));
    TOBR.putValue('BPM_NUMERO',fTOBPiece.getValue('GP_NUMERO'));
    TOBR.putValue('BPM_INDICEG',fTOBPiece.getValue('GP_INDICEG'));
  end;

end;

end.
