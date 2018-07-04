unit UPlannifchUtil;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls, CalcOleGescom,
{$IFDEF EAGLCLIENT}
     HPdfPrev,UtileAGL,Maineagl,
{$ELSE}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main, DB, EdtEtat,EdtDoc,
{$IFNDEF V530} EdtREtat,EdtRDoc, {$ENDIF}
{$ENDIF}
     SysUtils, Dialogs, Utiltarif, SaisUtil, UtilPGI, AGLInit, FactUtil,
     Math, StockUtil, EntGC, Classes, HMsgBox, FactNomen,  TiersUtil, 
{$IFDEF CHR}   HRReglement, {$ENDIF}
{$IFNDEF V530} uRecupSQLModele, {$ENDIF}
     Echeance, UtilGC, UtilArticle, ParamSoc, factcomm,FactTob, forms,uEntCommun;

type
CleligneDevCha = record
                 NaturePiece,Souche : String3 ;
                 DatePiece   : TDateTime ;
                 NumeroPiece,NumLigne : Integer ;
                 Indice,NoPersp : integer ;
                 Niv1,Niv2,Niv3,Niv4,Niv5 : integer; // pour les detail d'ouvrages
                 end;

function DecodeLienDevCHA (CleEntree : string) : CleligneDevCha;
procedure DetruitLienDevisChantier (TOBPiece : TOB);
function ExistReplacePou : boolean;
procedure LoadLesLienDEVCha (TOBPiece,TOBOUvrage,TOBLienDEVCHA : TOB);
procedure ValideLesLienDevCha(TOBPiece,TOBOuvrage,TOBLienDEVCHA : TOB) ;
function existeLienDevCha (TOBPiece,TOBL : TOB; Ecran : TForm) : boolean;
function ControleParagSuprimable (Ecran : Tform ; TobPiece,TOBPiece_O,TOBLienDEVCHA : TOB;GS : THgrid;Ligne : integer) : boolean;
procedure DeprepareDevis (TOBPiece : TOB);
procedure MiseEnPlaceLienReappro (TOBPiece,TOBLien : TOB);
function ExistsChantier (TOBDevis : TOB; var TOBChantier : TOB) : boolean;
procedure PurgeLesLiens (TOBLIen : TOB);
function EncodeLienDevCHADel (TOBPL : TOB ) : string;
procedure SupprimeLienDevCha (TOBL,TOBlien : TOB);

implementation
uses factouvrage,NomenUtil,Facture,UtilBTPgestChantier;


function EncodeLienDevCHADel (TOBPL : TOB ) : string;
begin
  if TOBPL.NomTable = 'LIGNE' then
     begin
//       DD:=TOBPL.GetValue('GL_DATEPIECE') ; StD:=FormatDateTime('ddmmyyyy',DD) ;
       Result:='%;'+TOBPL.GetValue('GL_NATUREPIECEG')+ ';'+TOBPL.GetValue('GL_SOUCHE')+';'
               +IntToStr(TOBPL.GetValue('GL_NUMERO'))+';'+IntToStr(TOBPL.GetValue('GL_INDICEG'))+';';
     end else if TOBPL.Nomtable = 'PIECE' then
     begin
       Result:='%;'+TOBPL.GetValue('GP_NATUREPIECEG')+ ';'+TOBPL.GetValue('GP_SOUCHE')+';'
               +IntToStr(TOBPL.GetValue('GP_NUMERO'))+';'+IntToStr(TOBPL.GetValue('GP_INDICEG'))+';';
     end;
end;

function DecodeLienDevCHA (CleEntree : string) : CleligneDevCha;
Var StC : String ;
begin
FillChar(Result,Sizeof(Result),#0) ; StC:=CleEntree;
Result.DatePiece:=EvalueDate(ReadTokenSt(StC)) ;
Result.NaturePiece:=ReadTokenSt(StC) ;
Result.Souche:=ReadTokenSt(StC) ;
Result.NumeroPiece:=StrToInt(ReadTokenSt(StC)) ;
Result.Indice:=StrToInt(ReadTokenSt(StC)) ;
end;

procedure DetruitLienDevisChantier (TOBPiece : TOB);
var Req : string;
    Nb : integer;
begin
Req := 'DELETE FROM LIENDEVCHA WHERE BDA_REFC LIKE "'+ EncodeLienDevCHADel (TOBPiece) +'"';
Nb := ExecuteSql (Req);
if Nb<0 then BEGIN V_PGI.IoError:=oeUnknown ; Exit ; END ;
end;

function ExistReplacePou : boolean;
var TheArticle : string;
begin
  Result := true;
  TheArticle := GetParamSoc ('SO_BTREPLPOURCENT');
  if TheArticle = '' then
  begin
    PgiBox (TraduireMemoire('Element de remplacement des articles de type pourcentage non défini'), 'Message' );
    result := false;
    exit;
  end;
  if not ExisteSql ('SELECT GA_ARTICLE FROM ARTICLE WHERE GA_CODEARTICLE="'+TheArticle+'"') then
  begin
    PgiBox (TraduireMemoire('Element de remplacement des articles de type pourcentage non défini'), 'Message');
    result := false;
    exit;
  end;
end;

procedure AddChampSupLDc (TOBLDC : TOB ; Withfilles : boolean=false);
begin
  TOBLDC.addChampSupValeur ('INDICELIENDEVCHA',0,WithFilles);
  TOBLDC.AddChampSupValeur ('USED','-');
end;

procedure PositionneNumLien (TobOuvrage : TOB; NumLien : integer);
var TOBLN : TOB;
    Indice : integer;
begin
  if TOBOuvrage = nil then exit;
  if TOBOuvrage.fieldExists('INDICELIENDEVCHA') then TobOuvrage.putValue('INDICELIENDEVCHA',NumLien);
  for Indice := 0 to TOBOuvrage.detail.count -1 do
  begin
    TOBLN := TOBOuvrage.detail[Indice];
    if TOBLN.detail.count > 0 then
    begin
    if TOBLN.fieldExists('INDICELIENDEVCHA') then TobLN.putValue('INDICELIENDEVCHA',NumLien);
    PositionneNumLien (TOBLN,NumLien);
    end;
  end;
end;

procedure LoadLesLienDEVCha (TOBPiece,TOBOUvrage,TOBLienDEVCHA : TOB);
var Indice,IndiceDevCha : Integer;
    NaturePiece,REq : string;
    TOBLDC,TOBL : TOB;
    NumLien : integer;
    QQ : TQuery;
begin
  NumLien := 1;
  NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
  if (Naturepiece <> VH_GC.AFNatAffaire) and (Naturepiece <> GetParamSoc ('SO_BTNATCHANTIER')) then Exit;
  if NaturePiece = VH_GC.AFNatAffaire then
  begin
    // positionné sur le devis
    Req := 'SELECT * FROM LIENDEVCHA WHERE BDA_REFD="'+EncodeLienDevCHA(TOBPiece)+'"';
    QQ := OpenSql (Req,true,-1,'',true);
    TOBLienDevCha.LoadDetailDB ('LIENDEVCHA','','',QQ,false,true);
    ferme (QQ);
    for Indice := 0 To TOBLienDevCha.detail.count -1 do
    begin
      TOBLDC := TOBLienDevCha.detail[Indice];
      AddChampSupLDc (TOBLDC);
      if TOBLDC.GetValue('BDA_N1D')=0 then
      begin
        // Rattache au TOBPiece
        TOBL := TOBPiece.findfirst(['GL_NUMLIGNE'],[TOBLDC.GetValue('BDA_NUMLD')],true);
        if TOBL <> nil then
        begin
          TOBL.PutValue('INDICELIENDEVCHA',NumLien);
          TOBLDC.PutValue('INDICELIENDEVCHA',NumLien);
          TOBLDC.PutValue ('USED','X');
          inc(NumLien);
        end;
      end else
      begin
        // on met a jour la ligne de l'ouvrage
        TOBL := TOBPiece.findfirst(['GL_NUMLIGNE'],[TOBLDC.GetValue('BDA_NUMLD')],true);
        if TOBL <> nil then
        begin
          TOBL.PutValue('INDICELIENDEVCHA',NumLien);
          if TOBL.GetValue('GL_INDICENOMEN') > 0 then
          	PositionneNumLien (TobOuvrage.detail[TOBL.GetValue('GL_INDICENOMEN')-1],Numlien);
        end;
        // Rattache au TOBOuvrage
        TOBL := TOBOuvrage.findfirst(['BLO_NUMLIGNE','BLO_N1','BLO_N2','BLO_N3','BLO_N4','BLO_N5'],
                                     [TOBLDC.GetValue('BDA_NUMLD'),TOBLDC.GetValue('BDA_N1D'),
                                     TOBLDC.GetValue('BDA_N2D'),TOBLDC.GetValue('BDA_N3D'),
                                     TOBLDC.GetValue('BDA_N4D'),TOBLDC.GetValue('BDA_N5D')],true);
        if TOBL <> nil then
        begin
          TOBL.PutValue('INDICELIENDEVCHA',NumLien);
          TOBLDC.PutValue('INDICELIENDEVCHA',NumLien);
          TOBLDC.PutValue ('USED','X');
          inc(NumLien);
        end;
      end;
    end;
  end else
  begin
    // positionné sur le Chantier
    Req := 'SELECT * FROM LIENDEVCHA WHERE BDA_REFC="'+EncodeLienDevCHA(TOBPiece)+'" ORDER BY BDA_NUMLC' ;
    QQ := OpenSql (Req,true,-1,'',true);
    TOBLienDevCha.LoadDetailDB ('LIENDEVCHA','','',QQ,false,true);
    ferme (QQ);
    for IndiceDevCha := 0 To TOBLienDevCha.detail.count -1 do AddChampSupLDc (TOBLienDevCha.detail[IndiceDevCha]);
    for Indice := 0 To TOBPiece.detail.count -1 do
    begin
      TOBL := TOBPiece.detail[Indice];
      for IndiceDevCha := 0 To TOBLienDevCha.detail.count -1 do
      begin
        TOBLDC := TOBLienDevCha.detail[IndiceDevCha];
        if TOBLDC.GetValue('BDA_NUMLC') > TOBL.GetValue('GL_NUMLIGNE') then break;
        if TOBLDC.GetValue('BDA_NUMLC') = TOBL.GetValue('GL_NUMLIGNE') then
        begin
          TOBL.PutValue('INDICELIENDEVCHA',NumLien);
          TOBLDC.PutValue('INDICELIENDEVCHA',NumLien);
          TOBLDC.Putvalue ('USED','X');
        end;
      end;
    inc(NumLien);
    end;
  end;
end;

procedure ValideLesLienDevCha(TOBPiece,TOBOuvrage,TOBLienDEVCHA : TOB) ;
var Indice,IndDevCha : Integer;
    NaturePiece : string;
    TOBLDC,TOBL: TOB;
    Pb : boolean;
begin
  Pb := false;
  NaturePiece := TOBPiece.GetValue('GP_NATUREPIECEG');
  if (Naturepiece <> VH_GC.AFNatAffaire) and (Naturepiece <> GetParamSoc ('SO_BTNATCHANTIER')) then Exit;
  if TOBLienDEVCha.detail.count = 0 then exit;

  if Naturepiece = VH_GC.AFNatAffaire then
  begin
    // maj a partir de piece
    for Indice := 0 to TOBLienDevCha.detail.count -1 do
    begin
      TOBLDC := TOBLienDevCha.detail[Indice];
      if TOBLDC.GetValue('BDA_N1D') > 0 then
      begin
        TOBL := TOBOuvrage.FindFirst (['INDICELIENDEVCHA'],[TOBLDC.GetValue('INDICELIENDEVCHA')],true);
        if (TOBL <> nil) then
        begin
          TOBLDC.PutValue('BDA_NUMLD',TOBL.GetValue('BLO_NUMLIGNE'));
          TOBLDC.PutValue('BDA_N1D',TOBL.GetValue('BLO_N1'));
          TOBLDC.PutValue('BDA_N2D',TOBL.GetValue('BLO_N2'));
          TOBLDC.PutValue('BDA_N3D',TOBL.GetValue('BLO_N3'));
          TOBLDC.PutValue('BDA_N4D',TOBL.GetValue('BLO_N4'));
          TOBLDC.PutValue('BDA_N5D',TOBL.GetValue('BLO_N5'));
        end;
      end else
      begin
        TOBL := TOBPiece.FindFirst (['INDICELIENDEVCHA'],[TOBLDC.GetValue('INDICELIENDEVCHA')],true);
        if (TOBL <> nil) then TOBLDC.PutValue('BDA_NUMLD',TOBL.GetValue('GL_NUMLIGNE'));
      end;
    end;
    if (ExecuteSql('DELETE FROM LIENDEVCHA WHERE BDA_REFD LIKE "'+EncodeLienDevCHADel(TOBPiece)+'"') < 0) then Pb := true;
  end else
  begin
    // maj a partir de piece
    for Indice := 0 to TOBPiece.detail.count -1 do
    begin
      TOBL := TOBPiece.detail[Indice];
      for IndDevCha := 0 to TOBLienDevCha.detail.count -1 do
      begin
        TOBLDC := TOBLienDevCha.detail[IndDevCha];
        if TOBLDC.GetValue('INDICELIENDEVCHA') > TOBL.GetValue('INDICELIENDEVCHA') then break;
        if TOBLDC.GetValue('INDICELIENDEVCHA') = TOBL.GetValue('INDICELIENDEVCHA') then TOBLDC.PutValue('BDA_NUMLC',TOBL.GetValue('GL_NUMLIGNE'));
      end;
    end;
    if ExecuteSql('DELETE FROM LIENDEVCHA WHERE BDA_REFC LIKE "'+EncodeLienDevCHADel(TOBPiece)+'"') < 0  then Pb := true;
  end;
  if not pb then
  begin
    TOBLienDevCha.SetAllModifie (true);
    //if not TOBLienDEVCHA.InsertDB (nil,true) then Pb := true;
    if not TOBLienDEVCHA.InsertOrUpdateDB (true) then Pb := true;
  end;
  if Pb then
  BEGIN
    MessageValid := 'Erreur / Lien devis-chantier';
    PgiError(MessageValid);
    V_PGI.IoError:=oeUnknown ;
  END;
end;

function ControleParagSuprimable (Ecran : Tform ; TobPiece,TOBPiece_O,TOBLienDEVCHA : TOB;GS : THgrid;Ligne : integer) : boolean;
var TOBL : TOB;
    Indice,NiveauParag : Integer;
    TypePar : string;
    XX : TFFacture;
    Warning,Gestionreliquat : boolean;
begin
  result := true;
  XX := TFFacture(Ecran);
  if TOBPiece_O.fieldExists('GP_NATUREPIECEG') then
  begin
  	GestionReliquat := (GetInfoParPiece (TOBPiece_O.GetValue('GP_NATUREPIECEG'),'GPP_RELIQUAT')='X');
  end else
  begin
  	GestionReliquat := false;
  end;

//  if (TobLienDEVCHA = nil) or (TOBLienDevCha.detail.count = 0) then exit; // pas la peine d'aler plus loin
  TOBL := GetTOBLigne (TobPiece,ligne);
  NiveauParag := strtoint(Copy(TOBL.GetValue('GL_TYPELIGNE'),3,1));
  TypePar := copy(TOBL.GetValue('GL_TYPELIGNE'),2,1);
  for Indice := Ligne + 1 to GS.rowcount -1 do
  begin
    TOBL := GetTOBLigne (TOBPiece,Indice);
    if TOBL = nil then exit;
    if TOBL.GetValue('GL_TYPELIGNE')='T'+TypePar+inttoStr(Niveauparag) then break;
    if copy(TOBL.GetValue('GL_TYPELIGNE'),1,2)='T'+TypePar then continue; // debut de phase ou paragraphe de niveau inférieure
    if copy(TOBL.GetValue('GL_TYPELIGNE'),1,2)='D'+TypePar then continue; // fin de phase ou paragraphe de niveau inférieure
    if (TobLienDEVCHA <> nil) and (TOBLienDevCha.detail.count > 0) and (existeLienDevCha (TOBpiece,TOBL,Ecran)) then
    begin
      result := false;
      break;
    end;
  	if not CanModifyLigne(TobL,XX.TransfoPiece, TobPiece_O,Warning,GestionReliquat,XX.Action) then
    begin
    	pgiBox ('Une ou plusieurs lignes ont déjà été livré. Suppression impossible');
      result := false;
      break;
    end;
  end;
end;

function existeLienDevCha (TOBPiece,TOBL : TOB; Ecran : TForm) : boolean;
begin
  result := false;
(*
  MODIF BRL DU 3/11/04 : MIS EN COMMENTAIRES SUITE DEMANDE SNTE
  PLUS DE CONTROLE DE LIEN DEVIS<=>PREVISIONS
  if (TOBL.FieldExists ('INDICELIENDEVCHA')) and (TOBL.GetValue('INDICELIENDEVCHA') > 0) then
  begin
    if (TobPiece.GetValue('GP_NATUREPIECEG')=VH_GC.AFNatAffaire) then pgiBox ('Impossible la prévision de chantier a été générée',ecran.Caption)
                                                                 else pgiBox ('Impossible Il existe un lien avec un devis',ecran.Caption);
    result := true;
    exit;
  end;
*)
end;

procedure DeprepareDevis (TOBPiece : TOB);
var QQ : Tquery;
    TOBLiens,TOBDocuments,TheDoc,UneLigne : TOB;
    Req : string;
    Indice : integer;
    Cledoc : R_cledoc;
    nb : integer;
begin
  TOBLiens := TOB.Create ('LES LIENS',nil,-1);
  TOBDocuments := TOB.Create ('LES DOCS',nil,-1);
  Req := 'SELECT * FROM LIENDEVCHA WHERE BDA_REFC="'+ EncodeLienDevCHA (TOBPiece) +'"';
  QQ := OpenSql (req,true,-1,'',true);
  TRY
    if not QQ.eof then
    begin
      TOBLiens.LoadDetailDB ('LIENDEVCHA','','',QQ,false);
      for Indice := 0 to TOBLiens.detail.count -1 do
      begin
        UneLigne := TOBLiens.detail[Indice];
        DecodeRefPiece (UneLigne.GetValue('BDA_REFD'),cledoc);
        TheDoc := TOBDocuments.findFirst (['GP_NATUREPIECEG','GP_SOUCHE','GP_NUMERO','GP_INDICEG'],
                                          [cledoc.NaturePiece,cledoc.Souche,inttostr(cledoc.NumeroPiece),inttostr(cledoc.Indice)],
                                          true);
        if TheDoc = nil then
        begin
          TheDoc := TOB.Create ('PIECE',TOBDocuments,-1);
          TheDoc.PutValue ('GP_NATUREPIECEG',cledoc.NaturePiece);
          TheDoc.PutValue('GP_SOUCHE',cledoc.souche);
          TheDoc.PutValue('GP_NUMERO',cledoc.NumeroPiece);
          TheDoc.PutValue('GP_INDICEG',cledoc.Indice);
          TheDoc.LoadDB (true);
          nb := ExecuteSQL ('UPDATE AFFAIRE SET AFF_PREPARE="-" WHERE AFF_AFFAIRE="'+TheDoc.GetValue('GP_AFFAIREDEVIS')+'"');
          if nb < 0 then begin V_PGI.IoError := oeUnknown; Exit; END;
        end;
      end;
    end;
  FINALLY
    TOBDocuments.free;
    TOBLIens.free;
    Ferme(QQ);
  END;
end;

procedure AjouteLeLienReappro (TOBLien,TOBL : TOB);
var TOBLI : TOB;
begin
  TOBLI := TOBLien.FindFirst (['LIENLIGNE'],[TOBL.GetValue('LIENLIGNE')],true);
  while TOBLI <> nil do
  begin
    TOBLI.PutValue ('USED','X');
    TOBLI.PutValue('BDA_REFC',EncodeLienDevCHA(TOBL));
    TOBLI.PutValue('BDA_NUMLC',TOBL.getValue('GL_NUMLIGNE'));
    TOBLI := TOBLIEN.FindNext (['LIENLIGNE'],[TOBL.GetValue('LIENLIGNE')],true);
  end;
end;

procedure MiseEnPlaceLienReappro (TOBPiece,TOBLien : TOB);
var Indice : integer;
    TOBL : TOB;
begin
  for Indice := 0 to TOBPiece.detail.count -1 do
  begin
    TOBL := TOBPiece.detail[Indice];
    AjouteLeLienReappro (TOBLien,TOBL);
  end;
end;


procedure PurgeLesLiens (TOBLIen : TOB);
var Indice : integer;
    TOBLI : TOB;
begin
   Indice := 0;
   if (TOBLien = nil) then exit;
   if (TOBLien.detail.count = 0) then exit;
   repeat
     TOBLI := TOBLien.detail[Indice];
     if TOBLI.GetValue('USED') <> 'X' Then
     begin
       TOBLI.free;
     end else
     begin
       Inc(Indice);
     end;
   until Indice >= TOBLien.detail.count -1;
end;

function ExistsChantier (TOBDevis : TOB; var TOBChantier : TOB) : boolean;
var QQ : TQuery;
begin
  result := true;
  if Trim(TOBDevis.getValue('GP_AFFAIRE')) = '' then exit;
  QQ := OpenSql ('SELECT * FROM PIECE WHERE GP_AFFAIRE="'+
                  TOBDevis.getValue('GP_AFFAIRE')+'" AND GP_NATUREPIECEG="'+
                  GetParamSoc('SO_BTNATCHANTIER')+'"',true,-1,'',true);
  if not QQ.eof Then
  begin
    TOBChantier.SelectDB ('',QQ);
    result := false;
  end;
  ferme (QQ);
end;


procedure SupprimeLienDevCha (TOBL,TOBlien : TOB);
var Indice,NumLien : integer;
    TOBLI : TOB;
begin
  Indice := 0;
  if (TOBLien = nil) then exit;
  if not TOBL.FieldExists('INDICELIENDEVCHA') then exit;

  NumLien := TOBL.GetValue('INDICELIENDEVCHA'); if NumLien = 0 then exit;
  if (TOBLien.detail.count = 0) then exit;
  repeat
    TOBLI := TOBLien.detail[Indice];
    if TOBLI.GetValue('INDICELIENDEVCHA') = NumLien Then
    begin
    	TOBLI.free;
    end else
    begin
    	Inc(Indice);
    end;
  until Indice >= TOBLien.detail.count -1;
end;

end.
