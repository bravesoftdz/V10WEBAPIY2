unit UtilsOuvragesPlat;

interface
uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, HCtrls, UTOB, Ent1, 
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB, Fe_Main,
{$ENDIF}
{$IFDEF BTP}
     BTPUtil,
{$ENDIF}
     FactTOB,FactArticle,FactPiece,FactOuvrage,
     SysUtils, Dialogs, SaisUtil, UtilPGI, AGLInit, FactUtil,
     Math, EntGC, Classes, NomenUtil,paramsoc,
     UtilArticle,OptimizeOuv,uEntCommun,UtilTOBPiece;

type

  TGenOuvPlat = class (TObject)
    private
      fTOBPiece : TOB;
      fTOBTiers : TOB;
      fTOBarticles : TOB;
      fTOBOuvrages : TOB;
      fTOBOuvragesPlat : TOB;
      fError : string;
      fTreatable : Boolean;
      OptimizeOuv : TOptimizeOuv;
      VenteAchat : string;
      DEV : RDevise;
      IndiceOuv : Integer;
      function isEcritBOP(TOBPiece: TOB; Nature: string): Boolean;
      function ChargeLapiece (cledoc : R_CLEDOC) : boolean;
      function GenOuvPlat : boolean;
      function EcritLesOuvPlats (cledoc : R_CLEDOC) : Boolean;
      procedure Reset;
    public
      property LibRetour : string read fError;
      constructor create;
      destructor destroy; override;
      function TraitePiece (cledoc : R_CLEDOC) : boolean;
  end;

implementation

{ TGenOuvPlat }

function TGenOuvPlat.isEcritBOP (TOBPiece : TOB; Nature : string) : Boolean;
var PassP : string;
    OkCPta : Boolean;
    TheCOndition : string;
    TheZone,TheValeur : string;
begin
  Result := false;
  VenteAchat := GetInfoParPiece(Nature, 'GPP_VENTEACHAT');
  DEV.Code := TOBPIECE.GetValue('GP_DEVISE');
  GetInfosDevise(DEV);
  //
  PassP := GetInfoParPiece(Nature, 'GPP_TYPEECRCPTA');
  OkCpta := ((PassP <> '') and (PassP <> 'RIE'));
  if (OkCpta) or (Pos(Nature,'FBP;FBT;ABT')>0) then BEGIN result := true; Exit; END;
  if Nature='FRC' then begin result:=True; exit; end; // ecriture systématique des pièces de frais dans LIGNEOUVPLAT : pour FEA POUCHAIN
  if (GetParamsocSecur('SO_BTECRITBOPBCE',false)) and (Nature='BCE') then begin Result := True; exit; end;
  if Nature<>'DBT' then exit;
  TheCondition := GetparamSoc('SO_BTCONDECRITBOPDEV');
  if TheCondition = '' then exit;
  if TheCondition = '1=1' then BEGIN result := true; exit; END; // on le veut systématiquement
  TheZone := READTOKENPipe (TheCondition,'=');
  Thevaleur := TheCondition;
  if TOBPiece.getValue(TheZone)=TheValeur then result := true;
end;

function TGenOuvPlat.ChargeLapiece(cledoc: R_CLEDOC): boolean;
var QQ : TQuery;
    cledocAff : r_cledoc;
    TOBAFormule,TOBConds : TOB;
begin
  Result := True;
  TOBConds := TOB.Create('XX XX',nil,-1);
  TOBAFormule := TOB.Create('YY YY',nil,-1);
  FillChar(cledocAff, Sizeof(cledocAff), #0);
  //
  if (cledoc.NaturePiece = '') or (cledoc.Souche = '') or (cledoc.NumeroPiece =0) then
  begin
    fError := 'Renseigner un document à traiter';
    Result := false;
    Exit;
  end;

  TRY
    TRY
      LoadPieceLignes(CleDoc, fTobPiece,true,false,false,false);
      PieceAjouteSousDetail(fTOBPiece,true,false,true);
      fTreatable :=  IsEcritBOP (fTOBpiece,cledoc.NaturePiece);
      if fTreatable then
      begin
        QQ := OpenSQL('SELECT * FROM TIERS LEFT JOIN TIERSCOMPL ON T_TIERS=YTC_TIERS WHERE T_TIERS="' + fTOBPiece.GetString('GP_TIERS') + '"', True,1, '', True);
        if not QQ.EOF then
        begin
          fTOBTiers.SelectDB('', QQ);
        end;
        Ferme(QQ);
        LoadLesOuvrages(fTOBPiece, fTOBOuvrages, fTOBArticles, Cledoc, IndiceOuv,OptimizeOuv);
        RecupTOBArticle (fTOBPiece,fTOBArticles,TOBAFormule,TOBConds,cledocAff,cledoc,VenteAchat);
      end else
      begin
        fError := 'Document non géré';
      end;
    EXCEPT
      on E: Exception do
      begin
        fError := 'Erreur durant la lecture du document';
        Result := false;
      end;
    end;
  FINALLY
    TOBConds.free;
    TOBAFormule.free;
  end;
end;

constructor TGenOuvPlat.create;
begin
  fTOBPiece := TOB.create ('PIECE',nil,-1);
  fTOBTiers := TOB.create ('TIERS',nil,-1);
  fTOBarticles := TOB.Create ('LES ARTICLES',nil,-1);
  fTOBOuvrages := TOB.Create ('LES OUVRAGES',nil,-1);
  fTOBOuvragesPlat := TOB.Create('LES OUV PLAT',nil,-1);
  OptimizeOuv := TOptimizeOuv.create;
end;

destructor TGenOuvPlat.destroy;
begin
  fTOBPiece.free;
  fTOBTiers.free;
  fTOBarticles.free;
  fTOBOuvrages.Free;
  fTOBOuvragesPlat.Free;
  OptimizeOuv.free;

  inherited;
end;

function TGenOuvPlat.EcritLesOuvPlats (cledoc : R_CLEDOC) : Boolean;
begin
  Result := True;
  try
    BEGINTRANS;
    ValideLesOuvPlat (fTOBOuvragesPlat ,fTOBPiece );
    ExecuteSQL('DELETE FROM LIGNEOUVPLAT WHERE '+WherePiece(cledoc,ttdLigneOUVP,true));
    fTOBOuvragesPlat.InsertDBByNivel(false);
    COMMITTRANS;
  except
    on e : Exception do
    begin
      ROLLBACK;
      fError := 'Erreur SQL : ' + E.Message;
      Result := false;
    end;
  end;
end;

function TGenOuvPlat.GenOuvPlat: boolean;
var TOBL,TOBPlat : TOB;
    i : Integer;
begin
  Result := True;
  TRY
    for i:=0 to fTOBpiece.detail.count -1 do
    BEGIN
      TOBL:=fTOBPiece.Detail[i] ;

      if TOBL=Nil then Break ;
      //
      if IsSousDetail(TOBL) then continue;
      //
      if (Pos(TOBL.GetString('GL_TYPELIGNE'),'ART;ARV;CEN')>0) and (not IsLigneFromCentralis(TOBL)) then
      begin
        // --------
        if IsOuvrage (TOBL) and (TOBL.GetInteger('GL_INDICENOMEN')<>0) then
        begin
          TOBPlat := AddMereLignePlat (fTOBOuvragesPlat,TOBL.GetValue('GL_NUMORDRE'));
          GetOuvragePlat (fTOBpiece,TOBL,fTOBOuvrages,TOBPlat,fTOBTiers,DEV,(fTOBPiece.getvalue('GP_NATUREPIECEG')<>'DBT'));
        end;
      end;
    END ;
  EXCEPT
    On E : Exception do
    begin
      fError := 'Erreur durant la préparation des ouvrages plat';
      Result := false;
    end;
  end;
end;

procedure TGenOuvPlat.Reset;              
begin
  fTOBPiece.ClearDetail; fTOBPiece.InitValeurs(false);
  fTOBTiers.InitValeurs(false);
  fTOBarticles.ClearDetail;
  fTOBOuvrages.ClearDetail;
  fTOBOuvragesPlat.ClearDetail;
  IndiceOuv := 1;
  OptimizeOuv.initialisation;
end;

function TGenOuvPlat.TraitePiece(cledoc: R_CLEDOC): boolean;
begin
  Reset;
  Result :=  ChargeLapiece (cledoc);
  if not result then exit;
  if Not fTreatable then BEGIN fError := 'Document non gérable';  Exit; end;
  Result := GenOuvPlat;
  if not Result then Exit;
  Result :=  EcritLesOuvPlats (cledoc);
  if Result then fError := rechdom('GCNATUREPIECEG',cledoc.NaturePiece,false)+' N° : '+InttoStr(cledoc.NumeroPiece)+' Traité';
end;

end.
