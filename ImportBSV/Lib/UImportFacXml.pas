unit UImportFacXml;

interface
Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
		 HPanel,
		 Hctrls,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     Variants,
     XMLDoc,xmlintf,UXMLvalidation,CBPPath,Graphics,
     UTOF ;

type
	TCodeStatusInt = (TCsOk,TCsNconforme,TCsTiersInconnu,TCsDateInvalide,TCsLivInvalide,TcsMtInvalide,TcsNoDetails,TcsBadChantier,TcsDocDejaGenere);
  TsetCodeStatusInt = set of TCodeStatusInt;

  TImportFachXml = class (TObject)
  private
  	frepertSto : string;
    fCodeStatus : TsetCodeStatusInt;
    fTOBDatas : TOb;
    fTOBGenere : TOB;
    //
    procedure CreeTOB;
    function ConstitueCodeChantier(Chantier: string): string;
    function ExisteChantier(CodeChantier: string): boolean;
    procedure InitTOB;
    function AddFille: TOB;
    function GetNomFournisseur(CodeFou: string): string;
    procedure ControleDoc(TOBL: TOB);
    procedure initTrait;
	public

    constructor create;
    destructor destroy; override;
    property TOBDatas : TOB read fTOBDatas;
    property TOBGenere : TOB read fTOBGenere write fTOBGenere;
    property RepertDef : string read frepertSto;
    property Statutimp : TsetCodeStatusInt read fCodeStatus;
    //
    procedure ExamineTOBDatas;

    procedure MoveToTreated (NomFic : string);
    procedure MoveToError (NomFic : string);
    procedure TraiteXml (NomFic : string);
    procedure ContitueMotif (MotifPiece : TStrings);
    //
  end;

implementation
uses Dialogs,BTPUtil,AffaireUtil,factGrp,ParamSoc;

{ TImportFachXml }


function  TImportFachXml.AddFille : TOB;
begin
	result := TOB.Create ('UNE FILLE',fTOBDatas,-1);
	result.AddChampSupValeur('NATUREPIECE','');
	result.AddChampSupValeur('NUMEROPIECE',0);
end;

constructor TImportFachXml.create;
begin
  frepertSto:= IncludeTrailingBackslash(GetparamSocSecur('SO_BTEMPLXMLFAC','C:\'));
  if frepertSto = '' then frepertSto := 'C:\';
  //
	CreeTob;
  fCodeStatus := [TcsOk];
end;

procedure TImportFachXml.CreeTOB;
begin
	fTOBDatas := TOB.Create ('UN DOCUMENT',nil,-1);
  fTOBDatas.AddChampSupValeur('CODEFOU','');
  fTOBDatas.AddChampSupValeur('TYPEDOC','');
  fTOBDatas.AddChampSupValeur('LIBDOC','');
  fTOBDatas.AddChampSupValeur('NOMFOU','');
  fTOBDatas.AddChampSupValeur('CHANTIER','');
  fTOBDatas.AddChampSupValeur('AFFAIRE','');
  fTOBDatas.AddChampSupValeur('NUMFACTUREFOU','');
  fTOBDatas.AddChampSupValeur('DATEFAC',Idate1900);
  fTOBDatas.AddChampSupValeur('MONTANTHT',0.0);
  fTOBDatas.AddChampSupValeur('TVA',0.0);
  fTOBDatas.AddChampSupValeur('MONTANTTTC',0.0);
  //
  fTOBGenere := TOB.Create('LES PIECES',nil,-1);
end;


function TImportFachXml.ConstitueCodeChantier (Chantier : string) : string;
var P0,P1,P2,P3,Avenant,Suite : string;
		Affaire : string;
begin
	Result := '';
	P0 := 'A';
  P1 := READTOKENPipe(Chantier,' ');
  Suite := READTOKENPipe(Chantier ,' ');
  if Suite <> '' then P2 := Suite;
  Suite := READTOKENPipe(Chantier ,' ');
  if Suite <> '' then P3 := Suite;
  Suite := READTOKENPipe(Chantier ,' ');
  if Suite <> '' then Avenant := Suite;
  IF avenant = '' then Avenant := '00';
	result := CodeAffaireRegroupe (P0,P1,P2,P3,Avenant,taModif,false,True,False);
end;

destructor TImportFachXml.destroy;
begin
  fTOBDatas.Free;
  fTOBGenere.Free;
  inherited;
end;

function TImportFachXml.ExisteChantier (CodeChantier : string) : boolean;
begin
	result := ExisteSql('SELECT AFF_AFFAIRE FROM AFFAIRE WHERE AFF_AFFAIRE="'+CodeChantier+'"');
end;

procedure TImportFachXml.initTrait;
begin
	InitTOB;
  fCodeStatus := [TCsOk];
end;

procedure TImportFachXml.InitTOB;
begin
  fTOBDatas.SetString('CODEFOU','');
  fTOBDatas.SetString('TYPEDOC','');
  fTOBDatas.SetString('LIBDOC','');
  fTOBDatas.SetString('NOMFOU','');
  fTOBDatas.SetString('CHANTIER','');
  fTOBDatas.SetString('AFFAIRE','');
  fTOBDatas.SetString('NUMFACTUREFOU','');
  fTOBDatas.SetDateTime('DATEFAC',Idate1900);
  fTOBDatas.SetDouble('MONTANTHT',0.0);
  fTOBDatas.SetDouble('TVA',0.0);
  fTOBDatas.SetDouble('MONTANTTTC',0.0);
  fTOBDatas.ClearDetail;
  //
  fTOBGenere.InitValeurs(false);
  fTOBGenere.ClearDetail;
end;

procedure TImportFachXml.MoveToError(NomFic: string);
var flocrepert : string;
begin
	flocrepert := ChangeFileExt(NomFic, '.BAD');
  RenameFile(NomFic,flocrepert);
end;

procedure TImportFachXml.MoveToTreated(NomFic: string);
var flocrepert : string;
begin
	flocrepert := ChangeFileExt(NomFic, '.OK');
  RenameFile(NomFic,flocrepert);
end;

Procedure TImportFachXml.ControleDoc (TOBL : TOB);
var QQ : TQuery;
		Sql : string;
begin
  Sql := 'SELECT GP_VIVANTE FROM PIECE WHERE GP_NATUREPIECEG="'+TOBL.GetString('NATUREPIECE')+'" AND '+
  			 'GP_NUMERO='+ IntToStr(TOBL.getInteger('NUMEROPIECE'));
  QQ := OpenSql (Sql,True,1,'',true);
  if not QQ.Eof then
  begin
    if QQ.fields[0].AsString <> 'X' then
    begin
      if TcsOk in fCodeStatus then  fCodeStatus:= fCodeStatus - [TCsOk];
      if not (TcsDocDejaGenere in fCodeStatus) then
      begin
        fCodeStatus := fCodeStatus + [TcsDocDejaGenere];
      end;
    end;
  end else
  begin
    if TcsOk in fCodeStatus then  fCodeStatus:= fCodeStatus - [TCsOk];
    if not (TCsLivInvalide in fCodeStatus) then
    begin
      fCodeStatus := fCodeStatus + [TCsLivInvalide];
    end;
  end;
  ferme (QQ);
end;

procedure TImportFachXml.TraiteXml(NomFic: string);
var XsdFile : string;
		XmlDoc : IXMLDocument ;
    XX,YY : IXMLNode;
    TOBL : TOB;
    NameSpace,ValInterm : string;
    II,JJ : Integer;
begin
	initTrait;
  // vérification de la structure du xml
  XsdFile :=IncludeTrailingBackslash (TCBPPath.GetCegidDataDistriStd)+'exportGLB.xsd';
  //
  if FileExists(XsdFile) then
  begin
    //
    NameSpace := 'Facture';
    TRY
      ValidateXMLFileName (NomFic,XsdFile,NameSpace);
    EXCEPT
  		if TcsOk in fCodeStatus then  fCodeStatus := fCodeStatus - [TCsOk];
      fCodeStatus := fCodeStatus + [TCsNconforme];
    END;
  end;
  //
  if TcsNConforme in fCodeStatus  then Exit;

  XmlDoc := NewXMLDocument();
  TRY
    XmlDoc.LoadFromFile(nomFic);
    if not XmlDoc.IsEmptyDoc then
    begin
      For II := 0 to Xmldoc.DocumentElement.ChildNodes.Count -1 do
      begin
        XX := XmlDoc.DocumentElement.ChildNodes[II];
        if XX.NodeName = 'CodeFournisseur' then
        begin
          fTOBDatas.SetString('CODEFOU',UTF8Decode(XX.NodeValue));
					ValInterm := GetNomFournisseur (UTF8Decode(XX.NodeValue));
          if ValInterm <> '' then fTOBDatas.SetString('NOMFOU',valInterm)
          									 else fTOBDatas.SetString('NOMFOU','Inconnu');

        end else if XX.NodeName = 'TypeFacture' then
        begin
          if UTF8Decode(XX.NodeValue) = 'F' then
          begin
            fTOBDatas.SetString('TYPEDOC','FF');
          end else
          begin
            fTOBDatas.SetString('TYPEDOC','AF');
          end;
  				fTOBDatas.SetString('LIBDOC',rechdom('GCNATUREPIECEG',fTOBDatas.GetString('TYPEDOC'),false));
        end else if XX.NodeName = 'DateFacture' then
        begin
          fTOBDatas.SetDateTime('DATEFAC',StrToDate(XX.NodeValue));
        end else if XX.NodeName = 'NumeroFacture' then
        begin
          fTOBDatas.SetString('NUMFACTUREFOU',UTF8Decode(XX.NodeValue));
        end else if XX.NodeName = 'CodeChantier' then
        begin
          if not VarIsNull(XX.NodeValue)  then
          begin
            fTOBDatas.SetString('CHANTIER',UTF8Decode(XX.NodeValue));
            fTOBDatas.SetString('AFFAIRE',ConstitueCodeChantier(UTF8Decode(XX.NodeValue)));
          end;
        end else if XX.NodeName = 'TotalHT' then
        begin
          ValInterm := StringReplace(UTF8Decode (XX.NodeValue),'.',',',[RfReplaceAll]);
          fTOBDatas.SetDouble('MONTANTHT',Abs(StrToFloat(ValInterm)));
        end else if XX.NodeName = 'TotalTVA' then
        begin
          ValInterm := StringReplace(UTF8Decode (XX.NodeValue),'.',',',[RfReplaceAll]);
          fTOBDatas.SetDouble('TVA',Abs(StrToFloat(ValInterm)));
        end else if XX.NodeName = 'TotalTTC' then
        begin
          ValInterm := StringReplace(UTF8Decode (XX.NodeValue),'.',',',[RfReplaceAll]);
          fTOBDatas.SetDouble('MONTANTTTC',Abs(StrToFloat(ValInterm)));
        end else if XX.NodeName = 'Commande' then
        begin
          (*
          For JJ := 0 TO XX.ChildNodes.Count -1 do
          begin
            YY := XX.ChildNodes[JJ];
            if YY.NodeName = 'NumeroCde' then
            begin
              TOBL := AddFille;
              TOBL.SetString('NATUREPIECE','CF');
              TOBL.SetString('NUMEROPIECE',YY.NodeValue);
            end;
          end;
          *)
        end else if XX.NodeName = 'Livraison' then
        begin
          For JJ := 0 TO XX.ChildNodes.Count -1 do
          begin
            YY := XX.ChildNodes[JJ];
            if YY.NodeName = 'NumeroBL' then
            begin
              TOBL := AddFille;
              TOBL.SetString('NATUREPIECE','BLF');
              TOBL.SetInteger('NUMEROPIECE',StrToInt(YY.NodeValue));
              ControleDoc (TOBL);
            end;
          end;
        end;
      end;
    end else
    begin
    	if TcsOk in fCodeStatus then  fCodeStatus:= fCodeStatus - [TCsOk];
      fCodeStatus := fCodeStatus + [TCsNconforme];
    end;
  FINALLY
  	XmlDoc:= nil;
  end;
  //
  ExamineTOBDatas;
end;

procedure TImportFachXml.ExamineTOBDatas;
begin
  if TcsNconforme in fCodeStatus then Exit;
  if ControleDateDocument (DateToStr(TOBDatas.GetDateTime ('DATEFAC')),TOBDatas.GetString('TYPEDOC'))<>0 then
  begin
  	if TcsOk in fCodeStatus then  fCodeStatus:= fCodeStatus - [TCsOk];
    fCodeStatus := fCodeStatus + [TCsDateInvalide];
  end;
  if fTOBDatas.getString('CHANTIER') <> '' then
  begin
    if not ExisteChantier(fTOBDatas.getString('AFFAIRE')) then
    begin
    	if TcsOk in fCodeStatus then  fCodeStatus:= fCodeStatus - [TCsOk];
      fCodeStatus := fCodeStatus + [TcsBadChantier];
    end;
  end;
  if (fTOBDatas.GetDouble('MONTANTTTC')=0) or (fTOBDatas.GetDouble('MONTANTHT')=0) then
  begin
   	if TcsOk in fCodeStatus then  fCodeStatus:= fCodeStatus - [TCsOk];
    fCodeStatus := fCodeStatus + [TcsMtInvalide];
  end;
  if fTOBDatas.detail.count = 0 then
  begin
  	if TcsOk in fCodeStatus then  fCodeStatus:= fCodeStatus - [TCsOk];
    fCodeStatus := fCodeStatus + [TcsNoDetails];
  end;
end;

function TImportFachXml.GetNomFournisseur(CodeFou: string): string;
var QQ : TQuery;
begin
	QQ := OpenSQL('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+CodeFou+'" AND T_NATUREAUXI="FOU"',True,1,'',true);
  if NOt QQ.Eof then
  begin
  	Result := QQ.Fields[0].AsString;
  end else
  begin
  	Result := '';
    fCodeStatus := fCodeStatus + [TCsTiersInconnu];
  end;
  ferme (QQ);
end;


procedure TImportFachXml.ContitueMotif(MotifPiece: TStrings);
begin
  if TCsNconforme in Statutimp then
  begin
    MotifPiece.Add('Document non conforme');
  end;
  if TCsTiersInconnu in Statutimp then
  begin
    MotifPiece.Add('Fournisseur Inconnu');
  end;
  if TCsDateInvalide in Statutimp then
  begin
    MotifPiece.Add('Date Invalide');
  end;
  if TCsLivInvalide in Statutimp then
  begin
    MotifPiece.Add('Document(s) constituant inexistant');
  end;
  if TcsMtInvalide in Statutimp then
  begin
    MotifPiece.Add('Montants nuls');
  end;
  if TcsNoDetails in Statutimp then
  begin
    MotifPiece.Add('pas de réception(s) associée(s)');
  end;
  if TcsBadChantier in Statutimp then
  begin
    MotifPiece.Add('Chantier invalide');
  end;
  if TcsDocDejaGenere in Statutimp then
  begin
    MotifPiece.Add('Document dèjà généré');
  end;


end;

end.
