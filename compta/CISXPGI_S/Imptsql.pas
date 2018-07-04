unit ImptSql;


interface

uses Classes,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet, Variants, ADODB,  {$ENDIF}
UScript, Dialogs, DB, Forms, HRichOle;

type
	TTypeSortie = (tsAucune, tsParadox, tsASCII, tsASCIIODBC, tsODBC);
	TDestSortie = (tsUser, tsTemp, tsAliasODBC);
	TModeAscii = (maFixe, maDelimite);

	TScriptRequete = class(TScript)
	private
		FRequete    : TStringList;
{$IFNDEF DBXPRESS}
		FTypeTable  : TTableType;
{$ENDIF}
		FAliasName  : String;
		FTypeSortie : TTypeSortie;
		FNomFichier : String;
		FDestSortie : TDestSortie ;
		FModeAscii  : TModeAscii;
		FAliasSortieODBC : string;
		FQuery      : TQuery;
		slCle       : TStringList;
		slNomMemo   : TStringList;
		slValeur    : TStringList;
                FMotdepasse : string;
                FUtilisateur: string;
                Fdomaine     : string;
                Flibelle     : string;
                FChemin      : string;
                FDriverName  : string;
	protected
		procedure DataCalcField(Sender : TDataset);
	public
		constructor Create(AOwner : TComponent); override;
		destructor Destroy; override;

		function Executer(AOnNextRecord : TNextRecordEvent) : Boolean; override;
		function DropComment(S : string):string;
		function ExtraireMemo(const sreq : string): string;
                //procedure AssignSql(ZSQL: THSQLMemo);
                procedure AssignSql(ZSQL: THRichEditOLE);
                // procedure AssignSqlMemo(var ZSQL: THSQLMemo; FR : TStringList);
                procedure AssignSqlMemo(var ZSQL: THRichEditOLE; FR : TStringList);
	published
		property Requete : TStringList read FRequete write FRequete;
{$IFNDEF DBXPRESS}
		property TypeTable : TTableType read FTypeTable write FTypeTable default ttParadox;
{$ENDIF}
		property AliasName : String read FAliasName write FAliasName;
		property TypeSortie : TTypeSortie read FTypeSortie write FTypeSortie default tsAucune;
		property DestSortie : TDestSortie read FDestSortie write FDestSortie default tsUser;
		property NomFichier : String read FNomFichier write FNomFichier;
		property ModeAscii : TModeAscii read FModeAscii write FModeAscii;
		property AliasSortieODBC : string read FAliasSortieODBC write FAliasSortieODBC;
		property Motdepasse : String read FMotdepasse write FMotdepasse;
		property Utilisateur : String read FUtilisateur write FUtilisateur;
		property Domaine : String read FDomaine write FDomaine;
		property Libelle : String read FLibelle write FLibelle;
		property RChemin : String read FChemin write FChemin;
		property DriverName : String read FDriverName write FDriverName;
	end;

procedure ModifierSchema(const aAliasName, aFileName : String);
procedure SplitStr(AS1:String; sep : char; var AS2 : array of string; var ACount:integer);
function LoadScriptRequete(AScriptName:String) : TScriptRequete;

implementation

uses bde, SysUtils, Windows, UDMIMP;

procedure SplitStr(AS1:String; sep : char; var AS2 : array of string; var ACount:integer);
var P, P1 : PChar;
	  N : integer;
begin
	ACount := 0;
	if AS1 = '' then exit;
	P := PChar(AS1); N:=0;
	while P <> nil do
	begin
		P1 := P;
		P := StrScan(P, sep);
		if P <> nil then
			SetString(AS2[N], P1, P-P1)
		else AS2[N] := P1;
		inc(N);
		if P <> nil then inc(P);
	end;
	ACount := N;
end;

procedure ModifierSchema(const aAliasName, aFileName : String);
var SL : TStringList;
	FSchemaName : String;
begin
    if aAliasName = '' then exit;
	SL := TStringList.Create;
{$IFNDEF DBXPRESS}
	FSchemaName := ChangeFileExt(Session.FindDatabase(aAliasName).Directory+'\'+aFileName, '.sch');
{$ELSE}
	FSchemaName := ChangeFileExt(Session.PrivateDir+'\'+aFileName, '.sch');
{$ENDIF}
	SL.LoadFromFile(FSchemaName);
	//	SL.LoadFromFile(ChangeFileExt(args.GetDBPath(aAliasName)+aFileName, '.sch'));
	SL.Values['FILETYPE'] := 'VARYING';
	SL.Values['DELIMITER'] := '"';
	SL.Values['SEPARATOR'] := ';';
	SL.SaveToFile(FSchemaName);
	SL.Free;
end;

constructor TScriptRequete.Create(AOwner : TComponent);
begin
	inherited Create(AOwner);
	FRequete	:= TStringList.Create;
	FAliasName	:= 'GLOBAL';
	FQuery		:= TQuery.Create(Self);
	ParName		:= 'SQL';
	slCle       := TStringList.Create;
	slNomMemo   := TStringList.Create;
	slValeur    := TStringList.Create;
end;

destructor TScriptRequete.Destroy;
var N : Integer;
begin
	slCle.Free;
	slCle := nil;
	for N:=0 to slNomMemo.Count-1 do
		slNomMemo.Objects[N].Free;
	slNomMemo.Free;
	slNomMemo := nil;
	slValeur.Free;
	FRequete.Free;
    FQuery.Free;
	inherited Destroy;
end;

function TScriptRequete.DropComment(S : string):string;
var
	sReste  : string;
	I, n : integer;
begin // retirer les commentaires et renvoyer une nouvelle chaines correcte
	result := '';
	I := pos('/*', S); // chercher le '/*'
	if I > 0 then
	begin
		Result := copy(S, 1, I-1); // prendre ce qu'il y a avant
		sReste := copy(S, I+2, Length(S));
		n := pos('*/',  sReste);// rechercher le '*/'
		if n > 0 then
			Result := Result + DropComment(copy(sReste, n+2, Length(S)))
		else Result := S;
	end
	else Result := S;
end;

// extraire les noms de champ memo dans la requête
//
// slNomMemo : liste des champs memo utilise
//
function TScriptRequete.ExtraireMemo(const sreq : string): string;
var n, m, I, J, P, MI : integer;
	mot, mot1 : array [0..1023] of string;
	nMemo, nMemo1, scle, nReq : string;
	SQM : TsqlMemo;
	slChamp : TStringList;
begin
	slChamp := TStringList.Create;
	for I := 0 to slNomMemo.Count-1 do
	  slNomMemo.objects[I].Free;
	slNomMemo.Clear;
	slCle.Clear;
	if sReq = '' then exit;
	P := pos('SELECT',sReq);
	if P = 0 then exit ;
	nreq := trim(Copy(sreq, P+6, pos('FROM', sreq)-(P+6) ));
	splitstr(nReq, ',', mot, n); // découpage de la requete
	if n < 1 then exit;	// pas de parametres
	MI := 1;
	for I := 0 to N-1 do
	begin
		splitstr(mot[I], '!', mot1, m);
		if m > 1  then // il y a un memo
		begin
		  nMemo := trim(mot1[0]);
		  nMemo1 := trim(mot1[1]);
		  J := slNomMemo.indexof(nMemo);
		  if (mot1[1] > '') and (mot1[1][1] = '"') then
			  scle := trim(copy(mot1[1], 2, length(mot1[1])-2) )
		  else
			  scle := mot1[1];
		  sCle := 'C'+IntToStr(MI)+'='+sCle;
		  if  J = -1 then
		  begin //enregistrer le NomMemo si pas encore fait
			 SQM := TSqlMemo.create;
			 SQM.NomMemo := nMemo;
			 SQM.NomCle.Add(scle);
			 SLNomMemo.AddObject(nMemo, SQM);
			 if SLChamp.indexof(nMemo) = -1 then
			 	SLChamp.add(nMemo);
		  end
		  else TsqlMemo(SLNomMemo.Objects[J]).NomCle.Add(sCle) ;//enregistrer la cle dans le sqlMemo
		  SLCle.Add(sCle); // enregistrer le nom de la cle
		  Inc(MI);
		end
		else SLChamp.Add(Trim(mot[I]));
	end;

	// reconstruire la requete
	nreq := copy(sReq, 1, P+6) + ' ';
	for I := 0 to slChamp.count-1 do
	begin
		 nreq := nreq + SLChamp[I];
		 if I <> slChamp.count-1 then
			nreq := nreq + ', ';
	end;
	nreq := nreq + ' '+ copy(sReq, pos('FROM', sreq)-1 ,  length(sReq));
	Result := nreq; // voici la nouvelle requete
	slChamp.Free;
end;

// Affectation des sous-champs MEMO

procedure TScriptRequete.DataCalcField(Sender : TDataset);
var     I, N : integer;
	sName, NomChamp : String;
	SM : TSQLMemo;
begin
	// retrouver le memo
	for N := 0 to Sender.FieldCount -1 do
	begin
		sName := Sender.Fields[N].FieldName;
		NomChamp := slCle.Values[sName];
		//		K := slCle.IndexOf(sName);
		if  NomChamp = '' then continue;
		//		if  K < 0 then continue;
		// retrouver le champ Memo
		for I := 0 to slNomMemo.count-1 do
		begin  // chercher dans Nomcle des objects de slNomMemo
			SM := TSQLMemo(slNomMemo.Objects[I]);
			if SM.NomCle.IndexOf(sName+'='+NomChamp) <> -1 then
			begin // le champ memo est trouvé
				slValeur.Text := Sender.FieldByName(SM.NomMemo).AsString;
				Sender.Fields[N].Value := SLValeur.Values[NomChamp];
			end;
		end;
	end;
end;

function TScriptRequete.Executer(AOnNextRecord : TNextRecordEvent) : Boolean;
var T : TTable;
	B : TBookmark;
	M, N, MI, cntTuples : Integer;
	NomChamp : String;
	OL : TOutputList;
	aFMTNumber, aNewFMTNumber : FMTNumber;
	SL : TStringList;
	Qry : TQuery;
	_DB, _DBin : TDatabase;

	slTrace : TStringList;

	procedure GestionMemo;
	var I, M : Integer;
	begin
	  Qry := TQuery.Create(Self);
	  Qry.DatabaseName := FQuery.DatabaseName;
	  MI := 1;
	  for I := 0 to FQuery.FieldCount-1 do
	  begin // parcours des champs du query
		  with FQuery.fields[I] do
		  begin
			N := SLNomMemo.Indexof(fieldName);
			if N = -1 then // c'est pas un des champs memo
			begin
			  ImportAddField(Qry, fieldName, dataType, datasize, Required, false, true);
			end
			else
			begin // c'est un champ memo
				 ImportAddField(Qry, fieldName, dataType, datasize, Required, false, false);
				 with TSqlMemo(SLNomMemo.objects[N]) do
					for M := 0 to Nomcle.Count-1 do
					begin
					  NomChamp := Format('C%d', [M+1]);
					  ImportAddField(Qry, NomChamp, ftstring, 50, false, true, true);
					  Inc(MI);
					  //	ImportAddField(Qry, NomCle[M], ftstring, 20, false, true, true);
					end;
				 Qry.OnCalcFields := DataCalcField;
				 Qry.AutoCalcFields := true;
			end;
		  end;
	  end;
	  Qry.Sql.Text := FQuery.Sql.Text;
	  try Qry.Open; except end;
	  FQuery.EnableControls;
	  FQuery.Free;

	  FQuery := Qry;
	  FQuery.DisableControls;
	  FQuery.AutoCalcFields := true;
	  FQuery.OnCalcfields := DataCalcField;
	end;
begin
	slTrace := TStringList.Create;
	_DB := TDatabase.Create(nil);
	_DBin := TDatabase.Create(nil);
	FQuery.DisableControls;
	try
	  try
		slTrace.Add('TScriptRequete.Executer'#13'AliasName='+FAliasName);
		OnNextRecord := AOnNextRecord;
		FQuery.Close;
		//slTrace.Add(IntToStr(Integer(TypeTable)));
		//slTrace.Add('ttParadox='+IntToStr(Integer(ttParadox)));
		//ShowMessage(Name);
		//ShowMessage(Variable.Text);
		if Variable.Values['_LoginPrompt'] = '1' then
		begin
			_DBin.Params.Add('USER NAME='+Variable.Values['_USER NAME']);
			_DBin.Params.Add('PASSWORD='+Variable.Values['_PASSWORD']);
			_DBin.DatabaseName := '_ODBCin';
{$IFNDEF  DBXPRESS}
			_DBin.AliasName := FAliasName;
{$ENDIF}
			_DBin.LoginPrompt := false;
			_DBin.Open;
			FQuery.DatabaseName := '_ODBCin';
		end
		else
			FQuery.DatabaseName := FAliasName;
		SL := TStringList.Create;
		for N:=0 to FRequete.Count-1 do
		  SL.Add(InterpreteVar(FRequete[N], Variable));
		//  FQuery.SQL.Assign(SL);
		FQuery.SQL.Text := ExtraireMemo(UPPERCASE(SL.Text));
		slTrace.Add(FQuery.SQL.Text);
		SL.Free;
		try
		  OL.data := nil;
		  //if iTrace > 1 then ShowMessage('TScriptRequete.Executer(1)');
		  if Assigned(AOnNextRecord) then
			  AOnNextRecord(Self, 0, nil, false, OL);
		  //if iTrace > 1 then ShowMessage('TScriptRequete.Executer(Query.Open)'#13+FQuery.SQL.Text);
		  FQuery.Open;
		  // rajouter les champs de la requête SLChamp contient tous les champs de la requête
		  if slNomMemo.count > 0 then
			  GestionMemo;	// s'il y a des champs Memos dans la requête
		  slTrace.Add('TScriptRequete.Executer(Query.Open) OK');
		except
		  on E : Exception do
		  begin
			  ShowMessage('Exec. requete !'#13#13+E.message);
			  exit;
		  end;
		end;

		if FTypeSortie <> tsAucune then
		begin
		  DbiGetNumberFormat(aFMTNumber);	// modifiaction du separateur de decimal
		  aNewFMTNumber := aFMTNumber;		//  DecimalSeparator ne marche pas
		  aNewFMTNumber.cDecimalSeparator := '.';	// pour les tables ASCII
		  DbiSetNumberFormat(aNewFMTNumber);
		  T := TTable.Create(nil);
		  B := FQuery.GetBookmark;
		  try
			T.TableName := ExtractFileName(FNomFichier);
{$IFNDEF  DBXPRESS}
			if FTypeSortie = tsParadox then
			  T.TableType := ttParadox
			else
			  T.TableType := ttASCII;
{$ENDIF}
			if TypeSortie = tsParadox then
			begin
				if DestSortie = tsUSER then
					T.DatabaseName := 'USER'
				else T.DatabaseName := 'TEMP';
			end
			else if TypeSortie = tsASCII then
			begin
				if DestSortie = tsUSER then
					T.DatabaseName := 'USER'
				else T.DatabaseName := 'TEMP';
			end
			else if TypeSortie = tsODBC then
			begin
			  if Variable.Values['_LoginPrompt'] = '1' then
			  begin
				_DB.Params.Add('USER NAME='+Variable.Values['_USER NAME']);
				_DB.Params.Add('PASSWORD='+Variable.Values['_PASSWORD']);
				_DB.DatabaseName := '_ODBC';
{$IFNDEF  DBXPRESS}
				_DB.AliasName := AliasSortieODBC;
{$ENDIF}
				_DB.LoginPrompt := false;
//                _DB.Params.Values ['user name']:='ADMIN';
//                _DB.Params.Values ['password']:='ADMIN';
				_DB.Open;
				T.DatabaseName := '_ODBC';
			  end
			  else
				  T.DatabaseName := AliasSortieODBC;
			end;

			slTrace.Add('AliasODBC='+T.DatabaseName);
			//	ShowMessage(slNomMemo.Text);
			if TypeSortie <> tsParadox then
			begin
			  with FQuery.FieldDefs do
				for N:=0 to Count-1 do
				  with Items[N] do
				  begin
					if DataType = ftMemo then
						T.FieldDefs.Add(Name, ftString, 1024, false)
					else
						T.FieldDefs.Add(Name, DataType, Size, false);
				  end;
			end
			else	// les blobs paradox sont conserves telles quelles
			begin
			  if slNomMemo.Count > 0 then
				for m := 0 to FQuery.FieldCount-1 do
				begin
					if FQuery.fields[m].visible then
					begin
						//	NomChamp := Format('C%d', [M+1]);
						ImportAddField(T, FQuery.fields[m].fieldName, FQuery.fields[m].datatype, FQuery.fields[m].size, false, false, FQuery.fields[m].visible);
					end;
				end
				else
					T.FieldDefs.Assign(FQuery.FieldDefs);
			end;
			slTrace.Add('TypeSortie : '+IntToStr(Integer(TypeSortie))+' : '+IntToStr(Integer(tsODBC)));
      {$IFNDEF  DBXPRESS}
                        T.CreateTable;
      {$ENDIF}
			if FTypeSortie = tsParadox then
				T.EmptyTable
			else
			begin
				if ModeAscii = maDelimite then   // TCurrencyField
				begin
				  Sleep(2000); // suite au pb avec Le fichier .SCH
				  ModifierSchema(T.DatabaseName, ExtractFileName(NomFichier));
				end;
			end;

			T.Open;
			cntTuples := 0;
			slTrace.Add(IntToStr(FQuery.FieldCount));

			while not FQuery.Eof do
			begin
				LaCallback(nil, true, OL);
				T.Append;
				for N:=0 to FQuery.FieldCount-1 do
				  if (slNomMemo.Count > 0) then
				  begin
					  if (FQuery.Fields[N].Visible) then
						try
							T.FieldbyName(FQuery.Fields[N].FieldName).Value := FQuery.Fields[N].Value;
						except
							// si tu ne trouves pas le champs tu laisses tombé
						end;
				  end
				  else
					  T.Fields[N].Value := FQuery.Fields[N].Value;
				T.Post;
				Inc(cntTuples);
				if Assigned(AOnNextRecord) then
				  AOnNextRecord(Self, 1, nil, false, OL);
				FQuery.Next;
			end;
			slTrace.Add(Format('%d tuple(s) transféré(s)', [cntTuples]));
			slTrace.Add('TScriptRequete.Executer(Close)');
		  finally
			DbiSetNumberFormat(aFMTNumber);
			T.Free;
			slTrace.Add('TScriptRequete.Executer(Fin)');
			if assigned(AOnNextRecord) then AOnNextRecord(Self, 2, nil, false, OL);
			FQuery.FreeBookmark(B);
		  end;
		end;
	  except
		// slTrace.Add('TScriptRequete.Executer'#13+Exception(ExceptObject).Message);
		ShowMessage('TScriptRequete.Executer'#13+Exception(ExceptObject).Message);
		raise;
	  end;
	finally
	  FQuery.EnableControls;
	  _DBin.Free;
	  _DB.Free;
	  if iTrace >= 50 then
		ShowMessage(slTrace.Text);
	  slTrace.Free;
	end;
end;

function LoadScriptRequete(AScriptName:String) : TScriptRequete;
var
	Stream1, AStreamTable : TmemoryStream;
	ABlobField            : TField;
	ATblCorField          : TBlobField;
	AScript               : TScript;
	S1                    : String;
{$IFDEF  DBXPRESS}
	tbCharger             : TADOTable;
{$ELSE}
	tbCharger             : TTable;
{$ENDIF}
	iDeuxPoint            : Integer;

begin
	Result := nil;

    if not Assigned(DMImport) then DMImport := TDMImport.Create(Application);
    tbCharger := DMImport.GzImpReq;
{$IFDEF  DBXPRESS}
    DMImport.Cmd.CommandText := 'SELECT * FROM PGZIMPREQ WHERE CLE="'+ AScriptName +'"';
    DMImport.Cmd.Execute;
{$ENDIF}

    if iDeuxPoint > 0 then
        AScriptName := Copy(AScriptName, iDeuxPoint+1, 99);

    with tbCharger do begin
            if not Active then Open;
{$IFDEF  DBXPRESS}
            if not EOF then begin
{$ELSE}
            if not FindKey([AScriptName]) then begin
{$ENDIF}
                showMessage('La requête '+aScriptName+' n''est pas une requête personnalisée');
                exit;
            end;
            ABlobField := FieldByName('PARAMETRES');
            ATblCorField := FieldByName('TBLCOR') as TBlobField;
            //Stream1 := TBlobStream.Create(TBlobField(ABlobField), bmRead);
            //AStreamTable := TBlobStream.Create(ATblCorField, bmRead);
            Stream1 := TmemoryStream.create;
            TBlobField(ABlobField).SaveToStream (Stream1);
            Stream1.Seek (0,0);

            AStreamTable := TmemoryStream.create;
            TBlobField(ATblCorField).SaveToStream (AStreamTable);
            AStreamTable.Seek (0,0);

            try AScript := LoadScriptFromStream(Stream1, AStreamTable);
            except begin
                AStreamTable.Free;
                Stream1.Free;
                ShowMessage('Le script est endommagé.');
                Close;
                exit; end;
            end;
            AScript.Destination := poGlobal;

            AStreamTable.Free;
            Stream1.Free;
            AScript.FComment := FieldByName('COMMENT').AsString;
            try
                AScript.NiveauAcces := FieldByName('MODIFIABLE').AsInteger;
            except
                AScript.NiveauAcces := 0;    // non modifiable
            end;
            Close;
    end;
	Result := TScriptRequete(Ascript);
end;

//procedure TScriptRequete.AssignSql(ZSQL: THSQLMemo);
procedure TScriptRequete.AssignSql(ZSQL: THRichEditOLE);
var
N   : integer;
begin
(*      if (ZSQL.text <> '') and (ZSQL.lines.Count = 0)  then
      begin
               FRequete.clear;
		       FRequete.Add(ZSQL.Text)
      end
      else
*)
               FRequete.clear;
      for N:=0 to ZSQL.lines.Count-1 do
		  FRequete.Add(ZSQL.lines[N]);
end;


procedure TScriptRequete.AssignSqlMemo(var ZSQL: THRichEditOLE; FR : TStringList);
//procedure TScriptRequete.AssignSqlMemo(var ZSQL: THSQLMemo; FR : TStringList);
var
N   : integer;
begin
      ZSQL.clear;
      for N:=0 to FR.Count-1 do
		  ZSQL.lines.Add(FR[N]);
end;

initialization
	RegisterClasses([TScriptRequete]);

end.
