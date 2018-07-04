unit USql;

interface
uses DB, ADODB, Variants, ActiveX, ComObj, AdoInt, OleDB,classes,SysUtils,WSNomadeDecl;

function GetTopAppels (ressource,Database,Server,StPlus : WideString) : WideString;
function GetDataAppel (CodeAppel,Database,Server : WideString) : WideString;

implementation
uses Ulog,Udefinitions;

function GetTopAppels (ressource,Database,Server,StPlus : WideString) : WideString;
var QQ: TADOQuery;
		Sql,UneLigne,Conn,EmplacementIni,NomIni : String;
begin
  Result := '';
  EmplacementIni := getCheminIni;
  NomIni := IncludeTrailingPathDelimiter(EmplacementIni)+'Definitions.ini';
	Conn := 'Provider=SQLOLEDB;Password=ADMIN;User ID=ADMIN;Initial Catalog='+Database+';Data Source='+Server+';';
  //
	ecritLog (EmplacementIni,'Chaine Connexion = '+Conn);
  Sql := 'SELECT TOP 10 AFF_PRIOCONTRAT,AFF_ETATAFFAIRE,'+
				 'T2.T_LIBELLE AS CLIENT,'+
         'AFF_TIERS,AFF_DATECREATION,AFF_DATEFIN,AFF_LIBELLE,ADR_CONTACT,'+
				 'ADR_TELEPHONE,ADR_LIBELLE,ADR_ADRESSE1,ADR_ADRESSE2,ADR_CODEPOSTAL,'+
         'ADR_VILLE,AFF_DATESIGNE,AFF_RESPONSABLE,AFF_AFFAIRE '+
				 'FROM BTMULAPPELS '+
         'LEFT OUTER JOIN TIERS T2 ON AFF_TIERS=T2.T_TIERS '+
         'WHERE AFF_AFFAIRE0 LIKE ''W%'' AND AFF_ETATAFFAIRE IN (''AFF'',''ECR'') '+
         'AND AFF_STATUTAFFAIRE = ''APP'' AND AFF_RESPONSABLE='''+ ressource +''' '+ stPlus + ' '+
         'ORDER BY AFF_ETATAFFAIRE DESC, AFF_PRIOCONTRAT ASC ,AFF_DATECREATION DESC ';
	ecritLog (EmplacementIni,'Requete = '+Sql);

	QQ := TADOQuery.Create(nil);
  QQ.ConnectionString := Conn;
  QQ.SQL.clear;
  QQ.SQL.Add(Sql);
  TRY
    QQ.Open;
    if not QQ.Eof then
    begin
      QQ.First;
      while not QQ.Eof do
      begin
        UneLigne := QQ.findfield('AFF_AFFAIRE').AsString+SepChamp+
        						QQ.findfield('AFF_PRIOCONTRAT').AsString+SepChamp+
        						QQ.findfield('AFF_ETATAFFAIRE').AsString+SepChamp+
                    QQ.findfield('CLIENT').AsString+SepChamp+
        						DateToStr(QQ.findfield('AFF_DATECREATION').AsDateTime)+SepChamp+
                    QQ.findfield('AFF_LIBELLE').AsString+SepLigne;

				ecritLog (EmplacementIni,'Une ligne = '+UneLigne);
        Result := Result  + UneLigne;
        QQ.Next;
      end;
    end;
  	QQ.Close;
  finally
    QQ.Free;
  end;
end;

function GetDataAppel (CodeAppel,Database,Server : WideString) : WideString;
var QQ: TADOQuery;
		Sql,UneLigne,Conn,EmplacementIni,NomIni : String;
begin
  Result := '';
  EmplacementIni := getCheminIni;
  NomIni := IncludeTrailingPathDelimiter(EmplacementIni)+'Definitions.ini';
	Conn := 'Provider=SQLOLEDB;Password=ADMIN;User ID=ADMIN;Initial Catalog='+Database+';Data Source='+Server+';';
  //
	ecritLog (EmplacementIni,'Chaine Connexion = '+Conn);
  Sql := 'SELECT AFF_PRIOCONTRAT,AFF_ETATAFFAIRE,'+
				 'T2.T_LIBELLE AS CLIENT,'+
         'AFF_TIERS,AFF_DATECREATION,AFF_DATEDEBUT,AFF_DATEFIN,AFF_LIBELLE,ADR_CONTACT,'+
				 'ADR_TELEPHONE,ADR_LIBELLE,ADR_ADRESSE1,ADR_ADRESSE2,ADR_CODEPOSTAL,'+
         'ADR_VILLE,AFF_DATESIGNE,AFF_RESPONSABLE,AFF_AFFAIRE '+
				 'FROM BTMULAPPELS '+
         'LEFT OUTER JOIN TIERS T2 ON AFF_TIERS=T2.T_TIERS '+
         'WHERE AFF_AFFAIRE = '''+CodeAppel+'''';
	ecritLog (EmplacementIni,'Requete = '+Sql);

	QQ := TADOQuery.Create(nil);
  QQ.ConnectionString := Conn;
  QQ.SQL.clear;
  QQ.SQL.Add(Sql);
  TRY
    QQ.Open;
    if not QQ.Eof then
    begin
      QQ.First;
      while not QQ.Eof do
      begin
        UneLigne := QQ.findfield('AFF_AFFAIRE').AsString+SepChamp+
        						QQ.findfield('AFF_PRIOCONTRAT').AsString+SepChamp+
        						QQ.findfield('AFF_ETATAFFAIRE').AsString+SepChamp+
                    DateToStr(QQ.findfield('AFF_DATECREATION').AsDateTime)+SepChamp+
                    DateToStr(QQ.findfield('AFF_DATEDEBUT').AsDateTime)+SepChamp+
                    DateToStr(QQ.findfield('AFF_DATEFIN').AsDateTime)+SepChamp+
                    QQ.findfield('AFF_LIBELLE').AsString+SepChamp+
                    QQ.findfield('ADR_CONTACT').AsString+SepChamp+
                    QQ.findfield('ADR_TELEPHONE').AsString+SepChamp+
                    QQ.findfield('ADR_LIBELLE').AsString+SepChamp+
                    QQ.findfield('ADR_ADRESSE1').AsString+SepChamp+
                    QQ.findfield('ADR_ADRESSE2').AsString+SepChamp+
                    QQ.findfield('ADR_CODEPOSTAL').AsString+SepChamp+
                    QQ.findfield('ADR_VILLE').AsString+SepChamp+
                    DateToStr(QQ.findfield('ADR_DATESIGNE').AsDateTime)+SepChamp+
                    QQ.findfield('CLIENT').AsString+SepChamp+
                    QQ.findfield('AFF_LIBELLE').AsString+SepLigne;

				ecritLog (EmplacementIni,'Une ligne = '+UneLigne);
        Result := Result  + UneLigne;
        QQ.Next;
      end;
    end;
  	QQ.Close;
  finally
    QQ.Free;
  end;
end;

end.
