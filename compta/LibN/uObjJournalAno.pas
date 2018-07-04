unit uObjJournalAno;

interface

uses
  HCtrls,
  SysUtils,   // FreeAndNil
{$IFDEF EAGLCLIENT}
{$ELSE}
  db,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
uTob;

type

  TObjJournalAno = class (TObject)

  private
    FLastError    : integer;
    FLastErrorMsg : string;
    FStCodeExo    : string;
    FStJournalAno : string;

    FTobListePiece : Tob;

    procedure   TraitementDoublon(vTobPiece : Tob; vInNumeroPiece : integer);
    procedure   DetruitPieceAno(vTobPiece : Tob);
    procedure   SupprimePieceAno( vInNumeroPiece : integer; vStEtab, vStDevise : string);

  public
    constructor Create;
    destructor  Destroy; override;
    procedure   Execute;

    property LastError : integer read FLastError write FLastError;
    property LastErrorMsg : string read FLastErrorMsg write FLastErrorMsg;
    property StCodeExo : string read FStCodeExo write FStCodeExo;

  end;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ParamSoc,   // GetParamSocSecur
  uTobDebug,  // TobDebug
  uLibWindows,// IIF
  Ent1,       // VH^
  HEnt1;      // TraduireMemoire


{ TObjJournalAno }

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/11/2006
Modifié le ... : 20/11/2006
Description .. :
Mots clefs ... :
*****************************************************************}
constructor TObjJournalAno.Create;
begin
  FLastError    := 0;
  FLastErrorMsg := '';

  FStJournalANO := GetParamSocSecur('SO_JALOUVRE', '');

  FTobListePiece := Tob.Create('', nil, -1);
  FTobListePiece.LoadDetailFromSql('SELECT DISTINCT E_DEVISE, E_ETABLISSEMENT, ' +
    'E_NUMEROPIECE FROM ECRITURE WHERE ' +
    'E_JOURNAL = "'+ FStJournalAno + '" AND E_EXERCICE = "'+ FStCodeExo + '" AND ' +
    'E_ECRANOUVEAU = "OAN" ' +
    'ORDER BY E_DEVISE, E_ETABLISSEMENT, E_NUMEROPIECE', False);

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/11/2006
Modifié le ... : 22/11/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TObjJournalAno.Execute;
var lQuery : TQuery;
    lTobEcrAno : Tob;
    lStSql : string;
    lStDevise : string;
    lStEtab : string;
begin
  lTobEcrAno := Tob.Create('', nil, -1);
  try
    lQuery := OpenSQL('SELECT DISTINCT E_DEVISE, E_ETABLISSEMENT, E_NUMEROPIECE ' +
                      'FROM ECRITURE WHERE ' +
                      'E_JOURNAL = "' + FStJournalANo + '" AND E_EXERCICE = "'+ FStCodeExo + '" AND ' +
                      'E_ECRANOUVEAU = "OAN" ' +
                      'ORDER BY E_DEVISE, E_ETABLISSEMENT, E_NUMEROPIECE', True);

    if not lQuery.Eof then
    begin
      lStDevise := '';
      lStEtab := '';

      lQuery.First;
      while not lQuery.Eof do
      begin
        if (lStDevise <> lQuery.FindField('E_DEVISE').AsString) or
           (lStEtab <> lQuery.FindField('E_ETABLISSEMENT').AsString) then
        begin
          lStDevise := lQuery.FindField('E_DEVISE').AsString;
          lStEtab   := lQuery.FindField('E_ETABLISSEMENT').AsString;

          lStSql := 'SELECT * FROM ECRITURE WHERE ' +
                    'E_ECRANOUVEAU = "OAN" AND ' +
                    'E_JOURNAL = "' + FStJournalAno + '" AND ' +
                    'E_EXERCICE = "' + FStCodeExo + '" AND ' +
                    'E_ETABLISSEMENT = "' + lStEtab + '" AND ' +
                    'E_DEVISE = "' + lStDevise + '" ' +
                    'ORDER BY E_GENERAL, E_AUXILIAIRE, E_ETABLISSEMENT, E_DEVISE, E_NUMEROPIECE';

          lTobEcrAno.ClearDetail;
          lTobEcrAno.LoadDetailDBFromSQL('ECRITURE', lStSql);
          TraitementDoublon(lTobEcrAno, lQuery.FindField('E_NUMEROPIECE').AsInteger);

          // Suppression de l'A-Nouveaux PGI
          SupprimePieceAno( lQuery.FindField('E_NUMEROPIECE').AsInteger, lStEtab, lStDevise);

          lTobEcrAno.InsertDb(nil);
        end
        else
        begin
          // Piece de même de couple établissement devise à supprimer puisqu'on fusionne
          lStSql := 'SELECT ##TOP 2## * FROM ECRITURE WHERE ' +
                    'E_ECRANOUVEAU = "OAN" AND ' +
                    'E_JOURNAL = "' + FStJournalAno + '" AND ' +
                    'E_EXERCICE = "' + FStCodeExo + '" AND ' +
                    'E_ETABLISSEMENT = "' + lStEtab + '" AND ' +
                    'E_DEVISE = "' + lStDevise + '" AND ' +
                    'E_NUMEROPIECE = ' + IntToStr(lQuery.FindField('E_NUMEROPIECE').AsInteger) + ' ' +
                    'ORDER BY E_GENERAL, E_AUXILIAIRE, E_NUMLIGNE';

          lTobEcrAno.ClearDetail;
          lTobEcrAno.LoadDetailDBFromSQL('ECRITURE', lStSql);
          DetruitPieceAno(lTobEcrAno);

          // Suppression de l'A-Nouveaux PGI
          SupprimePieceAno( lQuery.FindField('E_NUMEROPIECE').AsInteger,
                            lQuery.FindField('E_ETABLISSEMENT').AsString,
                            lQuery.FindField('E_DEVISE').AsString);

          lTobEcrAno.InsertDb(nil);
        end;

        lQuery.Next;
      end;
    end;

  finally
    FreeAndNil(lTobEcrAno);
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/11/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TObjJournalAno.TraitementDoublon(vTobPiece : Tob; vInNumeroPiece : integer);
var i : integer;
    lStGene : string;
    lStAuxi : string;
    lTobTemp : Tob;
    lTobi    : Tob;
    lSolde, lSoldeDev : Double;
    lNumLigne : integer;
begin
  lStGene := '';
  lStAuxi := '';
  vTobPiece.Detail.Sort('E_GENERAL;E_AUXILIAIRE;');

  lTobTemp := nil;
  for i := vTobPiece.Detail.Count - 1 downto 0 do
  begin
    lTobi  := vTobPiece.Detail[i];

    if (lTobi.GetString('E_GENERAL') <> lStGene) or
       (lTobi.GetString('E_AUXILIAIRE') <> lStAuxi) then
    begin
      lTobTemp       := lTobi;
      lStGene        := lTobi.GetString('E_GENERAL');
      lStAuxi        := lTobi.GetString('E_AUXILIAIRE');
      //lStRefPointage := lTobi.GetString('E_REFPOINTAGE');
    end
    else
    begin
      lSolde := lTobi.GetDouble('E_DEBIT') + lTobTemp.GetDouble('E_DEBIT') -
                lTobi.GetDouble('E_CREDIT') - lTobTemp.GetDouble('E_CREDIT');

      lSoldeDev := lTobi.GetDouble('E_DEBITDEV') + lTobTemp.GetDouble('E_DEBITDEV') -
                   lTobi.GetDouble('E_CREDITDEV') - lTobTemp.GetDouble('E_CREDITDEV');

      if lSolde > 0 then
      begin
        lTobTemp.PutValue('E_DEBIT', lSolde);
        lTobTemp.PutValue('E_DEBITDEV', lSoldeDev);
        lTobTemp.PutValue('E_CREDIT', 0);
        lTobTemp.PutValue('E_CREDITDEV', 0);
        lTobTemp.PutValue('TOTAL', lSolde);
        lTobTemp.PutValue('TOTALDEV', lSoldeDev);
      end
      else
      begin
        lTobTemp.PutValue('E_DEBIT', 0);
        lTobTemp.PutValue('E_DEBITDEV', 0);
        lTobTemp.PutValue('E_CREDIT', Abs(lSolde));
        lTobTemp.PutValue('E_CREDITDEV', Abs(lSoldeDev));
        lTobTemp.PutValue('TOTAL', lSolde);
        lTobTemp.PutValue('TOTALDEV', lSoldeDev);
      end;

      lTobi.Free;
    end;
  end;
  // FIN GCO

  lNumLigne := 0;
  for i := 0 to vTobPiece.Detail.Count - 1 do
  begin
    lTobTemp := vTobPiece.Detail[i];
    lTobTemp.PutValue('E_NUMEROPIECE', vInNumeroPiece);
    Inc(lNumLigne,1);
    lTobTemp.PutValue('E_NUMLIGNE', lNumLigne);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/11/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TObjJournalAno.DetruitPieceAno(vTobPiece: Tob);
var i : integer;
begin
  for i := 0 to vTobPiece.Detail.Count-1 do
  begin
    vTobPiece.Detail[i].PutValue('E_GENERAL', VH^.Cpta[fbGene].Attente);
    vTobPiece.Detail[i].PutValue('E_AUXILIAIRE', '');
    vTobPiece.Detail[i].PutValue('E_LIBELLE', TraduireMemoire('Détruite le ' + DateToStr(Date) + ' par la PURGE'));
    vTobPiece.Detail[i].PutValue('E_DEBIT', IIF(i = 0, 1, -1));
    vTobPiece.Detail[i].PutValue('E_CREDIT', 0);
    vTobPiece.Detail[i].PutValue('E_DEBITDEV', 0);
    vTobPiece.Detail[i].PutValue('E_CREDIT', 0);
    vTobPiece.Detail[i].PutValue('E_ENCAISSEMENT', IIF(i =0, 'ENC', 'DEC'));
    vTobPiece.Detail[i].PutValue('E_VALIDE', 'X');
    vTobPiece.Detail[i].PutValue('E_COUVERTURE', 0);
    vTobPiece.Detail[i].PutValue('E_COUVERTUREDEV', 0);
    vTobPiece.Detail[i].PutValue('E_LETTRAGE', '');
    vTobPiece.Detail[i].PutValue('E_LETTRAGEDEV', '-');
    vTobPiece.Detail[i].PutValue('E_TAUXDEV', 1);
    vTobPiece.Detail[i].PutValue('E_QTE1', 0);
    vTobPiece.Detail[i].PutValue('E_QTE2', 0);
    vTobPiece.Detail[i].PutValue('E_ECHE', '-');
    vTobPiece.Detail[i].PutValue('E_ANA', '-');
    vTobPiece.Detail[i].PutValue('E_NUMECHE', 0);
    vTobPiece.Detail[i].PutValue('E_TYPEMVT', 'DIV');
    vTobPiece.Detail[i].PutValue('E_TVAENCAISSEMENT', '-');
    vTobPiece.Detail[i].PutValue('E_CONTREPARTIEGEN', '');
    vTobPiece.Detail[i].PutValue('E_CONTREPARTIEAUX', '');
    vTobPiece.Detail[i].PutValue('E_CREERPAR', 'DET');
    vTobPiece.Detail[i].PutValue('E_DATEMODIF', Date);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/11/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TObjJournalAno.SupprimePieceAno( vInNumeroPiece : integer; vStEtab, vStDevise : string);
var lStsql : string;
begin
  lStSql := 'DELETE FROM ECRITURE WHERE ' +
              'E_EXERCICE = "' + FStCodeExo + '" AND ' +
              'E_JOURNAL = "' + FStJournalAno + '" AND ' +
              'E_DEVISE = "' + vStDevise + '" AND ' +
              'E_ETABLISSEMENT = "' + vStEtab + '" AND ' +
              'E_NUMEROPIECE = ' + IntToStr(vInNumeroPiece);

  // Suppression des écritures du Journal d'A-Nouveaux pour la pièce
  if ExecuteSQL( lStSql ) = 0 then
  begin
    FLastErrorMsg := 'Suppression impossible des écritures du journal d''A-Nouveaux' +#13#10 +
                     'Pièce : ' + IntToStr(vInNumeroPiece) + ' - ' +
                     'Etablissement : ' + vStEtab + ' - ' +
                     'Devise : ' + vStDevise;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/11/2006
Modifié le ... : 20/11/2006
Description .. :
Mots clefs ... :
*****************************************************************}
destructor TObjJournalAno.Destroy;
begin
  FreeAndNil(FTobListePiece);
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////

end.
