unit MajPontTax;

//Le ponts suivant est traités dans ce source :
//Pont ancienne table des taux -> nouvelle table des taux

interface

uses
  Classes,
  Dialogs,
  cbpType;

Type
  IMajPontTax = Interface
  ['{79B7AF2D-2257-4588-A55D-0FD9722F583B}']
    function TransferTaxOldToNew : boolean;
  end;



  function CreerObjetTransfertTax(const CodePays : hstring='') : IMajPontTax;


implementation
uses
  Windows,SysUtils,
{$IFNDEF EAGLCLIENT}
  uDbxDataSet,
{$ENDIF EAGLCLIENT}
  Hctrls,
  HEnt1,
  UTOB;

Type
  //Pont ancienne table des taux -> nouvelle table des taux
  //Pont nouvelle table des taux -> ancienne table des taux
  TPontTax = class(TInterfacedObject, IMajPontTax)
  private
    FTAbleTax : string;
    FPrefTax : string;
    FPaysFranceReference : hstring;
    FPaysFrance : hstring;
    function TreatOneEnregTaxOldToNew (FTobTaxNew, FTobTaxOld  : Tob) : boolean;
  public
    constructor Create (const CodePays : hstring = '');
    destructor Destroy; override;
    function   TransferTaxOldToNew : boolean;
  end;

//===============================================================================
//===============================================================================
{ TPontTax }

constructor TPontTax.Create  (const CodePays : hstring = '');
begin
  inherited Create;
  FTAbleTax:='YYTXTAUX';
  FPrefTax := 'YTD';
  FPaysFranceReference := '250';
  FPaysFrance := CodePays;
end;

destructor TPontTax.Destroy;
begin
  inherited;
end;

function TPontTax.TransferTaxOldToNew: boolean;
Var
  FTobOldInfoTax : Tob;
  FTobNewInfoTax : Tob;
  FTobCherche    : Tob;
  StSQL : hstring;
  OQuery : TQuery;
  bError : Boolean;
  ind : integer;
begin
  Result := False;
  if FPaysFrance <> FPaysFranceReference then
  begin
    Result := True;
    exit;
  end;

  bError := False;

  //Chargement des données de TXCPTTVA
  try
    try
      FTobOldInfoTax := Tob.Create('TXCPTTVA', NIL, - 1 );
      StSQL := 'SELECT * FROM TXCPTTVA WHERE TV_TVAOUTPF IN ("TX1", "TX2")';
      OQuery := OpenSQL (StSQL,True) ;
      if not OQuery.Eof then
        FTobOldInfoTax.LoadDetailDB('TXCPTTVA', '', '', OQuery, False);
    except
      bError := True;
    end;
  finally
    Ferme (OQuery);
  end;

  //Si erreur dans le chargement des données de TXCPTTVA, sortie de la fonction (la fonction renvoie False)
  if bError then
  begin
    FreeAndNil(FTobOldInfoTax);
    exit;
  end;

  //Chargement des données de la nouvelle table des taux et comptes
  //du pays FRANCE (code pays = 250)
  try
    try
      FTobNewInfoTax := Tob.Create(FTAbleTax, NIL, - 1 );
      StSQL := 'SELECT * FROM ' +FTAbleTax+ ' WHERE ' +
               FPrefTax+ '_CATEG IN ("TX1", "TX2") AND ' +
               FPrefTax+ '_CODEPAYS = "' + FPaysFrance + '"';
      OQuery := OpenSQL (StSQL,True) ;
      if not OQuery.Eof then
        FTobNewInfoTax.LoadDetailDB(FTAbleTax, '', '', OQuery, False);
    except
      bError := True;
    end;
  finally
    Ferme (OQuery);
  end;

  //Si erreur dans le chargement des données de la nouvelle table des taux et comptes,
  //sortie de la fonction (la fonction renvoie False)
  if bError then
  begin
    FreeAndNil(FTobOldInfoTax);
    FreeAndNil(FTobNewInfoTax);
    exit;
  end;

  //Si le chargement des données s'est bien passée, continuation du traitement
  if FTobOldInfoTax.Detail.Count = 0 then
  begin
    // Aucun enregistrement trouvé dans la table TXCPTTVA
    FreeAndNil(FTobOldInfoTax);
    FreeAndNil(FTobNewInfoTax);
    Result := True;
    exit;
  end;
  for ind := 0 to FTobOldInfoTax.Detail.Count - 1 do
  begin
    FTobCherche := FTobNewInfoTax.FindFirst([FPrefTax+'_CODEPAYS',
                   FPrefTax+'_CATEG', FPrefTax+'_VALEUR01',
                   FPrefTax+'_VALEUR02'],
                   [FPaysFrance, FTobOldInfoTax.Detail[ind].GetValue('TV_TVAOUTPF'),
                   FTobOldInfoTax.Detail[ind].GetValue('TV_REGIME'),
                   FTobOldInfoTax.Detail[ind].GetValue('TV_CODETAUX')], True);
    if FtobCherche = nil then
    begin
      FtobCherche:=Tob.Create(FTableTax,FTobNewInfoTax,-1);
      FtobCherche.InitValeurs(False);
    end ;
    bError := TreatOneEnregTaxOldToNew (FTobCherche, FTobOldInfoTax.Detail[ind]);
    if bError then
    begin
      // Erreur lors de la mise à jour des enregistrements taxes de la table
      FreeAndNil(FTobOldInfoTax);
      FreeAndNil(FTobNewInfoTax);
      exit;
    end;
  end;

  //Vérification des éléments supprimés dans la nouvelle table pour suppression
  //dans l'ancienne table des taux
  for ind := 0 to FTobNewInfoTax.Detail.Count - 1 do
  begin
    FTobCherche := FTobOldInfoTax.FindFirst(['TV_TVAOUTPF',
                   'TV_REGIME',
                   'TV_CODETAUX'],
                   [FTobNewInfoTax.Detail[ind].GetValue(FPrefTax+'_CATEG'),
                   FTobNewInfoTax.Detail[ind].GetValue(FPrefTax+'_VALEUR01'),
                   FTobNewInfoTax.Detail[ind].GetValue(FPrefTax+'_VALEUR02')], True);

    if (FTobCherche = nil) and
       ((FTobNewInfoTax.Detail[ind].GetValue(FPrefTax+'_CATEG') = 'TX1') or
       (FTobNewInfoTax.Detail[ind].GetValue(FPrefTax+'_CATEG') = 'TX2')) then
    begin
      FTobNewInfoTax.Detail[ind].DeleteDB; //gestion de la supression d'enregistrements
    end;
  end;

  FreeAndNil(FTobOldInfoTax);
  FreeAndNil(FTobNewInfoTax);
  Result := True;
end;

function TPontTax.TreatOneEnregTaxOldToNew(FTobTaxNew, FTobTaxOld: Tob): boolean;
var
  strCleConcat : hstring;
begin
  Result := False;
  try
    FTobTaxNew.PutValue(FPrefTax+ '_CODEPAYS', FPaysFrance);
    FTobTaxNew.PutValue(FPrefTax+ '_CATEG',FTobTaxOld.GetValue('TV_TVAOUTPF'));
    FTobTaxNew.PutValue(FPrefTax+ '_VALEUR01',FTobTaxOld.GetValue('TV_REGIME'));
    FTobTaxNew.PutValue(FPrefTax+ '_VALEUR02',FTobTaxOld.GetValue('TV_CODETAUX'));
    //Traitement soumis TPF
    if FTobTaxOld.GetValue('TV_TVAOUTPF') = 'TX2' then
    begin
      FTobTaxNew.PutValue(FPrefTax+ '_VALEUR03','X');
    end;
    strCleConcat := FTobTaxNew.GetValue(FPrefTax+ '_VALEUR01') + ';' +
                  FTobTaxNew.GetValue(FPrefTax+ '_VALEUR02') + ';' +
                  FTobTaxNew.GetValue(FPrefTax+ '_VALEUR03') +';;;';
    FTobTaxNew.PutValue(FPrefTax+ '_IDNOMCHPSPREF', strCleConcat);
    //Libellé
    if FTobTaxOld.GetValue('TV_TVAOUTPF') = 'TX1' then
    begin
      FTobTaxNew.PutValue(FPrefTax+ '_LIBELLE', 'TVA');
    end
    else if FTobTaxOld.GetValue('TV_TVAOUTPF') = 'TX2' then
    begin
      FTobTaxNew.PutValue(FPrefTax+ '_LIBELLE', 'TPF');
    end;
    //Gestion taux
    FTobTaxNew.PutValue(FPrefTax+ '_TXMT', 'X');
    //Taux achat et vente appliqué
    FTobTaxNew.PutValue(FPrefTax+ '_TAUXACH',FTobTaxOld.GetValue('TV_TAUXACH'));
    FTobTaxNew.PutValue(FPrefTax+ '_TAUXVTE',FTobTaxOld.GetValue('TV_TAUXVTE'));
    //Taux achat et vente applicable
    FTobTaxNew.PutValue(FPrefTax+ '_TAUXACHDEF',FTobTaxOld.GetValue('TV_TAUXACH'));
    FTobTaxNew.PutValue(FPrefTax+ '_TAUXVTEDEF',FTobTaxOld.GetValue('TV_TAUXVTE'));
    //Compte facture achat et vente
    FTobTaxNew.PutValue(FPrefTax+ '_CHP2CPTA',FTobTaxOld.GetValue('TV_CPTEACH'));
    FTobTaxNew.PutValue(FPrefTax+ '_CHP1CPTA',FTobTaxOld.GetValue('TV_CPTEVTE'));
    //Compte avoir achat et vente
    FTobTaxNew.PutValue(FPrefTax+ '_CHP4CPTA',FTobTaxOld.GetValue('TV_CPTEACH'));
    FTobTaxNew.PutValue(FPrefTax+ '_CHP3CPTA',FTobTaxOld.GetValue('TV_CPTEVTE'));
    //Compte encaissement achat et vente
    FTobTaxNew.PutValue(FPrefTax+ '_CHP6CPTA',FTobTaxOld.GetValue('TV_ENCAISACH'));
    FTobTaxNew.PutValue(FPrefTax+ '_CHP5CPTA',FTobTaxOld.GetValue('TV_ENCAISVTE'));
    //Compte avoir encaissement achat et vente
    FTobTaxNew.PutValue(FPrefTax+ '_CHP8CPTA',FTobTaxOld.GetValue('TV_ENCAISACH'));
    FTobTaxNew.PutValue(FPrefTax+ '_CHP7CPTA',FTobTaxOld.GetValue('TV_ENCAISVTE'));
    // GC_20090828_DM_GC16995 DEBUT
    // Compte FAR/FAE Vente et Achat
    FTobTaxNew.PutValue(FPrefTax+ '_CHP10CPTA',FTobTaxOld.GetValue('TV_CPTACHFARFAE'));
    FTobTaxNew.PutValue(FPrefTax+ '_CHP9CPTA',FTobTaxOld.GetValue('TV_CPTVTEFARFAE'));
    // GC_20090828_DM_GC16995 FIN
    FTobTaxNew.InsertOrUpdateDB;
  except
    Result := True;
  end;
end;

function CreerObjetTransfertTax(const CodePays : hstring='') : IMajPontTax;
begin
  Result:= TPontTax.Create(CodePays);
end;


Initialization

finalization

end.
