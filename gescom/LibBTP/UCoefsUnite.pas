unit UCoefsUnite;

interface
Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$else}
{$ENDIF}
     uTob,
     Hctrls,
     sysutils,
     HEnt1;
type

	TChangeUnite = class (TObject)
  private
  	fTOBGestUnite : TOB;
    fTOBUnites : TOB;
    //
    procedure ChargeUnites;
		procedure CreeTableCouples;
    procedure CreeCouple (TOBU1,TOBU2 : TOB);
    procedure LoadCoefs;
    //
  public
    constructor create;
    destructor destroy; override;
    property TOBunites : TOB read fTOBUnites;
    property TOBCoefs : TOB read fTOBGestUnite;
    function GetCoef (UniteDepart,UniteFin : string) : double;
    procedure SetCoef (UniteDepart,uniteFin : string; valeur : Double);
  end;


implementation

procedure TChangeUnite.ChargeUnites;
VAR QQ : TQuery;
begin
  QQ := OpenSQL('SELECT * FROM MEA WHERE GME_QUALIFMESURE="PIE"',True,-1,'',true);
  if Not QQ.Eof then fTOBUnites.LoadDetailDB('MEA','','',QQ,false);
  ferme (QQ);
end;


procedure TChangeUnite.CreeTableCouples;
var indice,Col : Integer;
		TOBU1,TOBU2 : TOB;
begin
	for Indice := 0 to fTOBUnites.detail.Count -1 do
  begin
    TOBU1 := fTOBUnites.detail[Indice];
		for Col := 0 to fTOBUnites.Detail.count -1 do
    begin
    	TOBU2 := fTOBUnites.detail[col];
			CreeCouple (TOBU1,TOBU2);
    end;
  end;
end;

procedure TChangeUnite.CreeCouple(TOBU1, TOBU2: TOB);
var TOBCC : TOB;
begin
	TOBCC := TOB.Create('BCHANGEUNITE',fTOBGestUnite,-1);
  TOBCC.SetString('BCU_UNITEDEPART',TOBU1.GetString('GME_MESURE'));
  TOBCC.SetString('BCU_UNITEARRIVE',TOBU2.GetString('GME_MESURE'));
  TOBCC.SetDouble('BCU_COEFCHANGE',1.0);
  TOBCC.AddChampSupValeur('LIBUNITEDEPART',RechDom('GCQUALUNITPIECE',TOBU1.GetString('GME_MESURE'),false)) ;
  TOBCC.AddChampSupValeur('LIBUNITEARRIVE',RechDom('GCQUALUNITPIECE',TOBU2.GetString('GME_MESURE'),false)) ;
end;

constructor TChangeUnite.create;
begin
  fTOBGestUnite := TOB.Create ('LES COUPLES UNITE',nil,-1);
  fTOBUnites 	 := TOB.Create('LES UNITES',nil,-1);
  ChargeUnites;
  CreeTableCouples;
  LoadCoefs;
end;

destructor TChangeUnite.destroy;
begin
  fTOBGestUnite.free;
  fTOBUnites.Free;
  inherited;
end;

procedure TChangeUnite.LoadCoefs;
var TLocals,TOBL : TOB;
		QQ : TQuery;
    Indice : Integer;
begin
	TLocals := TOB.Create('LES COUPLES',nil,-1);
  QQ := OpenSQL('SELECT * FROM BCHANGEUNITE',True,-1,'',true);
  if Not QQ.Eof then TLocals.LoadDetailDB('CHANGEUNITE','','',QQ,false);
  ferme (QQ);
  for Indice := 0 to TLocals.detail.count -1 do
  begin
		TOBL := fTOBGestUnite.FindFirst(['BCU_UNITEDEPART','BCU_UNITEARRIVE'],
    																[TLocals.detail[Indice].GetString('BCU_UNITEDEPART'),
                                     TLocals.detail[Indice].GetString('BCU_UNITEARRIVE')],true);
    if TOBL <> nil then
    begin
			TOBL.SetDouble('BCU_COEFCHANGE',TLocals.detail[Indice].GetDouble('BCU_COEFCHANGE'));
    end;
  end;
  TLocals.free;
end;

function TChangeUnite.GetCoef(UniteDepart, UniteFin: string): double;
var TOBL : TOB;
begin
	Result := 1;
  TOBL := fTOBGestUnite.FindFirst(['BCU_UNITEDEPART','BCU_UNITEARRIVE'],[UniteDepart,UniteFin],true);
  if TOBL <> nil then
  begin
    Result := TOBL.GetDouble('BCU_COEFCHANGE');
  end;
end;

procedure TChangeUnite.SetCoef(UniteDepart, uniteFin: string;valeur: Double);
var TOBL : TOB;
begin
  TOBL := fTOBGestUnite.FindFirst(['BCU_UNITEDEPART','BCU_UNITEARRIVE'],[UniteDepart,UniteFin],true);
  if TOBL <> nil then
  begin
    TOBL.SetDouble('BCU_COEFCHANGE',valeur);
  end;
  TOBL := fTOBGestUnite.FindFirst(['BCU_UNITEDEPART','BCU_UNITEARRIVE'],[UniteFin,UniteDepart],true);
  if TOBL <> nil then
  begin
    TOBL.SetDouble('BCU_COEFCHANGE',ARRONDI(1/valeur,4));
  end;

end;

end.
