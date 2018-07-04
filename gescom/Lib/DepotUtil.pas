unit DepotUtil;

interface

Uses HEnt1, HCtrls, UTOB, Ent1, LookUp, Controls, ComCtrls, StdCtrls, ExtCtrls,
     FactContreM,
{$IFDEF EAGLCLIENT}
     Maineagl,
{$ELSE}
     DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
{$ENDIF}
     SysUtils, Dialogs, EntGC,FactUtil, StockUtil;

Procedure ChargeUnDepot (TOBA : TOB; QuelDepot:string='');
Procedure LoadTOBDispoContreM ( TOBL, TOBCATA : TOB ; Force : boolean; QuelDepot : string='') ;
//function RecupPrixDepot (CodeArticle : String; QuelDepot:string = '') : TOB;
//function RecupPrixDepotLigne (TobA : TOB; stDepot : string) : TOB;

implementation

Procedure ChargeUnDepot (TOBA : TOB; QuelDepot:string='');
Var TSql : TQuery;
    iInd : integer;
    RefUnique, Depot, StWhereDispo, StWhereDispoLot : String;
    TOBDesLots, TOBDL, TOBDepot : TOB;
begin
    if TOBA = Nil then Exit;
    if TobA.FindFirst (['GQ_DEPOT'], [QuelDepot], False) = nil then
    begin
        //if ((TOBA.Detail.Count>0) and (Not Force)) then Exit ;
        {Chargement des dispos}
        //TOBA.ClearDetail;
        RefUnique := TOBA.GetValue ('GA_ARTICLE');
        if QuelDepot = '' then StWhereDispo := '' else StWhereDispo := ' AND GQ_DEPOT="' + QuelDepot + '"';
        TSql := OpenSQL ('Select * from DISPO WHERE GQ_ARTICLE="' + RefUnique + '" AND GQ_CLOTURE="-"' + StWhereDispo, True,-1,'',true);
        if Not TSql.EOF then
        begin
            TOBA.LoadDetailDB ('DISPO', '', '', TSql, True);
            DispoChampSupp (TOBA);
        end;
        Ferme (TSql) ;

        {Chargement des lots}
        if TOBA.GetValue ('GA_LOT') = 'X' then
        begin
            TOBDesLots := TOB.Create ('', Nil, -1);
            if QuelDepot = '' then StWhereDispoLot := '' else StWhereDispoLot := ' AND GQL_DEPOT="' + QuelDepot + '"';
            TSql := OpenSQL ('SELECT * FROM DISPOLOT WHERE GQL_ARTICLE="' + RefUnique + '"' + StWhereDispoLot, True,-1,'',true);
            if Not TSql.EOF then TOBDesLots.LoadDetailDB ('DISPOLOT', '', '', TSql, False);
            Ferme (TSql) ;
            for iInd := TOBDesLots.Detail.Count - 1 downto 0 do
            begin
                TOBDL := TOBDesLots.Detail [iInd];
                Depot := TOBDL.GetValue ('GQL_DEPOT');
                TOBDepot := TOBA.FindFirst (['GQ_DEPOT'], [Depot], False);
                if TOBDepot <> Nil then TOBDL.ChangeParent (TOBDepot, -1);
            end;
            TOBDesLots.Free ;
        end;
    end;
end;

Procedure LoadTOBDispoContreM ( TOBL, TOBCATA : TOB ; Force : boolean; QuelDepot : string='') ;
Var QQ   : TQuery ;
    RefCata,RefFour,CodeDepot : String ;
    CodeTiers,StWhereDispo,WhereAvecCli,WhereSansCli,WhereCli:String ;
    TobDCM:Tob;
BEGIN
if (TOBCATA=Nil) or (TOBL=Nil) then Exit ;
{Chargement des dispos contreM}
RefCata:=TOBCATA.GetValue('GCA_REFERENCE') ;
RefFour:=TOBCATA.GetValue('GCA_TIERS') ;
CodeTiers := GetCodeClientDCM(TOBL);
if Queldepot <> '' then CodeDepot := StringReplace(QuelDepot,'"','',[rfReplaceAll])
else CodeDepot := TOBL.GetValue('GL_DEPOT');

//Recherche si TOBDispoContreM déjà chargé pour ce catalogue et ce client
TOBDCM:=TOBCata.FindFirst( ['GQC_DEPOT','GQC_REFERENCE','GQC_FOURNISSEUR','GQC_CLIENT'],[CodeDepot,RefCata,RefFour,CodeTiers],False);
if TOBDCM<>Nil then
   begin
   if Force then //On efface tout avant de recharger
      begin
      While TOBDCM<>nil do
        Begin
        TOBDCM.free;
        TOBDCM:=TOBCata.FindFirst(['GQC_DEPOT','GQC_REFERENCE','GQC_FOURNISSEUR','GQC_CLIENT'],[CodeDepot,RefCata,RefFour,CodeTiers],False);
        WhereAvecCli:='GQC_CLIENT="'+CodeTiers+'"';
        end;
      end
      else //On ne recharge pas
      WhereAvecCli:='';
   end else WhereAvecCli:='GQC_CLIENT="'+CodeTiers+'"';

//Recherche si TOBDispoContreM déjà chargé pour ce catalogue et sans client d'affecté
TOBDCM:=TOBCata.FindFirst(['GQC_DEPOT','GQC_REFERENCE','GQC_FOURNISSEUR','GQC_CLIENT'],[CodeDepot,RefCata,RefFour,''],False);
if TOBDCM<>Nil then
   begin
   if Force then //On efface tout avant de recharger
      begin
      While TOBDCM<>nil do
        Begin
        TOBDCM.free;
        TOBDCM:=TOBCata.FindFirst( ['GQC_DEPOT','GQC_REFERENCE','GQC_FOURNISSEUR','GQC_CLIENT'],[CodeDepot,RefCata,RefFour,''],False);
        WhereSansCli:='GQC_CLIENT=""';
        end;
      end
      else //On ne recharge pas
      WhereSansCli:='';
   end else WhereSansCli:='GQC_CLIENT=""';

WhereCli:=WhereAvecCli;
if WhereSansCli<>'' then
   if WhereCli<>'' then WhereCli:=WhereCli+' OR '+WhereSansCli else WhereCli:=WhereSansCli;

if WhereCli<>'' then
   begin
   WhereCli:='('+WhereCli+')';
   StWhereDispo:=' AND GQC_DEPOT="'+CodeDepot+'"';
   QQ:=OpenSQL('Select * from DISPOCONTREM WHERE GQC_REFERENCE="'+RefCata+'" AND GQC_FOURNISSEUR="'+RefFour+
            '" AND '+WhereCli+StWhereDispo,True,-1,'',true);
   if (Not QQ.EOF)  then
      BEGIN
      TOBCATA.LoadDetailDB('DISPOCONTREM','','',QQ,true) ;
      TOBCATA.Detail[0].AddChampSupValeur('RESERVECLI', 0, True);
      END;
   Ferme(QQ) ;
   end;
END ;

{ function RecupPrixDepot (CodeArticle : String; QuelDepot:string = '') : TOB;
Var QPrix : TQuery ;
    SQL : String ;
    TobDisp : TOB;
begin
    TobDisp := Tob.Create ('DISPO', nil, -1);
    if QuelDepot = '' then QuelDepot := VH_GC.GCDepotDefaut ;
// Force le code article sur 18 positions minimums (unicité)
    CodeArticle := CodeArticle + StringOfChar( ' ', 18 - Length( CodeArticle ));
// Retourne les prix de l'article si il est unique
// sinon une moyennne
    SQL := 'SELECT A.NB NB_ARTICLE, SUM(GQ_PMAP)/A.NB GQ_PMAP, SUM(GQ_PMRP)/A.NB GQ_PMRP,' +
            'SUM(GQ_DPA)/A.NB GQ_DPA, SUM(GQ_DPR)/A.NB GQ_DPR ' +
            'FROM DISPO, (SELECT COUNT(*) NB FROM DISPO ' +
                 'WHERE GQ_ARTICLE LIKE "' + CodeArticle + '%" ' +
                 'AND GQ_CLOTURE="-" ' +
                 'AND GQ_DEPOT = "' + QuelDepot + '") A ' +
            'WHERE GQ_ARTICLE like "' + CodeArticle + '%" ' +
            'AND GQ_CLOTURE="-" ' +
            'AND GQ_DEPOT = "' + QuelDepot + '" ' +
            'GROUP BY A.NB';
    QPrix := OpenSQL (SQL, True);
    if Not QPrix.EOF then TobDisp.SelectDB ('', QPrix)
    else TobDisp.AddChampSupValeur ('NB_ARTICLE', 0, False);
    Ferme (QPrix) ;
    Result := TobDisp;
end ;

function RecupPrixDepotLigne (TobA : TOB; stDepot : string) : TOB;
begin
    Result := RecupPrixDepot (TobA.GetValue ('GA_ARTICLE'), stDepot);
    if Result.GetValue ('NB_ARTICLE') = 0 then // Attention que RecupPrixDepot retourne bien toujours une Tob !
    begin
        Result.Free;
        Result := RecupPrixDepot (TOBA.GetValue('GA_ARTICLE'));
    end;
end ; }

end.
