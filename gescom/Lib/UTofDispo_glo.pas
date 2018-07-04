unit UTofDispo_glo;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF,
{$IFDEF EAGLCLIENT}
      UTob,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,db,DBGrids,
{$ENDIF}
      Vierge;
Type
     TOF_DISPO_GLO = Class (TOF)
       private
           LibChampsDim,ChampsDim : TStringList ;
           function TrouveArgumentArticle(Argument: String): string;
           procedure InitListeChampDim ;
           procedure OnChangeListe(Sender: TObject);
       public
           procedure OnLoad ; override ;

     END ;

implementation

procedure TOF_Dispo_Glo.OnLoad;
var Fliste : Tlistbox;
    i : integer ;
    FF : TFVierge;
    QQ : TQuery ;
    SQL,stArticle : string;
    Physique,ReserCli,ReserFou : Double ;
begin
inherited ;
FF := TFVierge(Ecran) ;

if FF is TFVierge then
  begin
    Fliste:=Tlistbox(FF.FindComponent('LDEPOT')) ;
    InitListeChampDim  ;
    for i:=0 to ChampsDim.count-1 do
    begin
         Fliste.items.Add(LibChampsDim[i]) ;
    end;
    If not Assigned(Fliste.OnClick) then
       Fliste.OnClick:=OnChangeListe;


    StArticle:=TrouveArgumentArticle(FF.FArgument);
    setcontroltext('ARTICLE',StArticle) ;


    SQL := 'SELECT SUM(GQ_PHYSIQUE) as S1, SUM(GQ_RESERVECLI) AS S2' ;
    SQL := SQL + ',SUM(GQ_RESERVEFOU) as S3' ;
    SQL := SQL + ' FROM DISPODIM WHERE GQ_ARTICLE="'+StArticle+'"';
    QQ :=OpenSQL(SQL,true) ;
    if not QQ.eof then
       begin
            Physique:=QQ.Findfield('S1').AsFloat ;
            ReserCli:=QQ.Findfield('S2').AsFloat ;
            ReserFou:=QQ.Findfield('S3').AsFloat ;
            setcontroltext('PHY_GLO',FloatToStrF(Physique,ffFixed,20,2)) ;
            setcontroltext('RES_CLIENT_GLO',FloatToStrF(ReserCli,ffFixed,20,2)) ;
            setcontroltext('RES_FOUR_GLO',FloatToStrF(ReserFou,ffFixed,20,2)) ;
            setcontroltext('DISPO_GLO',FloatToStrF(Physique-ReserCli+ReserFou,ffFixed,20,2)) ;
       end
    else
       begin
            setcontroltext('PHY_GLO','') ;
            setcontroltext('RES_CLIENT_GLO','') ;
            setcontroltext('RES_FOUR_GLO','') ;
            setcontroltext('DISPO_GLO','') ;
       end;
       ferme(QQ);

    setcontroltext('PHY_DEP','');
    setcontroltext('RES_CLIENT_DEP','');
    setcontroltext('RES_FOUR_DEP','');
    setcontroltext('DISPO_DEP','');
   end;



end;

procedure TOF_Dispo_Glo.InitListeChampDim ;
begin
if LibChampsDim<>Nil then begin LibChampsDim.free ; LibChampsDim:=nil ; end ;
If ChampsDim<>Nil then begin ChampsDim.free ; ChampsDim:=nil ; end ;
ChampsDim:=TStringList.create ; LibChampsDim:=TStringList.create ;
RemplirValCombo('GCDEPOT','','',LibChampsDim,ChampsDim,False,False) ;
end;

procedure TOF_Dispo_Glo.OnChangeListe(Sender: TObject) ;
var Fliste : Tlistbox;
    FF : TFVierge;
    QQ : TQuery ;
    SQL,StArticle,StDepot : string;
    FlagOk,i : integer;
    Physique,ReserCli,ReserFou : Double ;    
begin
    FF := TFVierge(Ecran);
    Fliste:=Tlistbox(Sender) ;
    FlagOk := 0;   StDepot:='';
    for i := 0 to Fliste.Items.Count - 1 do
    begin
       if Fliste.Selected[i]  then
       begin
          if FlagOk = 1 then
          begin
             StDepot := StDepot+' OR';
          end;
          StDepot := StDepot+' GQ_DEPOT="'+ChampsDim[i]+'"';
          FlagOk := 1;
       end;
    end;


    StArticle := TrouveArgumentArticle(FF.FArgument);
    SQL := 'SELECT SUM(GQ_PHYSIQUE) as S1, SUM(GQ_RESERVECLI) AS S2' ;
    SQL := SQL + ',SUM(GQ_RESERVEFOU) as S3' ;
    SQL := SQL + ' FROM DISPODIM WHERE GQ_ARTICLE="'+StArticle+'"';
    SQL := SQL + ' AND ('+StDepot+')';

    QQ :=OpenSQL(SQL,true) ;
    if not QQ.eof then
       begin
          Physique:=QQ.Findfield('S1').AsFloat ;
          ReserCli:=QQ.Findfield('S2').AsFloat ;
          ReserFou:=QQ.Findfield('S3').AsFloat ;
          setcontroltext('PHY_DEP',FloatToStrF(Physique,ffFixed,20,2)) ;
          setcontroltext('RES_CLIENT_DEP',FloatToStrF(ReserCli,ffFixed,20,2)) ;
          setcontroltext('RES_FOUR_DEP',FloatToStrF(ReserFou,ffFixed,20,2)) ;
          setcontroltext('DISPO_DEP',FloatToStrF(Physique-ReserCli+ReserFou,ffFixed,20,2)) ;
       end
       else
       begin
          setcontroltext('PHY_DEP','0,00');
          setcontroltext('RES_CLIENT_DEP','0,00');
          setcontroltext('RES_FOUR_DEP','0,00');
          setcontroltext('DISPO_DEP','0,00');
       end;
       ferme(QQ);

end;


function TOF_Dispo_Glo.TrouveArgumentArticle(Argument: String): string;
var StArgument : string;
    i : integer;
begin
    StArgument := Argument ;
    i:=Pos('ARTICLE=',StArgument) ;

    if i>0 then
      begin
        system.Delete(StArgument,1,i+7) ;
        Result:=ReadTokenSt(StArgument);
      end
    else
      begin
        Result:='';
      end;

end;
Initialization
registerclasses([TOF_Dispo_glo]);
end.
