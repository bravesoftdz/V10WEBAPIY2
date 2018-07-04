unit UtofRTAnaProsp;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF,GraphUtil,GRS1,TeEngine,Series,Chart,UTOB
{$IFDEF EAGLCLIENT}
     ,MaineAGL
{$ELSE}
     ,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}db,FE_Main
{$ENDIF}
        ;
Type
     TOF_RTAnaProsp = Class (TOF)
     private
        FGraph : TFGRS1 ;
        QQ : TQuery ;
        TobG : Tob ;
        Function  GetNumCol( Nom : string ) : integer ;
        procedure ZoomGraphe (Sender: TCustomChart; Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

     public
        procedure OnArgument(Arguments : String ) ; override ;
        procedure OnLoad ; override ;
        procedure OnUpdate ; override ;
        Procedure OnClose ; override ;

END ;


implementation  

procedure  TOF_RTAnaProsp.OnArgument(Arguments : String ) ;
begin
inherited ;
Fgraph:=TFGRS1(Ecran) ;
SetControlText('AXE',' ');
SetControlText('REGROUPEMENT','RPE_TYPEPERSPECTIV');
end;


procedure  TOF_RTAnaProsp.OnLoad ;
begin
inherited ;
TobG := TOB.Create ('', nil, -1);
end;

procedure  TOF_RTAnaProsp.OnClose ;
begin
inherited ;
TobG.Free ;
end;

Function  TOF_RTAnaProsp.GetNumCol( Nom : string ) : integer ;
var i : integer ;
begin
// DBR : Result := -1;
i := TobG.detail[0].GetNumChamp(Nom) ;
if i >=1000 then result:=i-1000 else result:=i;
end;

procedure  TOF_RTAnaProsp.OnUpdate ;
var stWhere,stAnd, stSQL, Regroupement,stXWhere : string ;
    ColVal1,ColVal2 : Array of Integer ;
    TypeSerie : TChartSeriesClass ;
    TitreGraph1 , TitreGraph2 :string ;
    iAxe : integer ;
begin
inherited ;
TitreGraph1:='' ;
iAxe:=0;

stWhere:='';
stXWhere:=' WHERE ' ;
stAnd:='';
if GetControlText('RPE_INTERVENANT') <> '' then begin stwhere:=stWhere+stXWhere+stAND+' (RPE_INTERVENANT="'+GetControlText('RPE_INTERVENANT')+'")';stAnd:=' AND ' ; stXWhere:='' ; end;
if GetControlText('RPE_TYPETIERS') <> '' then begin stwhere:=stwhere+stXWhere+stAnd+' (RPE_TYPETIERS="'+GetControlText('RPE_TYPETIERS')+'")';stAnd:=' AND '; stXWhere:='' ; end;
if GetControlText('RPE_AUXILIAIRE') <> '' then begin stwhere:=stwhere+stXWhere+stAnd+' (RPE_AUXILIAIRE like "'+GetControlText('RPE_AUXILIAIRE')+'%")';stAnd:=' AND '; stXWhere:='' ; end;
if GetControlText('RPE_ETATPER') <> '' then begin stwhere:=stwhere+stXWhere+stAnd+' (RPE_ETATPER="'+GetControlText('RPE_ETATPER')+'")';stAnd:=' AND '; stXWhere:='' ; end;
if GetControlText('RPE_OPERATION') <> '' then begin stwhere:=stwhere+stXWhere+stAnd+' (RPE_OPERATION like "'+GetControlText('RPE_OPERATION')+'%")';stAnd:=' AND '; stXWhere:='' ; end;

FGraph.FListe.VidePile(False);
Regroupement:=GetControlText('REGROUPEMENT') ;
if Regroupement <> '' then
    begin
    stSQL:='SELECT '+Regroupement+',SUM(RPE_MONTANTPER) as Prevu,sum(RPE_MONTANTPER*(RPE_POURCENTAGE/100)) as Realisation'+
               ' FROM PERSPECTIVES ';
    if stWhere<>'' then stSQL:=stSQL+stWhere;
    stSQL:=stSQL+' GROUP BY '+ Regroupement ;
    end
else
    begin
    stSQL:='SELECT RPE_AUXILIAIRE,RPE_LIBELLE,RPE_INTERVENANT,RPE_MONTANTPER,RPE_POURCENTAGE'+
               ',RPE_MONTANTPER*(RPE_POURCENTAGE/100) as Realisation,RPE_TYPETIERS,RPE_ETATPER,RPE_TYPEPERSPECTIV FROM PERSPECTIVES';
    if stWhere<>'' then stSQL:=stSQL+stWhere;
    end;

QQ := OpenSql (stSQL,TRUE);
TobG.ClearDetail ;
TobG.LoadDetailDB ('', '', '', QQ, False);
ferme(QQ) ;
if Tobg.detail.count=0 then exit;

TobG.PutGridDetail (FGraph.FListe, True, False, '');


TitreGraph2:=TitreGraph1+'(Réalisation)';
if getcontrolText('AXE')<>'' then  iAxe:=GetNumCol(getcontrolText('AXE'));

if Regroupement <> '' then
    begin
    if getcontrolText('CB_PREVU')='X' then begin SetLength( ColVal1, 1); ColVal1[0]:=GetNumCol('Prevu'); end ;
    if getcontrolText('CB_REALISATION')='X' then begin SetLength( ColVal2, 1); ColVal2[0]:=GetNumCol('Realisation'); end ;
    TypeSerie:=TPieSeries ;
    TitreGraph1:=' par '+Rechdom('RTREGROUPPER',Regroupement,False) ;
    end
else
    begin
    if getcontrolText('CB_PREVU')='X' then begin SetLength( ColVal1, 2);ColVal1[0]:=GetNumCol('RPE_MONTANTPER');ColVal1[1]:=GetNumCol('Realisation'); end;
    if getcontrolText('CB_REALISATION')='X' then begin SetLength( ColVal2, 1);ColVal2[0]:=GetNumCol('Realisation'); end ;
    TypeSerie:=TBarSeries ;
    end;

with FGraph do
  begin
  InitGR ('Perspectives');
  FChart1.Title.Text[0] := 'Perspectives';
  UpdateGraph (iAxe, 0, 1, TobG.Detail.Count, False,ColVal1, ColVal2,TypeSerie);

  FChart1.Title.Text.Add(TitreGraph1);
  FChart2.Title.Text.Add(TitreGraph2);

                      {TPieSeries, TLineSeries, TBarSeries}
     if not ParamUtilisateur then
        begin
        if FChart1.Series[0] is TPieSeries then
           begin
           FChart1.AxisVisible:=false;
           //FChart1.frame.visible:=false;
           FChart1.LeftAxis.visible:=false;
           FChart1.TopAxis.visible:=false;
           FChart1.RightAxis.visible:=false;
           FChart1.BottomAxis.visible:=false;
           FChart1.View3DWalls:=false;
           FChart1.View3D:=true;
           FChart1.Series[0].Marks.visible:=true;
           FChart1.Series[0].Marks.style:=smsLabelPercent;
           Fchart1.OnClickSeries:=ZoomGraphe ;
           end;
        if (High(ColVal1)>0) and (FChart1.Series[1] is TBarSeries) then
           begin
           FChart1.Series[1].active:=false;
           end;
        end;
  end;
end;

procedure  TOF_RTAnaProsp.ZoomGraphe (Sender: TCustomChart;
  Series: TChartSeries; ValueIndex: Integer; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin

AGLLanceFiche ('GC','GCCOMMERCIAL','',FGraph.FListe.cells[0,ValueIndex+1],'ACTION=CONSULTATION;MONOFICHE') ;
// PgiInfo ('Données ('+inttostr(ValueIndex)+')= '+ FGraph.FListe.cells[0,ValueIndex+1],'Donnée graphe') ;

end ;

Initialization
registerclasses([TOF_RTAnaProsp]);

end.
