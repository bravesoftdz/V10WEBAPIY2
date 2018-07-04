unit UTofDispo_Dim;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF,
{$IFDEF EAGLCLIENT}
      emul,UTob,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}db,DBGrids, mul,
{$ENDIF}
      HDimension,EntGC;
Type
     TOF_DISPO_DIM = Class (TOF)
        procedure OnUpdate ; override ;
        procedure OnArgument(Arguments : String ) ; override ;
     END ;

implementation

procedure TOF_Dispo_Dim.OnUpdate ;
var i:integer ;
    FF:TFMul ;
    C1:THLAbel ;
    QQ : TQuery ;
    SQL : string;
    Physique,ReserCli,ReserFou : Double ;
begin
inherited ;
FF:=TFMul(Ecran) ;
{$IFDEF EAGLCLIENT}
for i:=0 to FF.FListe.ColCount-1 do
    begin
    if copy(FF.FListe.Cells[i,0],1,3)='DIM' then
        begin
        C1:=THLabel(FF.FindComponent('LGRILLE'+copy(FF.FListe.Cells[i,0],4,1))) ;
        if Not C1.Visible then FF.Fliste.colwidths[i]:=0 ;
        FF.Fliste.cells[i,0]:=C1.Caption ;
        end;
    end;
{$ELSE}
for i:=0 to FF.FListe.Columns.Count-1 do
    begin
    if copy(FF.FListe.Columns[i].Title.caption,1,3)='DIM' then
        begin
        C1:=THLabel(FF.FindComponent('LGRILLE'+copy(FF.FListe.Columns[i].Title.caption,4,1))) ;
        FF.Fliste.columns[i].visible:=C1.visible ;
        FF.Fliste.columns[i].Field.DisplayLabel:=C1.Caption ;
        end;
    end;
{$ENDIF}

    SQL := 'SELECT SUM(GQ_PHYSIQUE) as S1, SUM(GQ_RESERVECLI) AS S2' ;
    SQL := SQL + ',SUM(GQ_RESERVEFOU) as S3' ;
    SQL := SQL + ' FROM DISPODIM WHERE '+FF.Q.Criteres;
    QQ :=OpenSQL(SQL,true) ;
    if not QQ.eof then
        begin
        Physique:=QQ.Findfield('S1').AsFloat ;
        ReserCli:=QQ.Findfield('S2').AsFloat ;
        ReserFou:=QQ.Findfield('S3').AsFloat ;
        setcontroltext('SUM_PHYS',FloatToStrF(Physique,ffFixed,20,2)) ;
        setcontroltext('SUM_RCLI',FloatToStrF(ReserCli,ffFixed,20,2)) ;
        setcontroltext('SUM_RFOU',FloatToStrF(ReserFou,ffFixed,20,2)) ;
        setcontroltext('SUM_QDIS',FloatToStrF(Physique-ReserCli+ReserFou,ffFixed,20,2)) ;
       end
    else
        begin
        setcontroltext('SUM_PHYS','') ;
        setcontroltext('SUM_RCLI','') ;
        setcontroltext('SUM_RFOU','') ;
        setcontroltext('SUM_QDIS','') ;
       end;
    ferme(QQ);
    
end;
procedure TOF_Dispo_Dim.OnArgument(Arguments : String ) ;
var  critere: string;
     ChampMul,ValMul : string;
     CodeArticle,DimMAsque,LibelleArticle : string;
      CC : TComponent;
      x,i:integer;
      C1:THValComboBox;
      C2:THLabel;
      QMasque : TQuery ;
      FF : TFMUL;
begin
inherited ;
FF:=TFMul(Ecran) ;
Repeat
 Critere:=uppercase(Trim(ReadTokenSt(Arguments))) ;
 if Critere<>'' then
    begin
    x:=pos('=',Critere);
    if x<>0 then
       begin
       ChampMul:=copy(Critere,1,x-1);
       ValMul:=copy(Critere,x+1,length(Critere));
       end;
       CC:=FF.FindComponent(ChampMul);
       If (CC is THEdit) then THEdit(CC).Text:=ValMul;
       If (CC is THValComboBox) then THValComboBox(CC).Value:=ValMul;
       if (CC is TCheckBox) then
           begin
           if (ValMul='X') then TCheckBox(CC).state:=cbChecked;
           if (ValMul='-') then TCheckBox(CC).state:=cbUnChecked;
           if (ValMul=' ') then TCheckBox(CC).state:=cbGrayed;
           end;
       if (CC is THLabel) then THLabel(CC).caption:=ValMul;

       if ChampMul='GA_DIMMASQUE' then DimMasque := ValMul;
       if ChampMul='GA_LIBELLE' then LibelleArticle := ValMul;
    end;
until  Critere='';

CodeArticle:= THEdit(FF.FindComponent('GA_CODEARTICLE')).Text;
FF.caption:='Dispo de l''article '+CodeArticle+' '+LibelleArticle;

If (DimMasque<>'') then
    BEGIN
    if ctxMode in V_PGI.PGIContexte
    then QMasque:=OpenSQL('SELECT GDM_TYPE1,GDM_TYPE2,GDM_TYPE3,GDM_TYPE4,GDM_TYPE5 FROM DIMMASQUE WHERE GDM_MASQUE="'+DimMasque+'" and GDM_TYPEMASQUE="'+VH_GC.BOTypeMasque_Defaut+'"',TRUE)
    else QMasque:=OpenSQL('SELECT GDM_TYPE1,GDM_TYPE2,GDM_TYPE3,GDM_TYPE4,GDM_TYPE5 FROM DIMMASQUE WHERE GDM_MASQUE="'+DimMasque+'"',TRUE) ;
    if not QMasque.EOF then
        begin
        for i:=1 to MaxDimension do
            begin
            C1:=THValComboBox(FF.FindComponent('GA_CODEDIM'+InttoStr(i)));
            C2:=THLabel(FF.FindComponent('LGRILLE'+InttoStr(i)));
            if (QMasque.Findfield('GDM_TYPE'+IntToStr(i)).AsString<>'') then
                begin
                C1.Visible:=TRUE;
                C1.DataType:='GCDIMENSION' ;
                C1.Plus:='GDI_TYPEDIM="DI'+InttoStr(i)+'" AND GDI_GRILLEDIM="'+QMasque.FindField('GDM_TYPE'+IntToStr(i)).AsString+'" ORDER BY GDI_RANG' ;
                C2.Visible:=TRUE;
                C2.Caption:= RechDom('GCGRILLEDIM'+IntToStr(i),QMasque.Findfield('GDM_TYPE'+IntToStr(i)).AsString,FALSE) ;   //metre TRUE qd OK
                end
            else begin
                C1.Visible:=FALSE;
                C2.Visible:=FALSE;
                end;
            end;
        end;
    Ferme(QMasque);
    end else
    for i:=1 to MaxDimension do
        begin
        C1:=THValComboBox(FF.FindComponent('GA_CODEDIM'+InttoStr(i)));
        C2:=THLabel(FF.FindComponent('LGRILLE'+InttoStr(i)));
        C1.Visible:=FALSE;
        C2.Visible:=FALSE;
        end;
end;


Initialization
registerclasses([TOF_Dispo_dim]);
end.
