{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 17/03/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : QBPLOI (QBPLOI)
Mots clefs ... : TOM;QBPLOI
*****************************************************************}
Unit QBPLOI_TOM ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     FE_MAIn,
     Fiche,
     DBCtrls,
{$else}
     eFiche,
     MainEagl,
     UtileAGL,
{$ENDIF}
     sysutils,
     HCtrls,
     HEnt1,
     UTOM,
     UTob,HDB,UGraph,SynScriptBP;

Type
  TOM_QBPLOI = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
     OkNew,OkValid:boolean;
     code,DateCell,codeJoinR:hString;
     OldMaille:string;
     ValCell:double;
     LigneSolde,ColSolde:integer;
     QT: TQTob;
     procedure RemplitGrid;
     procedure RemplitGridHisto;
     procedure MAJTableGrid;
     function  ChercheCodeRestrictionLoi:hString;
     procedure BtnHisto(Sender: TObject);         
     procedure BtnSolde(Sender: TObject);
     procedure RecalculTot(F:THGrid);
     procedure CellExit(Sender: TObject; var ACol,ARow: Longint;
               var Cancel: Boolean);     
     procedure CellEnter(Sender: TObject; var ACol,ARow: Longint;
               var Cancel: Boolean);
     procedure BtnGraph(Sender: TObject);
     procedure ChangeMaille( Sender: TObject );
     { EVI / Force le statut de la fiche pour passer dans le UpdateRecord }
     procedure BValider_OnClick(Sender: TObject);
    end ;


Implementation
uses HTB97,  UUtil,
     BPEclatement,
     QUFGBPLOI_TOF,
     BPBasic,
     BPFctSession,
     HMsgBox,BPUtil;

const
  TexteMessage : array[1..2] of string =  (
    {1}'Pour le calcul d''initialisation global : %s ',
    {2}'Il n''est pas possible de définir des lois à la semaine pour une période supérieure à 1 an.'
    ) ;


procedure TOM_QBPLOI.OnNewRecord ;
begin
  Inherited ;
  OkNew:=true;
  OldMaille:='NEW';
  RemplitGrid;
end ;

procedure TOM_QBPLOI.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_QBPLOI.OnUpdateRecord ;
begin
  Inherited ;
  MAJTableGrid;
  if not OkValid
   then LastError:=1
   else LastError:=0;
  LastErrorMsg:=TraduireMemoire('La somme des pourcentages doit être égale à 100.');
end ;

procedure TOM_QBPLOI.OnAfterUpdateRecord ;
var codeW:hString;
begin
  Inherited ;
  if BPOkOrli
   //-----------------> ORLI
   then codeW:=''
   //ORLI <-----------------
   else codeW:=' QBO_CODESESSION="'+getfield('QBO_CODESESSION')+'" AND ';
  if OkValid
   then MExecuteSql('UPDATE QBPLOI SET '+code+
                   ' WHERE '+codeW+' QBO_CODELOI="'+getfield('QBO_CODELOI')+'"',
                   'QBPLOI_TOM (OnAfterUpdateRecord).');
end ;

procedure TOM_QBPLOI.RecalculTot(F:THGrid);
var i:Integer ;
    somme:double ;
begin
 somme:=0 ;
 for i:=0 to F.RowCount-1 do
  somme:=somme+Valeur(F.Cells[1,i]) ;
 SetcontrolText('EDTTOTAL',Format('%10.4f',[somme]));
 SetcontrolText('EDTSOLDE',Format('%10.4f',[100-somme]));
end ;

procedure TOM_QBPLOI.CellExit(Sender: TObject; var ACol,ARow: Longint;
          var Cancel: Boolean);
var G:THGrid ;
begin
 G:=THGrid(Sender) ;
 RecalculTot(G) ;
 
end;


procedure TOM_QBPLOI.MAJTableGrid;
var i:integer;
    {$IFDEF MODE} val:hString; //EVI_TEMP_V800
    {$ELSE} val:String; {$ENDIF}
    somme:double;
begin
 code:='';
 somme:=0;

 for i := 0 to QT.laTob.Detail.Count-1 do
  begin
   QT.laTob.GetLigneGrid(THGrid(getcontrol('GRIDLOI')),i,'LADATE;VAL1');
   val:=QT.laTob.getValue('VAL1');
   somme:=somme+VALEUR(val);
   RemplaceVirguleParPoint(val);
   if code=''
    then code:=' QBO_BPM'+IntToStr(i+1)+'="'+val+'" '
    else code:=code+', QBO_BPM'+IntToStr(i+1)+'="'+val+'" ';
  end;


 OkValid:=false;
 if (VALEUR(Format('%10.4f',[somme]))=100)
  then OkValid:=true;
end;

procedure TOM_QBPLOI.RemplitGrid;
var datedeb,datefin:TDateTime;
    aD,numsemD,aF,numsemF,diff:integer;
    intermediaire : integer;
    yF,mF,dF,yD,mD,dD:word;
    i,II,JJ:integer;
    val,somme,valChp,ecart:double;
    perD,perF:TperiodeP;
    numTD,anTD,numTF,anTF:integer;
begin
 if BPOkOrli
  then
   //-----------------> ORLI
   begin
    datedeb:=StrToDateTime(GetcontrolText('QBO_DATEDEBECLAT'));
    datefin:=StrToDateTime(GetcontrolText('QBO_DATEFINECLAT'));
   end
   //ORLI <-----------------
  else
   begin
    datedeb:=StrToDateTime(GetcontrolText('EDTDATEDEB'));
    datefin:=StrToDateTime(GetcontrolText('EDTDATEFIN'));
   end;
 numsemD:= numSemaine(datedeb,aD);
 DecodeDate(datedeb,yD,mD,dD);
 numsemF:= numSemaine(datefin,aF);
 DecodeDate(datefin,yF,mF,dF);

 if QT <> nil
  then QT.Free;
 QT :=TQTob.create(['LADATE'],['VAL1']);
 THGrid(getcontrol('GRIDLOI')).RowCount:=0;

 somme:=0;

 { EVI / Cas semaine avec une période de session supérieur à 1 an -> Exit }
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='2'
 then
 begin
   if (datefin - datedeb) > 367 then
   begin
     PGIINFO(Format(TexteMessage[1] + TexteMessage[2],[#13#10]));
     THDBValComboBox(getcontrol('QBO_MAILLE')).Value := '3';
     Exit;
   end;
 end;

 //cas semaine
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='2'
  then
   begin
    if aF=aD
     then diff:=numsemF-numsemD+1
     else diff:=(52-numsemD)+numsemF;
    if datedeb<10
     then diff:=53;
    if diff>=53
     then diff:=52;

    if OkNew
     then
      begin
       val:=(1/diff)*100;
       ecart:=0;
       for i:=1 to diff do
        begin
         valChp:=VALEUR(Format('%10.4f',[val+ecart]));
         setField('QBO_BPM'+IntToStr(i),valChp);
         ecart:=ecart+(val-valChp);
        end;
       for i:=diff+1 to 52 do
        setField('QBO_BPM'+IntToStr(i),0);
      end;
    i:=1;
    while (i<=diff) do
     begin
      QT.addValeur([DonneMailleSuivantDelai('2',datedeb,i-1)], [getfield('qbo_bpm'+IntToStr(i))]);
      somme:=somme+getfield('qbo_bpm'+IntToStr(i));
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');
   end;

 { EVI / Modification pour lois avec période supérieure à 2 ans }
 //cas quinzaine
if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='3'
  then
   begin
     diff:=0;
     if yF=yD then diff:=mF-mD+1
     else
     begin
       for intermediaire := (yD+1) to (yF-1) do diff := diff + 12;
       diff:= diff+(12-mD)+mF;
     end;

    diff:=diff*2;

    if OkNew
     then
      begin
       val:=(1/diff)*100;
       ecart:=0;
       for i:=1 to diff do
        begin
         valChp:=VALEUR(Format('%10.4f',[val+ecart]));
         setField('QBO_BPM'+IntToStr(i),valChp);
         ecart:=ecart+(val-valChp);
        end;
       for i:=diff+1 to 52 do
        setField('QBO_BPM'+IntToStr(i),0);
      end;
    i:=1;
    JJ:=0;
    while (i<=diff/2) do
     begin
      if dD<16
       then
        begin
         if BPOkOrli
          then
           //-----------------> ORLI
           begin
            II:=(2*i)-2;
            JJ:=(2*i)-1;
           end
           //ORLI <-----------------
          else
           begin
            II:=(i-1);
            JJ:=jj+1;
           end;
         QT.addValeur([DonneMailleSuivantDelai('3',datedeb,II)],[getfield('qbo_bpm'+IntToStr(JJ))]);
         if BPOkOrli
          then
           //-----------------> ORLI
           begin
            II:=(2*i)-1;
            JJ:=(2*i);
           end
           //ORLI <-----------------
          else
           begin
            II:=(i-1);
            JJ:=jj+1;
           end;
        QT.addValeur([DonneMailleSuivantDelai('5',datedeb,II)],[getfield('qbo_bpm'+IntToStr(JJ))]);
        end
       else
        begin
         if BPOkOrli
          then
           //-----------------> ORLI
           begin
            II:=(2*i)-2;
            JJ:=(2*i)-1;
           end
           //ORLI <-----------------
          else
           begin
            II:=(i-1);
            JJ:=i;
           end;
         QT.addValeur([DonneMailleSuivantDelai('5',datedeb,II)], [getfield('qbo_bpm'+IntToStr(JJ))]);
         if BPOkOrli
          then
           //-----------------> ORLI
           begin
            II:=(2*i)-1;
            JJ:=(2*i);
           end
           //ORLI <-----------------
          else
           begin
            II:=(i-1);
            JJ:=i+1;
           end;
         QT.addValeur([DonneMailleSuivantDelai('3',datedeb,II)], [getfield('qbo_bpm'+IntToStr(JJ))]);
        end;

      somme:=somme+getfield('qbo_bpm'+IntToStr(jj-1))+getfield('qbo_bpm'+IntToStr(jj));
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),false,false,'');
   end;

 //cas mois
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='4'
  then
   begin
    if yF=yD
     then diff:=mF-mD+1
     else diff:=(12-mD)+mF;
    if datedeb<10
     then diff:=12;

    if OkNew
     then
      begin
       val:=(1/diff)*100;
       ecart:=0;
       for i:=1 to diff do
        begin
         valChp:=VALEUR(Format('%10.4f',[val+ecart]));
         setField('QBO_BPM'+IntToStr(i),valChp);
         ecart:=ecart+(val-valChp);
        end;
       for i:=diff+1 to 52 do
        setField('QBO_BPM'+IntToStr(i),0);
      end;

    i:=1;
    while (i<=diff) do
     begin
      QT.addValeur([DonneMailleSuivantDelai('4',datedeb,i-1)], [getfield('qbo_bpm'+IntToStr(i))]);
      somme:=somme+getfield('qbo_bpm'+IntToStr(i));
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');
   end;

 //cas mois 4-4-5
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='5'
  then
   begin
    PerD:=FindPeriodeP(datedeb);
    PerF:=FindPeriodeP(datefin);
    diff:=PerF.periode-PerD.periode+1;

    if datedeb<10
     then diff:=12;

    if OkNew
     then
      begin
       val:=(1/diff)*100;
       ecart:=0;
       for i:=1 to diff do
        begin
         valChp:=VALEUR(Format('%10.4f',[val+ecart]));
         setField('QBO_BPM'+IntToStr(i),valChp);
         ecart:=ecart+(val-valChp);
        end;
       for i:=diff+1 to 52 do
        setField('QBO_BPM'+IntToStr(i),0);
      end;

    i:=1;
    while (i<=diff) do
     begin
      QT.addValeur([DonneMailleSuivantDelai('6',datedeb,i-1)], [getfield('qbo_bpm'+IntToStr(i))]);
      somme:=somme+getfield('qbo_bpm'+IntToStr(i));
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');
   end;

 //cas trimestre
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='6'
  then
   begin
    DonneNumTrimestre(datedeb,numTD,anTD);
    DonneNumTrimestre(datefin,numTF,anTF);
    if anTF=anTD
     then diff:=numTF-numTD+1
     else diff:=(4-numTD)+numTF;

    if datedeb<10
     then diff:=4;

    if OkNew
     then
      begin
       val:=(1/diff)*100;
       ecart:=0;
       for i:=1 to diff do
        begin
         valChp:=VALEUR(Format('%10.4f',[val+ecart]));
         setField('QBO_BPM'+IntToStr(i),valChp);
         ecart:=ecart+(val-valChp);
        end;
       for i:=diff+1 to 52 do
        setField('QBO_BPM'+IntToStr(i),0);
      end;

    i:=1;
    while (i<=diff) do
     begin
      QT.addValeur([DonneMailleSuivantDelai('7',datedeb,i-1)], [getfield('qbo_bpm'+IntToStr(i))]);
      somme:=somme+getfield('qbo_bpm'+IntToStr(i));
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');
   end;

 //cas quadrimestre
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='7'
  then
   begin
    DonneNumQuadrimestre(datedeb,numTD,anTD);
    DonneNumQuadrimestre(datefin,numTF,anTF);
    if anTF=anTD
     then diff:=numTF-numTD+1
     else diff:=(3-numTD)+numTF;

    if datedeb<10
     then diff:=3;

    if OkNew
     then
      begin
       val:=(1/diff)*100;
       ecart:=0;
       for i:=1 to diff do
        begin
         valChp:=VALEUR(Format('%10.4f',[val+ecart]));
         setField('QBO_BPM'+IntToStr(i),valChp);
         ecart:=ecart+(val-valChp);
        end;
       for i:=diff+1 to 52 do
        setField('QBO_BPM'+IntToStr(i),0);
      end;

    i:=1;
    while (i<=diff) do
     begin
      QT.addValeur([DonneMailleSuivantDelai('8',datedeb,i-1)], [getfield('qbo_bpm'+IntToStr(i))]);
      somme:=somme+getfield('qbo_bpm'+IntToStr(i));
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');
   end;

 SetcontrolText('EDTTOTAL',Format('%10.4f',[somme]));
 SetcontrolText('EDTSOLDE',Format('%10.4f',[100-somme]));

end;

procedure TOM_QBPLOI.RemplitGridHisto;
var datedeb,datefin:TDateTime;
    aD,numsemD,aF,numsemF,diff:integer;
    yF,mF,dF,yD,mD,dD:word;
    i,j,II:integer;
    somme:double;
    perD,perF:TperiodeP;
    numTD,anTD,numTF,anTF:integer;
begin
 dateDeb:=StrToDateTime0(GetcontrolText('JOURD'));
 dateFin:=StrToDateTime0(GetcontrolText('JOURF'));
 if not BPOkOrli
  then
   begin
    dateDeb:=StrToDateTime0(GetcontrolText('EDTDATEDEBREF'));
    dateFin:=StrToDateTime0(GetcontrolText('EDTDATEFINREF'));
   end;
 numsemD:= numSemaine(datedeb,aD);
 DecodeDate(datedeb,yD,mD,dD);
 numsemF:= numSemaine(datefin,aF);
 DecodeDate(datefin,yF,mF,dF);

 if QT <> nil
  then QT.Free;
 QT :=TQTob.create(['LADATE'],['VAL1']);
 THGrid(getcontrol('GRIDLOI')).RowCount:=0;

 somme:=0;

 //cas semaine
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='2'
  then
   begin
    if aF=aD
     then diff:=numsemF-numsemD+1
     else diff:=(52-numsemD)+numsemF;
    if diff>=53
     then diff:=52;

    i:=1;
    while (i<=diff) do
     begin
      QT.addValeur([DonneMailleSuivantDelai('2',datedeb,i-1)], [getfield('qbo_bpm'+IntToStr(i))]);
      somme:=somme+getfield('qbo_bpm'+IntToStr(i));
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');
   end;

 //cas quinzaine
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='3'
  then
   begin
    if yF=yD
     then diff:=mF-mD+1
     else diff:=(12-mD)+mF;
    diff:=diff*2;
    i:=1;
    j:=1;
    while (i<=diff/2) do
     begin
      if dD<16
       then
        begin
         if BPOkOrli
          //-----------------> ORLI
          then II:=(2*i)-2
          //ORLI <-----------------
          else II:=(i-1);
         QT.addValeur([DonneMailleSuivantDelai('3',datedeb,II)], [getfield('qbo_bpm'+IntToStr(j))]);
         somme:=somme+getfield('qbo_bpm'+IntToStr(j));
         j:=j+1;
         if BPOkOrli
          //-----------------> ORLI
          then II:=(2*i)-1
          //ORLI <-----------------
          else II:=(i-1);
         QT.addValeur([DonneMailleSuivantDelai('5',datedeb,II)], [getfield('qbo_bpm'+IntToStr(j))]);
         somme:=somme+getfield('qbo_bpm'+IntToStr(j));
         j:=j+1;
        end
       else
        begin   
         if BPOkOrli
          //-----------------> ORLI
          then II:=(2*i)-2
          //ORLI <-----------------
          else II:=(i-1);
         QT.addValeur([DonneMailleSuivantDelai('5',datedeb,II)], [getfield('qbo_bpm'+IntToStr(j))]);
         somme:=somme+getfield('qbo_bpm'+IntToStr(j));
         j:=j+1;     
         if BPOkOrli
          //-----------------> ORLI
          then II:=(2*i)-1
          //ORLI <-----------------
          else II:=(i-1);
         QT.addValeur([DonneMailleSuivantDelai('3',datedeb,II)], [getfield('qbo_bpm'+IntToStr(j))]);
         somme:=somme+getfield('qbo_bpm'+IntToStr(j));
         j:=j+1;
        end;
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),false,false,'');
   end;

 //cas mois
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='4'
  then
   begin
    if yF=yD
     then diff:=mF-mD+1
     else diff:=(12-mD)+mF;

    i:=1;
    while (i<=diff) do
     begin
      QT.addValeur([DonneMailleSuivantDelai('4',datedeb,i-1)], [getfield('qbo_bpm'+IntToStr(i))]);
      somme:=somme+getfield('qbo_bpm'+IntToStr(i));
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');
   end;

 //cas mois 4-4-5
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='5'
  then
   begin
    PerD:=FindPeriodeP(datedeb);
    PerF:=FindPeriodeP(datefin);
    diff:=PerF.periode-PerD.periode+1;

    i:=1;
    while (i<=diff) do
     begin
      QT.addValeur([DonneMailleSuivantDelai('6',datedeb,i-1)], [getfield('qbo_bpm'+IntToStr(i))]);
      somme:=somme+getfield('qbo_bpm'+IntToStr(i));
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');
   end;

 //cas trimestre
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='6'
  then
   begin
    DonneNumTrimestre(datedeb,numTD,anTD);
    DonneNumTrimestre(datefin,numTF,anTF);
    if anTF=anTD
     then diff:=numTF-numTD+1
     else diff:=(4-numTD)+numTF;

    i:=1;
    while (i<=diff) do
     begin
      QT.addValeur([DonneMailleSuivantDelai('7',datedeb,i-1)], [getfield('qbo_bpm'+IntToStr(i))]);
      somme:=somme+getfield('qbo_bpm'+IntToStr(i));
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');
   end;

 //cas quadrimestre
 if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='7'
  then
   begin
    DonneNumQuadrimestre(datedeb,numTD,anTD);
    DonneNumQuadrimestre(datefin,numTF,anTF);
    if anTF=anTD
     then diff:=numTF-numTD+1
     else diff:=(3-numTD)+numTF;

    i:=1;
    while (i<=diff) do
     begin
      QT.addValeur([DonneMailleSuivantDelai('8',datedeb,i-1)], [getfield('qbo_bpm'+IntToStr(i))]);
      somme:=somme+getfield('qbo_bpm'+IntToStr(i));
      i:=i+1;
     end;
    QT.laTob.PutGridDetail(THGrid(getcontrol('GRIDLOI')),true,false,'');
   end;

 SetcontrolText('EDTTOTAL',Format('%10.4f',[somme]));
 SetcontrolText('EDTSOLDE',Format('%10.4f',[100-somme]));

end;

procedure TOM_QBPLOI.OnLoadRecord ;
begin
 Inherited ;
 if  THDBValComboBox(getcontrol('QBO_MAILLE')).Enabled = false then THDBValComboBox(getcontrol('QBO_MAILLE')).Value := '4';
 RemplitGrid;
end ;

procedure TOM_QBPLOI.OnChangeField ( F: TField ) ;
begin
 Inherited ;
end ;


procedure TOM_QBPLOI.OnArgument ( S: String ) ;
begin
  Inherited ;
  TToolbarButton97(GetControl('BValider')).OnClick := BValider_OnClick;
  THGrid(getcontrol('GRIDLOI')).ColFormats[1]:='#,####0.0000';
  //THGrid(getcontrol('GRIDLOI')).ColAligns[1]:=taRightJustify;
  THGrid(getcontrol('GRIDLOI')).ColTypes[1]:='F';
  OkNew:=false;
  TToolBarButton97(getcontrol('BTNHISTO')).OnClick := BtnHisto;    
  TToolBarButton97(getcontrol('BTNSOLDE')).OnClick := BtnSolde;
  THGrid(getcontrol('GRIDLOI')).OnCellExit:=CellExit;
  THGrid(getcontrol('GRIDLOI')).OnCellEnter:=CellEnter;
  TToolBarButton97(getcontrol('BTNGRAPH')).OnClick :=BtnGraph;
  THDBValComboBox(getcontrol('QBO_MAILLE')).OnChange:=ChangeMaille;

  if ContextBP=3 then
  begin
    THDBValComboBox(getcontrol('QBO_MAILLE')).Enabled := false;
    TToolBarButton97(getcontrol('BTNHISTO')).Enabled := false;
    TToolBarButton97(getcontrol('BTNHISTO')).Visible := false;
  end;

  //création de la tob
  QT :=TQTob.create(['LADATE'],['VAL1']);
end ;

{ EVI / Force le statut de la fiche pour passer dans le UpdateRecord }
procedure TOM_QBPLOI.BValider_OnClick(Sender: TObject);
var S, S2 : hString;
begin
  inherited;
  { EVI / Récupération du script }
  S := FuncSynMakeDate ( [GetControlText('JOURD'), ''],0 ) ;
  S2:= FuncSynMakeDate ( [GetControlText('JOURF'), ''],0 ) ;
  if S<>''  then SetField ('QBO_DATEDEBECLAT',StrToDateTime(S))
           else SetField ('QBO_DATEDEBECLAT','');
  if S2<>'' then SetField('QBO_DATEFINECLAT',StrToDateTime(S2))
           else SetField('QBO_DATEFINECLAT','');

  MAJTableGrid;
  if TFFiche( Ecran ).FTypeAction <> taCreat then TFFiche( Ecran ).QFiche.Edit;
  TFFiche(Ecran).Bouge(nbPost) ;
  Ecran.Close ;
end;

procedure TOM_QBPLOI.OnClose ;
begin
 Inherited ;
 //suppression de la tob
 if QT <> nil
  then QT.Free;
end ;

procedure TOM_QBPLOI.OnCancelRecord ;
begin
  Inherited ;
end ;

function TOM_QBPLOI.ChercheCodeRestrictionLoi: hString;
var CodeAxe1,CodeAxe2,CodeAxe3,CodeAxe4,CodeAxe5:hString;
    CodeAxe6,CodeAxe7,CodeAxe8,CodeAxe9,CodeAxe10:hString;
    ValAxe1,ValAxe2,ValAxe3,ValAxe4,ValAxe5:hString;      
    ValAxe6,ValAxe7,ValAxe8,ValAxe9,ValAxe10:hString;
    structure,codeSql,codeSession:hString;
    Q:TQuery;
    i:integer;
begin
 codeSql:='';
 codeJoinR:='';
 CodeAxe1:=THValComboBox(getcontrol('QBO_CODEAXER1')).Value;
 CodeAxe2:=THValComboBox(getcontrol('QBO_CODEAXER2')).Value;
 CodeAxe3:=THValComboBox(getcontrol('QBO_CODEAXER3')).Value;
 CodeAxe4:=THValComboBox(getcontrol('QBO_CODEAXER4')).Value;
 CodeAxe5:=THValComboBox(getcontrol('QBO_CODEAXER5')).Value;

 ValAxe1:=THDBEdit(getcontrol('QBO_VALAXER1')).Text;
 ValAxe2:=THDBEdit(getcontrol('QBO_VALAXER2')).Text;
 ValAxe3:=THDBEdit(getcontrol('QBO_VALAXER3')).Text;
 ValAxe4:=THDBEdit(getcontrol('QBO_VALAXER4')).Text;
 ValAxe5:=THDBEdit(getcontrol('QBO_VALAXER5')).Text;

 if ContextBP <> 1
  then
   begin
    CodeAxe6:=THValComboBox(getcontrol('QBO_CODEAXER6')).Value;
    CodeAxe7:=THValComboBox(getcontrol('QBO_CODEAXER7')).Value;
    CodeAxe8:=THValComboBox(getcontrol('QBO_CODEAXER8')).Value;
    CodeAxe9:=THValComboBox(getcontrol('QBO_CODEAXER9')).Value;
    CodeAxe10:=THValComboBox(getcontrol('QBO_CODEAXER10')).Value;

    ValAxe6:=THDBEdit(getcontrol('QBO_VALAXER6')).Text;
    ValAxe7:=THDBEdit(getcontrol('QBO_VALAXER7')).Text;
    ValAxe8:=THDBEdit(getcontrol('QBO_VALAXER8')).Text;
    ValAxe9:=THDBEdit(getcontrol('QBO_VALAXER9')).Text;
    ValAxe10:=THDBEdit(getcontrol('QBO_VALAXER10')).Text;
   end;

 if BPOkOrli
  then
   //-----------------> ORLI
   begin
    structure:=THDBEdit(getcontrol('QBO_CODESTRUCT')).Text;

    Q:=MOpenSql('SELECT * FROM QBPSTRUCTURE WHERE QBC_CODESTRUCT="'+structure+'"',
                'QBPLOI_TOM (ChercheCodeRestrictionLoi).',true);
    if not Q.eof
     then
      begin
       for i:=1 to 30 do
        begin
         if (Q.Fields[i].asString=CodeAxe1) and (CodeAxe1<>'')
          then codeSql:=CodeSql+' AND QBV_VALAXEP'+IntToStr(i)+'="'+ValAxe1+'" ';
         if (Q.Fields[i].asString=CodeAxe2) and (CodeAxe2<>'')
          then codeSql:=CodeSql+' AND QBV_VALAXEP'+IntToStr(i)+'="'+ValAxe2+'" ';
         if (Q.Fields[i].asString=CodeAxe3) and (CodeAxe3<>'')
          then codeSql:=CodeSql+' AND QBV_VALAXEP'+IntToStr(i)+'="'+ValAxe3+'" ';
         if (Q.Fields[i].asString=CodeAxe4) and (CodeAxe4<>'')
          then codeSql:=CodeSql+' AND QBV_VALAXEP'+IntToStr(i)+'="'+ValAxe4+'" ';
         if (Q.Fields[i].asString=CodeAxe5) and (CodeAxe5<>'')
          then codeSql:=CodeSql+' AND QBV_VALAXEP'+IntToStr(i)+'="'+ValAxe5+'" ';
         if (Q.Fields[i].asString=CodeAxe6) and (CodeAxe6<>'')
          then codeSql:=CodeSql+' AND QBV_VALAXEP'+IntToStr(i)+'="'+ValAxe6+'" ';
         if (Q.Fields[i].asString=CodeAxe7) and (CodeAxe7<>'')
          then codeSql:=CodeSql+' AND QBV_VALAXEP'+IntToStr(i)+'="'+ValAxe7+'" ';
         if (Q.Fields[i].asString=CodeAxe8) and (CodeAxe8<>'')
          then codeSql:=CodeSql+' AND QBV_VALAXEP'+IntToStr(i)+'="'+ValAxe8+'" ';
         if (Q.Fields[i].asString=CodeAxe9) and (CodeAxe9<>'')
          then codeSql:=CodeSql+' AND QBV_VALAXEP'+IntToStr(i)+'="'+ValAxe9+'" ';
         if (Q.Fields[i].asString=CodeAxe10) and (CodeAxe10<>'')
          then codeSql:=CodeSql+' AND QBV_VALAXEP'+IntToStr(i)+'="'+ValAxe10+'" ';
        end;
      end;
    ferme(Q);
   end
   //ORLI <-----------------
  else
   begin
    //cas pgi
    codeSession:=GetcontrolText('EDTSESSION');
    codeSql:=ChercheCodeRestrictionSession('','','',codeSession,codeJoinR);
    if CodeAxe1<>''
     then RechercheCodeRestrictionMultiValeurPGI('','Z11',CodeAxe1,ValAxe1,codeSql,codeJoinR);
    if CodeAxe2<>''
     then RechercheCodeRestrictionMultiValeurPGI('','Z12',CodeAxe2,ValAxe2,codeSql,codeJoinR);
    if CodeAxe3<>''
     then RechercheCodeRestrictionMultiValeurPGI('','Z13',CodeAxe3,ValAxe3,codeSql,codeJoinR);
    if CodeAxe4<>''
     then RechercheCodeRestrictionMultiValeurPGI('','Z14',CodeAxe4,ValAxe4,codeSql,codeJoinR);
    if CodeAxe5<>''
     then RechercheCodeRestrictionMultiValeurPGI('','Z15',CodeAxe5,ValAxe5,codeSql,codeJoinR);
    if ContextBP in [0,2,3]
     then
      begin
       if CodeAxe6<>''
        then RechercheCodeRestrictionMultiValeurPGI('','Z16',CodeAxe6,ValAxe6,codeSql,codeJoinR);
       if CodeAxe7<>''
        then RechercheCodeRestrictionMultiValeurPGI('','Z17',CodeAxe7,ValAxe7,codeSql,codeJoinR);
       if CodeAxe8<>''
        then RechercheCodeRestrictionMultiValeurPGI('','Z18',CodeAxe8,ValAxe8,codeSql,codeJoinR);
       if CodeAxe9<>''
        then RechercheCodeRestrictionMultiValeurPGI('','Z19',CodeAxe9,ValAxe9,codeSql,codeJoinR);
       if CodeAxe10<>''
        then RechercheCodeRestrictionMultiValeurPGI('','Z20',CodeAxe10,ValAxe10,codeSql,codeJoinR);
      end;
   end;
 result:=codeSql;
end;

procedure TOM_QBPLOI.BtnHisto(Sender: TObject);
var tabQteLoi:array [1..53] of double;
    i:integer;
    dateDebRef,dateFinRef:TDateTime;
    CodeRestriction:hString;
    somme:double;
begin
  if THDBValComboBox(getcontrol('QBO_MAILLE')).Value <> '' then
  begin
    if DS.State = DsBrowse then DS.Edit;
    for i:=1 to 53 do tabQteLoi[i]:=0;

    //récupère les dates
    dateDebRef:=StrToDateTime0(GetcontrolText('JOURD'));
    dateFinRef:=StrToDateTime0(GetcontrolText('JOURF'));
    if not BPOkOrli then
    begin
      dateDebRef:=StrToDateTime0(GetcontrolText('EDTDATEDEBREF'));
      dateFinRef:=StrToDateTime0(GetcontrolText('EDTDATEFINREF'));
    end;

    //cherche le code de restriction
    CodeRestriction:=ChercheCodeRestrictionLoi;

    //remplit le tableau des %
    InitialisationPrctDelaiParRapportHisto(THDBValComboBox(getcontrol('QBO_MAILLE')).Value,
                getfield('QBO_CODESESSION'),CodeRestriction,codeJoinR,dateDebRef,dateFinRef,tabQteLoi);

    //met à jour les champs avec le tableau des %
    somme:=0;
    for i:=1 to 52 do
    begin
      setField('QBO_BPM'+IntToStr(i),VALEUR(Format('%10.4f',[tabQteLoi[i+1]])));
      somme:=somme+VALEUR(Format('%10.4f',[tabQteLoi[i+1]]));
    end;

    if (100-somme)<>0 then setField('QBO_BPM1',VALEUR(Format('%10.4f',[tabQteLoi[2]+(100-somme)])));

    RemplitGridHisto;
  end;
end;



procedure TOM_QBPLOI.BtnGraph(Sender: TObject);
var i:integer;
    val,LaDate:hString;
begin
 QTL:=TQTob.create(['LADATE'],['VAL1']);
 for i := 0 to QT.laTob.Detail.Count-1 do
  begin
   QT.laTob.GetLigneGrid(THGrid(getcontrol('GRIDLOI')),i,'LADATE;VAL1');
   val:=QT.laTob.getValue('VAL1');
   LaDate:=QT.laTob.getValue('LADATE');
   QTL.addValeur([LaDate],[VALEUR(val)]);
  end;
 AglLanceFiche('Q', 'QUFGBPLOI', '', '', '');
 QTL.free;
end;

procedure TOM_QBPLOI.ChangeMaille(Sender: TObject);
begin   
 if OldMaille='' then OldMaille:=THDBValComboBox(getcontrol('QBO_MAILLE')).Value;
 if ((THDBValComboBox(getcontrol('QBO_MAILLE')).Value<>OldMaille) OR (THDBValComboBox(getcontrol('QBO_MAILLE')).Value=''))  then
 begin
   if THDBValComboBox(getcontrol('QBO_MAILLE')).Value='' then THDBValComboBox(getcontrol('QBO_MAILLE')).Value:='4';
   OkNew:=true;
   RemplitGrid;
 end;
 OldMaille:=THDBValComboBox(getcontrol('QBO_MAILLE')).Value;
end;

procedure TOM_QBPLOI.BtnSolde(Sender: TObject);
var val,somme:double;
    i:integer;
begin
 MAJTableGrid;
 QT.laTob.GetLigneGrid(THGrid(getcontrol('GRIDLOI')),LigneSolde,'LADATE;VAL1');
 val:=QT.laTob.getValue('VAL1');
 QT.laTob.SetDouble('VAL1',val+VALEUR(GetcontrolText('EDTSOLDE')));
 QT.laTob.PutLigneGrid(THGrid(getcontrol('GRIDLOI')),LigneSolde,false,false,'LADATE;VAL1');
 somme:=0;
 for i := 0 to QT.laTob.Detail.Count-1 do
  begin
   QT.laTob.GetLigneGrid(THGrid(getcontrol('GRIDLOI')),i,'LADATE;VAL1');
   val:=QT.laTob.getValue('VAL1');
   somme:=somme+val;
  end;
 
 SetcontrolText('EDTTOTAL',Format('%10.4f',[somme]));
 SetcontrolText('EDTSOLDE',Format('%10.4f',[100-somme]));
end;

procedure TOM_QBPLOI.CellEnter(Sender: TObject; var ACol, ARow: Integer;
  var Cancel: Boolean);
var G:THGrid ;
begin
 G:=THGrid(Sender) ;
 LigneSolde:=G.row;
 ColSolde:=G.col;
 ValCell:=VALEUR(G.cells[ColSolde,LigneSolde]);
 DateCell:=G.cells[0,LigneSolde];
end;

Initialization
  registerclasses ( [ TOM_QBPLOI ] ) ;
end.
