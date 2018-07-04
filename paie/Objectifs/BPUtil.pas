unit BPUtil;

interface

Uses HEnt1,Classes;

function BPCalculQteListProrata(ErrInit: double;ListQte: array of double; Rapp: double;var  ListeQteRes: array of double): double;
function ContextBP:integer;
function NbrValAff(session:string):integer;
function TableSalariesOK(session:string) : boolean;
procedure LibValAff(session:string; var TabLibValAff:array of hString);
procedure ReqValAffTree(session,axes:string; NbAxes:integer; var NbrValue:integer;
                    var ChpArbre,ChpArbreDetail,TableH,WhereH :string);
procedure ReqValAffCube(session,BaseS,axes:string; NbAxes:integer; var NbrValue:integer;
                    var TableR,WhereR,ChpReal :string);
procedure DonnePrefixe(NomSession : String; var PrfxH,PrfxR:string);
procedure MajCubeLibelleNewValues(Session:string;AxeSal:integer);
procedure VerifCoherenceSessions(const ListeBases,SessionName:string;var lstMess,lstMessState:TStringList);

implementation

Uses SysUtils,StrUtils,HCtrls,paramsoc,UtilPGI,HMsgBox,
    {$IFDEF PAIEGRH}EntPaie,{$ENDIF PAIEGRH}
    {$IFDEF EAGLCLIENT} UtileAGL, uTob,
    {$ELSE} {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    {$ENDIF}
     Uutil;

function BPCalculQteListProrata(ErrInit: double;ListQte: array of double; Rapp: double;var  ListeQteRes: array of double): double;
var i, Qte2: integer;
  Qte,Err1: double;
begin
  Err1 := ErrInit;
  for i := 0 to 19 do
  begin
    Qte := ListQte[i];  // double
    Qte2 := round(1.0* Qte * Rapp + Err1 );  // integer
    Qte2 := plusGrand(0, Qte2);
    Err1 := Err1 + (1.0* Qte * Rapp - Qte2);
    ListeQteRes[i]:= Qte2;
  end;
  Result:= Err1;
end;

function ContextBP:integer;
begin
  if ctxMode in V_PGI.PGIContexte then Result := 1
  else if ctxCompta in V_PGI.PGIContexte then Result := 2
  else if ctxPaie in V_PGI.PGIContexte then Result := 3
  else Result := 0; // Gescom
end;

function NbrValAff(session:string):integer ;
var Q:TQuery;
begin
 Result := 0;
 Q:=MOpenSql('SELECT QBS_NBVALAFF FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+session+'"',
             'BPUtil(NbValAff).',true);
 if not Q.Eof then Result := Q.fields[0].AsInteger;
 Ferme(Q);
end;

procedure DonnePrefixe(NomSession : String; var PrfxH,PrfxR:string);
var Q: TQuery;
begin
  Q:= OpenSQL('SELECT QBS_VALAFFH1,QBS_VALAFFR1 FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+NomSession+'"',True);
  if not Q.Eof then
  begin
    PrfxH := Q.fields[0].AsString ;
    PrfxR := Q.fields[1].AsString ;
  end;
  Ferme(Q)
end;

function TableSalariesOK(session:string) : boolean;
var Q:TQuery;
i:integer;
begin
  result := false;
  Q:=MOpenSql('SELECT QBS_NBVALAFF, QBS_VALAFFH1, QBS_VALAFFH2, QBS_VALAFFH3, QBS_VALAFFH4, QBS_VALAFFH5,'+
              ' QBS_VALAFFH6, QBS_VALAFFH7 FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+session+'"',
              'BPUtil(TableSalariesOK).',true);
  if not Q.Eof then
  begin
    for i:=1 to Q.Fields[0].AsInteger do
    begin
      if Q.Fields[i].AsString = 'PSA' then
      begin
        Result := true;
        break;
      end;
    end;
  end;
  Ferme(Q);
end;

procedure LibValAff(session:string; var TabLibValAff:array of hstring);
var Q:TQuery;
i:integer;
begin
  Q:=MOpenSql('SELECT QBS_NBVALAFF, QBS_VALAFFLIB1, QBS_VALAFFLIB2, QBS_VALAFFLIB3, QBS_VALAFFLIB4, QBS_VALAFFLIB5,'+
              ' QBS_VALAFFLIB6, QBS_VALAFFLIB7 FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+session+'"',
              'BPUtil(LibValAff).',true);
  if not Q.Eof then
  begin
    for i:=0 to 7 do TabLibValAff[i] := Q.Fields[i].AsString;
  end;
  Ferme(Q);
end;

function RecupListeChamps(ReqSql:string):string;
var Q1 : TQuery;
begin
  Result:='';
  Q1 := OpenSQL(ReqSql,True);
  while not Q1.eof do
  begin
    if Result = '' then Result := Q1.fields[0].AsString
    else Result := Result + ';' + Q1.fields[0].AsString;
    Q1.next;
  end;
  Result := Result + ';' ;
  ferme(Q1);
end;

function VerifTous(Champs : string;Table : integer) : string;
var i:integer;
begin
  if Uppercase(Champs) = Uppercase(TraduireMemoire('<<Tous>>')) then
  begin
    case Table of
     1 : Result := RecupListeChamps('SELECT CO_CODE FROM COMMUN WHERE CO_TYPE="BPP"'); //QUTBPPGPPU
     2 : for i := 1 to VH_Paie.PgNbSalLib do result:=result+'SM'+IntToSTr(i)+';'
    end;
  end else Result := Champs;
end;

procedure ReqValAffTree(session,axes:string; NbAxes:integer; var NbrValue:integer;
                    var ChpArbre,ChpArbreDetail,TableH,WhereH :string);
var i:integer;
    ChpsSum1,ChpsSum2,PrefixH:string;
    Q:TQuery;
    TabChpsSumH,TabChpsSumHDet : array [0..6] of hString;
begin
  axes := AnsiReplaceText(axes,'TRIM','');
  axes := AnsiReplaceText(axes,')','');
  axes := AnsiReplaceText(axes,'(','');

  Q:=MOpenSql('SELECT QBS_NBVALAFF, QBS_VALAFFH1,QBS_VALAFFH2,QBS_VALAFFH3,QBS_VALAFFH4,QBS_VALAFFH5,QBS_VALAFFH6,QBS_VALAFFH7,'+
              'QBS_VALAFFH1A,QBS_VALAFFH2A,QBS_VALAFFH3A,QBS_VALAFFH4A,QBS_VALAFFH5A,QBS_VALAFFH6A,QBS_VALAFFH7A,'+
              'QBS_VALAFFH1B,QBS_VALAFFH2B,QBS_VALAFFH3B,QBS_VALAFFH4B,QBS_VALAFFH5B,QBS_VALAFFH6B,QBS_VALAFFH7B,'+
              'QBS_VALAFFR1,QBS_VALAFFR2,QBS_VALAFFR3,QBS_VALAFFR4,QBS_VALAFFR5,QBS_VALAFFR6,QBS_VALAFFR7,'+
              'QBS_VALAFFR1A,QBS_VALAFFR2A,QBS_VALAFFR3A,QBS_VALAFFR4A,QBS_VALAFFR5A,QBS_VALAFFR6A,QBS_VALAFFR7A,'+
              'QBS_VALAFFR1B,QBS_VALAFFR2B,QBS_VALAFFR3B,QBS_VALAFFR4B,QBS_VALAFFR5B,QBS_VALAFFR6B,QBS_VALAFFR7B'+
              ' FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+session+'"',
              'BPUtil(ReqValAff).',true);

  if not Q.Eof then
  begin
    NbrValue := Q.fields[0].AsInteger;

    For i:=1 to NbrValue do
    begin

      //Valeur Historique Rubriques de Salare
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i].AsString = 'PSA' then
      begin
        ChpsSum1:=VerifTous(Q.fields[i+7].AsString,2);
        ChpsSum1:=AnsiReplaceText(ChpsSum1,';','+')  ;

        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SM1','PSA_SALAIREMOIS1');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SM2','PSA_SALAIREMOIS2');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SM3','PSA_SALAIREMOIS3');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SM4','PSA_SALAIREMOIS4');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SM5','PSA_SALAIREMOIS5');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SA1','PSA_SALAIRANN1');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SA2','PSA_SALAIRANN2');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SA3','PSA_SALAIRANN3');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SA4','PSA_SALAIRANN4');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SA5','PSA_SALAIRANN5');

        Delete(ChpsSum1,LastDelimiter('+',ChpsSum1),1)  ;

        if i=1 then
        begin
          TableH := 'SALARIES AS T1';
          PrefixH := 'PSA';
          TabChpsSumH[i-1]:='SUM('+ChpsSum1+')';
          TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
          WhereH:=' PSA_DATEENTREE<="[DATEDEBUT]" AND (PSA_DATESORTIE>="[DATEFIN]" OR PSA_DATESORTIE="'+USDateTime(iDate1900)+'")'
        end
        else
        begin
          TabChpsSumH[i-1]:='SUM('+ChpsSum1+')';
          TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
        end
      end;

      //Valeur Historique Effectifs Temps Plein
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i].AsString = 'ETP' then
      begin
        if i=1 then
        begin
          TableH := 'SALARIES AS T1';
          PrefixH := 'PSA';
          TabChpsSumH[i-1]:='SUM(PSA_HORAIREMOIS)/ETB_HORAIREETABL';
          TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
          WhereH:=' PSA_DATEENTREE<="[DATEDEBUT]" AND (PSA_DATESORTIE>="[DATEFIN]" OR PSA_DATESORTIE="'+USDateTime(iDate1900)+'")'
        end
        else
        begin
          TabChpsSumH[i-1]:='SUM(PSA_HORAIREMOIS)/ETB_HORAIREETABL';
          TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
        end
      end;



    end;
  end;
  Ferme(Q);

  ChpArbre:=TabChpsSumH[0]+',0';
  For i:=2 to 7 do
  begin
    if TabChpsSumH[i-1]='' then ChpArbre:=ChpArbre+',0' else ChpArbre:=ChpArbre+','+TabChpsSumH[i-1] ;
  end;
  ChpArbre := ChpArbre + ',';

  ChpArbreDetail:=TabChpsSumHDet[0]+',0';
  For i:=2 to 7 do
  begin
    if TabChpsSumHDet[i-1]='' then ChpArbreDetail:=ChpArbreDetail+',0' else ChpArbreDetail:=ChpArbreDetail+','+TabChpsSumHDet[i-1] ;
  end;
  ChpArbreDetail := ChpArbreDetail + ',';

end;


procedure ReqValAffCube(session,BaseS,axes:string; NbAxes:integer; var NbrValue:integer;
                    var TableR,WhereR,ChpReal :string);
var i,j,k:integer;
    ChpsSum1,ChpsSum2,ReqJoin,PrefixR,PPUCode,PPUCodeTmp,PPUAbr,PPUAbrTmp,HoraireETAB:string;
    Q,Q1:TQuery;
    TabChpsSumR : array [0..6] of hString;
    TabAxes : array [0..9] of string;
begin
  axes := AnsiReplaceText(axes,'TRIM','');
  axes := AnsiReplaceText(axes,')','');
  axes := AnsiReplaceText(axes,'(','');

  { EV5 / Gestion des booléens libres PAIE }
  j:=0;
  for i:= 1 to NbAxes do
  begin
    if copy(copy(axes,0,PosEx(',',axes)-1),0,13)<>'PSA_BOOLLIBRE' then
    begin
      TabAxes[j]:=copy(axes,0,PosEx(',',axes)-1);
      TabAxes[j] := AnsiReplaceText(TabAxes[j],'[PREFIXE]','');
      j:=j+1;
    end;
    Delete(axes,1,PosEx(',',axes));
  end;



  Q:=MOpenSql('SELECT QBS_NBVALAFF, QBS_VALAFFH1,QBS_VALAFFH2,QBS_VALAFFH3,QBS_VALAFFH4,QBS_VALAFFH5,QBS_VALAFFH6,QBS_VALAFFH7,'+
              'QBS_VALAFFH1A,QBS_VALAFFH2A,QBS_VALAFFH3A,QBS_VALAFFH4A,QBS_VALAFFH5A,QBS_VALAFFH6A,QBS_VALAFFH7A,'+
              'QBS_VALAFFH1B,QBS_VALAFFH2B,QBS_VALAFFH3B,QBS_VALAFFH4B,QBS_VALAFFH5B,QBS_VALAFFH6B,QBS_VALAFFH7B,'+
              'QBS_VALAFFR1,QBS_VALAFFR2,QBS_VALAFFR3,QBS_VALAFFR4,QBS_VALAFFR5,QBS_VALAFFR6,QBS_VALAFFR7,'+
              'QBS_VALAFFR1A,QBS_VALAFFR2A,QBS_VALAFFR3A,QBS_VALAFFR4A,QBS_VALAFFR5A,QBS_VALAFFR6A,QBS_VALAFFR7A,'+
              'QBS_VALAFFR1B,QBS_VALAFFR2B,QBS_VALAFFR3B,QBS_VALAFFR4B,QBS_VALAFFR5B,QBS_VALAFFR6B,QBS_VALAFFR7B'+
              ' FROM ' + GetBase (BaseS, 'QBPSESSIONBP') + ' WHERE QBS_CODESESSION="'+session+'"',
              'BPUtil(ReqValAff).',true);

  if not Q.Eof then
  begin
    NbrValue := Q.fields[0].AsInteger;

    For i:=1 to NbrValue do
    begin

      //Valeur Effectifs Temps Plein (Table PAIENCOURS)
{      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i+21].AsString = 'ETP' then
      begin
        ChpsSum1:=VerifTous(Q.fields[i+21+7].AsString,1);
        ChpsSum1:=AnsiReplaceText(ChpsSum1,';','+')  ;
        Delete(ChpsSum1,LastDelimiter('+',ChpsSum1),1)  ;
        if i=1 then
        begin
          TableR := 'PAIEENCOURS AS T1';
          PrefixR := 'PPU';
          TabChpsSumR[i-1]:='SUM(PPU_CHEURESTRAV)/ETB_HORAIREETABL';
          WhereR:=' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'
        end
        else
        begin
          if TableR = 'PAIEENCOURS AS T1' then TabChpsSumR[i-1]:='SUM(PPU_CHEURESTRAV)/ETB_HORAIREETABL'
          else
          begin
            ReqJoin := '';
            for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PPU'+TabAxes[k-1]+'=T1.'+PrefixR+TabAxes[k-1];
            TabChpsSumR[i-1]:='(SELECT SUM(PPU_CHEURESTRAV)/ETB_HORAIREETABL FROM PAIEENCOURS WHERE PPU_DATEDEBUT=T1.'+PrefixR+'_DATEDEBUT '+
                              'AND PPU_DATEFIN=T1.'+PrefixR+'_DATEFIN '+      ReqJoin + ')'
          end;
        end
      end;    }

      //Valeur Rubrique du salaire
{      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i+21].AsString = 'PSA' then
      begin
        ChpsSum1:=VerifTous(Q.fields[i+21+7].AsString,1);
        ChpsSum1:=AnsiReplaceText(ChpsSum1,';','+')  ;

        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SM1','PPU_CBRUT');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SM2','PPU_CBRUTFISCAL');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SM3','PPU_CNETIMPOSAB');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SM4','PPU_CNETAPAYER');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SM5','PPU_CCOUTSALARIE');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SA1','PPU_CCOUTPATRON');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SA2','PPU_CPLAFONDSS');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SA3','PPU_CBASESS');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SA4','PPU_CBASESS');
        ChpsSum1:=AnsiReplaceText(ChpsSum1,'SA5','PPU_CBASESS');

        Delete(ChpsSum1,LastDelimiter('+',ChpsSum1),1)  ;

        if i=1 then
        begin
          TableR := 'PAIEENCOURS AS T1';
          PrefixR := 'PPU';
          TabChpsSumR[i-1]:='SUM('+ChpsSum1+')';
          WhereR:=' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'
        end
        else
        begin
          if TableR = 'PAIEENCOURS AS T1' then TabChpsSumR[i-1]:='SUM('+ChpsSum1+')'
          else
          begin
            ReqJoin := '';
            for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PPU'+TabAxes[k-1]+'=T1.'+PrefixR+TabAxes[k-1];
            TabChpsSumR[i-1]:='(SELECT SUM('+ChpsSum1+') FROM PAIEENCOURS WHERE PPU_DATEDEBUT=T1.'+PrefixR+'_DATEDEBUT '+
                              'AND PPU_DATEFIN=T1.'+PrefixR+'_DATEFIN '+      ReqJoin + ')'
          end;
        end
      end; }

      //Valeur PAIENCOURS
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i+21].AsString = 'PPU' then
      begin
        ChpsSum1:=VerifTous(Q.fields[i+21+7].AsString,1);

        PPUCode:='';
        PPUAbr:='';
        Q1 := OpenSQL('SELECT CO_CODE,CO_ABREGE FROM ' + GetBase (BaseS, 'COMMUN') + ' WHERE CO_TYPE="BPP"',True);
        while not Q1.eof do
        begin
          if PPUCode = '' then PPUCode := Q1.fields[0].AsString
          else PPUCode := PPUCode + ';' + Q1.fields[0].AsString;
          if PPUAbr = '' then PPUAbr := Q1.fields[1].AsString
          else PPUAbr := PPUAbr + ';' + Q1.fields[1].AsString;
          Q1.next;
        end;
        PPUCode := PPUCode + ';' ;
        PPUAbr := PPUAbr + ';' ;
        ferme(Q1);

        While PPUCode <> '' do
        begin
          PPUCodeTmp:=ReadTokenSt(PPUCode);
          PPUAbrTmp:=ReadTokenSt(PPUAbr);
          ChpsSum1:=AnsiReplaceText(ChpsSum1,PPUCodeTmp,PPUAbrTmp);
        end;

        ChpsSum1:=AnsiReplaceText(ChpsSum1,';','+')  ;
        Delete(ChpsSum1,LastDelimiter('+',ChpsSum1),1)  ;
        if i=1 then
        begin
          TableR := GetBase (BaseS, 'PAIEENCOURS AS T1');
          PrefixR := 'PPU';
          TabChpsSumR[i-1]:='SUM('+ChpsSum1+')';
          WhereR:=' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'
        end
        else
        begin
          ReqJoin := '';
          for k:= 0 to j-1 do ReqJoin := ReqJoin + ' AND PPU'+TabAxes[k]+'=T1.'+PrefixR+TabAxes[k];
          TabChpsSumR[i-1]:='(SELECT SUM('+ChpsSum1+') FROM '+ GetBase (BaseS, 'PAIEENCOURS') + ' WHERE PPU_DATEDEBUT=T1.'+PrefixR+'_DATEDEBUT '+
                            'AND PPU_DATEFIN=T1.'+PrefixR+'_DATEFIN '+      ReqJoin + ')'
        end
      end;

      //Valeur Nature Rubrique (table HISTOBULLETIN)
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i+21].AsString = 'PHB' then
      begin
        ChpsSum1:=Q.fields[i+21+7].AsString;
        if Uppercase(ChpsSum1)=Uppercase(TraduireMemoire('<<Tous>>')) then ChpsSum1:='' else ChpsSum1:='AND PHB_NATURERUB="'+ChpsSum1+'"';
        ChpsSum2:=Q.fields[i+21+14].AsString;
        if Uppercase(ChpsSum2)=Uppercase(TraduireMemoire('<<Tous>>')) then ChpsSum2:='' else ChpsSum2:='AND PHB_RUBRIQUE IN ("'+ChpsSum2+'")';
        Delete(ChpsSum2,LastDelimiter(';',ChpsSum2),1);
        ChpsSum2:=AnsiReplaceText(ChpsSum2,';','","')  ;
        if i=1 then
        begin
          TableR := GetBase (BaseS, 'HISTOBULLETIN AS T1');
          PrefixR := 'PHB';
          TabChpsSumR[i-1]:='SUM(PHB_MTREM)';
          WhereR:=' PHB_DATEDEBUT>="[DATEDEBUT]" AND PHB_DATEFIN<="[DATEFIN]" '+ChpsSum1+' '+ChpsSum2
        end
        else
        begin
          ReqJoin := '';
          for k:= 0 to j-1 do ReqJoin := ReqJoin + ' AND PHB'+TabAxes[k]+'=T1.'+PrefixR+TabAxes[k];
          TabChpsSumR[i-1]:='(SELECT SUM(PHB_MTREM) FROM '+GetBase (BaseS, 'HISTOBULLETIN')+' WHERE PHB_DATEDEBUT=T1.'+PrefixR+'_DATEDEBUT '+
                            'AND PHB_DATEFIN=T1.'+PrefixR+'_DATEFIN ' + ChpsSum1+' '+ChpsSum2+ReqJoin + ')'
        end;
      end;

      //Valeur Cumuls (Table HISTOCUMSAL)
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i+21].AsString = 'PHC' then
      begin
        ChpsSum1:=Q.fields[i+21+7].AsString;
        if ChpsSum1=GetParamsocSecur('SO_PGLIENETP','Nada') then HoraireETAB:='/ETB_HORAIREETABL'
        else HoraireETAB:='';
        Delete(ChpsSum1,LastDelimiter(';',ChpsSum1),1);
        ChpsSum1:=AnsiReplaceText(ChpsSum1,';','","')  ;
        if Uppercase(ChpsSum1)=Uppercase(TraduireMemoire('<<Tous>>')) then ChpsSum1:='' else ChpsSum1:='AND PHC_CUMULPAIE IN ("'+ChpsSum1+'")';
        if i=1 then
        begin
          TableR := GetBase (BaseS, 'HISTOCUMSAL AS T1');
          PrefixR := 'PHC';
          TabChpsSumR[i-1]:='SUM(PHC_MONTANT)'+HoraireETAB;
          WhereR:=' PHC_DATEDEBUT>="[DATEDEBUT]" AND PHC_DATEFIN<="[DATEFIN]" '+ChpsSum1
        end
        else
        begin
          ReqJoin := '';
          for k:= 0 to j-1 do ReqJoin := ReqJoin + ' AND PHC'+TabAxes[k]+'=T1.'+PrefixR+TabAxes[k];
          TabChpsSumR[i-1]:='(SELECT SUM(PHC_MONTANT)'+HoraireETAB+' FROM '+GetBase (BaseS, 'HISTOCUMSAL')+' WHERE PHC_DATEDEBUT=T1.'+PrefixR+'_DATEDEBUT '+
                            ChpsSum1 + ' '+ ReqJoin + ')'
        end
      end;

      //Valeur Bulletins de salaire
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i+21].AsString = 'NBS' then
      begin
        if i=1 then
        begin
          TableR := GetBase (BaseS, 'PAIEENCOURS AS T1');
          PrefixR := 'PPU';
          TabChpsSumR[i-1]:='SUM(1)';
          WhereR:=' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'
        end
        else
        begin
          ReqJoin := '';
          for k:= 0 to j-1 do ReqJoin := ReqJoin + ' AND PPU'+TabAxes[k]+'=T1.'+PrefixR+TabAxes[k];
          TabChpsSumR[i-1]:='(SELECT SUM(1) FROM '+GetBase (BaseS, 'PAIEENCOURS')+' WHERE' +
                            ' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"' + ReqJoin + ')'
        end
      end;

      //Valeur Bulletins de salaire Fin de Mois
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i+21].AsString = 'EFM' then
      begin
        if i=1 then
        begin
          TableR := GetBase (BaseS, 'PAIEENCOURS AS T1');
          PrefixR := 'PPU';
          TabChpsSumR[i-1]:='SUM(IIF((PPU_DATEFIN="[DATEFIN]"),1,0))';
          WhereR:=' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'
        end
        else
        begin
            ReqJoin := '';
            for k:= 0 to j-1 do ReqJoin := ReqJoin + ' AND PPU'+TabAxes[k]+'=T1.'+PrefixR+TabAxes[k];
            TabChpsSumR[i-1]:='(SELECT SUM(IIF((PPU_DATEFIN="[DATEFIN]"),1,0)) FROM '+GetBase (BaseS, 'PAIEENCOURS')+' WHERE' +
                              ' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"' + ReqJoin + ')'
        end
      end;



    end;
  end;
  Ferme(Q);

  For i:=1 to 6 do
  begin
    if i=1 then
    begin
      if TabChpsSumR[i-1]='' then ChpReal:=ChpReal+'0,' else ChpReal:=ChpReal+TabChpsSumR[i-1]+',';
    end
    else if i = 2 then
    begin
      if TabChpsSumR[i-1]='' then ChpReal:='0,' + ChpReal else ChpReal:=TabChpsSumR[i-1]+','+ ChpReal;
    end
    else if i > 2 then
    begin
      if TabChpsSumR[i-1]='' then ChpReal:=ChpReal+'0,' else ChpReal:=ChpReal+TabChpsSumR[i-1]+',';
    end;
  end;
  if TabChpsSumR[6]='' then ChpReal:=','+ChpReal+'0' else ChpReal:=','+ChpReal+TabChpsSumR[6]+',';

end;

procedure MajCubeLibelleNewValues(Session:string;AxeSal:integer);
var i,j,k:integer;
  Q:TQuery;
  TabCode,TabLibelle:array of string;
  TabIIF:array[0..9] of string;
  ReqIIF,ReqTree,Libelle,ValAxe:string;
begin
   i:=0;
   Q:=OpenSql('SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE="PTK"',true);
   SetLength(TabCode,Q.RecordCount);
   SetLength(TabLibelle,Q.RecordCount);
   while not Q.eof do
   begin
     if ReqTree ='' then ReqTree :=  'QBR_VALEURAXE LIKE "'+Q.fields[0].asString+'%"'
     else ReqTree :=  ReqTree + ' OR QBR_VALEURAXE LIKE "'+Q.fields[0].asString+'%"';
     TabCode[i]:= Q.fields[0].asString;
     TabLibelle[i]:=Q.fields[1].asString;
     i:=i+1;
     Q.next;
   end;
   ferme(Q);

   i:=0;
   Q:=OpenSql('SELECT QBR_VALEURAXE FROM QBPARBRE WHERE QBR_CODESESSION="'+Session+'"'+
              ' AND ('+ReqTree+')',true);
   while not Q.eof do
   begin
     for j:=0 to Length(TabCode) do
     begin
       if Copy(Q.fields[0].asString,0,2)=TabCode[j] then
       begin
         Libelle := TabLibelle[j];
         break;
       end;
     end;
     TabIIF[i]:='IIF(QBQ_VALAXECT'+IntToStr(AxeSal-1)+'="'+Q.fields[0].asString+'","'+Libelle+'",[KO])';
     if ValAxe='' then ValAxe:='"'+Q.fields[0].asString+'"' else ValAxe:=ValAxe + ',"' +Q.fields[0].asString+'"';
     i:=i+1;

     if i=10 then
     begin
       ReqIIF:=TabIIF[0];
       for k:=1 to (i-1) do ReqIIF:=StringReplace(ReqIIF,'[KO]',TabIIF[k],[rfReplaceAll, rfIgnoreCase]);
       ReqIIF:=StringReplace(ReqIIF,'[KO]','""',[rfReplaceAll, rfIgnoreCase]);
       if ReqIIF<>'' then
       ExecuteSQL('UPDATE QBPCUBETMP SET QBQ_LIBVALAXECT'+IntToStr(AxeSal-1)+'='+ReqIIF+
                  ' WHERE QBQ_VALAXECT'+IntToStr(AxeSal-1)+' IN ('+ValAxe+')');
       ReqIIF:='';
       ValAxe:='';
       i:=0;

       for k:=0 to 9 do TabIIF[k]:='';
     end;

     Q.next;

   end;
   ferme(Q);

   ReqIIF:=TabIIF[0];
   for k:=1 to (i-1) do ReqIIF:=StringReplace(ReqIIF,'[KO]',TabIIF[k],[rfReplaceAll, rfIgnoreCase]);
   ReqIIF:=StringReplace(ReqIIF,'[KO]','""',[rfReplaceAll, rfIgnoreCase]);
   if ReqIIF<>'' then
   ExecuteSQL('UPDATE QBPCUBETMP SET QBQ_LIBVALAXECT'+IntToStr(AxeSal-1)+'='+ReqIIF+
             ' WHERE QBQ_VALAXECT'+IntToStr(AxeSal-1)+' IN ('+ValAxe+')');

end;

procedure VerifCoherenceSessions(const ListeBases,SessionName:string;var lstMess,lstMessState:TStringList);
var Q,QSoc:Tquery;
SocError:boolean;
i:integer;
sSocGroupe,sBase:string;
begin
  Q:= openSql('SELECT QBS_DATEDEBC,QBS_DATEFINC,QBS_DATEDEBREF,QBS_DATEFINREF,'+
              'QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4,QBS_CODEAXES5,'+
              'QBS_CODEAXES6,QBS_CODEAXES7,QBS_CODEAXES8,QBS_CODEAXES9,QBS_CODEAXES10,'+
              'QBS_CODEAXE1,QBS_VALEURAXE1,QBS_CODEAXE2,QBS_VALEURAXE2,QBS_CODEAXE3,QBS_VALEURAXE3,'+
              'QBS_CODEAXE4,QBS_VALEURAXE4,'+
              'QBS_NBVALAFF,QBS_VALAFFH1,QBS_VALAFFH1A,QBS_VALAFFH1B,'+
              'QBS_VALAFFR1,QBS_VALAFFR1A,QBS_VALAFFR1B,QBS_VALAFFH2,QBS_VALAFFH2A,QBS_VALAFFH2B,'+
              'QBS_VALAFFR2,QBS_VALAFFR2A,QBS_VALAFFR2B,QBS_VALAFFH3,QBS_VALAFFH3A,QBS_VALAFFH3B,'+
              'QBS_VALAFFR3,QBS_VALAFFR3A,QBS_VALAFFR3B,QBS_VALAFFH4,QBS_VALAFFH4A,QBS_VALAFFH4B,'+
              'QBS_VALAFFR4,QBS_VALAFFR4A,QBS_VALAFFR4B,QBS_VALAFFH5,QBS_VALAFFH5A,QBS_VALAFFH5B,'+
              'QBS_VALAFFR5,QBS_VALAFFR5A,QBS_VALAFFR5B,QBS_VALAFFH6,QBS_VALAFFH6A,QBS_VALAFFH6B,'+
              'QBS_VALAFFR6,QBS_VALAFFR6A,QBS_VALAFFR6B,QBS_VALAFFH7,QBS_VALAFFH7A,QBS_VALAFFH7B,'+
              'QBS_VALAFFR7,QBS_VALAFFR7A,QBS_VALAFFR7B,'+
              'QBS_VALAFFLIB1,QBS_VALAFFLIB2,QBS_VALAFFLIB3,QBS_VALAFFLIB4,QBS_VALAFFLIB5,'+
              'QBS_VALAFFLIB6,QBS_VALAFFLIB7,'+
              'QBS_SESSIONINIT FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+SessionName+'"',true);
  if not Q.eof then
  begin
    sSocGroupe := ListeBases;
    sBase := ReadTokenSt(sSocGroupe) ;
    repeat
      if sBase <> V_PGI.SchemaName then
      begin
        SocError:=false;
        Qsoc:= openSql('SELECT QBS_DATEDEBC,QBS_DATEFINC,QBS_DATEDEBREF,QBS_DATEFINREF,'+
                       'QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4,QBS_CODEAXES5,'+
                       'QBS_CODEAXES6,QBS_CODEAXES7,QBS_CODEAXES8,QBS_CODEAXES9,QBS_CODEAXES10,'+
                       'QBS_CODEAXE1,QBS_VALEURAXE1,QBS_CODEAXE2,QBS_VALEURAXE2,QBS_CODEAXE3,QBS_VALEURAXE3,'+
                       'QBS_CODEAXE4,QBS_VALEURAXE4,'+
                       'QBS_NBVALAFF,QBS_VALAFFH1,QBS_VALAFFH1A,QBS_VALAFFH1B,'+
                       'QBS_VALAFFR1,QBS_VALAFFR1A,QBS_VALAFFR1B,QBS_VALAFFH2,QBS_VALAFFH2A,QBS_VALAFFH2B,'+
                       'QBS_VALAFFR2,QBS_VALAFFR2A,QBS_VALAFFR2B,QBS_VALAFFH3,QBS_VALAFFH3A,QBS_VALAFFH3B,'+
                       'QBS_VALAFFR3,QBS_VALAFFR3A,QBS_VALAFFR3B,QBS_VALAFFH4,QBS_VALAFFH4A,QBS_VALAFFH4B,'+
                       'QBS_VALAFFR4,QBS_VALAFFR4A,QBS_VALAFFR4B,QBS_VALAFFH5,QBS_VALAFFH5A,QBS_VALAFFH5B,'+
                       'QBS_VALAFFR5,QBS_VALAFFR5A,QBS_VALAFFR5B,QBS_VALAFFH6,QBS_VALAFFH6A,QBS_VALAFFH6B,'+
                       'QBS_VALAFFR6,QBS_VALAFFR6A,QBS_VALAFFR6B,QBS_VALAFFH7,QBS_VALAFFH7A,QBS_VALAFFH7B,'+
                       'QBS_VALAFFR7,QBS_VALAFFR7A,QBS_VALAFFR7B,'+
                       'QBS_VALAFFLIB1,QBS_VALAFFLIB2,QBS_VALAFFLIB3,QBS_VALAFFLIB4,QBS_VALAFFLIB5,'+
                       'QBS_VALAFFLIB6,QBS_VALAFFLIB7,'+
                       'QBS_SESSIONINIT FROM '+GetBase (sBase, 'QBPSESSIONBP')+' WHERE QBS_CODESESSION="'+SessionName+'"',true);
        if not QSoc.Eof then
        begin
          For i:= 0 to 3 do
          begin
            //Vérification des dates
            if Q.fields[i].AsDateTime <> Qsoc.fields[i].AsDateTime then
            begin
              SocError:=true;
              lstMess.Add(TraduireMemoire( 'Base '+sBase+' : Erreur (Les dates de la session ne correspondent pas).' ));
              lstMessState.Add('clRed');
              break;
            end;
          end;
          if SocError=false then
          begin
            For i:= 4 to 13 do
            begin
              //Vérification des axes
              if Q.fields[i].AsString <> Qsoc.fields[i].AsString then
              begin
                SocError:=true;
                lstMess.Add(TraduireMemoire( 'Base '+sBase+' : Erreur (Les axes ne correspondent pas).' ));
                lstMessState.Add('clRed');
                break;
              end;
            end;
          end;
          if SocError=false then
          begin
            For i:= 14 to 21 do
            begin
              //Vérification des axes
              if Q.fields[i].AsString <> Qsoc.fields[i].AsString then
              begin
                SocError:=true;
                lstMess.Add(TraduireMemoire( 'Base '+sBase+' : Erreur (Les axes de restriction ne correspondent pas).' ));
                lstMessState.Add('clRed');
                break;
              end;
            end;
          end;
          if SocError=false then
          begin
            For i:= 22 to 64 do
            begin
              //Vérification des axes
              if Q.fields[i].AsString <> Qsoc.fields[i].AsString then
              begin
                SocError:=true;
                lstMess.Add(TraduireMemoire( 'Base '+sBase+' : Erreur (Les valeurs affichées ne correspondent pas).' ));
                lstMessState.Add('clRed');
                break;
              end;
            end;
          end;
          if SocError=false then
          begin
            For i:= 65 to 71 do
            begin
              //Vérification des axes
              if Q.fields[i].AsString <> Qsoc.fields[i].AsString then
              begin
                SocError:=true;
                lstMessState.Add(TraduireMemoire( 'Base '+sBase+' : Erreur (Les libellés des valeurs affichées ne correspondent pas).' ));
                lstMessState.Add('clRed');
                break;
              end;
            end;
          end;
          if SocError=false then
          begin
            //Vérification des axes
            if Q.fields[72].AsString <> Qsoc.fields[72].AsString then
            begin
              SocError:=true;
              if Qsoc.fields[72].AsString='-' then
              begin
                lstMess.Add(TraduireMemoire('Base '+sBase+' : OK (ATTENTION : La session n''est pas initialisée).' ));
                lstMessState.Add('$000080FF');
              end
              else
              begin
                lstMess.Add(TraduireMemoire( 'Base '+sBase+' : OK (ATTENTION : La session n''est pas initialisée).' ));
                lstMessState.add('$000080FF');
              end
            end;
          end;
          if SocError=false then
          begin
            lstMess.Add(TraduireMemoire( 'Base '+sBase+' : OK (Les sessions sont conformes).' ));
            lstMessState.add('clGreen');
          end;
        end
        else
        begin
          lstMess.Add(TraduireMemoire( 'Base '+sBase+' : Erreur (La session n''existe pas).' ));
          lstMessState.add('clRed');
        end;
        Ferme(Qsoc);
      end;
    sBase := ReadTokenSt(sSocGroupe) ;
    until sBase='';
  end else HShowMessage('1;Erreur;Session non trouvée.;W;O;O;O;','','');
  Ferme(Q);
end ;

end.
