unit BPUtil;

interface

Uses HCtrls,SysUtils,
    {$IFDEF EAGLCLIENT}
     UtileAGL,
     uTob,
     {$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     DBCtrls,
     {$ENDIF}
     Uutil,HEnt1;

function BPCalculQteListProrata(ErrInit: double;ListQte: array of double; Rapp: double;var  ListeQteRes: array of double): double;
function ContextBP:integer;
function NbrValAff(session:string):integer;
function TableSalariesOK(session:string) : boolean;
procedure LibValAff(session:string; var TabLibValAff:array of hString);
procedure ReqValAff(session,axes:string; NbAxes:integer; var NbrValue:integer;
                    var ChpArbre,ChpArbreDetail,TableH,WhereH,ChpHisto,TableR,WhereR,ChpReal :string);
procedure DonnePrefixe(NomSession : String; var PrfxH,PrfxR:string);

implementation

uses StrUtils;

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

procedure ReqValAff(session,axes:string; NbAxes:integer; var NbrValue:integer;
                    var ChpArbre,ChpArbreDetail,TableH,WhereH,ChpHisto,TableR,WhereR,ChpReal :string);
var i,k:integer;
    ChpsSum1,ChpsSum2,ReqJoin,PrefixH,PrefixR,ReqWhere:string;
    Q:TQuery;
    TabChpsSumH,TabChpsSumHDet,TabChpsSumR : array [0..6] of hString;
    TabAxes : array [0..9] of string;
    CodeAbrege,Code,Abrege:string;

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

    function RecupListeCodeAbrege(ReqSql:string):string;
    var Q1 : TQuery;
    begin
      Result:='';
      Q1 := OpenSQL(ReqSql,True);
      while not Q1.eof do
      begin
        if Result = '' then Result := Q1.fields[0].AsString
        else Result := Result + ';' + Q1.fields[0].AsString;
        Result := Result + ';' + Q1.fields[1].AsString;
        Q1.next;
      end;
      Result := Result + ';' ;
      ferme(Q1);
    end;

    function VerifTous(Champs : string;Table : integer) : string;
    begin
      if Uppercase(Champs) = Uppercase(TraduireMemoire('<<Tous>>')) then
      begin
        case Table of
          1 : Result := RecupListeChamps('SELECT CO_ABREGE FROM  COMMUN WHERE CO_TYPE="BPP"'); //QUTBPPGPPU
          2 : Result := RecupListeChamps('SELECT CO_ABREGE FROM COMMUN WHERE CO_TYPE="BPA"');  //QUTBPPGPSA
        end;
      end else Result := Champs;
    end;

begin
  axes := AnsiReplaceText(axes,'TRIM','');
  axes := AnsiReplaceText(axes,')','');
  axes := AnsiReplaceText(axes,'(','');

  for i:= 1 to NbAxes do
  begin
    TabAxes[i-1]:=copy(axes,0,PosEx(',',axes)-1);
    TabAxes[i-1] := AnsiReplaceText(TabAxes[i-1],'[PREFIXE]','');
    //TabAxes[i-1] := copy(TabAxes[i-1],4,Length(TabAxes[i-1]));
    Delete(axes,1,PosEx(',',axes));

  end;

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
    bEgin
      //Valeur Historique PAIENCOURS
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i].AsString = 'PPU' then
      begin
        ChpsSum1:=VerifTous(Q.fields[i+7].AsString,1);
        ChpsSum1:=AnsiReplaceText(ChpsSum1,';','+')  ;
        Delete(ChpsSum1,LastDelimiter('+',ChpsSum1),1)  ;
        if i=1 then
        begin
          TableH := 'PAIEENCOURS AS T1';
          PrefixH := 'PPU';
          TabChpsSumH[i-1]:='SUM('+ChpsSum1+')';
          TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
          WhereH:=' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'
        end
        else
        begin
          ReqJoin := '';
          for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PPU'+TabAxes[k-1]+'=T1.'+PrefixH+TabAxes[k-1];
          if Q.fields[1].AsString = 'PPU' then
          begin
            TabChpsSumH[i-1]:='SUM('+ChpsSum1+')';
            TabChpsSumHDet[i-1]:=TabChpsSumH[i-1]
          end
          else
          begin
            TabChpsSumH[i-1]:='(SELECT SUM('+ChpsSum1+') FROM PAIEENCOURS WHERE PPU_DATEDEBUT>="[DATEDEBUT]"'+
                              ' AND PPU_DATEFIN<="[DATEFIN]"' + ReqJoin + ')';
            TabChpsSumHDet[i-1]:='(SELECT SUM('+ChpsSum1+') FROM PAIEENCOURS WHERE PPU_DATEDEBUT=T1.'+PrefixH+'_DATEDEBUT '+
                                 'AND PPU_DATEFIN=T1.'+PrefixR+'_DATEFIN '+    ReqJoin + ')'
          end;
        end
      end;

      //Valeur Réalisé PAIENCOURS
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i+21].AsString = 'PPU' then
      begin
        ChpsSum1:=VerifTous(Q.fields[i+21+7].AsString,1);
        ChpsSum1:=AnsiReplaceText(ChpsSum1,';','+')  ;
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
          ReqJoin := '';
          for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PPU'+TabAxes[k-1]+'=T1.'+PrefixR+TabAxes[k-1];
          if Q.fields[22].AsString = 'PPU' then
          TabChpsSumR[i-1]:='SUM('+ChpsSum1+')'
          else
          TabChpsSumR[i-1]:='(SELECT SUM('+ChpsSum1+') FROM PAIEENCOURS WHERE PPU_DATEDEBUT=T1.'+PrefixR+'_DATEDEBUT '+
                            'AND PPU_DATEFIN=T1.'+PrefixR+'_DATEFIN '+      ReqJoin + ')'
        end
      end;

      //Valeur Historique Effectif Payés
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i].AsString = 'NBS' then
      begin
        if i=1 then
        begin
          TableH := 'PAIEENCOURS AS T1';
          PrefixH := 'PPU';
          TabChpsSumH[i-1]:='SUM(1)';
          TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
          WhereH:=' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'
        end
        else
        begin
          ReqJoin := '';
          for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PPU'+TabAxes[k-1]+'=T1.'+PrefixH+TabAxes[k-1];
          if Q.fields[1].AsString = 'NBS' then
          begin
            TabChpsSumH[i-1]:='SUM(1)';
            TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
          end
          else
          begin
            TabChpsSumH[i-1]:='(SELECT SUM(1) FROM PAIEENCOURS WHERE ' +
                              'PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'+ ReqJoin + ')';
            TabChpsSumHDet[i-1]:='(SELECT SUM(1) FROM PAIEENCOURS WHERE PPU_DATEDEBUT=T1.'+PrefixH+'_DATEDEBUT AND PPU_DATEFIN=T1.'+
                                 PrefixR+'_DATEFIN'+ ReqJoin + ')';
          end;
        end
      end;

      //Valeur Réalisé Effectif Payés
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i+21].AsString = 'NBS' then
      begin
        if i=1 then
        begin
          TableR := 'PAIEENCOURS AS T1';
          PrefixR := 'PPU';
          TabChpsSumR[i-1]:='SUM(1)';
          WhereR:=' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'
        end
        else
        begin
          ReqJoin := '';
          for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PPU'+TabAxes[k-1]+'=T1.'+PrefixR+TabAxes[k-1];
          if Q.fields[22].AsString = 'NBS' then
          TabChpsSumR[i-1]:='SUM(1)'
          else
          TabChpsSumR[i-1]:='(SELECT SUM(1) FROM PAIEENCOURS WHERE' +
                            ' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"' + ReqJoin + ')'
        end
      end;

      //Valeur Historique Effectif Fin de Mois
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i].AsString = 'EFM' then
      begin
        if i=1 then
        begin
          TableH := 'PAIEENCOURS AS T1';
          PrefixH := 'PPU';
          TabChpsSumH[i-1]:='SUM(IIF(([EFFECTIFFINMOIS]),1,0))';
          TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
          WhereH:=' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'
        end
        else
        begin
          ReqJoin := '';
          for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PPU'+TabAxes[k-1]+'=T1.'+PrefixH+TabAxes[k-1];
          if Q.fields[1].AsString = 'EFM' then
          begin
            TabChpsSumH[i-1]:='SUM(IIF(([EFFECTIFFINMOIS]),1,0))';
            TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
          end
          else
          begin
            TabChpsSumH[i-1]:='(SELECT SUM(IIF(([EFFECTIFFINMOIS]),1,0)) FROM PAIEENCOURS WHERE ' +
                              'PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'+ ReqJoin + ')';
            TabChpsSumHDet[i-1]:='(SELECT SUM(IIF((PPU_DATEFIN="[DATEFIN]"),1,0)) FROM PAIEENCOURS WHERE ' +
                                 'PPU_DATEDEBUT=T1.'+PrefixH+'_DATEDEBUT AND PPU_DATEFIN=T1.'+PrefixH+'_DATEFIN'+ ReqJoin + ')';
          end;
        end
      end;

      //Valeur Réalisé Effectif Fin de Mois
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i+21].AsString = 'EFM' then
      begin
        if i=1 then
        begin
          TableR := 'PAIEENCOURS AS T1';
          PrefixR := 'PPU';
            TabChpsSumR[i-1]:='SUM(IIF((PPU_DATEFIN="[DATEFIN]"),1,0))';
            WhereR:=' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'
        end
        else
        begin
          ReqJoin := '';
          for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PPU'+TabAxes[k-1]+'=T1.'+PrefixR+TabAxes[k-1];
          if Q.fields[22].AsString = 'EFM' then
          TabChpsSumR[i-1]:='SUM(IIF((PPU_DATEFIN="[DATEFIN]"),1,0))'
          else
          TabChpsSumR[i-1]:='(SELECT SUM(IIF((PPU_DATEFIN="[DATEFIN]"),1,0)) FROM PAIEENCOURS WHERE' +
                            ' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"' + ReqJoin + ')'
        end
      end;

      //Valeur Historique HISTOBULLETIN
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i].AsString = 'PHB' then
      begin
        ChpsSum1:=Q.fields[i+7].AsString;
        if Uppercase(ChpsSum1)=Uppercase(TraduireMemoire('<<Tous>>')) then ChpsSum1:='' else ChpsSum1:='AND PHB_NATURERUB="'+ChpsSum1+'"';
        ChpsSum2:=Q.fields[i+14].AsString;
        if Uppercase(ChpsSum2)=Uppercase(TraduireMemoire('<<Tous>>')) then ChpsSum2:='' else ChpsSum2:='AND PHB_RUBRIQUE IN ("'+ChpsSum2+'")';
        Delete(ChpsSum2,LastDelimiter(';',ChpsSum2),1);
        ChpsSum2:=AnsiReplaceText(ChpsSum2,';','","')  ;
        if i=1 then
        begin
          TableH := 'HISTOBULLETIN AS T1';
          PrefixH := 'PHB';
          TabChpsSumH[i-1]:='SUM(PHB_MTREM)';
          TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
          WhereH:=' PHB_DATEDEBUT>="[DATEDEBUT]" AND PHB_DATEFIN<="[DATEFIN]" '+ChpsSum1+' '+ChpsSum2
        end
        else
        begin
          ReqJoin := '';
          for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PHB'+TabAxes[k-1]+'=T1.'+PrefixH+TabAxes[k-1];
          TabChpsSumH[i-1]:='(SELECT SUM(PHB_MTREM) FROM HISTOBULLETIN WHERE PHB_DATEDEBUT>="[DATEDEBUT]" AND PHB_DATEFIN<="[DATEFIN]" '+
                           ChpsSum1+' '+ChpsSum2+ReqJoin + ')';
          TabChpsSumHDet[i-1]:='(SELECT SUM(PHB_MTREM) FROM HISTOBULLETIN WHERE PHB_DATEDEBUT=T1.'+PrefixH+'_DATEDEBUT '+
                               'AND PHB_DATEFIN=T1.'+PrefixR+'_DATEFIN '+ChpsSum1+' '+ChpsSum2+ReqJoin + ')'
        end;
      end;

      //Valeur Réalisé HISTOBULLETIN
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
          TableR := 'HISTOBULLETIN AS T1';
          PrefixR := 'PHB';
          TabChpsSumR[i-1]:='SUM(PHB_MTREM)';
          WhereR:=' PHB_DATEDEBUT>="[DATEDEBUT]" AND PHB_DATEFIN<="[DATEFIN]" '+ChpsSum1+' '+ChpsSum2
        end
        else
        begin
          ReqJoin := '';
          for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PHB'+TabAxes[k-1]+'=T1.'+PrefixR+TabAxes[k-1];
          TabChpsSumR[i-1]:='(SELECT SUM(PHB_MTREM) FROM HISTOBULLETIN WHERE PHB_DATEDEBUT=T1.'+PrefixR+'_DATEDEBUT '+
                            'AND PHB_DATEFIN=T1.'+PrefixR+'_DATEFIN ' + ChpsSum1+' '+ChpsSum2+ReqJoin + ')'
        end;
      end;

      //Valeur Historique HISTOCUMSAL
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i].AsString = 'PHC' then
      begin
        ChpsSum1:=Q.fields[i+7].AsString;
        Delete(ChpsSum1,LastDelimiter(';',ChpsSum1),1);
        ChpsSum1:=AnsiReplaceText(ChpsSum1,';','","');
        if Uppercase(ChpsSum1)=Uppercase(TraduireMemoire('<<Tous>>')) then ChpsSum1:='' else ChpsSum1:='AND PHC_CUMULPAIE IN ("'+ChpsSum1+'")';
        if i=1 then
        begin
          TableH := 'HISTOCUMSAL AS T1';
          PrefixH := 'PHC';
          TabChpsSumH[i-1]:='SUM(PHC_MONTANT)';
          TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
          WhereH:=' PHC_DATEDEBUT>="[DATEDEBUT]" AND PHC_DATEFIN<="[DATEFIN]" '+ChpsSum1
        end
        else
        begin
          ReqJoin := '';
          for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PHC'+TabAxes[k-1]+'=T1.'+PrefixH+TabAxes[k-1];
          TabChpsSumH[i-1]:='(SELECT SUM(PHC_MONTANT) FROM HISTOCUMSAL WHERE PHC_DATEDEBUT>="[DATEDEBUT]"'+
                            ' AND PHC_DATEFIN<="[DATEFIN]" ' + ChpsSum1 + ' '+ ReqJoin + ')';
          TabChpsSumHDet[i-1]:='(SELECT SUM(PHC_MONTANT) FROM HISTOCUMSAL WHERE PHC_DATEDEBUT=T1.'+PrefixH+'_DATEDEBUT '+
                               'AND PHC_DATEFIN=T1.'+PrefixR+'_DATEFIN '+ChpsSum1 + ' '+ ReqJoin + ')'
        end
      end;

      //Valeur Réalisé HISTOCUMSAL
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i+21].AsString = 'PHC' then
      begin
        ChpsSum1:=Q.fields[i+21+7].AsString;
        Delete(ChpsSum1,LastDelimiter(';',ChpsSum1),1);
        ChpsSum1:=AnsiReplaceText(ChpsSum1,';','","')  ;
        if Uppercase(ChpsSum1)=Uppercase(TraduireMemoire('<<Tous>>')) then ChpsSum1:='' else ChpsSum1:='AND PHC_CUMULPAIE IN ("'+ChpsSum1+'")';
        if i=1 then
        begin
          TableR := 'HISTOCUMSAL AS T1';
          PrefixR := 'PHC';
          TabChpsSumR[i-1]:='SUM(PHC_MONTANT)';
          WhereR:=' PHC_DATEDEBUT>="[DATEDEBUT]" AND PHC_DATEFIN<="[DATEFIN]" '+ChpsSum1
        end
        else
        begin
          ReqJoin := '';
          for k:= 1 to NbAxes do ReqJoin := ReqJoin + ' AND PHC'+TabAxes[k-1]+'=T1.'+PrefixR+TabAxes[k-1];
          TabChpsSumR[i-1]:='(SELECT SUM(PHC_MONTANT) FROM HISTOCUMSAL WHERE PHC_DATEDEBUT=T1.'+PrefixR+'_DATEDEBUT '+
                            ChpsSum1 + ' '+ ReqJoin + ')'
        end
      end;

      //Valeur Historique SALARIES
      ChpsSum1:='';ChpsSum2:='';
      if Q.fields[i].AsString = 'PSA' then
      begin
        ChpsSum1:=VerifTous(Q.fields[i+7].AsString,2);
        ChpsSum1:=AnsiReplaceText(ChpsSum1,';','+')  ;
        CodeAbrege:=RecupListeCodeAbrege('SELECT CO_CODE,CO_ABREGE FROM COMMUN WHERE CO_TYPE="BPA"');
        While CodeAbrege<>'' do
        begin
          Code:=ReadTokenSt(CodeAbrege);
          Abrege:=ReadTokenSt(CodeAbrege);
          ChpsSum1:=AnsiReplaceText(ChpsSum1,Code,Abrege);
        end;
        Delete(ChpsSum1,LastDelimiter('+',ChpsSum1),1)  ;
        if i=1 then
        begin
          TableH := 'PAIEENCOURS AS T1';
          PrefixH := 'PPU';
          ReqWhere:='';
          for k:= 1 to NbAxes do
          begin
            if ReqWhere = '' then ReqWhere := ReqWhere + ' AND PSA'+TabAxes[k-1]+'='+PrefixH+TabAxes[k-1]
            else ReqWhere := ReqWhere + ' AND PSA'+TabAxes[k-1]+'='+PrefixH+TabAxes[k-1];
          end;
          TabChpsSumH[i-1]:='(SELECT SUM('+ChpsSum1+')*[NBMOIS] FROM SALARIES WHERE '+ReqWhere+')';
          TabChpsSumHDet[i-1]:=TabChpsSumH[i-1];
          WhereH:=' PPU_DATEDEBUT>="[DATEDEBUT]" AND PPU_DATEFIN<="[DATEFIN]"'
        end
        else
        begin
          ReqJoin := '';
          for k:= 1 to NbAxes do
          begin
            if ReqJoin = '' then ReqJoin := 'PSA'+TabAxes[k-1]+'=T1.'+PrefixH+TabAxes[k-1]
            else ReqJoin := ReqJoin + ' AND PSA'+TabAxes[k-1]+'=T1.'+PrefixH+TabAxes[k-1];
          end;
          TabChpsSumH[i-1]:='(SELECT SUM('+ChpsSum1+')*[NBMOIS] FROM SALARIES WHERE ' + ReqJoin + ')';
          TabChpsSumHDet[i-1]:=TabChpsSumH[i-1]
        end
      end

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

  For i:=1 to 6 do
  begin
    if i=1 then
    begin
      if TabChpsSumR[i-1]='' then ChpReal:=ChpReal+'0,' else ChpReal:=ChpReal+TabChpsSumR[i-1]+',';
      if TabChpsSumHDet[i-1]='' then ChpHisto:=ChpHisto+',0' else ChpHisto:=ChpHisto+','+TabChpsSumHDet[i-1];
    end
    else if i = 2 then
    begin
      if TabChpsSumR[i-1]='' then ChpReal:='0,' + ChpReal else ChpReal:=TabChpsSumR[i-1]+','+ ChpReal;
      if TabChpsSumHDet[i-1]='' then ChpHisto:=',0' + ChpHisto else ChpHisto:=','+TabChpsSumHDet[i-1] + ChpHisto;
    end
    else if i > 2 then
    begin
      if TabChpsSumR[i-1]='' then ChpReal:=ChpReal+'0,' else ChpReal:=ChpReal+TabChpsSumR[i-1]+',';
      if TabChpsSumHDet[i-1]='' then ChpHisto:=ChpHisto+',0' else ChpHisto:=ChpHisto+','+TabChpsSumHDet[i-1];
    end;
  end;
  if TabChpsSumR[6]='' then ChpReal:=','+ChpReal+'0' else ChpReal:=','+ChpReal+TabChpsSumR[6];
  if TabChpsSumHDet[6]='' then ChpHisto:=ChpHisto+',0' else ChpHisto:=ChpHisto+','+TabChpsSumHDet[6];

end;

end.
