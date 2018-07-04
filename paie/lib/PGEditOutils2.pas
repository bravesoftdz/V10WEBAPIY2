{
PT1   13/09/2007 FC V_80 Prise en compte des habilitations dans MinMaxTablette
}
unit PGEditOutils2;

interface
uses
{$IFDEF EAGLCLIENT}
      UTOB,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      HCtrls
      ,P5Def,StrUtils
      ;

procedure RecupMinMaxTablette(Pref,NomTable,CodeTyp : string ;var Min,Max : string);


var

PGGlbEtabDe,PGGlBEtabA,PGGlbSalDe,PGGlbSalA : string;

implementation


procedure RecupMinMaxTablette(Pref,NomTable,CodeTyp : string ;var Min,Max : string);
Var QQuery : TQuery;
  Where,stHabilitation : String;
  Longueur : integer;
begin
Min:=''; Max:='';
if Length(CodeTyp)<3 then exit;
If Pref='CC' then
   begin
   QQuery:=OpenSql('SELECT MIN(CC_CODE),MAX(CC_CODE) FROM CHOIXCOD '+
   'WHERE CC_TYPE="'+CodeTyp+'"',True);
   If not QQuery.eof then // PORTAGECWAS
     Begin
     Min:=QQuery.Fields[0].asstring;
     Max:=QQuery.Fields[1].asstring;
     End;
   Ferme(QQuery);
   end;
If Pref='CO' then
   begin
   QQuery:=OpenSql('SELECT MIN(CO_CODE),MAX(CO_CODE) FROM COMMUN '+
   'WHERE CO_TYPE="'+CodeTyp+'"',True);
   if not QQuery.eof then //PORTAGECWAS
     Begin
     Min:=QQuery.Fields[0].asstring;
     Max:=QQuery.Fields[1].asstring;
     End;
   Ferme(QQuery);
   end;
If Pref='PG' then
   begin
   //DEB PT1
   Where := '';
   if NomTAble = 'SALARIES' then
     if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
       if copy(MonHabilitation.LeSQL, 1, 3) <> 'PSA' then
       begin
         Longueur := Length(MonHabilitation.LeSQL);
         Longueur := Longueur - 2;
         stHabilitation := MidStr(MonHabilitation.LeSQL, 4, Longueur);
         stHabilitation := 'PSA' + stHabilitation;
         Where := ' WHERE ' + stHabilitation;
       end
       else
         Where := ' WHERE ' + MonHabilitation.LeSQL;
   //FIN PT1
   QQuery:=OpenSql('SELECT MIN('+CodeTyp+'),MAX('+CodeTyp+') FROM '+NomTAble+'' + Where,True);
   If not QQuery.eof then //PORTAGECWAS
     Begin
     Min:=QQuery.Fields[0].asstring;
     Max:=QQuery.Fields[1].asstring;
     End;
   Ferme(QQuery);
   end;
end;

end.
