unit CalcOleGRC;

interface

uses SysUtils,
{$IFNDEF EAGLCLIENT}
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ELSE}
    Utob,
{$ENDIF}
    HCtrls,HEnt1
    ;
function RTCalcOLEEtat(sf,sp : string) : Variant ;
function rtchamptolibelle(sp : string) : string;
function RtLibellesIntervenants(sp : string) : string;
function RtDecoupeChamp (sp : string) : string;
function RtLibellesDestMail(sp : string) : string;
implementation


function rtchamptolibelle(sp : string) : string;
begin
result:=ChampToLibelle(sp);
//if (result[1]='.') or (result[1]='-') then result:='';
if (copy(result,1,1)='.') or (copy(result,1,1)='-') then result:='';
end;

function RtLibellesIntervenants(sp : string) : string;
var ListeNoms,libelle,CodeAuxiliaire : string;
    Q : TQuery;
    NoAction : integer;
begin
  result:=''; ListeNoms:='';
  CodeAuxiliaire := ReadTokenSt (sp);
  NoAction := Valeuri(ReadTokenSt (sp));
  Q:=OpenSql('SELECT ARS_LIBELLE FROM ACTIONINTERVENANT LEFT JOIN RESSOURCE on RAI_RESSOURCE = ARS_RESSOURCE  WHERE RAI_AUXILIAIRE = "'+CodeAuxiliaire+'" AND RAI_NUMACTION ='+IntToStr(NoAction),true,-1,'',true);
  while not Q.EOF do
  begin
    libelle:=Q.FindField('ARS_LIBELLE').AsString+';';
    if length(ListeNoms)+ length(libelle) < 80 then
        ListeNoms:=ListeNoms+libelle
    else
        break;
    Q.Next;
  end;
  Ferme (Q);
  result:=ListeNoms;
end;

function RtDecoupeChamp(sp : string) : string;
var Champ :  String;
Pos,Long,PosOri,LongOri : integer;
begin
Champ:=ReadTokenPipe(sp,'|') ;
result:=Champ;
if Champ <> '' then
    begin
    PosOri:=StrToInt(ReadTokenPipe(sp,'|')) ;
    LongOri:=StrToInt(ReadTokenPipe(sp,'|')) ;
    Pos:=PosOri;
    Long:=LongOri;
    if (Champ<>'') and (Pos<>0) and (Long<>0) then
        begin
        if PosOri = 1 then
           begin
           if Length(Champ) > Long then
           { on revient en arrière jusqu'au dernier blanc
             pour ne pas tronquer un mot }
              while Copy(Champ,Long,1) <> ' ' Do
              begin
                Dec(Long);
                if Long = 0 then
                begin
                  Long:=LongOri;
                  break;
                end;
              end;
           end;
        if PosOri > 1 then
        begin
           if Length(Champ) <= (PosOri -1) then
              begin
              result:='';
              exit;
              end;
           while Copy(Champ,Pos,1) <> ' ' Do //Dec(Pos);
           begin
              Dec(Pos);
              if Pos = 0 then
              begin
                Pos:=PosOri;
                break;
              end;
           end;
           { correction du long qui est en fait la position : 200
             alors que la vrai longueur est long - PosOri +1 }
           Long:=Long-PosOri+1;
           Long:=Long+(PosOri-Pos);
           { on saute le caractère blanc }
           if PosOri<>Pos then
             Inc(Pos);
        end;
        { limite d'impression = 120 caractères par ligne }
        if long > 118 then long:=118;          
        result:=Copy(Champ,Pos,Long);
        end;
    end;
end;

function RtLibellesDestMail(sp : string) : string;
var Critere,ListeNoms,libelle,ListeCodes,CodeAuxiliaire : string;
    Q : TQuery;
begin
  result:=''; ListeNoms:='';
  ListeCodes := ReadTokenPipe (sp,'|');
  CodeAuxiliaire := ReadTokenPipe (sp,'|');
  While ListeCodes <> '' do
  begin
      Critere:=uppercase(Trim(ReadTokenSt(ListeCodes))) ;
      if Critere<>'' then
      begin
      Q:=OpenSQL ('Select C_NOM from contact where C_TYPECONTACT = "T" AND C_AUXILIAIRE = "'+CodeAuxiliaire+'" AND C_NUMEROCONTACT ='+Critere,true,-1,'',true);
      if not Q.EOF then libelle:=Q.FindField('C_NOM').AsString+';';
      Ferme (Q);
      if length(ListeNoms)+ length(libelle) < 80 then
          ListeNoms:=ListeNoms+libelle
      else
          break;
      end;
  end;
  result:=ListeNoms;
end;

function RTCalcOLEEtat(sf,sp : string) : Variant ;
BEGIN
if copy(sf,1,2)<>'RT' then exit
else if sf='RTCHAMPTOLIBELLE' then Result:=rtchamptolibelle(sp)
else if sf='RTLIBELLESINTERVENANTS' then Result:=RtLibellesIntervenants(sp)
else if sf='RTDECOUPECHAMP' then Result:=RtDecoupeChamp(sp)
//else if copy(sf,1,5)='GCPDF' then Result:=GCGetPortsEdt(sf,sp)
else if sf='RTLIBELLESDESTMAIL' then Result:=RtLibellesDestMail(sp)
else Result := '';
END;

end.
