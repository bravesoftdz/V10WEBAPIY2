unit RecupNegSQL7;

interface
uses
    SysUtils, HEnt1, Hctrls, UTOB
   ,CbpMCD
   ,CbpEnumerator
    ;

procedure Ecrire_Datas(ItemInd : integer; NomChampTable : string; NumTable, NumeroChamp : integer;
                       ValeurSeq : variant; var SQL_Champs, SQL_Datas, SQL_Sep : string;
                       var BlobPresent : boolean);
procedure PrepareListe(ItemInd : integer; TobTable : TOB; var SQL_Champs : string);

implementation


procedure Ecrire_Datas(ItemInd : integer; NomChampTable : string; NumTable, NumeroChamp : integer;
                       ValeurSeq : variant; var SQL_Champs, SQL_Datas, SQL_Sep : string;
                       var BlobPresent : boolean);
//
//  ItemInd : valeur precisant quelle est la cible de la récupération.
//          O = Insert direct dans la base
//          1 = Fichier Sequentiel pour SQL Server 7
//
var
    st_trav1, st_trav2 : string;
		Mcd : IMCDServiceCOM;
		Table     : ITableCOM ;

begin

MCD := TMCD.GetMcd;
if not mcd.loaded then mcd.WaitLoaded();
//Table := Mcd.GetTable(Mcd.PrefixeToTable( ))

//st_trav1 := V_PGI.DEChamps[NumTable, NumeroChamp].Tipe;
st_trav1 := mcd.getField(NomChampTable).tipe;

if st_trav1 = 'BLOB' then
    begin
    BlobPresent := True;
    Exit;
    end;

if (st_trav1='INTEGER') or (st_trav1='SMALLINT') then st_trav2 := IntToStr(ValeurSeq)    else
if (st_trav1='DOUBLE')  or (st_trav1='RATE')     then st_trav2 := FloatToStr(ValeurSeq)  else
if (st_trav1='DATE')                             then st_trav2 := UsDateTime(ValeurSeq)  else
                                                      st_trav2 := ValeurSeq;
while Pos(',', st_trav2) <> 0 do st_trav2[Pos(',', st_trav2)] := '.';
if (st_trav1 = 'DATE') or (st_trav1 = 'COMBO') or (st_trav1 = 'BOOLEAN')or (st_trav1 = 'BLOB') then
//
//   ATTENTION : SPECIFIQUE POUR SQL 7
//
    if ItemInd = 0 then
        SQL_Datas  := SQL_Datas + SQL_Sep + '"' + st_trav2 + '"'
        else
        SQL_Datas  := SQL_Datas + SQL_Sep + st_trav2
else if (Pos('CHAR', st_trav1) <> 0) then
    begin
    while Pos('"', st_trav2) <> 0 do st_trav2[Pos('"', st_trav2)] := '''';
    ReadTokenPipe(st_trav1, '(');
    st_trav1 := Copy(st_trav1, 0, Pos(')', st_trav1) - 1);
    st_trav2 := Copy(st_trav2, 0, StrToInt(st_trav1));
//
//   ATTENTION : SPECIFIQUE POUR SQL 7
//
    if ItemInd = 0 then
        SQL_Datas  := SQL_Datas + SQL_Sep + '"' + st_trav2 + '"'
        else
        SQL_Datas  := SQL_Datas + SQL_Sep + st_trav2;
    end
    else
    SQL_Datas  := SQL_Datas + SQL_Sep + st_trav2;
//
//   ATTENTION : SPECIFIQUE POUR SQL 7
//
if ItemInd = 0 then SQL_Champs := SQL_Champs + SQL_Sep + V_PGI.DEChamps[NumTable, NumeroChamp].Nom;
//
//
//   ATTENTION : SPECIFIQUE POUR SQL 7
//
if ItemInd = 0 then
   SQL_Sep    := ', '
   else
   SQL_Sep    := #9;
end;

procedure PrepareListe(ItemInd : integer; TobTable : TOB; var SQL_Champs : string);
var
   i_ind1 : integer;

begin
SQL_Champs := '';
for i_ind1 := 1 to TobTable.NbChamps do
    begin
    if SQL_Champs = '' then
        SQL_Champs := TobTable.GetNomChamp(i_ind1)
        else
        SQL_Champs := SQL_Champs + #9 + TobTable.GetNomChamp(i_ind1);
    end;
end;

end.
