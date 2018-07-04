unit ZReleve;

//=======================================================
//=================== Clés primaires ====================
//=======================================================
// RBLIGECR : CRB_JOURNAL,  CRB_EXERCICE, CRB_DATECOMPTABLE,
//            CRB_NUMLIGNE, CRB_NUMINTERNE,

interface

uses Classes, Db, DbTables, SysUtils, ComCtrls, Forms, { +- Delphi -+ }
     HCtrls, HEnt1, HStatus,                           { +- Bibli  -+ }
     Ent1, ed_tools, SaisUtil, SaisComm, HCompte,      { +- Compta -+ }
     TZ, ZTypes ;                                      { +-  PFU   -+ }

type
  TZReleve = class
  private
    Lignes : TZF ;
    // Références du where pour l'objet
    FDecimales : Integer ;
    FOpt : ROpt ;
    FJal : string ;
  public
    constructor Create(Decimales : Integer; Jal : string; Opt : ROpt) ;
    destructor  Destroy ; override ;
    function    CreateRow(Q : TQuery; Row : LongInt) : TZF ;
    procedure   DeleteRow(Row : LongInt) ;
    function    GetRow(Row : LongInt) : TZF ;
//    function    FindCompte(var NumCompte, NumAux : string) : Integer ;
    function    Read : Boolean ;
    function    Write : Boolean ;
    function    Del : Boolean ;
  end ;

implementation

constructor TZReleve.Create(Decimales : Integer; Jal : string; Opt : ROpt) ;
begin
Lignes:=TZF.Create('VLIGNES', nil, -1) ;
FDecimales:=Decimales ;
FJal:=Jal ;
FOpt:=Opt ;
end ;

destructor TZReleve.Destroy ;
begin
if Lignes<>nil then Lignes.Free ;
end ;

function TZReleve.CreateRow(Q : TQuery; Row : LongInt) : TZF ;
var Ligne : TZF ; RowRef : LongInt ;
begin
if Row>Lignes.Detail.Count-1 then RowRef:=-1
                             else RowRef:=Row ;
if Q<>nil then Ligne:=TZF.CreateDB('RBLIGECR', Lignes, RowRef, Q)
          else Ligne:=TZF.Create('RBLIGECR', Lignes, RowRef) ;
Result:=Ligne ;
end ;

procedure TZReleve.DeleteRow(Row : LongInt) ;
var Ligne : TZF ;
begin
if Row>Lignes.Detail.Count-1 then Exit ;
Ligne:=TZF(Lignes.Detail[Row]) ; Ligne.Free ;
end ;

function TZReleve.GetRow(Row : LongInt) : TZF ;
begin
Result:=nil ; if Row>Lignes.Detail.Count-1 then Exit ; Result:=TZF(Lignes.Detail[Row]) ;
end ;

(*
function TZReleve.FindCompte(var NumCompte, NumAux : string) : Integer ;
var i : integer ;
begin
Result:=-1 ;
NumCompte:=Trim(NumCompte) ; if NumCompte='' then Exit ;
NumCompte:=BourreLaDonc(NumCompte, fbGene) ;
if NumAux<>'' then NumAux:=BourreLaDonc(NumAux, fbAux) ;
for i:=0 to Ecrs.Count-1 do
  if (TZF(Ecrs[i]).GetValue('HB_COMPTE1')=NumCompte) and
     (TZF(Ecrs[i]).GetValue('HB_COMPTE2')=NumAux) then
     begin Result:=i ; Exit ; end ;
end ;
*)

function TZReleve.Read : Boolean ;
var QLig : TQuery ;
begin
QLig:=OpenSQL('SELECT * FROM RBLIGECR WHERE CRB_JOURNAL="'+FJal+'"'
             +' AND CRB_EXERCICE="'+FOpt.Exo+'"'
             +' AND CRB_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, 1))+'"'
             +' AND CRB_DATECOMPTABLE<="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"'
             +' ORDER BY CRB_JOURNAL, CRB_EXERCICE, CRB_NUMLIGNE', TRUE) ;
Result:=not QLig.EOF ;
while not QLig.EOF do
  begin
  // Charger le TZF
  CreateRow(QLig, Lignes.Detail.Count+1) ;
  QLig.Next ;
  end ;
Ferme(QLig) ;
end ;

function TZReleve.Write : Boolean ;
//var i : Integer ;
begin
Result:=TRUE ;
Lignes.SetAllModifie(TRUE) ;
Lignes.InsertDBByNivel(TRUE) ;
(*
InitMove(Lignes.Detail.Count, TraduireMemoire('Ecriture du relevé')) ;
for i:=0 to Lignes.Detail.Count-1 do
    begin
    MoveCur(FALSE) ;
    if TZF(Lignes.Detail[i]).GetValue('BADROW')<>'X' then TZF(Lignes.Detail[i]).InsertDB(nil) ;
    end ;
FiniMove ;
*)
end ;

function TZReleve.Del : Boolean ;
var Where : string ;
begin
Result:=TRUE ;
Where:=' CRB_JOURNAL="'+FJal+'"'
      +' AND CRB_EXERCICE="'+FOpt.Exo+'"'
      +' AND CRB_DATECOMPTABLE>="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, 1))+'"'
      +' AND CRB_DATECOMPTABLE<="'+USDateTime(EncodeDate(FOpt.Year, FOpt.Month, FOpt.MaxJour))+'"' ;
if ExecuteSQL('DELETE FROM RBLIGECR WHERE'+Where)<=0 then
   begin Result:=FALSE ; V_PGI.IOError:=oeSaisie ; Exit ; end ;
end ;

end.
