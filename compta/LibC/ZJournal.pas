unit ZJournal;

//=======================================================
//=================== Clés primaires ====================
//=======================================================
// JOURNAL : J_JOURNAL

interface

uses SysUtils, //Classes, Db, DbTables, ComCtrls, Forms, { +- Delphi -+ }
     UTob, //HCtrls, HEnt1,                              { +- Bibli  -+ }
     Ent1, //ed_tools, SaisUtil, SaisComm, HCompte,      { +- Compta -+ }
     {$IFNDEF EAGLCLIENT}
     DB,DBTables,
     {$ENDIF}
     SaisUtil,  // pour le GetNewNumJal
     HCtrls // pour le OpenSQL
     ;                                                { +-  PFU   -+ }

type
 // cet objet n'est pas utilise dans la compta
 TZJournal = class
  private
  Journal : TOB ;
  FCode : string ;
  function    GetCount : Integer ;
  public
  constructor Create(Code : string) ;
  destructor  Destroy ; override ;
  function    GetValue(Nom : string) : Variant ;
  procedure   PutValue(Nom : string ; Valeur : Variant) ;
  function    Load : Boolean ;
  property    Count : Integer read GetCount ;
  end ;

(* TZListJournal = class
  private
   FTOBJournal        : TOB ; // TOB contenant l'ensemble des journaux chargé
   FTOBLigneJournal   : TOB ; // pointeur sur la ligne courante
   FInIndex           : integer ; // index de la ligne courante
  public
   constructor Create ;
   destructor  Destroy ; override ;
   function    Load(CodeJournal : string) : boolean ; // charge un journal soit depuis la base ou la tob FTOBJournal et positionne FTOBLigneJournal
   function    GetValueByIndex(Nom : string; Index : integer) : Variant ;
   function    GetValue(Nom : string) : Variant ;
   function    GetNumJal(vDtDateComptable : TDateTime) : integer; // retourne le prochain numero de journal
   function    GetTabletteNature : string; // retourne le nom de la tablette des natures de piece en fonction de la nature du journal
  end ;

  *)

implementation

constructor TZJournal.Create(Code : string) ;
begin
FCode:=Code ;
end ;

destructor TZJournal.Destroy ;
begin
if Journal<>nil then Journal.Free ;
end ;

function TZJournal.GetCount : Integer ;
begin
Result:=Journal.Detail.Count ;
end ;

function TZJournal.GetValue(Nom : string) : Variant ;
begin
Result:=#0 ;
if Journal<>nil then Result:=Journal.GetValue(Nom) ;
end ;

procedure TZJournal.PutValue(Nom : string ; Valeur : Variant) ;
begin
if Journal<>nil then Journal.PutValue(Nom, Valeur) ;
end ;

function TZJournal.Load : Boolean ;
var Key : string ;
begin
Journal:=TOB.CreateDB('JOURNAL', nil, -1, nil) ;
Key:='"'+FCode+'"' ;
Journal.SelectDB(Key, nil) ;
Result:=TRUE ;
end ;

//*****************************************************************************//
 (*
constructor TZListJournal.Create;
begin
FTOBJournal:=TOB.Create('', nil, -1) ;
end ;

destructor TZListJournal.Destroy ;
begin
FTOBJournal.Free ; FTOBJournal := nil ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... : 14/02/2002
Description .. : chargement des info du journal :
Suite ........ : - soit depuis la tob journal s'ila déjà été charger
Suite ........ : - soit depuis la base
Suite ........ : et place le pointeur de ligne courant sur cette ligne
Mots clefs ... :
*****************************************************************}
function TZListJournal.Load( CodeJournal : string) : boolean ;
var Q : TQuery ; i : integer ;
begin
Result:=false ;
CodeJournal:=Trim(CodeJournal) ; if CodeJournal='' then Exit ;
for i:=0 to FTOBJournal.Detail.Count-1 do
  if UpperCase(FTOBJournal.Detail[i].GetValue('J_JOURNAL'))=UpperCase(CodeJournal) then
   begin Result:=true ; FTOBLigneJournal:=FTOBJournal.Detail[i] ; FInIndex:=i ; Exit ; end ;
Q:=OpenSQL('SELECT * FROM JOURNAL WHERE J_JOURNAL="'+UpperCase(CodeJournal)+'"',true);
if FTOBJournal.LoadDetailDB('JOURNAL', '', '', Q, TRUE) then
begin FInIndex:=FTOBJournal.Detail.Count-1 ; FTOBLigneJournal:=FTOBJournal.Detail[FInIndex] ; result:=true ; end ;
Ferme(Q) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /
Description .. : retourne la valeur du champ <nom> de la fille <index>
Mots clefs ... :
*****************************************************************}
function TZListJournal.GetValueByIndex(Nom : string; Index : integer) : Variant ;
begin
Result:=FTOBJournal.Detail[Index].GetValue(Nom) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /
Description .. : retourne la valeur du champ <nom> pour le journal courant
Mots clefs ... :
*****************************************************************}
function TZListJournal.GetValue(Nom : string) : Variant ;
begin
Result:=FTOBLigneJournal.GetValue(Nom) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /
Description .. : retourne le prochain numero de piece du journal en fonction
Suite ........ : de la date comptable
Mots clefs ... :
*****************************************************************}
function TZListJournal.GetNumJal(vDtDateComptable : TDateTime) : integer;
begin
 result:=GetNewNumJal(GetValue('J_JOURNAL'),true,vDtDateComptable,true,GetValue('J_COMPTEURNORMAL'),GetValue('J_MODESAISIE'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 14/02/2002
Modifié le ... :   /  /
Description .. : Retourne le nom de la tablette pour la nature de piece ne
Suite ........ : fonction du type de journal
Mots clefs ... :
*****************************************************************}
function TZListJournal.GetTabletteNature : string;
begin
 case CaseNatJal(GetValue('J_NATUREJAL')) of
  tzJVente       : result:='ttNatPieceVente' ;
  tzJAchat       : result:='ttNatPieceAchat' ;
  tzJBanque      : result:='ttNatPieceBanque' ;
  tzJEcartChange : result:='ttNatPieceEcartChange' ;
  tzJOD          : result:='ttNaturePiece' ;
 end ; // case
end;
    *)
end.
