unit DicoAF;

interface
Uses classes,HEnt1,
{$IFDEF EAGLCLIENT}

{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}

sysutils,Utob,hmsgbox,HCtrls,Ent1,UiUtil;

// Fonctions de traduction vocab affaire / Gestion interne
function  TraduitGA(phrase : string) : string ;
Procedure PGIBoxAF (Msg, Titre : String);
procedure PGIInfoAF (Msg, Titre : String);
function  PGIAskCancelAF (Msg, Titre : string):integer ;
function  PGIAskAF (Msg, Titre : string):integer ;

// Fonction de gestion des niveaux de gestion d'affaires  et des modules
Function VersionGAAvancee : Boolean;

// Fonction de modification/création de ligne dans une tablette
function ModifierTablette(TypeTable,TypeCode,code,Libelle:string; LngCode:integer):boolean;

         // fct utiliser pour les groupe de confidentialité
function GIGereCritGroupeConf : string;

implementation
uses EntGC;

function TraduitMotGA(mot : string; NumDico : integer =0) : string ;
var i : integer ;
    TOBDet : TOB;
begin
result:=mot ;
If NumDico<>0 Then
    result:=VH_GC.AFTOBTraduction.Detail[NumDico].GetValue('CC_LIBRE')
Else
    for i:=1 to VH_GC.AFTOBTraduction.Detail.Count-1 do
        BEGIN
        TobDet := VH_GC.AFTOBTraduction.Detail[i];
        if AnsiUppercase(TobDet.GetValue('CC_LIBELLE'))=AnsiUppercase(mot) then
            result:=TobDet.GetValue('CC_LIBRE') ;
        END;

if result<>mot then
   begin
    if (length(mot)>1) and (ord(mot[1])<=ord('Z')) and (ord(mot[2])>ord('Z')) then
       result:=FirstMajuscule(result)
    else if ord(mot[1])<=ord('Z') then result:=AnsiUppercase(result);
    End ;
end ;

function TraduitGA(phrase : string) : string ;
var mot,mot2,StFind : string ;
    i, ipos : integer;
    TobDet : TOB;
begin
Result := phrase;
if (VH_GC.AFTOBTraduction.Detail.Count = 0) then begin
   TraduireMemoire(Result);// mcd 28/01/02 pour traduction autre langue
   Exit;
   end;
// Recherche automatique des expressions du dico
for i:=0 to VH_GC.AFTOBTraduction.Detail.Count-1 do
    Begin
    TobDet := VH_GC.AFTOBTraduction.Detail[i];
    StFind := AnsiUppercase(TobDet.GetValue('CC_LIBELLE'));
    ipos:=Pos(StFind,AnsiUppercase(Phrase));
    While ipos>0 do
        BEGIN
        mot := Copy(phrase, ipos, Length(StFind));
        mot2:=TraduitMotGA(mot,i) ;
        Phrase:=FindEtReplace(phrase,mot,mot2,true) ;
        ipos:=Pos(StFind,AnsiUppercase(phrase));
        End ;                   
    End;
Result := Phrase;
TraduireMemoire(Result); // mcd 28/01/02 pour traduciton autre langue
end ;

Procedure PGIBoxAF (Msg, Titre : String);
Begin
If (Titre ='') Then Titre := TitreHalley;
PGIBox(TraduitGA(Msg),TraduitGA(titre));
End;

procedure PGIInfoAF (Msg, Titre : string);
Begin
If (Titre ='') Then Titre := TitreHalley;
PGIInfo (TraduitGA(Msg),TraduitGA(titre));
End;

function PGIAskCancelAF (Msg, Titre : string) :integer ;
Begin
If (Titre ='') Then Titre := TitreHalley;
Result := PGIAskCancel(TraduitGA(Msg),TraduitGA(titre));
End;

function PGIAskAF (Msg, Titre : string):integer ;
Begin
If (Titre ='') Then Titre := TitreHalley;
Result := PGIAsk(TraduitGA(Msg),TraduitGA(titre));
End;


Function VersionGAAvancee : Boolean;
BEGIN
Result := true;
if (ctxAffaire in V_PGI.PGIContexte) and (ctxNegoce in V_PGI.PGIContexte ) then Result := True;
// A voir en fonction des seria mises en place par la suite
END;


{***********A.G.L.***********************************************
Auteur  ...... : Pierre LENORMAND
Créé le ...... : 26/09/2000
Modifié le ... : 26/09/2000
Description .. : Modification ou création d une ligne dans une tablette
Suite ........ :
Mots clefs ... : MODIFICATION;TABLETTE
*****************************************************************}
function ModifierTablette(TypeTable,TypeCode,code,Libelle:string; LngCode:integer):boolean;
var TTablette : TOB;
Req, table:string;
begin
Result:=false;

If (TypeTable = 'YX') then table:='CHOIXEXT' Else
If (TypeTable = 'CC') then table:='CHOIXCOD' Else
If (TypeTable = 'CO') then table:='COMMUN' Else Exit;

TTablette := TOB.CREATE (table, NIL, -1);
try
TTablette.initValeurs;
Req:='SELECT '+TypeTable+'_CODE FROM '+table+' WHERE '+TypeTable+'_TYPE="'+TypeCode+'" AND '+TypeTable+'_CODE="'+code+'"' ;
if ExisteSQL(Req) then
    begin
    TTablette.PutValue (TypeTable+'_TYPE', TypeCode);
    TTablette.PutFixedStValue(TypeTable+'_CODE', code, 1,LngCode, tctrim, true);
    TTablette.PutValue (TypeTable+'_ABREGE', copy(libelle,1,17));
    TTablette.PutValue (TypeTable+'_LIBELLE', libelle);
    TTablette.InsertOrUpdateDB;
    end
else
    begin
    TTablette.PutValue (TypeTable+'_TYPE', TypeCode);
    TTablette.PutFixedStValue(TypeTable+'_CODE', code, 1,LngCode, tctrim, true);
    TTablette.PutValue (TypeTable+'_ABREGE', copy(libelle,1,17));
    TTablette.PutValue (TypeTable+'_LIBELLE', libelle);
    TTablette.InsertDB(nil);
    end;

Result:=true;
finally
TTablette.Free;
end;

End;


//NCX 21/12/01
//Penser de modifier à l'identique si modif de GereCritGroupeConf....
// dans GalOutils du DP
function GIGereCritGroupeConf : string;
Var Q: TQuery;
    critere, plus : String;
begin
  // si superviseur, le crit. restera vide car on voit tout
  // 27/07/01 sauf en Mode pme : cacher les dossiers qui ont une conf.
  critere := '';
  // pour ne sortir aucuns groupes sauf <<Tous>>
  plus := 'CC_CODE=""';
  // 1er critère = les dossiers sans confidentialité...
  critere := '';
  // dans la combo DOS_GROUPECONF, on ne met
  // que les groupes auxquels appartient le user
  Q := OpenSQL('SELECT UCO_GROUPECONF FROM USERCONF WHERE UCO_USER="'+V_PGI.User+'"',True);
  While Not Q.Eof do
    begin
    plus := plus + ' OR CC_CODE="'+ Q.FindField('UCO_GROUPECONF').AsString +'"';
    critere := critere + Q.FindField('UCO_GROUPECONF').AsString+ ';' ;  //mcd 19/09/03 ordre des champs changé (mise XXX; au liue de ;XXX plantage guillement droit manquant
    // Next OK, on doit lire tous les groupes auxquels appartient le user
    Q.Next ;
    end;
  Ferme(Q);
  //On retourne le critère du ThValMultiComboBox et sa variable "Plus" .
  Result := critere+';ComboPlus;'+Plus;
end;  

end.

