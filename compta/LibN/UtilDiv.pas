{***********UNITE*************************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /    
Description .. : Regroupement de différentes fonctions génériques.
Mots clefs ... : FONCTIONS;
*****************************************************************}
unit UtilDiv;

interface
uses Classes,Hctrls,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
   {$IFDEF MODENT1}
   CPTypeCons,
   {$ENDIF MODENT1}
     HEnt1,Ent1,Utob,SysUtils ;

function RecupMessageAlmacom(Erreur: integer): string ;
function TrouveArgument(Argument: String;TypeArg : string): string;  overload ;
function TrouveArgument(Argument,TypeArg : string ; Default : Variant ; Nullable : Boolean = FALSE  ): Variant ; overload ;
function GetChoixCode(CCType,CCCode : string ) : string ;
procedure SetChoixCode(CCType,CCCode,CCLibelle : string ) ;
procedure VideStringList ( L : TStringList ) ;
procedure MemSet(var Chaine : string;Lettre : char;Nombre : integer) ;
function QuelEstLefb(Budget: boolean;typerub,axe: string) : TFichierBase ;
Function QuelSontLesfb(TypRub,Axe : String ; Var fb1,fb2 : TFichierBase ; OnBudget : Boolean) : Byte ;
Function AnalyseCompteRubrique(Rub : String) : Tob ;
procedure AnalyseRubrique(Rub : String ; LaTob: Tob;Var Tab: array of string );
procedure QuelEstLaTable(fb: TFichierBase;var Pref,Champ,Table: string) ;
function EnleveEntreParentheses(Chaine: String): string ;
function CommentExoDateDT(dd: TDateTime): integer ;

implementation
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 16/05/2000
Modifié le ... :   /  /    
Description .. : Identifie les fichiers de base d'une rubrique
Mots clefs ... : RUBRIQUE; FICHIERS DE BASE;
*****************************************************************}
Function QuelSontLesfb(TypRub,Axe : String ; Var fb1,fb2 : TFichierBase ; OnBudget : Boolean) : Byte ;
Var fbAnal : TFichierBase ;
BEGIN
fb1:=fbNone ; fb2:=fbNone ;
fbAnal:=fbNone ;
if (TypRub='ANA') Or (TypRub='BSE') Or (TypRub='G/A') Or (TypRub='A/G') then
  begin
  if Axe[1]='A' then
    begin
    case Axe[2] of
      '1' : If OnBudget Then fbAnal:=fbBudSec1 else fbAnal:=fbAxe1 ;
      '2' : If OnBudget Then fbAnal:=fbBudSec2 else fbAnal:=fbAxe2 ;
      '3' : If OnBudget Then fbAnal:=fbBudSec3 else fbAnal:=fbAxe3 ;
      '4' : If OnBudget Then fbAnal:=fbBudSec4 else fbAnal:=fbAxe4 ;
      '5' : If OnBudget Then fbAnal:=fbBudSec5 else fbAnal:=fbAxe5 ;
      end ;
    end ;
  end ;
if OnBudget Then
   begin
   If TypRub='GEN' Then fb1:=fbBudGen else
      If TypRub='BGE' Then fb1:=fbBudGen else
         If TypRub='ANA' Then fb1:=fbAnal else
            If TypRub='BSE' Then fb1:=fbAnal else
               If TypRub='G/A' Then begin If OnBudget Then fb1:=fbBudGen else fb1:=fbGene ; fb2:=fbAnal ; end else
                  If TypRub='A/G' Then begin If OnBudget Then fb2:=fbBudGen else fb2:=fbGene ; fb1:=fbAnal ; end ;
   end else
   begin
   If TypRub='GEN' Then fb1:=fbGene else
      If TypRub='TIE' Then fb1:=fbAux else
         If TypRub='ANA' Then begin fb1:=fbAnal ; end else
            If TypRub='G/A' Then begin fb1:=fbGene ; fb2:=fbAnal ; end else
               If TypRub='A/G' Then begin fb2:=fbGene ; fb1:=fbAnal ; end else
                  If TypRub='G/T' Then begin If OnBudget Then fb1:=fbBudGen else fb1:=fbGene ; fb2:=fbAux ; end else
                     If TypRub='T/G' Then begin If OnBudget Then fb2:=fbBudGen else fb2:=fbGene ; fb1:=fbAux ; end ;
   end ;
Result:=1 ;
end ;
{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /
Description .. : Renvoi le fichier de base en fonction du type de rubrique et de l'axe
Mots clefs ... : RUBRIQUE;FICHIER BASE;
*****************************************************************}
function QuelEstLefb(Budget: boolean;typerub,axe: string) : TFichierBase ;
var Fb2: TFichierBase ;
begin
QuelSontLesfb(TypeRub,Axe,Result,fb2,Budget) ;
//if Budget then begin case typerub[1]of'A':case Axe[2]of'1':Result:=fbBudsec1 ;'2':Result:=fbBudsec2;'3':Result:=fbBudsec3;'4':Result:=fbBudsec4;'5':Result:=fbBudsec5;end;'G':Result:=fbBudgen ;end ;end
//          else begin case typerub[1]of'A':case Axe[2]of'1':Result:=fbAxe1 ;'2':Result:=fbAxe2;'3':Result:=fbAxe3;'4':Result:=fbAxe4;'5':Result:=fbAxe5;end;'G':Result:=fbGene;'T':Result:=fbAux ;end ;end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... : 15/05/2000
Description .. : Renvoi le paramètre "TypeArgument" qui se trouve dans la chaine Argument
Mots clefs ... : CHAINE;AGLLANCEFICHE;ARGUMENT
*****************************************************************}
function TrouveArgument(Argument: String;TypeArg : string): string; overload ; 
var StArgument : string;
    i,lg : integer;
begin
lg:=Length(TypeArg)-1 ; StArgument := Argument ; i:=Pos(TypeArg,StArgument) ;
if i>0 then
  begin
  system.Delete(StArgument,1,i+lg) ;
  Result:=ReadTokenSt(StArgument);
  end
  else
  Result:='';
end;

function sLeft(s:string;Count:integer):string;
var   L:integer;
begin
L:=length(s);
if (L<=Count) or (s='') then
   begin
   result:=s;
   exit;
   end;
result:=Copy(s, 1, Count);
end;
//////////////////////////////////////////////////////////////////////////////
function  sRight(s:string;Count:integer):string;
var   L:integer;
begin
L:=length(s);
if (Count>L) or (s='') then
   begin
   result:=s;
   exit;
   end;
result:=Copy(s, L-Count+1, Count);
end;
//////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : X.Maluenda
Créé le ...... : 05/04/2004
Modifié le ... :   /  /
Description .. : Renvoie le paramètre "typeArg" qui se trouve dans la
Suite ........ : chaine "Argument", on peut y specifier la valeur par défaut
Suite ........ : et si elle peut être "nulle" ou pas.....
Mots clefs ... :
*****************************************************************}
function TrouveArgument(Argument,TypeArg : string ; Default : Variant ; Nullable : Boolean = FALSE  ): Variant ; overload ;
var ii,lg : integer ;
    StArgument : String ;
begin
//On done la valeur par défaut
Result:=Default ;
//on "adapte" les strings
if sleft(Argument,1)<>';' then Argument:=';'+Argument;
StArgument:=Uppercase(Argument);

TypeArg:=Uppercase(typeArg) ;
if sright(TypeArg,1)<>'=' then TypeArg:=TypeArg+'=' ;
if sleft(TypeArg,1)<>';' then TypeArg:=';'+TypeArg ;

//On essaie de trouver l'argument....
lg:=Length(TypeArg)-1 ; ii:=Pos(TypeArg,StArgument) ;
if ii>0 then
  begin
  system.Delete(Argument,1,ii+lg) ;
  Result:=ReadTokenSt(Argument);
  if (not Nullable) and (trim(Result)='') then Result:=Default ;
  end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... : 15/05/2000
Description .. : Récupère le libellé d'un enregistrement de la table choixcode
Mots clefs ... : CHOIXCODE;LIBELLE;LECTURE
*****************************************************************}
function GetChoixCode(CCType,CCCode : string ) : string ;
var Q : TQuery ;
begin
Q:=OpenSql('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+CCType+'" AND CC_CODE="'+CCCode+'"',True) ;
if Not Q.Eof then Result:=Q.findfield('CC_LIBELLE').AsString else Result:='0' ;
ferme(Q) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /    
Description .. : MAJ ou INSERT un enregistrement dans CHOIXCODE
Mots clefs ... : CHOIXCODE;LIBELLE;ENREGISTRE
*****************************************************************}
procedure SetChoixCode(CCType,CCCode,CCLibelle : string ) ;
var Q : TQuery ;
begin
Q:=OpenSql('SELECT * FROM CHOIXCOD WHERE CC_TYPE="'+CCType+'" AND CC_CODE="'+CCCode+'"',False) ;
if Not Q.Eof then
  begin Q.Edit ; Q.FindField('CC_LIBELLE').AsString:=CCLibelle ;  end
  else
  begin
  Q.Insert ; InitNew(Q) ;
  Q.FindField('CC_TYPE').AsString:=CCType ; //GRB
  Q.FindField('CC_CODE').AsString:=CCCode ; //AUT
  Q.FindField('CC_LIBELLE').AsString:=CCLibelle ;
  end ;
Q.Post ;
ferme(Q) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /    
Description .. : Fait le "free" des objects d'une TSTRINGLIST
Mots clefs ... : STRINGLIST;
*****************************************************************}
procedure VideStringList ( L : TStringList ) ;
Var i : integer ;
BEGIN
if L=Nil then Exit ; if L.Count<=0 then Exit ;
for i:=0 to L.Count-1 do If L.Objects[i]<>NIL Then L.Objects[i].Free ;
L.Clear ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 15/05/2000
Modifié le ... :   /  /    
Description .. : Place "Nombre" caractère(s) "Lettre" dans "Chaine"
Mots clefs ... : CHAINE;
*****************************************************************}
procedure MemSet(var Chaine : string;Lettre : char;Nombre : integer) ;
var i : integer ;
begin Chaine:='' ; for i:=1 to Nombre do Chaine:=Chaine+Lettre ; end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 16/05/2000
Modifié le ... : 16/05/2000
Description .. : Analyse le contenu des rubriques et renvoi une Tob contenant tous les comptes
Mots clefs ... : RUBRIQUES;TOB;
*****************************************************************}
procedure AnalyseRubrique(Rub : String ; LaTob: Tob;Var Tab: array of string );
var TobTmp: Tob;
    St,UneRub: string ;
    TabRub,TabRes: array[0..3] of string ;
    i,Nb: integer ;
begin
Tab[0]:='' ;Tab[1]:='' ;Tab[2]:='' ;Tab[3]:='' ;
if (Rub='')Or(LaTob=nil)then Exit ;
TobTmp:=LaTob.FindFirst(['RB_RUBRIQUE'],[Rub],True) ;
if TobTmp=nil then exit ;
TabRub[0]:=TobTmp.GetValue('RB_COMPTE1') ;
TabRub[1]:=TobTmp.GetValue('RB_EXCLUSION1') ;
TabRub[2]:=TobTmp.GetValue('RB_COMPTE2') ;
TabRub[3]:=TobTmp.GetValue('RB_EXCLUSION2') ;
//if (TobTmp.GetValue('RB_RUBDERUB')<>'X') then begin for i:=0 to 3 do Tab[i]:=TabRub[i] ; Exit ; end ;
if (TobTmp.GetValue('RB_CLASSERUB')<>'RDR') then begin for i:=0 to 3 do Tab[i]:=TabRub[i] ; Exit ; end ;
//Rubrique de Rubriques
for i:=0 to 1 do
  repeat
    UneRub:=ReadTokenStV(TabRub[i]) ;
    AnalyseRubrique(UneRub,LaTob,TabRes) ;
    if i=0 then Nb:=0 else Nb:=1 ;
    Tab[0]:=Tab[0]+TabRes[0+Nb] ; //
    Tab[1]:=Tab[1]+TabRes[1-Nb] ; // On inverse les comptes qui vont
    Tab[2]:=Tab[2]+TabRes[2+Nb] ; // dans les exclusions lorsque
    Tab[3]:=Tab[3]+TabRes[3-Nb] ; // la rubrique est exclus
  until St='' ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 16/05/2000
Modifié le ... : 16/05/2000
Description .. : Analyse le contenu d'une rubrique et renvoi une Tob contenant tous les comptes
Mots clefs ... : RUBRIQUES;TOB;
*****************************************************************}
Function AnalyseCompteRubrique(Rub : String) : Tob ;
var TobRub,TobTmp,TobRes: Tob ;
    OnTBL,OnBud: Boolean ;
    fb1,fb2: TFichierBase ;
    TabRes,Where: array[0..3] of string ;
    Sql,LeAnd,Pref1,Champ1,Table1,Pref2,Champ2,Table2: string ;
    i: integer ;
    Q: TQuery ;
begin
Result:=nil ;
if Rub='' then Exit ;
TobRub:=TOB.Create('§RUBRIQUE',nil,-1) ;
TobRub.LoadDetailDB('RUBRIQUE','','',nil,True,True) ;
TobTmp:=TobRub.FindFirst(['RB_RUBRIQUE'],[Rub],True) ;
if TobTmp=nil then exit ;
//OnTBL:=(TobTmp.GetValue('RB_TABLELIBRE')='BUD') ;
OnTBL:=(TobTmp.GetValue('RB_CLASSERUB')<>'TLI') ;
OnBud:=(TobTmp.GetValue('RB_NATRUB')='BUD') ;
QuelSontLesfb(TobTmp.GetValue('RB_TYPERUB'),TobTmp.GetValue('RB_AXE'),fb1,fb2,OnBud) ;
AnalyseRubrique(Rub,TobRub,TabRes) ;
Where[0]:=AnalyseCompte(TabRes[0],fb1,False,OnTBL) ;
TabRes[1]:=EnleveEntreParentheses(TabRes[1]) ;
Where[1]:=AnalyseCompte(TabRes[1],fb1,True,OnTBL) ;
Where[2]:=AnalyseCompte(TabRes[2],fb2,False,OnTBL) ;
TabRes[3]:=EnleveEntreParentheses(TabRes[3]) ;
Where[3]:=AnalyseCompte(TabRes[3],fb2,True,OnTBL) ;
QuelEstLaTable(fb1,Pref1,Champ1,Table1);
QuelEstLaTable(fb2,Pref2,Champ2,Table2);
//Fabrication du select !!
Sql:='SELECT '+Champ1 ;
if Champ2<>'' then Sql:=Sql+','+Champ2 ;
Sql:=Sql+' FROM '+Table1 ;
if Table2<>'' then Sql:=Sql+','+Table2 ;
LeAnd:=' WHERE ' ;
for i:=0 to 3 do if Where[i]<>'' then begin Sql:=Sql+LeAnd+Where[i] ; LeAnd:=' AND ' end ;
TobRes:=TOB.Create('$COMPTE',nil,-1) ;
Q:=OpenSql(Sql,True) ;
TobRes.LoadDetailDB('$COMPTE','','',Q,True,True) ;
ferme(Q) ;
Result:=TobRes ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 16/05/2000
Modifié le ... : 16/05/2000
Description .. : Renvoi la table correspondant à un fichierdebase
Mots clefs ... : FICHIERSBASE;TABLE;
*****************************************************************}
procedure QuelEstLaTable(fb: TFichierBase;var Pref,Champ,Table: string) ;
begin
Pref:='' ; Champ:='' ; Table:='' ;
case fb of
    fbGene :               begin Pref:='G_' ;  Champ:='G_GENERAL' ;    Table:='GENERAUX' ; end ;
    fbAux  :               begin Pref:='T_' ;  Champ:='T_AUXILIAIRE' ; Table:='TIERS' ;   end ;
    fbImmo :               begin Pref:='I_' ;  Champ:='I_IMMO' ;       Table:='IMMO'      end ;
    fbBudgen :             begin Pref:='BG_' ; Champ:='BG_BUDGENE' ;   Table:='BUDGENE'   end ;
    fbBudSec1..fbBudSec5 : begin Pref:='BS_' ; Champ:='BS_BUDSECT' ;   Table:='BUDSECT'   end ;
    fbBudJal :             begin Pref:='BJ_' ; Champ:='BJ_BUDJAL' ;    Table:='BUDJAL'    end ;
    fbJal :                begin Pref:='J_' ;  Champ:='J_JOURNAL' ;    Table:='JOURNAL'   end ;
    fbAxe1..fbAxe5 :       begin Pref:='S_' ;  Champ:='S_SECTION' ;    Table:='SECTION'   end ;
    end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 16/05/2000
Modifié le ... : 16/05/2000
Description .. : Enleve le contenu des parentheses dans une chaine
Mots clefs ... : CHAINE;
*****************************************************************}
function EnleveEntreParentheses(Chaine: String) :string ;
var i,Okok: integer ;
begin
Okok:=0 ;   Result:='' ;
for i:=1 to length(Chaine) do
  begin
  if Chaine[i]='(' then inc(Okok) ;
  if Okok=0 then Result:=Result+Chaine[i];
  if Chaine[i]=')' then dec(Okok) ;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 16/05/2000
Modifié le ... : 16/05/2000
Description .. : Renvoi le type d'exercice (0:pas dans la base, 1:cloturé, 2:periode cloturée, 3:ouvert)
Mots clefs ... : CHAINE;
*****************************************************************}
function CommentExoDateDT(dd: TDateTime): integer ;
var i: integer ;
begin
result:=0 ;
if (dd>=VH^.EnCours.Deb) and (dd<=VH^.EnCours.Fin) then Result:=3 else
   if (dd>=VH^.Suivant.Deb) and (dd<=VH^.Suivant.Fin) then Result:=3 else
      if (dd>=VH^.Precedent.Deb) and (dd<=VH^.Precedent.Fin) then Result:=1 else
         for i:=1 To 5 do if (dd>=VH^.ExoClo[i].Deb) And (dd<=VH^.ExoClo[i].Fin) then begin Result:=1 ; break ; end ;
if (result=3) and (VH^.DateCloturePer>0) and (dd<=VH^.DateCloturePer) then begin Result:=2 ; Exit ; end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 27/07/2000
Modifié le ... :   /  /
Description .. : Recupération du message d'erreur renvoyé par la DLL 
Suite ........ : Almacom. 
Mots clefs ... : ALMACOM
*****************************************************************}
function RecupMessageAlmacom(Erreur: integer): string ;
begin
Result:='Erreur n°'+IntToStr(Erreur) ;
//case Erreur of
//  2701:  Result:=Result+' Erreur réception données';
//  2702:  Result:=Result+' Fin de données érronées';
//  etc ...
//  else
//    Result:=Result+' inconnue';
//  end ;
TraduireMemoire(Result) ;
end ;
end.
