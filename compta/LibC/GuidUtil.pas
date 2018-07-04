unit GuidUtil;

interface

uses Classes, ExtCtrls, Forms, Controls, SysUtils, StdCtrls, Windows, Graphics,
     HCompte, HCtrls, HEnt1,Formule,hmsgbox, Buttons,Ent1
     {$IFDEF MODENT1}
     , CPTypeCons
     {$ENDIF MODENT1}
{$IFDEF EAGLCLIENT}
     ,uTob
{$ELSE}
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
{$ENDIF}
     ;

Function G_Croix(St : String) : Boolean ;
procedure G_Renum(F : THGrid) ;
Function  G_LigneVide (F : THGrid ; Ligne : Integer ) : boolean ;
Function  CtrlFormule(F:THGrid;Lig:integer; Anal : Boolean) : integer;
Function  TestRefLigne(St:string; Lig,NbLigne:Integer): boolean;
Function  TestChamps(F:THGrid; St:string ; Anal : Boolean): boolean;
Function  MotReserve(St:string):boolean;
Function  OkVentil(CpteGene,NumAxe : String):boolean;
Function  IsClose(Cpte,Table:string):boolean;
Function  IsInteger(St: string): boolean;
Function  IsMontant ( St : string ; Gene : boolean ) : boolean ;
function  CCalculeCodeGuide( vTypeGuide : string ) : string ;

Type TVentGuide = Class
       Ventil : Array [1..MaxAxe] of TStringList ;
       Constructor Create ;
       Destructor Destroy ; override ;
     END ;

implementation


Constructor TVentGuide.Create ;
Var i : integer ;
BEGIN
Inherited Create ;
For i:=1 to MaxAxe do Ventil[i]:=TStringList.Create ;
END ;

Destructor TVentGuide.Destroy ;
Var i : integer ;
BEGIN
For i:=1 to MaxAxe do Ventil[i].Clear ;
For i:=1 to MaxAxe do Ventil[i].Free ;
Inherited Destroy ;
END ;

Function G_LigneVide (F : THGrid ; Ligne : Integer ) : boolean ;
Var i : integer ;
begin
G_LigneVide:=FALSE ;
For i:=2 to F.ColCount-1 do
  if (i mod 2)=0 then
   if F.Cells[i,Ligne]<>'' then Exit ;
G_LigneVide:=TRUE ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 25/01/2006
Modifié le ... :   /  /    
Description .. : - LG - 25/01/2005 - FB 17363 - pb qd on saisisait que ':'
Suite ........ : suppresion d'un apram inutilie
Mots clefs ... : 
*****************************************************************}
Function CtrlFormule(F:THGrid;Lig:integer;Anal : Boolean) : integer;
var i    : integer;
{$IFDEF VER150}
    St : Widestring;
{$ELSE}
    St : string;
{$ENDIF}

BEGIN
{*** TestFormule : teste l'égalité de crochets, accolades, parenthèses et guillemets******
**** TestRefLigne : teste la validité de la référence à une ligne du guide de saisie******
**** TestChamp : teste la validité des noms de champs utilisés (uniqnt champs de mvt)*****
**** Reste à faire : autres champs, mots réservés, opérations}
CtrlFormule:=-1;
for i:=1 to F.ColCount-1 do
 if (i mod 2)=0 then
   BEGIN
   St:=F.Cells[i,Lig] ;
   if St<>'' then
      BEGIN
      if Not TestFormule(St) then
         BEGIN
         F.Col:=i ; F.SetFocus ;
         result := 12 ;
         //THMsgBox(F.Owner.FindComponent('Messages')).Execute(12,'',''); Exit;
         END;
      if Not TestRefLigne(St,Lig,F.RowCount-1) then
         BEGIN
         F.Col:=i ; F.SetFocus ;
         result := 13 ;
         //THMsgBox(F.Owner.FindComponent('Messages')).Execute(13,'',''); Exit;
         END;
      if Not TestChamps(F,St,Anal) then
         BEGIN
         F.Col:=i ; F.SetFocus ;
         result := 15 ;
         //THMsgBox(F.Owner.FindComponent('Messages')).Execute(15,'',''); Exit;
         END;
      END;
   END;
END;

function ControlePrefixePourGuide( vStPrefixe : string ; Anal : boolean) : boolean;
begin
 if not Anal then
  result:= (vStPrefixe='E') or (vStPrefixe='J') or (vStPrefixe='G') or (vStPrefixe='T')
   else
    result:= (vStPrefixe='Y');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent GENDREAU
Créé le ...... : 27/03/2002
Modifié le ... :   /  /
Description .. : LG* modif suite a chg du tobm + correction du controle sur
Suite ........ : les champs
Mots clefs ... :
*****************************************************************}
Function TestChamps(F:THGrid; St:string ; Anal : Boolean): boolean;
var NomChamp,Prefix,StReste, PChamp : string;
    i,PosO,PosF,PosL,LReste,PosD    : integer;
BEGIN
TestChamps:=false; St:=UpperCase(St) ;
if Not Anal then BEGIN Prefix:='E'; END
            else BEGIN Prefix:='Y'; END;
StReste:=St; NomChamp:='';
Repeat
   PosO:=Pos('[',StReste); LReste:=Length(StReste);
   if PosO>0 then
      BEGIN
      // récupération du champ
      PosL:=Pos(':',StReste); PosF:=Pos(']',StReste);
      if (PosL>0) and (PosL<PosF) then NomChamp:=Copy(StReste,PosO+1,PosL-(PosO+1))
                                  else NomChamp:=Copy(StReste,PosO+1,PosF-(PosO+1));
      if PosF<LReste then StReste:=Copy(StReste,PosF+1,LReste) else StReste:='';
      for i:=1 to 2 do BEGIN PosD:=Pos('"',NomChamp);if PosD>0 then NomChamp:=Copy(NomChamp,PosD+1,LReste);END;
      if Not MotReserve(NomChamp) then
         BEGIN      // test de la validité du champ
          if (Pos('_',NomChamp)>0) then PChamp:=Copy(NomChamp,1,Pos('_',NomChamp)-1)
                                   else PChamp:='';
          if (ChampToNum(NomChamp)<0) or not ( ControlePrefixePourGuide(copy(NomChamp,1,1),Anal)) then exit;
         END;
      END;
Until (StReste='') or (PosO<=0);
TestChamps:=true;
END;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 29/10/2002
Modifié le ... : 29/10/2002
Description .. : - 29/10/2002 - ajout des mots reserves 
Suite ........ :  - TVA : compte de tva
Mots clefs ... :  
*****************************************************************}
Function MotReserve(St:string):boolean;
BEGIN
MotReserve:=True ;
St:=uppercase(Trim(St)) ;
if Copy(St,1,4)='AUTO' then Exit ;
if ((St='SOLDE') or (St='TVA') or (St='TPF') or (St='INTITULE')) then Exit ;
if ((St='TVANOR') or (St='TVARED')) then Exit ; if (St='TVA') then Exit ;
MotReserve:=False ;
END;

Function TestRefLigne(St:string; Lig,NbLigne:Integer): boolean;
var deb,long : integer;
    St1            : string;
    LigRef : double ;
BEGIN
TestRefLigne:=false;
LigRef:=0;
{Compatibilité :1 et :L1}
While Pos(':L',St)>0 do Delete(St,Pos(':L',St)+1,1) ;
if Pos(':',St)>0 then
   repeat
     BEGIN
     deb:=Pos(':',St)+1; St1:=Copy(St,deb,5);long:=Pos(']',St1)-1;
     if long>0 then St1:=Copy(St,deb,long);
      if ((IsInteger(St1)) or (St1='-1')) then
       LigRef:=Valeur(St1);
     if ((LigRef<=0) and (LigRef<>-1)) or (LigRef>NbLigne) or (LigRef=Lig) then Exit ;
     St:=Copy(St,deb,Length(St));
     END
   until Pos(':',St)<=0 ;
TestRefLigne:=true;
END;

Function IsInteger(St: string): boolean;
var i:integer;
BEGIN
IsInteger:=false;
for i:=1 to Length(St) do if Not(St[i] in ['0'..'9']) then exit;
IsInteger:=true;
END;

Function IsMontant ( St : string ; Gene : boolean ) : boolean ;
var i    : integer;
    Good : boolean;
    SetNum : Set of Char ;
BEGIN
SetNum:=['0'..'9','.',',','+','-',' ',#160] ;
if Gene then SetNum:=SetNum+[VH^.Cpta[fbGene].Cb,'A'..'Z'] ; //n°1741
Good:=TRUE ;
For i:=1 to Length(St) do
   BEGIN
   Good:=(St[i] in SetNum) ;
   if Not Good then break ;
   END;
IsMontant:=Good;
END;


Function OkVentil(CpteGene,NumAxe : String):boolean;
var Q   : TQuery;
    SQL : String;
BEGIN
OkVentil:=False; if CpteGene='' then Exit ;
SQL:='Select G_VENTILABLE'+NumAxe+' From GENERAUX Where G_GENERAL="'+CpteGene+'"';
Q:=OpenSQl(SQL,true);
if Not Q.EOF then
   BEGIN
   if Q.FindField('G_VENTILABLE'+NumAxe).AsString<>'X' then BEGIN Ferme(Q) ; exit ; END;
   END else
   BEGIN
   Ferme(Q) ; Exit ;
   END ;
Ferme(Q) ;
OkVentil:=True;
END;

Function IsClose(Cpte,Table:string):boolean;
var Q                : TQuery;
    SQL,FClose,FCpte : String;
BEGIN
IsClose:=true;
if Table='GENERAUX' then
   BEGIN
   FClose:='G_FERME'; FCpte:='G_GENERAL'
   END else if Table='TIERS' then
   BEGIN
   FClose:='T_FERME'; FCpte:='T_AUXILIAIRE'
   END else if Table='SECTION' then
   BEGIN
   FClose:='S_FERME'; FCpte:='S_SECTION'
   END;
SQL:='Select '+FClose+' from '+Table+' where '+FCpte+'="'+Cpte+'"';
Q:=OpenSQL(SQL,true);
if Not Q.EOF then if Q.FindField(FClose).AsString='X' then BEGIN Ferme(Q) ; Exit; END ;
Ferme(Q);
IsClose:=false;
END;

procedure G_Renum(F : THGrid) ;
Var i : integer ;
begin
For i:=1 to F.RowCount-1 do F.Cells[0,i]:=IntToStr(i) ;
Screen.Cursor:=Syncrdefault ;
end ;

Function G_Croix(St : String) : Boolean ;
begin
St:=Copy(St,1,1) ;
if St<>'' then
   BEGIN
   St:=UpperCase(St) ;
   Result:=(St='X') ;
   END else Result:=FALSE ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 09/10/2006
Modifié le ... : 16/10/2006
Description .. : Incrémentation d'une valeur sur 36 positions
Mots clefs ... :
*****************************************************************}
function _NextCar(ValIn : Char) : Char;
begin

  if Ord(ValIn)=57 then
      Result:=Chr(65)   // 9 --> A
  else
    if Ord(ValIn)<>90 then
        Result:=Chr(Ord(ValIn)+1)
    else
        Result:=Chr(48);  // Z --> 0

end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 09/10/2006
Modifié le ... : 16/10/2006
Description .. : Incrémentation d'une chaîne sur 36 positions
Mots clefs ... :
*****************************************************************}
Function _NextCode(ValIn : String; taille : integer) : String;
var
  ValOut : String ;
  i, j  : integer ;
  mov : boolean;
begin

 ValOut:='';
 ValIn:=Uppercase(ValIn+'Z');

 for i:=1 to taille do
 begin
      mov:=true;
      for j:=i+1 to taille+1 do      // si tous les suivants à Z...
          If ValIn[j]<>'Z' then mov:=false;

      If mov then
        ValOut:=ValOut+_NextCar(ValIn[i]) // ...on incrémente
      else
        ValOut:=ValOut+ValIn[i];
 end;

 Result:=ValOut;

end;



function CCalculeCodeGuide( vTypeGuide : string ) : string ;
var
 lQ : TQuery ;
begin

 result := '001' ;

 lQ := OpenSQL('Select MAX(GU_GUIDE) from GUIDE Where GU_TYPE="' + vTypeGuide + '" Order by 1',True) ;

 if not lQ.EOF then
  begin
   result := lQ.Fields[0].AsString ;
   if result <> '' then
    result := _NextCode(result, 3)
     else
      result := '' ;

   while length(result) < 3 do
    result := '0' + result ;

  end ;

 ferme(lQ) ;

end ;



end.
