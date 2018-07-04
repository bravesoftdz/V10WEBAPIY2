unit SQLParts;

interface

uses ComCtrls ;

(*
Ce sont d'outils diverses qui permettent gérer les differentes parties d'un sentence SQL
*)
Function RecupSQLParts ( SQL : String ; Tree : TTreeNodes ) : Boolean ;
Function RecupSQLfromparts ( Tree : TTreeNodes ) : String ;
Function CherchePart(Tree : TTreeNodes ; Quoi : String ) : TTreeNode ;
Procedure Recupnoeud ( Tree : TTreeNodes ; Source : TTreeNode ; OkClear : Boolean ) ;
Procedure RecopieTreeNodes ( Source,Dest : TTreeNodes ) ;
procedure Recupenfants (Tree : TTreeNodes ; var Source,Pere : TTreenode ) ;
Function InsertPart(Tree : TTreenodes ; Parent : TTreeNode ; St : String ) :TTreeNode ;
Function SupprimeCRLF( St: String ) : String ;
procedure fraccioner ( OkFirst : Boolean ; Txt : String ; var SQL : String ; Tree : TTreeNodes ; Pere : TTreeNode  ) ;
Function ZIPText ( St : String ) : String ;

implementation

uses Sysutils, HCtrls, UTOB, MC_Lib ;
//////////////////////////////////////////////////////////////////////////////
Procedure VideTree ( Tree : TTreenodes ) ;
var Node : TTreenode ;
Begin
Node:=Tree.GetFirstNode ;
while assigned(Node) do
  Begin
  if assigned(Node.Data) then
     Begin
     TObject(Node.Data).Free ;
     Node.Data:=nil ;
     End ;
  Node:=Node.GetNext ;
  End ;
End ;
//////////////////////////////////////////////////////////////////////////////
Function ZIPText ( St : String ) : String ;
Var Last : Char ;
    i    :  integer ;
Begin
St:=uppercase(Trim(St)) ;
Last:=#0 ; i:=Length(St) ;
while (i>0) do
  Begin
  if (St[i]=' ') and (Last=#32) then system.delete(St,i,1) ;
  last:=St[i] ;
  dec(i) ;
  End ;
result:=St ;
End ;
//////////////////////////////////////////////////////////////////////////////
Const Select       = 'ORDER BY;PLAN;UNION;HAVING;GROUP BY;WHERE;FROM;SELECT;' ;  //DISTINC;ALL; // AS
      Insert       = 'VALUES;(;INSERT INTO' ;
      Delete       = 'WHERE;DELETE FROM '  ;
      Update       = 'WHERE;SET;UPDATE' ;

      Idx_Select = 1 ;
      Idx_Insert = 2 ;
      Idx_Delete = 3 ;
      Idx_update = 4 ;

      ConditionJ   = 'ON;' ;
      Virgule      = ',;' ;

Var   Instructions : Array[1..4] of string = ('SELECT', 'INSERT INTO','DELETE FROM','UPDATE') ;
      Join         : array[1..9] of string = ('INNER JOIN','LEFT OUTER JOIN','RIGHT OUTER JOIN','FULL OUTER JOIN','LEFT JOIN','RIGHT JOIN','FULL JOIN','OUTER JOIN','JOIN') ;

//////////////////////////////////////////////////////////////////////////////
Function Find_Instruc ( SQL : String ; Tab : Array of String ) : Integer ;
Var i,J  : Integer ;
    St   : String ;
Begin
result:=-1 ;
For i:=low(Tab) to High(Tab) do
    Begin
    St:=Tab[i] ;
    j:=pos(St,SQL) ;
    if (j>0) and ((j<Result) or (Result=-1)) then Result:=j ;
    End ;
End ;
//////////////////////////////////////////////////////////////////////////////
Function First_Instruc ( SQL : String ; Tab : Array of String ; Var St : String ) : Integer ;
Var i  : Integer ;
Begin
result:=-1 ;
For i:=low(Tab) to High(Tab) do
    Begin
    St:=Tab[i] ;
    if St=Copy(SQL,1,length(St)) then Begin Result:=i+1 ; break ; End ;
    End ;
End ;
//////////////////////////////////////////////////////////////////////////////
Function Find_InstrucToken ( St,Tokens : String ) : Integer ;
Var txt : String ;
    i   : Integer ;
Begin
St:=Uppercase(Trim(St)) ;
Result:=-1 ; i:=0 ;
if St='' then exit ;
while (Tokens<>'') and (Result=-1) do
  Begin
  Txt:=uppercase(Trim(readTokenSt(Tokens))) ; inc(i) ;
  if St=Txt then Result:=i ;
  End ;
End ;
//////////////////////////////////////////////////////////////////////////////
Procedure FractionSansOrdre(Tab : Array of string ; Var SQL : String ; Tree : TTreeNodes ; Pere : TTreeNode  ) ;
var mot,st,Stt      : String ;
    index,j         : Integer ;
    Node            : TTreeNode ;
Begin
j:=pos(' ',SQL) ; if j<=0 then j:=length(SQL)+1 ;
Stt:=Copy(SQL,1,j-1) ; system.delete(SQL,1,j) ; SQL:=Trim(SQL) ;
Tree.AddChild(pere,Stt) ; index:=9999 ;
while (SQL<>'') and (Index>0) do
  Begin
  index:=first_instruc(SQL,Tab,mot) ;
  if index>0 then
    Begin
    Stt:=trim(Copy(SQL,1,length(Mot))) ; System.delete(SQL,1,length(Mot)) ; SQL:=Trim(SQL) ;
    St:=SQL ;
    if (Mot='OR') and (Pere.parent<>nil) then Pere:=pere.Parent ;
    Node:=Tree.AddChild(Pere,Stt) ;
    FractionSansOrdre(Tab,St,Tree,Node) ;
       Begin
       j:=pos(' ',SQL) ;
       if J>0 then
         Begin
         System.delete(SQL,1,j) ; SQL:=trim(SQL) ;
         j:=Find_Instruc(SQL,tab)-1 ; if j<=0 then j:=length(SQL) ;
         St:=trim(Copy(SQL,1,j)) ; System.delete(SQL,1,j) ; SQL:=Trim(SQL) ;
         Fraccioner(TRUE,ConditionJ,St,Tree,Node) ;
         End ;
       End ;
    End ;
  End ;
End ;
//////////////////////////////////////////////////////////////////////////////
procedure fraccioner ( OkFirst : Boolean ; Txt : String ; var SQL : String ; Tree : TTreeNodes ; Pere : TTreeNode  ) ;
Var Mot,St,Stt,Txtt,OldTxt  : String ;
    j,last                  : Longint ;
    Node                    : TTreeNode ;
    Aucun,OkFirst2          : Boolean ;
    vv                      : TCS ;
Begin
OldTxt:=Txt ;
while (Txt<>'') and (Txt<>';') do
  Begin
  Mot:=ReadTokenSt(Txt) ; Last:=99999 ;
  Aucun:=TRUE ;
  while ((OLdTxt=',;') or ((aucun) and (OldTxt<>',;'))) and (Last>0) do
    Begin
    Last:=0 ;
    repeat
      j:=pos(Mot,Copy(SQL,Last+1,length(SQL))) ;
      if j>0 then Last:=Last+j ;
    until ((j<=0) and (Not OkFirst)) or (OkFirst) ;
    if Last>0 then
       Begin
       Aucun:=FALSE ;
       if Mot<>',' then
          Begin
          St:=Copy(SQL,Last,Length(SQL)) ;
          system.Delete(SQL,Last,Length(SQL)) ; SQL:=Trim(SQL) ;
          Stt:=Mot ;
          System.Delete(St,1,Length(Stt)) ; St:=trim(St) ;
          End else
          Begin
          Stt:=trim(Copy(SQL,1,last-1)) ; System.delete(SQL,1,last) ; SQL:=Trim(SQL) ;
          St:=SQL ;
          End ;
       if (Mot=',') then
          Begin
          vv:=TCS.Create ;
          vv.nom:='mot' ;
          vv.Valeur:=',' ;
          Node:=Tree.AddChildObject(Pere,Stt,vv)
          End else
          if (Mot='ON') then Node:=Tree.AddChild(Pere,Stt) else Node:=Tree.AddChildFirst(Pere,Stt) ;
       Txtt:='' ; okFirst2:=TRUE ;
       if (Stt='FROM')   then Fractionsansordre(Join,St,Tree,Node) ;
       if Stt='SELECT'   then Txtt:=Virgule else
       if Stt='('        then Begin System.Delete(St,length(St),1) ; Txtt:=Virgule ; End else
       if Stt='VALUES'   then Begin System.Delete(St,1,1) ; System.Delete(St,length(St),1) ; Txtt:=Virgule ; End else
       if Stt='SET'      then Txtt:=Virgule else
//       if Stt='UNION'    then Begin Txtt:=Select ; OkFirst2:=FALSE ; End else
       if Stt='GROUP BY' then Txtt:=Virgule else
       if Stt='ORDER BY' then Txtt:=Virgule else
       if (pos('JOIN',St)>0) then Fractionsansordre(Join,St,Tree,Node) ;
       ST:=Trim(St) ;
       if Txtt<>'' then Fraccioner(Okfirst2,Txtt,St,Tree,Node) else
          if (Mot<>',') and (Stt<>'FROM') then Tree.AddChildFirst(Node,St) ;
       End ;
    End ;
  End ;
if SQL<>'' then
   Begin
   //XMG 07/04/04 début
   if Mot='ON' then
      Tree.AddChild(Pere.item[0],SQL) else
   if Mot=',' then
      Tree.AddChild(Pere,SQL)
   else
      Tree.AddChildFirst(Pere,SQL) ;
   //XMG 07/04/04 fin
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////
Function PartsFromIndex(Index : Integer ; Var OkFirst: Boolean ) : String ;
Begin
Result:='' ;
OkFirst:=TRUE ;
Case index of
  idx_Select : Begin OkFirst:=FALSE ; Result:=Select ; End ;
  idx_Insert : Result:=Insert ;
  idx_Delete : Result:=Delete ;
  idx_Update : Result:=Update ;
  END ;
End ;
//////////////////////////////////////////////////////////////////////////////
Function SupprimeCRLF( St: String ) : String ;
Const CCrlf = #13+#10+';'+#13+';'+#10+';' ;
Var crlf,cc  : String ;
    j        : Longint ;
Begin
CrLf:=CCrLf ;
while (Trim(CrLf)<>'') and (Trim(CrLf)<>';') do
  Begin
  cc:=ReadTokenSt(CrLf) ;
  repeat
    j:=pos(cc,St) ; if j>0 then Begin System.delete(St,j,length(cc)) ; System.Insert(' ',st,j) ; End ;
  until j<=0 ;
  End ;
Result:=St ;
End ;
//////////////////////////////////////////////////////////////////////////////
Function RecupSQLParts ( SQL : String ; Tree : TTreeNodes ) : Boolean ;
Var St,Txt     : String ;
    Index      : Longint ;
    Ok,OkFirst : Boolean ;
Begin
Result:=FALSE ;
Ok:=(Tree<>nil) ;
SQL:=SupprimeCRLF(SQL) ;
SQL:=ZipText(SQL) ;
Ok:=(Ok) and (SQL<>'') ;
Index:=First_Instruc(SQL,Instructions,St) ;
Ok:=(Ok) and (index>0) ;
if not Ok then exit ;
Tree.BeginUpdate ;
Tree.Clear ;
if Ok then
   Begin
   Txt:=PartsFromIndex(Index,OkFirst) ;
   Fraccioner(OkFirst,Txt,SQL,Tree,Nil) ;
   End ;
Tree.EndUpdate ;
Result:=Ok ;
End ;
//////////////////////////////////////////////////////////////////////////////
Function Rajoute ( R,st,Courant : String ; var dernier : String ; vv : TCS ) : String ;
Var Pref,Suf : String ;
Begin
pref:=' ' ; Suf:='' ;
if Assigned(vv) then suf:=VString(vv.Valeur) ;
if courant='(' then
   begin
   if Courant<>Dernier then Begin Pref:='' ; Dernier:=Courant; End else
   if st='VALUES' then begin dernier:=Courant ; pref:=') ' ; Suf:=' (' ; end ;
   end ;
if Courant='VALUES' then
   if St='' then Pref:=')' ;
Result:=R+Pref+St+Suf ;
End ;
//////////////////////////////////////////////////////////////////////////////
Function RecupSQLfromparts ( Tree : TTreeNodes ) : String ;
Var St,Dernier,Courant : String ;
    Node               : TTreeNode ;
    vv                 : TCS ;
Begin
Result:='' ;
if Tree=nil then exit ;
if tree.count<=0 then exit ;
Node:=Tree.GetFirstNode ; Dernier:='' ; Courant:='' ;
while Node<>nil do
  Begin
  St:=Node.Text ; vv:=TCS(Node.Data) ;
  Result:=Rajoute(Result,St,Courant,Dernier,vv) ;
  if Node.Parent=nil then Begin Dernier:=Courant ; Courant:=St ; End ;
  Node:=Node.GetNext ;
  End ;
Result:=trim(rajoute(Result,'',Courant,dernier,nil)) ;
end ;
//////////////////////////////////////////////////////////////////////////////
Function CherchePart(Tree : TTreeNodes ; Quoi : String ) : TTreeNode ;
Var Node : TTreeNode ;
Begin
Result:=nil ;
Quoi:=uppercase(Trim(Quoi)) ;
Node:=Tree.GetFirstNode ;
while (Node<>nil) and (Result=nil) do
  Begin
  if Node.Text=Quoi then Result:=Node ;
  node:=Node.GetNext ;
  End ;
End ;
//////////////////////////////////////////////////////////////////////////////
procedure Recupenfants (Tree : TTreeNodes ; var Source,Pere : TTreenode ) ;
Var fils,Nouveau : TTreenode ;
    vv           : TCS ;
Begin
if (Tree=nil) or (Source=nil) then exit ;
Tree.BeginUpdate ;
Fils:=Source.GetFirstChild ;
while Fils<>nil do
 begin
 Nouveau:=Tree.AddChild(Pere,'') ;
 Nouveau.Assign(Fils) ;
 if assigned(Fils.data) then
    Begin
    vv:=TCS.Create ;
    vv.nom:=TCS(Fils.data).Nom ;
    vv.Valeur:=TCS(Fils.data).Valeur ;
    Nouveau.data:=vv ;
    End ;
 if Fils.HasChildren then
    RecupEnfants(Tree,Fils,Nouveau) ;
 Fils:=Fils.getnextsibling ;
 end ;
Tree.EndUpdate ;
end;
//////////////////////////////////////////////////////////////////////////////
Procedure Recupnoeud ( Tree : TTreeNodes ; Source : TTreeNode ; OkClear : Boolean ) ;
Var Pere : TTreenode ;
    vv   : TCS ;
Begin
if Tree=nil then exit ;
Tree.beginupdate ;
if OkClear then Begin VideTree(Tree) ; Tree.clear ; End ;
if Source<>nil then
   Begin
   Pere:=Tree.addChild(nil,'') ;
   Pere.Assign(Source) ;
   if assigned(Source.data) then
      Begin
      vv:=TCS.Create ;
      vv.nom:=TCS(Source.data).Nom ;
      vv.Valeur:=TCS(Source.data).Valeur ;
      Pere.data:=vv ;
      End ;
   if Source.HasChildren then
      RecupEnfants(Tree,Source,pere) ;
   End ;
tree.endupdate ;
End ;
//////////////////////////////////////////////////////////////////////////////
Procedure RecopieTreeNodes ( Source,Dest : TTreeNodes ) ;
var NodeS : TTreeNode ;
Begin
NodeS:=Source.GetFirstNode ;
Dest.BeginUpdate ;
VideTree(Dest) ; Dest.Clear ;
while Assigned(NodeS) do
  Begin
  RecupNoeud(Dest,NodeS,FALSE) ;
  NodeS:=NodeS.GetNextSibling ;
  End ;
Dest.EndUpdate ;
End ;
//////////////////////////////////////////////////////////////////////////////
Type TTriRecord = Record
                  St : String ;
                  i  : Integer ;
                  End ;
//////////////////////////////////////////////////////////////////////////////
Procedure Sort (var Table : Array of TTriRecord ; L,H : Integer ) ;
Var i,j : INteger ;
    w,x : TTriRecord ;
Begin
i:=L ; j:=H ;
x:=Table[(i+j) div 2] ;
repeat
  while Table[i].i<x.i do inc(i) ;
  While x.i<Table[j].i do dec(j) ;
  if i<=j then Begin w:=Table[i] ; Table[i]:=Table[j] ; Table[j]:=w ; inc(i) ; Dec(j) ; End ;
until i>j ;
if l<j then Sort(Table,l,j) ;
if i<h then Sort(Table,i,H) ;
End ;
//////////////////////////////////////////////////////////////////////////////
procedure TriArbre ( Tree : TTreeNodes ) ;
Var Node              : TTreeNode ;
    St,Parts          : String ;
    Index,i,Combien   : Integer ;
    Work              : TTreeView ;
    OkFirst           : Boolean ;
    Table             : Array of TTriRecord ;
Begin
if Tree=nil then exit ;
node:=Tree.GetFirstNode ; index:=-1 ;
While (node<>Nil) and (index=-1) do
  Begin
  St:=Node.Text ;
  Index:=first_instruc(ST,Instructions,St) ;
  Node:=Node.GetNextSibling ;
  End ;
if Index>0 then
   Begin
   Parts:=PartsFromIndex(Index,OkFirst) ;
   Table:=Copy(Table,0,0) ;
   node:=Tree.GetFirstNode ; i:=-1 ;
   While (node<>Nil) do
     Begin
     inc(i) ;
     SetLength(Table,Length(Table)+1) ;
     Table[i].St:=Node.Text ;
     Table[i].i:=0 ;
     Node:=Node.GetNextSibling ;
     End ;
   Combien:=Length(Table) ;
   for i:=0 to combien-1 do Table[i].i:=-Find_InstrucToken(Table[i].St,Parts) ;
   sort(Table,1,Combien-1) ;
   Work:=TTreeView.Create(Tree.Owner) ;
   try
      Work.Parent:=Tree.Owner.Parent ;
      Work.Items.BeginUpdate ;
      for i:=0 to combien-1 do
        Begin
        Node:=cherchePart(Tree,Table[i].St) ;
        RecupNoeud(Work.Items,Node,(i=0)) ;
        End ;
      RecopieTreeNodes(Work.Items,Tree) ;
      Work.Items.EndUpdate ;
    Finally
      VideTree(Work.Items) ;
      Work.Items.Clear ;
      Work.Free ;
    End ;
   End ;
End ;
//////////////////////////////////////////////////////////////////////////////
Function InsertPart(Tree : TTreenodes ; Parent : TTreeNode ; St : String ) :TTreeNode ;
Begin
Tree.BeginUpdate ;
if Assigned(Parent) then Tree.addChild(Parent,St)
   else Tree.Add(Parent,St) ;
triArbre(Tree) ;
Result:=CherchePart(Tree,St) ;
Tree.EndUpdate ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

end.


