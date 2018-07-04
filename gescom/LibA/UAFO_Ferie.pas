unit UAFO_Ferie;

interface
uses  Classes,
{$IFDEF EAGLCLIENT}
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
      sysutils,UTob,ParamSoc,HCtrls
		;
type TJourFerie = Class
    Private
    FTobFerie : tob;
    Procedure ChargeJourFerie;
    Public
    Constructor Create;
    Destructor Destroy;override;
    function TestJourFerie(DateJour :TDateTime):boolean;
    property TobFerie : Tob Read FTobFerie;
    End;

// Fonction calendrier indépendantes.
Function  TestJourFermeture(Queljour:Integer;JF:array of integer):boolean;
Procedure ChargeJourFermeture(Var JF:array of integer);


implementation

//**************** Fonctions sur l'objet Jour férié  *******************************//
Procedure TJourFerie.ChargeJourFerie;
Var QQ:TQUERY;
 SQL : string;
begin
If FTobFerie <> Nil Then Begin FTobFerie.Free; FTobFerie:=Nil; End;
FTobFerie:=Tob.Create('Liste jours Ferie',Nil,-1);
QQ := nil;
Try
    //mcd 03/03/03 ajout test sur STD ou dossier
Sql := 'SELECT AJF_JOURFIXE,AJF_JOUR,AJF_MOIS,AJF_ANNEE FROM JOURFERIE';
Sql := Sql + ' WHERE ##AJF_PREDEFINI##';
QQ:=OpenSQL(Sql,True,-1,'',true);
if Not QQ.EOF then
    FTobFerie.LoadDetailDB('JOURFERIE','','',QQ,True)
else
    begin
    FTobFerie.Free;
    FTobFerie := nil;
    end;

Finally
   Ferme(QQ);
   End;
end;

function TJourFerie.TestJourFerie(DateJour :TDateTime):boolean;
Var TobDet : TOB;
    i :integer;
    An,mois,jour : word;
    DateFerie:TDateTime;
    bDateValide:boolean;
begin
result:=False;
DateFerie:=0;
If (TobFerie <> Nil) Then
   Begin
   for i:=0 to TobFerie.Detail.Count-1 do
       BEGIN
       TobDet:=TobFerie.Detail[i];
       If (TobDet.GetValue('AJF_JOURFIXE')='X') Then Decodedate(DateJour,An,mois,Jour)
       else An:= TobDet.GetValue('AJF_ANNEE');

       try
       bDateValide:=true;
       DateFerie := EncodeDate(An,TobDet.GetValue('AJF_MOIS'),TobDet.GetValue('AJF_JOUR'));
       except
       bDateValide:=false;
       end;

       if bDateValide then
         If(DateJour = DateFerie) Then
            Begin result:=True; Break; end;
       End;
   end;
end;


Constructor TJourFerie.Create;
Begin
ChargeJourFerie;
End;

Destructor TJourFerie.Destroy;
Begin
FTobFerie.Free;
FTobFerie:=nil;
inherited;
End;

// Fonctions spécifiques sur les jours de fermeture
Function  TestJourFermeture(Queljour:Integer;JF:array of integer):boolean;
begin
Result:=False;
If (JF[QuelJour-1]=-1) then Result:=True;
end;

Procedure ChargeJourFermeture (Var JF:array of integer);
Var i:integer;

begin
JF[0]:=GetParamSoc('SO_JOURFERMETURE');
For i:=1 To 6 do
   JF[i]:=GetParamSoc('SO_JOURFERMETURE'+IntToStr(i));
end;


end.
