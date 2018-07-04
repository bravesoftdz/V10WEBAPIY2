unit UTofChoixTrans_Act;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,vierge,
{$IFDEF EAGLCLIENT}
       Maineagl,
{$ELSE}
   dbTables, db, FE_Main,
{$ENDIF}
     HCtrls,HEnt1,HMsgBox,UTOF, entgc;

Type
     TOF_CHOIXTRANS_ACT = Class (TOF)
        procedure OnArgument(stArgument : String ) ; override ;
     public
     END ;
Function AFLanceFiche_ChoixTransF_ACt(Argument:string):variant;

implementation

procedure TOF_CHOIXTRANS_ACT.OnArgument(stArgument : String );
var
RBProp:TRadioButton;
critere, champ, valeur:string;
x:integer;
Begin
Inherited;
RBProp:=TRadioButton(GetControl('RBPROPOSITION'));

// on lit dans l''argument, le code proposition passé
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    X:=pos(':',Critere);
    if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;

    if (Champ='PROPOSITION') then
	        begin
   	        RBProp.Caption := RBProp.Caption +' '+ Valeur;
	        end;

    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
End;

Function AFLanceFiche_ChoixTransF_ACt(Argument:string):variant;
begin
result:=AGLLanceFiche ('AFF','CHOIXTRANS_ACT','','',Argument);
end;


Initialization
registerclasses([TOF_CHOIXTRANS_ACT]);


end.
