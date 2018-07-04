{***********UNITE*************************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 13/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : YYDTLINKS_MUL ()
Mots clefs ... : TOF;YYDTLINKS_MUL
*****************************************************************}
Unit YYDTLINKS_MUL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} 
     mul,
     Fe_Main,
{$else}
     eMul, 
     uTob,
     MainEagl,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF ,LicUtil, YdataTypeLInks_Tom; 

Type
  TOF_YYDTLINKS_MUL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (StArgument : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Procedure YYLanceFiche_Mul_YYdtLinks (starg:string);
Implementation

procedure TOF_YYDTLINKS_MUL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_YYDTLINKS_MUL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_YYDTLINKS_MUL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_YYDTLINKS_MUL.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_YYDTLINKS_MUL.OnArgument (StArgument : String ) ;
var critere ,champ,val: string;
    X : integer;
    Cegid : boolean;
begin
  Inherited ;
  Cegid :=False;
 SetControlText ('ARGUMENT',StArgument); // on conserve les paramètres pour les passer à la fiche de saisie
 Critere:=(Trim(ReadTokenSt(stArgument)));
 While (Critere <>'') do
 begin
  if Critere<>'' then
    begin
    X:=pos('=',Critere);
    if x<>0 then
    begin
      Champ:=copy(Critere,1,X-1);
      Val:=Copy (Critere,X+1,length(Critere)-X);
    end
    else begin   //pour paramètre CEGID
      Champ :=Critere;
      Val:='';
      end;
    if Champ = 'PREDEFINI' then
      begin
      val := StringReplace (val, '+', ';', [rfReplaceAll]);
      SetControlText ('YDL_PREDEFINI',val);
      end
    else if Champ ='CEGID' then Cegid :=true
    else if Champ = 'RACINEAPPLI'then SetControlText ('YDL_CODEHDTLINK_',val);
    end;
  Critere := (Trim(ReadTokenSt(stArgument)));
 end;
 if (V_PGI.PassWord = CryptageSt(DayPass(V_PGI.DateEntree))) or  (Cegid) then
   begin // dans ce cas , tout est permis. ne tient pas compte des paramètres passé
   SetCOntrolText ('YDL_PREDEFINI','');
   SetCOntrolText ('YDL_CODEHDTLINK_','');
   end;
end ;

procedure TOF_YYDTLINKS_MUL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_YYDTLINKS_MUL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_YYDTLINKS_MUL.OnCancel () ;
begin
  Inherited ;
end ;

procedure YYLanceFiche_Mul_YYdtLinks (starg:string);
begin
     AGLLanceFiche ('YY','YYDTLINKS_MUL', '', '', stArg);
end;

Initialization
  registerclasses ( [ TOF_YYDTLINKS_MUL ] ) ; 
end.
