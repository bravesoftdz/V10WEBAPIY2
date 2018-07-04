unit UTofAfChoixEcheAffair;


interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
      Maineagl,
{$ELSE}
   {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db,  FE_Main,
{$ENDIF}
      HCtrls,UTOF,AffaireUtil,HEnt1,EntGC,HTB97,UTOB,Stat,M3FP;

Type
     Tof_AFChoixEcheAffair = Class (TOF)
        procedure OnArgument(stArgument : String ) ; override ;
        procedure OnUpdate ; override ;
      public
        Devise          :  string;
        sMtPeriodique   :  string;
        sMtGlobal       :  string;
        sGenerauto      :  string;
        Choix1          :  TRadioButton;
        Choix2          :  TRadioButton;
        Choix3          :  TRadioButton;
     END ;

Function AFLanceFiche_ChoixEcheAff(Argument:string):variant;

implementation


procedure Tof_AFChoixEcheAffair.OnArgument(stArgument : String );
Var
   Critere, Champ, valeur  : String;
   x : integer;
Begin
Inherited;
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'DEVISE'         then Devise := Valeur         else
        if Champ = 'MTPERIODIQUE'   then sMtPeriodique := Valeur  else
        if Champ = 'MTGLOBAL'       then sMtGlobal := Valeur ;
        if Champ = 'AFF_GENERAUTO'  then sGenerauto := Valeur;
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;

Choix1 := TRadioButton(GetControl('CHOIX1'));
Choix2 := TRadioButton(GetControl('CHOIX2'));
Choix3 := TRadioButton(GetControl('CHOIX3'));

if (Choix1<>nil) then
   Choix1.Caption := Choix1.Caption + ' ' + sMtPeriodique + ' ' + Devise + ' sur chaque échéance' ;
if (Choix2<>nil) then
   begin
   Choix2.Caption := Choix2.Caption + ' ' + sMtGlobal + ' ' + Devise ;
   // PA le 31/08/2001 si fact. sur lignes d'affaires ou contrat l'option 2 non visible
   if (sGenerAuto = 'CON') then Choix2.Visible :=False;
   end;
End;

procedure Tof_AFChoixEcheAffair.OnUpdate;
Begin
Inherited;
End;

Function AFLanceFiche_ChoixEcheAff(Argument:string):variant;
begin
result:=AGLLanceFiche ('AFF','AFCHOIXECHEAFFAIR','','',Argument);
end;

Initialization
registerclasses([Tof_AFChoixEcheAffair]);
end.
