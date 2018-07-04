unit UTofAFTiers_Mul;

interface
uses  windows,Classes,UTOF, DicoBTP, Hent1,SysUtils,forms,HCtrls,UtilPGI,
      HMsgBox,Ent1,AGlINit,UtilGc,UtofAftraducChampLibre
{$IFDEF EAGLCLIENT}
      ,Maineagl,emul
{$ELSE}
		 ,mul,FE_Main
{$ENDIF}
		;
Type
     TOF_AfTiers_MUL = Class (TOF_AFTRADUCCHAMPLIBRE)
        procedure OnArgument (stArgument : String ) ; override ;
     END ;
 Procedure AFLanceFiche_Mul_tiers(Range,Argument:string);
 Procedure AFLanceFiche_Mul_tiers_Multi(Range,Argument:string);

implementation
uses UFonctionsCBP;

procedure TOF_AfTiers_MUL.OnArgument(stArgument : String );
var      Critere, Champ, valeur : string;
         x : integer;
BEGIN
Inherited;

// mcd 24/04/02If ctxTempo in V_PGI.PGIContexte  then
If not(ctxSCot in V_PGI.PGIContexte)  then
    begin
    SetControlVisible ('TT_MOISCLOTURE' , False);
    SetControlVisible ('T_MOISCLOTURE' , False);
    SetControlVisible ('TT_MOISCLOTURE_' , False);
    SetControlVisible ('T_MOISCLOTURE_' , False);
    end;
	// on lit dans l'argument, l'action eventuelle
	// et le code tiers passer
Critere:=(Trim(ReadTokenSt(stArgument)));
While (Critere <>'') do
    BEGIN
    if Critere<>'' then
        BEGIN
        X:=pos(':',Critere);
        if x=0 then  X:=pos('=',Critere);
        if x<>0 then
           begin
           Champ:=copy(Critere,1,X-1);
           Valeur:=Copy (Critere,X+1,length(Critere)-X);
           end;
        if Champ = 'ORIGINE'    then begin
                SetControlText('XXOrigine',critere);
                SetControlVisible('BInsert',False);	// il ne faut pas permettre la création, tout est grisé !!!
                end
       else if Champ = 'MULTI'    then begin   // mcd 05/03/02 pour gérer accès multi fiches
                TFMul(Ecran).caption:='Client multi fiches';
                SetControlChecked('Multi',True);
                SetControlVisible('PRESCLIENT',False);
                SetControlVisible('PSTATCLIENT',False);
                SetControlVisible('PZONE',False);
                TFMUL(Ecran).Dbliste:='AFMULTIERS';
                if TfMul(Ecran).Q <> NIL then TfMul(Ecran).Q.Liste  := 'AFMULTIERS';
                UpdateCaption(TFMul(Ecran)) ;
                end
        else if Champ = 'ACTION'       then begin
                 SetControlText('XXAction',critere);
                  if (StringToAction(GetcontrolText('XXAction')) = taConsult) then begin
                          SetControlVisible('Binsert',False);
                          end;
                 end ;
        END;
    Critere:=(Trim(ReadTokenSt(stArgument)));
    END;
if Not ExJaiLeDroitConcept(TConcept(gcCLICreat),False) then
  BEGIN
  SetControlVisible('BINSERT',False) ;
  END ;
  {$IFDEF CCS3}
  if (getcontrol('PZONE') <> Nil) then SetControlVisible ('PZONE', False);
  if (getcontrol('PRESCLIENT') <> Nil) then SetControlVisible ('PRESCLIENT', False);
  {$ENDIF}
END;


Procedure AFLanceFiche_Mul_tiers(Range,Argument:string);
begin
     AGLLanceFiche('AFF','AFTIERS_MUL',Range,'',Argument);
end;

Procedure AFLanceFiche_Mul_tiers_Multi(Range,Argument:string);
begin
     AGLLanceFiche('AFF','AFTIERS_MULTI',Range,'',Argument);
end;


Initialization
registerclasses([TOF_AfTiers_MUL]);
end.
