unit UTofAFTiers_Rech;

interface
uses  HMsgBox,Ent1,windows,Classes,UTOF,stdctrls,
{$IFDEF EAGLCLIENT}
      eMul,MaineAGL,
{$ELSE}
       HDB, mul, db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}

 DicoBTP, Hent1,UtofAftiers_Mul,SysUtils,HCtrls,HTB97;

Type
     TOF_AfTiers_RECH = Class (TOF_AFTiers_Mul)
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnUpdate                 ; override ;
     END ;
 Function AFLanceFiche_REch_tiers(Range,Argument:string):variant;

implementation
uses UFonctionsCBP;

procedure TOF_AfTiers_RECH.OnUpdate ;
begin
Inherited ;
(*if (TFMul(Ecran).FListe.DataSource.DataSet.RecordCount = 1) then
  begin
  if (GetCheckBoxState('PARLIBELLE') = cbChecked)  then
	  TFMul(Ecran).Retour:='CODE:'+GetControlText('T_TIERS')+';LIBELLE:'+GetControlText('T_LIBELLE')+';'
  else
	  TFMul(Ecran).Retour:=GetControlText('T_TIERS');

  //TFMul(Ecran).close;
  end;
//        ChercheClick;
  *)
end ;

procedure TOF_AfTiers_RECH.OnArgument(stArgument : String );
var x,y : integer;
    Critere, tmp : string;
BEGIN
Inherited;
          // ATTENTION, ne fonctionne pas si la piece accepte plusieurs
          // nature client : seule la 1ere nature sera permise sur le
          // bouton création !!!!
          // ATTENTION, ne pas utiliser depuis le DP, ou voir si pb
          // l'argument passé depuis la facture est '(T_NATUREAUXI="CLI")'
tmp := stArgument;
Critere:=(Trim(ReadTokenSt(tmp)));
While (Critere <>'') do
   begin
   x:= Pos('T_NATUREAUXI=',Critere) ;
   if x<>0 then
      begin
      y := Pos('T_NATUREAUXI=',Copy(Critere,X+18,Length(Critere)-X-18)) ;
      if y = 0 then begin// 1 seule nature d'auxi...
         Setcontroltext('T_NATUREAUXI',copy(Stargument,x+14,3)) ;
         SetControlVisible ('T_NATUREAUXI',False);
         SetControlVisible ('TT_NATUREAUXI',False);
         end
      else
         begin
         SetControlVisible ('T_NATUREAUXI',True); SetControlVisible ('TT_NATUREAUXI',True);
         end;
      Setcontroltext('XX_WHERE',Critere) ;
      end;
   Critere:=(Trim(ReadTokenSt(tmp)));
   end;
if Not ExJaiLeDroitConcept(TConcept(gcCLICreat),False) then
  BEGIN
  SetControlVisible('BINSERT',False) ;
  END ;
  {$IFDEF CCS3}
  if (getcontrol('PZONES') <> Nil) then SetControlVisible ('PZONES', False);
  if (getcontrol('PRESCLIENT') <> Nil) then SetControlVisible ('PRESCLIENT', False);
  {$ENDIF}
END;


Function AFLanceFiche_REch_tiers(Range,Argument:string):variant;
begin
result:=AGLLanceFiche('AFF','AFTIERS_RECH',range,'',argument);
end;


Initialization
registerclasses([TOF_AfTiers_RECH]);
end.
