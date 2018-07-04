{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/07/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : PRINTCALREGLE ()
Mots clefs ... : TOF;PRINTCALREGLE
*****************************************************************}
Unit UTOFPRINTCALREGLE    ;

Interface

Uses StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
{$IFDEF EAGLCLIENT}
   MainEagl,  EQrs1 ,
{$ELSE}
   dbTables, db,FE_Main, Qrs1,
{$ENDIF}
      HCtrls,HEnt1,HMsgBox,UTOM,Dialogs,TiersUtil,EntGC,
      DicoAF,HTB97,Vierge, UTOB, UTOF, AGLInit,UTOFAFTRADUCCHAMPLIBRE ;

Type
  TOF_PRINTCALREGLE = Class (TOF_AFTRADUCCHAMPLIBRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  Private
    ressource : string;
    salarie   : string;
    standard  : string;
    typecal   : string;
    ComboRupt1: THValComboBox;
    CBSTAND   : TCheckBox;
    pref, antipref : string;
    //BValider  : TToolbarButton97;
    procedure CBRUPTURE1_OnChange(sender : TObject);
    procedure CBSTAND_OnClick(sender : TObject) ;
    procedure Ressource_OnChange(sender : TObject);
    //procedure Salarie_OnChange(sender : TObject);
  end ;

Procedure AFLanceFiche_Edit_Calendrier;
Procedure AFLanceFiche_Edit_CalRegle;
Implementation

procedure TOF_PRINTCALREGLE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PRINTCALREGLE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PRINTCALREGLE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PRINTCALREGLE.OnLoad ;
begin
  //PgiInfo(GetControlText('XX_WHERE'),'SAV');
  //PgiInfo(GetControlText('ANTIXXWHERE'),'SAV');
  If GetControlText('XX_RUPTURE1') = '' then
    begin;
     SetControlText('XX_RUPTURE1',pref+'_RESSOURCE');
     ComboRupt1.Value := pref+'_RESSOURCE';
     ComboRupt1.Enabled := True;
     //SetControlText(pref+'_RESSOURCE',Ressource);
    end
  else
    begin
     //SetControlText('XX_WHERE',GetControlText('XX_RUPTURE1')+'<>"***"');
    end;
  If CBSTAND.checked = false then SetControlText('XX_WHERE',GetControlText('XX_WHERE')+' AND '+GetControlText('XX_RUPTURE1')+'<>"***"');
  SetControlText('ANTIXXWHERE',StringReplace(GetcontrolText('XX_WHERE'),pref+'_',AntiPref+'_',[rfReplaceAll]));
  //SetControlText('DATESPE',GetControlText('DATE'));
  Inherited ;
end ;

procedure TOF_PRINTCALREGLE.OnArgument (S : String ) ;
var
champ, valeur  : string;
begin
  Inherited ;
  //
   pref := GetControlText('INDENTIFPREF');
   if pref ='ACA' then AntiPref :='ACG' else AntiPref:='ACA';
   ComboRupt1 := ThValcomboBox(GetControl('CBRUPTURE1'));
   ComboRupt1.OnChange := CBRUPTURE1_OnChange;
   CBSTAND := TCheckBox(GetControl('CBSTAND'));
   CBSTAND.OnClick := CBSTAND_OnClick;
   THEdit(GetControl(pref+'_SALARIE')).OnChange := Ressource_OnChange;
   THEdit(GetControl(pref+'_RESSOURCE')).OnChange := Ressource_OnChange;
   
  //
  if S <>'' then //Lancement depuis la Fiche AFCALENDIERREGLE , sinon on imprime tout
   begin
    While S<>'' do
     begin
       Valeur := ReadTokenSt(S);
       Champ := ReadTokenPipe(Valeur,'=');
       If Champ='RESSOURCE' then Ressource := Valeur
       else If Champ='SALARIE' then Salarie := Valeur
       else if champ = 'NOFILTRE' then  TFQRS1(Ecran).FiltreDisabled:=true   //mcd 12/02/03
       else if Champ = 'STANDARD' then Standard := Valeur
       else if Champ = 'TYPECAL' then TypeCal := Valeur ;
     end;
    end;
    ComboRupt1 := ThValcomboBox(GetControl('CBRUPTURE1'));
    If TypeCal = 'SAL' then
       begin
        ComboRupt1.Value := pref+'_SALARIE';
        SetControlText(pref+'_SALARIE',salarie);
        SetControlText('XX_RUPTURE1',pref+'_SALARIE');
        SetControlVisible('XX_VARIABLE1',False);
        SetControlText('XX_VARIABLE1','SALARIE');
       end
    else
    if TypeCal = 'RES' then
       begin
        ComboRupt1.Value := pref+'_RESSOURCE';
        SetControlText(pref+'_RESSOURCE',Ressource);
        SetControlVisible('XX_VARIABLE1',False);
        SetControlText('XX_VARIABLE1','RESSOURCE');
       end
    else
    If TypeCal ='STD' then
       begin
        SetControlEnabled('CBRUPTURE1',False);
        SetControlText(pref+'_STANDCALEN',standard);
        SetControlText('XX_RUPTURE1',pref+'_STANDCALEN');
        SetControlText('XX_VARIABLE1','STANDCALEN');
        TcheckBox(GetControl('CBSTAND')).Checked := True;
        SetControlVisible('RUPTURE',false);  //mcd 08/01/03 on ne peut rien saisir sur cet ongelt dans ce cas d'appel
       end
    else
       begin
        SetControlText('XX_RUPTURE1',pref+'_RESSOURCE');
        ComboRupt1.Value := pref+'_RESSOURCE';
        ComboRupt1.Enabled := True;
        SetControlText(pref+'_RESSOURCE',Ressource);
        SetControlVisible('XX_VARIABLE1',False);
        SetControlText('XX_VARIABLE1','RESSOURCE');
       end;
  //
  If CBSTAND.checked = true then SetControltext('XX_WHERE',pref+'_RESSOURCE="***" AND '+pref+'_SALARIE="***" AND '+pref+'_STANDCALEN<>"***"' ) else
                                 SetControltext('XX_WHERE','NOT ('+pref+'_RESSOURCE="***" AND '+pref+'_SALARIE="***")');
end ;

procedure TOF_PRINTCALREGLE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PRINTCALREGLE.CBRUPTURE1_OnChange(sender : TObject) ;
begin
  SetControlText('XX_RUPTURE1',ComboRupt1.value);
  SetControlText('XX_VARIABLE1',copy(GetControlText('XX_RUPTURE1'),5,length(GetControlText('XX_RUPTURE1'))));
end;

procedure TOF_PRINTCALREGLE.CBSTAND_OnClick(sender : TObject) ;
begin
If CBSTAND.checked = true then
  begin
    SetControlText('XX_WHERE',pref+'_RESSOURCE="***" AND '+pref+'_SALARIE="***"');
    SetControlText('XX_RUPTURE1',pref+'_STANDCALEN');
    ComboRupt1.enabled := False;
    SetControlVisible('XX_RUPTURE1',True);
    ComboRupt1.visible :=False;
    SetControlText(pref+'_RESSOURCE','');
    SetControlEnabled(pref+'_RESSOURCE',False);
    SetControlText(pref+'_SALARIE','');
    SetControlEnabled(pref+'_SALARIE',False);
    SetControlText('XX_VARIABLE1','STANDCALEN');
    SetControlVisible('RUPTURE',false);  //mcd 08/01/03 on ne peut rien saisir sur cet ongelt dans ce cas d'appel
  end
else
  begin
   SetControlText('XX_WHERE','NOT ('+pref+'_RESSOURCE="***" AND '+pref+'_SALARIE="***")');
   SetControlText('XX_RUPTURE1','');
   ComboRupt1.enabled := True;
   SetControlVisible('XX_RUPTURE1',False);
   ComboRupt1.visible :=True;
   SetControlEnabled(pref+'_RESSOURCE',True);
   SetControlEnabled(pref+'_SALARIE',True);
   SetControlText('XX_VARIABLE1',copy(GetControlText('XX_RUPTURE1'),5,length(GetControlText('XX_RUPTURE1'))));
   SetControlVisible('RUPTURE',True);  //mcd 08/01/03 pour OK si mis en invisibel avant
  end;

end;

procedure TOF_PRINTCALREGLE.RESSOURCE_OnChange(sender : Tobject);
begin
SetControlEnabled('ACG_SALARIE',GetControlText(pref+'_RESSOURCE')<>'');
end;

{procedure  TOF_PRINTCALREGLE.SALARIE_OnChange(sender : Tobject);
begin
SetControlEnabled('ACG_RESSOURCE',GetControlText(pref+'_SALARIE')<>'');
end;
}

Procedure AFLanceFiche_Edit_Calendrier;
begin
AGLLanceFiche ('AFF','AFPRINTCALENDRIER','','','');
end;
 Procedure AFLanceFiche_Edit_CalRegle;
begin
AGLLanceFiche ('AFF','AFPRINTCALREGLE','','','');
end;

Initialization
  registerclasses ( [ TOF_PRINTCALREGLE ] ) ;
end.

 