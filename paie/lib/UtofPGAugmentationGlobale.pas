{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 04/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGAUGMGLOBALE ()
Mots clefs ... : TOF;PGAUGMGLOBALE
*****************************************************************}
Unit UtofPGAugmentationGlobale ;

Interface

Uses StdCtrls,Controls,Classes,
{$IFNDEF EAGLCLIENT}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul,
{$else}
     eMul,uTob,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,Vierge,ParamSoc,EntPaie; 

Type
  TOF_PGAUGMGLOBALE = Class (TOF)
    procedure OnUpdate                 ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    IArrondiAugm,PrecisionArrondi,PctAugmDec : Integer;
    procedure ChangeMontantGlob(Sender : Tobject);
    Function FormatageDouble(Chaine : String) : Double;
    Function ArrondiAugm(Montant : Double) : Double;
  end ;

   TOF_PGAUGM_XLS = Class (TOF)
   procedure OnUpdate                 ; override ;
   procedure OnArgument (S : String ) ; override ;
  end;

Implementation

procedure TOF_PGAUGMGLOBALE.OnUpdate ;
begin
  Inherited ;
            TFVierge(Ecran).Retour := GetControlText('PCTFIXE')+ ';' + GetControlText('PCTVARIABLE')+';'+GetcontrolText('MOTIFAUGM');
end ;

procedure TOF_PGAUGMGLOBALE.OnArgument (S : String ) ;
var Edit : THEdit;
    Fixe,Variable : String;
    Q : TQuery;
    CodeArrondi : String;
begin
  Inherited ;
        CodeArrondi := GetParamSoc('SO_PGAUGMSALARR');
        PctAugmDec := GetParamSoc('SO_PGAUGMPCTDEC');
        Q := OpenSQL('SELECT * FROM COMMUN WHERE CO_TYPE="PRR" AND CO_CODE="'+CodeArrondi+'"',True);
        If Not Q.Eof then
        begin
             IArrondiAugm := StRtoInt(Q.FindField('CO_ABREGE').AsString);
             PrecisionArrondi := StRtoInt(Q.FindField('CO_LIBRE').AsString);
        end
        else
        begin
             IArrondiAugm := 0;
             PrecisionArrondi := 0;
        end;
        Fixe := ReadTokenPipe(S,';');
        Variable := ReadTokenPipe(S,';');
        Edit := THEdit(Getcontrol('PCTFIXE'));
        If Edit <> Nil then Edit.OnExit := ChangeMontantGlob;
        Edit := THEdit(Getcontrol('PCTVARIABLE'));
        If Edit <> Nil then Edit.OnExit := ChangeMontantGlob;
        Edit := THEdit(Getcontrol('PCTTOTAL'));
        If Edit <> Nil then Edit.OnExit := ChangeMontantGlob;
        Edit := THEdit(Getcontrol('NEWFIXE'));
        If Edit <> Nil then Edit.OnExit := ChangeMontantGlob;
        Edit := THEdit(Getcontrol('NEWVARIABLE'));
        If Edit <> Nil then Edit.OnExit := ChangeMontantGlob;
        Edit := THEdit(Getcontrol('NEWTOTAL'));
        If Edit <> Nil then Edit.OnExit := ChangeMontantGlob;
        setControlText('ANCFIXE',Fixe);
        setControlText('ANCVARIABLE',Variable);
        setControlText('ANCTOTAL',FloatToStr(StrToFloat(Variable)+StrToFloat(Fixe)));
        setControlText('PCTFIXE','0');
        setControlText('PCTVARIABLE','0');
        setControlText('PCTTOTAL','0');
        setControlText('NEWFIXE',Fixe);
        setControlText('NEWVARIABLE',vARIABLE);
        setControlText('NEWTOTAL',FloatToStr(StrToFloat(Variable)+StrToFloat(Fixe)));
        setControlText('DIFFIXE','0');
        setControlText('DIFVARIABLE','0');
        setControlText('DIFTOTAL','0');
        TFVierge(Ecran).Retour := '0;0';
        SetControlText('MOTIFAUGM','');
end ;

procedure TOF_PGAUGMGLOBALE.OnClose ;
begin
  Inherited ;

end ;

procedure TOF_PGAUGMGLOBALE.ChangeMontantGlob(Sender : Tobject);
var Anc,New,Pct,Autre,AutreAnc : Double;
begin
        If (THEdit(sender).Name='NEWFIXE') then
        begin
                If IsNumeric(GetControlText('NEWFIXE')) and IsNumeric(GetControlText('PCTFIXE')) then
                begin
                        Anc := FormatageDouble(GetControlText('ANCFIXE'));
                        New := FormatageDouble(GetControlText('NEWFIXE'));
                        Autre := FormatageDouble(GetControlText('NEWVARIABLE'));
                        AutreAnc := FormatageDouble(GetControlText('ANCVARIABLE'));
                        If (Anc <> 0) and (New <> 0) Then Pct := Arrondi(((New - Anc)/Anc)*100,PctAugmDec)
                        else Pct := 0;
                        If (ArrondiAugm((Anc * (Pct/100))+ Anc)) = New then Exit;
                        SetControlText('PCTFIXE',FloatToStr(Pct));
                        SetControlText('NEWTOTAL',FloatToStr(New + Autre));
                        If (Anc + AutreAnc) <> 0 then SetControlText('PCTTOTAL',FloatToStr(Arrondi((((New + Autre) - (Anc+AutreAnc))/(Anc+AutreAnc))*100,PctAugmDec)))
                        else SetControlText('PCTTOTAL','0');
                        SetControlText('DIFFIXE',FloatToStr(New - Anc));
                        SetControlText('DIFTOTAL',FloatToStr((New + Autre) - (Anc + AutreAnc)));
                end
                else
                begin
                     SetControlText('NEWFIXE',GetControlText('ANCFIXE'));
                     SetControlText('PCTFIXE','0');
                     SetControlText('NEWTOTAL',GetControlText('ANCTOTAL'));
                     Anc := FormatageDouble(GetControlText('NEWTOTAL'));
                     New := FormatageDouble(GetControlText('ANCTOTAL'));
                     If (Anc <> 0) and (New <> 0) then SetControlText('PCTTOTAL',FloatToStr(Arrondi(((New - Anc)/Anc)*100,PctAugmDec)))
                     else SetControlText('PCTTOTAL','0');
                     SetControlText('DIFFIXE','0');
                     SetControlText('DIFTOTAL',FloatToStr(New - Anc));
                end;
        end;
        If (THEdit(sender).Name='PCTFIXE') then
        begin
                If IsNumeric(GetControlText('NEWFIXE')) and IsNumeric(GetControlText('PCTFIXE')) then
                begin
                        Anc := FormatageDouble(GetControlText('ANCFIXE'));
                        Pct := FormatageDouble(GetControlText('PCTFIXE'));
                        SetControlText('NEWFIXE',FloatToStr(ArrondiAugm(Anc + (Anc * (Pct /100)))));
                        New := FormatageDouble(GetControlText('NEWFIXE'));
                        SetControlText('DIFFIXE',FloatToStr(New - Anc));
                        Autre := FormatageDouble(GetControlText('NEWVARIABLE'));
                        AutreAnc := FormatageDouble(GetControlText('ANCVARIABLE'));
                        SetControlText('NEWTOTAL',FloatToStr(New + Autre));
                        If (Anc + AutreAnc) <> 0 then SetControlText('PCTTOTAL',FloatToStr(Arrondi((((New + Autre) - (Anc+AutreAnc))/(Anc+AutreAnc))*100,PctAugmDec)))
                        else SetControlText('PCTTOTAL','0');
                        SetControlText('NEWTOTAL',FloatToStr(New + Autre));
                        If (Anc + AutreAnc) <> 0 then SetControlText('PCTTOTAL',FloatToStr(Arrondi((((New + Autre) - (Anc+AutreAnc))/(Anc+AutreAnc))*100,PctAugmDec)))
                        else SetControlText('PCTTOTAL','0');
                        SetControlText('DIFTOTAL',FloatToStr((New + Autre) - (Anc + AutreAnc)));
                end
                ELSE
                begin
                     SetControlText('NEWFIXE',GetControlText('ANCFIXE'));
                     SetControlText('DIFFIXE','0');
                     SetControlText('PCTFIXE','0');
                     SetControlText('NEWTOTAL',GetControlText('ANCTOTAL'));
                     Anc := FormatageDouble(GetControlText('NEWTOTAL'));
                     New := FormatageDouble(GetControlText('ANCTOTAL'));
                     If (Anc <> 0) and (New <> 0) then SetControlText('PCTTOTAL',FloatToStr(Arrondi(((New - Anc)/Anc)*100,PctAugmDec)))
                     else SetControlText('PCTTOTAL','0');
                     SetControlText('DIFTOTAL',FloatToStr(New - Anc));
                end;
        end;

        If (THEdit(sender).Name='NEWVARIABLE') then
        begin
                If IsNumeric(GetControlText('NEWVARIABLE')) and IsNumeric(GetControlText('PCTVARIABLE')) then
                begin
                        Anc := FormatageDouble(GetControlText('ANCVARIABLE'));
                        New := FormatageDouble(GetControlText('NEWVARIABLE'));
                        Autre := FormatageDouble(GetControlText('NEWFIXE'));
                        AutreAnc := FormatageDouble(GetControlText('ANCFIXE'));
                        If (Anc <> 0) and (New <> 0) Then Pct := Arrondi(((New - Anc)/Anc)*100,PctAugmDec)
                        else Pct := 0;
                        If (ArrondiAugm((Anc * (Pct/100))+ Anc)) = New then Exit;
                        SetControlText('VARIABLE',FloatToStr(Pct));
                        SetControlText('DIFVARIABLE',FloatToStr(New - Anc));
                        SetControlText('DIFTOTAL',FloatToStr((New + Autre) - (Anc + AutreAnc)));
                end
                else
                begin
                     SetControlText('NEWVARIABLE',GetControlText('ANCVARIABLE'));
                     SetControlText('PCTVARIABLE','0');
                     SetControlText('NEWTOTAL',GetControlText('ANCTOTAL'));
                     Anc := FormatageDouble(GetControlText('NEWTOTAL'));
                     New := FormatageDouble(GetControlText('ANCTOTAL'));
                     If (Anc <> 0) and (New <> 0) then SetControlText('PCTTOTAL',FloatToStr(Arrondi(((New - Anc)/Anc)*100,PctAugmDec)))
                     else SetControlText('PCTTOTAL','0');
                     SetControlText('DIFVARIABLE','0');
                     SetControlText('DIFTOTAL',FloatToStr(New - Anc));
                end;
        end;
        If (THEdit(sender).Name='PCTVARIABLE') then
        begin
                If IsNumeric(GetControlText('NEWVARIABLE')) and IsNumeric(GetControlText('PCTVARIABLE')) then
                begin
                        Anc := FormatageDouble(GetControlText('ANCVARIABLE'));
                        Pct := FormatageDouble(GetControlText('PCTVARIABLE'));
                        SetControlText('NEWVARIABLE',FloatToStr(ArrondiAugm(Anc + (Anc * (Pct /100)))));
                        New := FormatageDouble(GetControlText('NEWVARIABLE'));
                        Autre := FormatageDouble(GetControlText('NEWFIXE'));
                        AutreAnc := FormatageDouble(GetControlText('ANCFIXE'));
                        SetControlText('NEWTOTAL',FloatToStr(New + Autre));
                        If (Anc + AutreAnc) <> 0 then SetControlText('PCTTOTAL',FloatToStr(Arrondi((((New + Autre) - (Anc+AutreAnc))/(Anc+AutreAnc))*100,PctAugmDec)))
                        else SetControlText('PCTTOTAL','0');
                        SetControlText('DIFVARIABLE',FloatToStr(New - Anc));
                        SetControlText('DIFTOTAL',FloatToStr((New + Autre) - (Anc + AutreAnc)));
                end
                else
                begin
                SetControlText('NEWVARIABLE',GetControlText('ANCVARIABLE'));
                SetControlText('PCTVARIABLE','0');
                SetControlText('NEWTOTAL',GetControlText('ANCTOTAL'));
                Anc := FormatageDouble(GetControlText('NEWTOTAL'));
                New := FormatageDouble(GetControlText('ANCTOTAL'));
                If (Anc <> 0) and (New <> 0) then SetControlText('PCTTOTAL',FloatToStr(Arrondi(((New - Anc)/Anc)*100,PctAugmDec)))
                else SetControlText('PCTTOTAL','0');
                SetControlText('DIFVARIABLE','0');
                SetControlText('DIFTOTAL',FloatToStr(New - Anc));
                end;
        end;

        If (THEdit(sender).Name='NEWTOTAL') then
        begin
                If IsNumeric(GetControlText('NEWTOTAL')) and IsNumeric(GetControlText('PCTTOTAL')) then
                begin
                        Anc := FormatageDouble(GetControlText('ANCTOTAL'));
                        New := FormatageDouble(GetControlText('NEWTOTAL'));
                        If (Anc <> 0) and (New <> 0) Then Pct := Arrondi(((New - Anc)/Anc)*100,PctAugmDec)
                        else Pct := 0;
                        If (ArrondiAugm((Anc * (Pct/100))+ Anc)) = New then Exit;
                        SetControlText('PCTTOTAL',FloatToStr(Pct));
                        SetControlText('DIFTOTAL',FloatToStr(New - Anc));
                end;
        end;
        If (THEdit(sender).Name='PCTTOTAL') then
        begin
                If IsNumeric(GetControlText('NEWTOTAL')) and IsNumeric(GetControlText('PCTTOTAL')) then
                begin
                        Anc := FormatageDouble(GetControlText('ANCTOTAL'));
                        Pct := FormatageDouble(GetControlText('PCTTOTAL'));
                        SetControlText('NEWTOTAL',FloatToStr(ArrondiAugm(Anc + (Anc * (Pct /100)))));
                        New := FormatageDouble(GetControlText('NEWTOTAL'));
                        SetControlText('DIFTOTAL',FloatToStr(New - Anc));
                end;
        end;
end;

Function TOF_PGAUGMGLOBALE.FormatageDouble(Chaine : String) : Double;
var Longueur,Indice : Integer;
    St : String;
begin
         If Chaine = '' then
         begin
              Result := 0;
              Exit;
         end;
        longueur:=Length (Chaine);
        indice:=1;
        repeat
        if (Chaine [Indice]<>' ') then
        St := St + Chaine [Indice];
        inc (Indice);
        until (Indice=Longueur+1);
        result := StrToFloat(St);
end;

Function TOF_PGAUGMGLOBALE.ArrondiAugm(Montant : Double) : Double;
var Calcul : Double;
begin
     If IArrondiAugm >= 0 then
     begin
          If PrecisionArrondi = 5 then Calcul := 5 * (Arrondi((Montant/5),IArrondiAugm))
          else Calcul := Arrondi(Montant,IArrondiAugm);
     end
     else
     begin
          Calcul := 10 * (Arrondi((Montant/10),0));
     end;
     Result := Calcul;
end;

procedure TOF_PGAUGM_XLS.OnUpdate;
begin
     TFVierge(Ecran).Retour := GetControlText('FICHIER');
end;

procedure TOF_PGAUGM_XLS.OnArgument;
begin
     TFVierge(Ecran).Retour := '';
end;

Initialization
  registerclasses ( [ TOF_PGAUGMGLOBALE,TOF_PGAUGM_XLS ] ) ;
end.

