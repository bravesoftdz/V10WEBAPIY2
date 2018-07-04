{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 13/01/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGAUGMDETAILSAL ()
Mots clefs ... : TOF;PGAUGMDETAILSAL
*****************************************************************}
Unit UTofPGAugmDetailSalarie ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     uTob, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     ParamSoc,
     UTOF ; 

Type
  TOF_PGAUGMDETAILSAL = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    private
    IArrondiAugm,PrecisionArrondi,PctAugmDec : Integer;
    Function ArrondiAugm(Montant : Double) : Double;
  end ;

Implementation

procedure TOF_PGAUGMDETAILSAL.OnArgument (S : String ) ;
var Annee,Salarie,Etat : String;
    Q : TQuery;
    FixeAv,PctF,PctV,FixeAp,VarAv,VarAp,Pct,PctFS,PctVS,PctFR,PctVR : Double;
    CalcFixe,CalcVar : Double;
    DateS,DateR,DateD,DateF,DateV : TDateTime;
    IntFixe,IntVar : Boolean;
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
  Salarie := ReadTokenPipe(S,';');
  Annee := ReadTokenPipe(S,';');
  Q := OpenSQL('SELECT * FROM BUDGETPAIE '+
  'WHERE PBG_SALARIE="'+Salarie+'" AND PBG_TYPEBUDG="AUG" '+
  'AND PBG_ANNEE="'+Annee+'"',True);
  If Not Q.Eof then
  begin
       FixeAv := Q.FindField('PBG_FIXEAV').AsFloat;
       PctF := Q.FindField('PBG_PCTFIXE').AsFloat;
       PctV := Q.FindField('PBG_PCTVARIABLE').AsFloat;
       FixeAp := Q.FindField('PBG_FIXEAP').AsFloat;
       VarAv := Q.FindField('PBG_VARIABLEAV').AsFloat;
       VarAp := Q.FindField('PBG_VARIABLEAP').AsFloat;
       PctFS := Q.FindField('PBG_PCTFIXESAISIE').AsFloat;
       PctVS := Q.FindField('PBG_PCTVARSAISIE').AsFloat;
       PctFR := Q.FindField('PBG_PCTFIXERESP').AsFloat;
       PctVR := Q.FindField('PBG_PCTVARRESP').AsFloat;
       Etat := Q.FindField('PBG_ETATINTAUGM').AsString;
       DateS := Q.FindField('PBG_DATESAISIE').AsDateTime;
       DateR := Q.FindField('PBG_VALIDELE').AsDateTime;
       DateD := Q.FindField('PBG_ACCEPTELE').AsDateTime;
       IntFixe := (Q.FindField('PBG_INTFIXE').AsString = 'X');
       IntVar := (Q.FindField('PBG_INTVARIABLE').AsString = 'X');
       DateF := Q.FindField('PBG_DATEFIXE').AsDateTime;
       DateV := Q.FindField('PBG_DATEVAR').AsDateTime;
       //SAISIE
       SetControltext('FIXEAV',FloatToStr(FixeAv));
       CalcFixe := ArrondiAugm(FixeAv + (PctFS * FixeAv/100));
       SetControltext('FIXEAP',FloatToStr(CalcFixe));
       SetControltext('PCTF',FloatToStr(PctFS));
       SetControltext('PCTV',FloatToStr(PctVS));
       SetControltext('VARAV',FloatToStr(VarAv));
       CalcVar := ArrondiAugm(VarAv + (PctVS * VarAv/100));
       SetControltext('VARAP',FloatToStr(CalcVar));
       SetControltext('TOTAV',FloatToStr(FixeAV + varAv));
       If (FixeAV + varAv) <> 0 then Pct := (((CalcFixe + CalcVar) - (FixeAV + varAv)) / (FixeAV + varAv)) * 100
       else Pct := 0;
       SetControltext('PCTT',FloatToStr(Arrondi(pct,PctAugmDec)));
       SetControltext('TOTAP',FloatToStr(CalcFixe + CalcVar));
       SetControltext('DATESAISIE',DateToStr(DateS));
       //RESONSABLE
       If Etat > '002' then
       begin
            SetControltext('FIXEAVR',FloatToStr(FixeAv));
            CalcFixe := ArrondiAugm(FixeAv + (PctFR * FixeAv/100));
            SetControltext('FIXEAPR',FloatToStr(CalcFixe));
            SetControltext('PCTFR',FloatToStr(PctFR));
            SetControltext('PCTVR',FloatToStr(PctVR));
            SetControltext('VARAVR',FloatToStr(VarAV));
            CalcVar := ArrondiAugm(VarAv + (PctVR * VarAv/100));
            SetControltext('VARAPR',FloatToStr(CalcVar));
            SetControltext('TOTAVR',FloatToStr(FixeAV + varAv));
            If (FixeAV + varAv) <> 0 then Pct := (((CalcFixe + CalcVar) - (FixeAV + varAv)) / (FixeAV + varAv)) * 100
            else Pct := 0;
            SetControltext('PCTTR',FloatToStr(Arrondi(pct,PctAugmDec)));
            SetControltext('TOTAPR',FloatToStr(CalcFixe + CalcVar));
            SetControltext('DATERESP',DateToStr(DateR));
            SetControltext('DATEDRH',DateToStr(DateD));
            If Etat > '005' then
            begin
                 If PctFS <> PctFr then SetControlText('ETATRESP','005')
                 else SetControlText('ETATRESP','003');
            end
            else SetControlText('ETATRESP',Etat);
       end
       else
       begin
            SetControlText('ETATRESP','');
       end;

       //DRH
       If Etat > '005' then
       begin
            SetControltext('FIXEAVD',FloatToStr(FixeAv));
            SetControltext('FIXEAPD',FloatToStr(FixeAp));
            SetControltext('PCTFD',FloatToStr(PctF));
            SetControltext('PCTVD',FloatToStr(PctV));
            SetControltext('VARAVD',FloatToStr(VarAV));
            SetControltext('VARAPD',FloatToStr(VarAp));
            SetControltext('TOTAVD',FloatToStr(FixeAV + varAv));
            If (FixeAV + varAv) <> 0 then Pct := (((FixeAP + varAP) - (FixeAV + varAv)) / (FixeAV + varAv)) * 100
            else Pct := 0;
            SetControltext('PCTTD',FloatToStr(Arrondi(pct,PctAugmDec)));
            SetControltext('TOTAPD',FloatToStr(FixeAP + varAP));
            SetControlText('ETATDRH',Etat);
            If IntFixe then
            begin
                 SetControlChecked('CINTEGREF',True);
                 SetControlCaption('CINTEGREF','Salaire fixe intégré dans bulletin le '+DateToStr(DateF));
            end
            else
            begin
                 SetControlChecked('CINTEGREF',False);
                 SetControlCaption('CINTEGREF','Salaire fixe intégré dans bulletin ');
            end;
            If IntVar then
            begin
                 SetControlChecked('CINTEGREVAR',True);
                 SetControlCaption('CINTEGREVAR','Salaire variable intégré dans bulletin le '+DateToStr(DateV));
            end
            else
            begin
                 SetControlChecked('CINTEGREVAR',False);
                 SetControlCaption('CINTEGREVAR','Salaire variable intégré dans bulletin');
            end;
       end
       else
       begin
            SetControlText('ETATDRH','');
       end;
  end;
  Ferme(Q);
  Ecran.Caption := 'Augmentation pour l''année '+Annee+' du salarié '+ RechDom('PGSALARIE',Salarie,False);
  UpdateCaption(Ecran);
end ;

Function TOF_PGAUGMDETAILSAL.ArrondiAugm(Montant : Double) : Double;
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


Initialization
  registerclasses ( [ TOF_PGAUGMDETAILSAL ] ) ;
end.

