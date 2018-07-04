{***********UNITE*************************************************
Auteur  ...... : EV5
Créé le ...... : xx/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFVBPEVOLUTION2 ()
Mots clefs ... : TOF;QUFVBPEVOLUTION2
*****************************************************************}
Unit QUFVBPEVOLUTION2_TOF ;

Interface

Uses UTOF,uTob;

Type
  TOF_QUFVBPEVOLUTION2 = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    function AffecteMenus(T : TOB) : integer;
    private
    public
    procedure ComboBox_OnChange(Sender:TObject);
    procedure NbValAff_OnChange(Sender:TObject);
  end ;


var AxeDate,NbDates,numnoeud,NodeNiveau,Axe:integer;
    LastLevel:boolean;
    FonctionOrig,CodeSession,CodeAxe,LibAxe,ValeurCodeAxe,ValeurLibAxe : string;
    CodeAxePrec,LibAxePrec,ValeurLibAxePrec:string;
    DatesSession,NumDateOK:string ;
    TabDatesNewVal : array of string;
    TabValuesAxe,TabCodeAxe : array[1..10] of string;
    NameCurrentAxe:string;
    TobArbre:Tob;
    DateDeb,DateFin:TDateTime;

Implementation

Uses  Controls,Classes,HEnt1,Graphics,
      {$IFNDEF EAGLCLIENT}
      {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
      {$else}eMul, {$ENDIF}
      sysutils,ComCtrls,HCtrls,StrUtils,
      HMsgBox,vierge,Uutil,BPBasic,BPUtil,BPFctSession;

procedure TOF_QUFVBPEVOLUTION2.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPEVOLUTION2.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPEVOLUTION2.OnUpdate ;
var AxeKO,DatesCheckedNum,DatesChecked,ValuesAxe:string;
i,rep,nbdates:integer;
okval:boolean;
begin
  Inherited ;

  AxeKO:='';
  nbdates:=0;
  for i := 1 to StrToInt(GetControlText('NBVALAFF')) do
  begin
    if i=AxeDate then
    begin
      if ValuesAxe='' then ValuesAxe:='DATES' else ValuesAxe := ValuesAxe + '|' + 'DATES'
    end
    else if ((i+1=AxeDate) OR (i=StrToInt(GetControlText('NBVALAFF')))) then
    begin
      if THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntTostr(i),2))).value='' then
      begin
        if THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntTostr(i),2))).text='' then
        begin
          AxeKO:=GetControlText('THE_CODEAXES'+IntToStr(i));
          break
        end;
//        if ValuesAxe='' then ValuesAxe:=THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntTostr(i),2))).text
//        else ValuesAxe := ValuesAxe + '|' + THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntTostr(i),2))).text
        if ValuesAxe='' then ValuesAxe:=TabValuesAxe[i]
        else ValuesAxe := ValuesAxe + '|' + TabValuesAxe[i];
      end
      else
      begin
        if ValuesAxe='' then ValuesAxe:=THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntTostr(i),2))).value
        else ValuesAxe := ValuesAxe + '|' + THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntTostr(i),2))).value
      end
    end
    else
    begin
      if ValuesAxe='' then ValuesAxe:=THValComboBox(GetControl('COMBOBOXAXE'+MetZero(IntTostr(i),2))).value
      else ValuesAxe := ValuesAxe + '|' + THValComboBox(GetControl('COMBOBOXAXE'+MetZero(IntTostr(i),2))).value
    end
  end;

  for i := StrToInt(GetControlText('NBVALAFF'))+1 to 10 do
  begin
    if ValuesAxe='' then ValuesAxe:='VIDE'
    else ValuesAxe := ValuesAxe + '|VIDE'
  end;

  if AxeKO='' then
  begin
    DatesCheckedNum:=THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntToStr(AxeDate),2))).Value;

    if DatesCheckedNum='' then
    begin
      if THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntToStr(AxeDate),2))).Text='' then
      begin
        AxeKO:=GetControlText('THE_CODEAXES'+IntToStr(AxeDate));
      end
      else if THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntToStr(AxeDate),2))).Text=TraduireMemoire('<<Tous>>') then
      begin
        For i:=0 to Length(TabDatesNewVal)-1 do
        begin
          if i=0 then DatesChecked:=TabDatesNewVal[i]
          else DatesChecked:=DatesChecked + '|' + TabDatesNewVal[i];
          nbdates:=nbdates+1;
        end;
      end
    end
    else
    begin
      While DatesCheckedNum<>'' do
      begin
        if DatesChecked='' then DatesChecked:=TabDatesNewVal[StrToInt(ReadTokenSt(DatesCheckedNum))-1]
        else DatesChecked:=DatesChecked + '|' + TabDatesNewVal[StrToInt(ReadTokenSt(DatesCheckedNum))-1];
        nbdates:=nbdates+1;
      end
    end
  end;

  OkVal:=false;
  if AxeKO= '' then
  begin
    for i:=1 to 7 do
    begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
      if i <> 7 then
      begin
        if GetControlText('EDTSAISIEEVOL'+IntToStr(i)) <> '' then begin OkVal:=true ; break end;
        if GetControlText('EDTPRCTEVOL'+IntToStr(i)) <> '' then begin OkVal:=true ; break end;
        if GetControlText('EDTVALEVOL'+IntToStr(i)) <> '' then begin OkVal:=true ; break end;
      end
      else
      begin
        if GetControlText('EDTSAISIQTE') <> '' then begin OkVal:=true ; break end;
        if GetControlText('EDTEVOLQTEPRCT') <> '' then begin OkVal:=true ; break end;
        if GetControlText('EDTEVOLVALQTE') <> '' then begin OkVal:=true ; break end;
      end
    end;
  end;

  if AxeKO<>'' then
  begin
    HShowmessage('1;Modification;L''axe '+AxeKO+' n''est pas renseigné.;I;O;O;O', '', '');
    SetControlText('EdtFerme','NON');
    TFVierge(ecran).Retour := '';
  end
  else if OkVal=false then
  begin
    HShowmessage('1;Modification;Aucune valeur de modification n''est renseignée.;I;O;O;O', '', '');
    SetControlText('EdtFerme','NON');
    TFVierge(ecran).Retour := '';
  end
  else
  begin
    SetControlText('EdtFerme','OUI');
    rep := HShowmessage('1;Modification;Vous allez modifier ' + IntToStr(nbdates) + ' valeurs dates pour tous les'+
                          #13#10 +' sous-niveaux sélectionnés. Etes-vous sûr ?;Q;YN;N;N', '', '');

    if rep = mrYes
    then TFVierge(ecran).Retour:='DATESCHECKED='+DatesChecked+';SAISIE1='+GetControlText('EDTSAISIEEVOL1')+
       ';SAISIE2='+GetControlText('EDTSAISIEEVOL2')+';SAISIE3='+GetControlText('EDTSAISIEEVOL3')+
       ';SAISIE4='+GetControlText('EDTSAISIEEVOL4')+';SAISIE5='+GetControlText('EDTSAISIEEVOL5')+
       ';SAISIE6='+GetControlText('EDTSAISIEEVOL6')+';SAISIEQTE='+GetControlText('EDTSAISIQTE')+
       ';EVOLPRCT1='+GetControlText('EDTPRCTEVOL1')+';EVOLPRCT2='+GetControlText('EDTPRCTEVOL2')+
       ';EVOLPRCT3='+GetControlText('EDTPRCTEVOL3')+';EVOLPRCT4='+GetControlText('EDTPRCTEVOL4')+
       ';EVOLPRCT5='+GetControlText('EDTPRCTEVOL5')+';EVOLPRCT6='+GetControlText('EDTPRCTEVOL6')+
       ';EVOLPRCTQTE='+GetControlText('EDTEVOLQTEPRCT')+';EVOLVAL1='+GetControlText('EDTVALEVOL1')+
       ';EVOLVAL2='+GetControlText('EDTVALEVOL2')+';EVOLVAL3='+GetControlText('EDTVALEVOL3')+
       ';EVOLVAL4='+GetControlText('EDTVALEVOL4')+';EVOLVAL5='+GetControlText('EDTVALEVOL5')+
       ';EVOLVAL6='+GetControlText('EDTVALEVOL6')+';EVOLVALQTE='+GetControlText('EDTEVOLVALQTE')+
       ';NIVEAU='+GetControlText('NBVALAFF')+';VALEURSAXES='+ValuesAxe
    else
    begin
      TFVierge(ecran).Retour := '';
      SetControlText('EdtFerme','NON');
    end;
  end;
end ;

function TOF_QUFVBPEVOLUTION2.AffecteMenus(T : TOB) : integer;
var CodeAxe:string;
begin
  CodeAxe:=TabCodeAxe[StrToInt(AnsiRightStr(NameCurrentAxe,2))];
  if AnsiLeftStr(NameCurrentAxe,1)='M' then
  begin
    if CodeAxe='' then THMultiValComboBox(GetControl(NameCurrentAxe)).Items.Add(T.GetString('QBR_VALEURAXE'))
    else THMultiValComboBox(GetControl(NameCurrentAxe)).Items.Add(DonneLibelleValeurAxe(CodeAxe,T.GetString('QBR_VALEURAXE')));
    THMultiValComboBox(GetControl(NameCurrentAxe)).Values.Add(T.GetString('QBR_VALEURAXE'));
  end
  else
  begin
    if CodeAxe='' then     THValComboBox(GetControl(NameCurrentAxe)).Items.Add(T.GetString('QBR_VALEURAXE'))
    else THValComboBox(GetControl(NameCurrentAxe)).Items.Add(DonneLibelleValeurAxe(CodeAxe,T.GetString('QBR_VALEURAXE')));
    THValComboBox(GetControl(NameCurrentAxe)).Values.Add(T.GetString('QBR_VALEURAXE'));
  end;
  if TabValuesAxe[StrToInt(AnsiRightStr(NameCurrentAxe,2))]=''
  then TabValuesAxe[StrToInt(AnsiRightStr(NameCurrentAxe,2))]:=T.GetString('QBR_VALEURAXE')
  else TabValuesAxe[StrToInt(AnsiRightStr(NameCurrentAxe,2))]:=TabValuesAxe[StrToInt(AnsiRightStr(NameCurrentAxe,2))] + '","' + T.GetString('QBR_VALEURAXE');
  result:=0;
end;

procedure TOF_QUFVBPEVOLUTION2.OnLoad ;
var Q:TQuery;
i:integer;
DateCur:TDateTime;
begin
  dateDebFinSesssion(codeSession,DateDeb,DateFin);
  SetLength(TabDatesNewVal,DonneNbIntervalleMailleDateDebDateFin('4',DateDeb,DateFin));

  Q:=OpenSQL('SELECT QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4,QBS_CODEAXES5,QBS_CODEAXES6,'
             +'QBS_CODEAXES7,QBS_CODEAXES8,QBS_CODEAXES9,QBS_CODEAXES10 FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+CodeSession+'"',true);
  for i := 1 to 10 do
  begin
    if Q.Fields[i-1].AsString<> '' then
    begin
      SetControlVisible('THL_CODEAXES'+IntToStr(i),TRUE);
      SetControlVisible('THE_CODEAXES'+IntToStr(i),TRUE);
      SetControlVisible('COMBOBOXAXE'+MetZero(IntToStr(i),2),FALSE);
      SetControlEnabled('COMBOBOXAXE'+MetZero(IntToStr(i),2),FALSE);
      SetControlVisible('MCOMBOBOXAXE'+MetZero(IntToStr(i),2),TRUE);
      if i=1 then SetControlEnabled('MCOMBOBOXAXE'+MetZero(IntToStr(i),2),TRUE)
      else SetControlEnabled('MCOMBOBOXAXE'+MetZero(IntToStr(i),2),FALSE);
      SetControlText('THE_CODEAXES'+IntToStr(i),DonneLibelleAxe(Q.Fields[i-1].AsString));
      THValComboBox(GetControl('COMBOBOXAXE'+MetZero(IntToStr(i),2))).ItemIndex :=0;
      TabCodeAxe[i]:=Q.Fields[i-1].AsString;
    end
    else
    begin
      SetControlVisible('THL_CODEAXES'+IntToStr(i),TRUE);
      SetControlVisible('THE_CODEAXES'+IntToStr(i),TRUE);
      THEdit(GetControl('THE_CODEAXES'+IntToStr(i))).Color := clBtnHighlight ;
      SetControlVisible('COMBOBOXAXE'+MetZero(IntToStr(i),2),FALSE);
      SetControlVisible('MCOMBOBOXAXE'+MetZero(IntToStr(i),2),TRUE);
      SetControlText('THE_CODEAXES'+IntToStr(i),TraduireMemoire('Dates'));
      DateCur := DateDeb;
      NbDates:=1;
      While DateCur<DateFin do
      begin
        THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntToStr(i),2))).Items.Add(DateTimeToStr(DateCur));
        THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntToStr(i),2))).Values.Add(MetZero(IntToStr(NbDates),3));
        TabDatesNewVal[NbDates-1]:=DateTimeToStr(DateCur);
        AxeDate:=i;
        DateCur:=DateCur+32;
        DateCur:=DEBUTDEMOIS(DateCur);
        NbDates:=NbDates+1;
      end;
      break;
    end;
  end;

  if i<>10 then THSpinEdit(GetControl('NBVALAFF')).MaxValue:=i-1
  else
  begin
    SetControlVisible('THL_CODEAXES'+IntToStr(i+1),TRUE);
    SetControlVisible('THE_CODEAXES'+IntToStr(i+1),TRUE);
    THEdit(GetControl('THE_CODEAXES'+IntToStr(i+1))).Color := clBtnHighlight ;
    SetControlVisible('COMBOBOXAXE'+MetZero(IntToStr(i+1),2),FALSE);
    SetControlVisible('MCOMBOBOXAXE'+MetZero(IntToStr(i+1),2),TRUE);
    SetControlText('THE_CODEAXES'+IntToStr(i+1),TraduireMemoire('Dates'));
    DateCur := DateDeb;
    NbDates:=1;
    While DateCur<DateFin do
    begin
      THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntToStr(i+1),2))).Items.Add(DateTimeToStr(DateCur));
      THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntToStr(i+1),2))).Values.Add(MetZero(IntToStr(NbDates),3));
      TabDatesNewVal[NbDates-1]:=DateTimeToStr(DateCur);
      AxeDate:=i+1;
      DateCur:=DateCur+32;
      DateCur:=DEBUTDEMOIS(DateCur);
      NbDates:=NbDates+1;
    end;
    THSpinEdit(GetControl('NBVALAFF')).MaxValue:=i;
  end;

  Ferme(Q);

  TobArbre := Tob.Create ( 'Arbre de la session', nil, -1 );

  TobArbre.LoadDetailDBFromSQL('QBPARBRE','SELECT QBR_VALEURAXE,QBR_NIVEAU,'+
  'QBR_VALAXENIV1,QBR_VALAXENIV2,QBR_VALAXENIV3,QBR_VALAXENIV4,QBR_VALAXENIV5,QBR_VALAXENIV6,QBR_VALAXENIV7,'+
  'QBR_VALAXENIV8,QBR_VALAXENIV9,QBR_VALAXENIV10 FROM QBPARBRE '+
  'WHERE  QBR_CODESESSION="'+CodeSession+'"');

  Axe:=1;
  NameCurrentAxe:='MCOMBOBOXAXE01';
  TabValuesAxe[1]:='';
  TobArbre.ParcoursTraitement(['QBR_NIVEAU'],['1'],false, AffecteMenus);
end ;

procedure TOF_QUFVBPEVOLUTION2.OnArgument (S : String ) ;
var i,ValeurAffiche,NbValAff:integer;
TabValAff,TabLibelle:array [0..7] of hString;
begin
  Inherited ;
  CodeSession:=TrouveArgument(S,'SESSION','');
  ValeurAffiche:=DonneValeurAffiche(codesession);
  for i:=1 to 7 do THTabSheet(GetControl('TAB'+IntToStr(i))).tabVisible:=false;
  Case ContextBP of
    0,1 : begin //Mode-GC
          case ValeurAffiche of
            1 : begin
                  TabValAff[0]:='PTC';
                  THTabSheet(GetControl('TAB1')).tabVisible:=true;
                end;
            2 : begin
                  TabValAff[0]:='QTE';
                  THTabSheet(GetControl('TAB2')).tabVisible:=true;
                end;
             3 : begin
                  TabValAff[0]:='PHT';
                  THTabSheet(GetControl('TAB3')).tabVisible:=true;
                end;
            4 : begin
                  TabValAff[0]:='UTC';
                  THTabSheet(GetControl('TAB4')).tabVisible:=true;
                end;
            5 : begin
                  TabValAff[0]:='UHT';
                  THTabSheet(GetControl('TAB5')).tabVisible:=true;
                end;
            6 : begin
                  TabValAff[0]:='PAH';
                  THTabSheet(GetControl('TAB6')).tabVisible:=true;
                end;
            7 : begin
                  TabValAff[0]:='MAR';
                  THTabSheet(GetControl('TAB7')).tabVisible:=true;
                end
          end
        end;
    2 : begin //Compta
          case ValeurAffiche of
            1 : begin
                  TabValAff[0]:='DC1';
                  THTabSheet(GetControl('TAB1')).tabVisible:=true;
                  THTabSheet(GetControl('TAB1')).Caption := 'Débit - Crédit';
                end;
            3 : begin
                  TabValAff[0]:='CD1';
                  THTabSheet(GetControl('TAB3')).tabVisible:=true;
                  THTabSheet(GetControl('TAB3')).Caption := 'Crédit - Débit';
                end;
          end
        end;
    3 : begin //Paie
          LibValAff(codeSession,TabLibelle);
          NbValAff := StrToInt(TabLibelle[0]);
          for i:=1 to NbValAff do
          begin               //Tab1 = Val2, Tab2 = Val1, Tab3 = Val3...
            if i = 1 then THTabSheet(GetControl('TAB2')).tabVisible := true
            else if i = 2 then THTabSheet(GetControl('TAB1')).tabVisible := true
            else THTabSheet(GetControl('TAB'+IntToStr(i))).tabVisible := true;
            if i = 1 then THTabSheet(GetControl('TAB2')).Caption := TabLibelle[i]
            else if i = 2 then THTabSheet(GetControl('TAB1')).Caption := TabLibelle[i]
            else THTabSheet(GetControl('TAB'+IntToStr(i))).Caption := TabLibelle[i]
          end;
          if NbValAff > 1 then
          begin
            TPageControl(getControl( 'PAGE')).Pages[1].PageIndex := 0;
            TPageControl(getControl( 'PAGE')).ActivePage := THTabSheet(GetControl('TAB2'));
          end;
        end
  end; //CASE

  THMultiValComboBox(GetControl('MCOMBOBOXAXE01')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE02')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE03')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE04')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE05')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE06')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE07')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE08')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE09')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE10')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE11')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('COMBOBOXAXE01')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('COMBOBOXAXE02')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('COMBOBOXAXE03')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('COMBOBOXAXE04')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('COMBOBOXAXE05')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('COMBOBOXAXE06')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('COMBOBOXAXE07')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('COMBOBOXAXE08')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('COMBOBOXAXE09')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('COMBOBOXAXE10')).OnChange := ComboBox_OnChange;
  THMultiValComboBox(GetControl('COMBOBOXAXE11')).OnChange := ComboBox_OnChange;

  THSpinEdit(GetControl('NBVALAFF')).OnChange := NbValAff_OnChange;

end ;

procedure TOF_QUFVBPEVOLUTION2.ComboBox_OnChange(Sender:TObject);
var NameAxe:string;
TabChamp,TabValues: array [1..11] of string;
i:integer;
begin
  NameAxe := TControl( Sender ).name;

  Axe:=StrToInt(AnsiRightStr(NameAxe,2));
  if ((Axe=AxeDate) OR (Axe+1=AxeDate)) then exit;
  if Axe=StrToInt(GetControlText('NBVALAFF')) then exit;

  if ((Axe+1=StrToInt(GetControlText('NBVALAFF'))) OR (Axe+1=AxeDate) OR (Axe+1=AxeDate-1)) then
  begin
    NameCurrentAxe:='MCOMBOBOXAXE'+MetZero(IntToStr(Axe+1),2);
    THMultiValComboBox(GetControl(NameCurrentAxe)).Values.Clear;
    THMultiValComboBox(GetControl(NameCurrentAxe)).Items.Clear;
  end
  else
  begin
    NameCurrentAxe:='COMBOBOXAXE'+MetZero(IntToStr(Axe+1),2);
    THValComboBox(GetControl(NameCurrentAxe)).Values.Clear;
    THValComboBox(GetControl(NameCurrentAxe)).Items.Clear;
  end;

  for i:=1 to Axe do
  begin
    TabChamp[i]:='QBR_VALAXENIV'+IntToStr(i);
    if ((i=AxeDate) OR (i=AxeDate-1)) then TabValues[i]:=GetControlText('MCOMBOBOXAXE'+MetZero(IntToStr(i),2))
    else TabValues[i]:=GetControlText('COMBOBOXAXE'+MetZero(IntToStr(i),2));
  end;

  For i:=1 to 11 do
  begin
    if ((TabValues[i]='') OR (TabValues[i]=TraduireMemoire('<<Tous>>'))) then
    begin
      TabChamp[i]:='';
      TabValues[i]:=''
    end;
  end;

  TabValuesAxe[Axe+1]:='';
  TobArbre.ParcoursTraitement(['QBR_NIVEAU',TabChamp[1],TabChamp[2],TabChamp[3],TabChamp[4],TabChamp[5]
  ,TabChamp[6],TabChamp[7],TabChamp[8],TabChamp[9],TabChamp[10],TabChamp[11]],[Axe+1,TabValues[1]
  ,TabValues[2],TabValues[3],TabValues[4],TabValues[5],TabValues[6],TabValues[7],TabValues[8]
  ,TabValues[9],TabValues[10],TabValues[11]],false, AffecteMenus);

  if ((Axe+1=StrToInt(GetControlText('NBVALAFF'))) OR (Axe+1=AxeDate) OR (Axe+1=AxeDate-1))
  then THMultiValComboBox(GetControl(NameCurrentAxe)).Text := ''
  else THValComboBox(GetControl(NameCurrentAxe)).Text := '';

  if ((Axe+1<>StrToInt(GetControlText('NBVALAFF'))) AND (Axe+1<>AxeDate) AND (Axe+1<>AxeDate-1)) then
  begin
    THValComboBox(GetControl(NameCurrentAxe)).ItemIndex:=0;
    ComboBox_OnChange(THValComboBox(GetControl(NameCurrentAxe)));
  end;

end ;

procedure TOF_QUFVBPEVOLUTION2.NbValAff_OnChange;
var i : integer;
TabChamp,TabValues: array [1..11] of string;
begin
  for i:=1 to StrToInt(GetControlText('NBVALAFF'))-1 do
  begin
    THEdit(GetControl('THE_CODEAXES'+IntToStr(i))).Color := clBtnHighlight ;
    if ((i=AxeDate) OR (i=AxeDate-1)) then
    begin
      SetControlVisible('COMBOBOXAXE'+MetZero(IntToStr(i),2), false);
      SetControlEnabled('COMBOBOXAXE'+MetZero(IntToStr(i),2), false);
      SetControlVisible('MCOMBOBOXAXE'+MetZero(IntToStr(i),2), true);
      SetControlEnabled('MCOMBOBOXAXE'+MetZero(IntToStr(i),2), true);
      TabValues[i]:=GetControlText('MCOMBOBOXAXE'+MetZero(IntToStr(i),2));
    end
    else
    begin
      SetControlVisible('COMBOBOXAXE'+MetZero(IntToStr(i),2), true);
      SetControlEnabled('COMBOBOXAXE'+MetZero(IntToStr(i),2), true);
      SetControlVisible('MCOMBOBOXAXE'+MetZero(IntToStr(i),2), false);
      SetControlEnabled('MCOMBOBOXAXE'+MetZero(IntToStr(i),2), false);
      TabValues[i]:=GetControlText('COMBOBOXAXE'+MetZero(IntToStr(i),2));
    end;
    TabChamp[i]:='QBR_VALAXENIV'+IntToStr(i);
  end;

  THEdit(GetControl('THE_CODEAXES'+GetControlText('NBVALAFF'))).Color := clBtnHighlight ;
  SetControlVisible('COMBOBOXAXE'+MetZero(GetControlText('NBVALAFF'),2), false);
  SetControlEnabled('COMBOBOXAXE'+MetZero(GetControlText('NBVALAFF'),2), false);
  SetControlVisible('MCOMBOBOXAXE'+MetZero(GetControlText('NBVALAFF'),2), true);
  SetControlEnabled('MCOMBOBOXAXE'+MetZero(GetControlText('NBVALAFF'),2), true);

  for i:=StrToInt(GetControlText('NBVALAFF'))+1 to THSpinEdit(GetControl('NBVALAFF')).MaxValue do
  begin
    THEdit(GetControl('THE_CODEAXES'+IntToStr(i))).Color := cl3DLight;
    SetControlVisible('COMBOBOXAXE'+MetZero(IntToStr(i),2), false);
    SetControlEnabled('COMBOBOXAXE'+MetZero(IntToStr(i),2), false);
    SetControlVisible('MCOMBOBOXAXE'+MetZero(IntToStr(i),2), true);
    SetControlEnabled('MCOMBOBOXAXE'+MetZero(IntToStr(i),2), false);
    TabChamp[i]:='';
    TabValues[i]:='';
  end;

  For i:=1 to 11 do
  begin
    if ((TabValues[i]='') OR (TabValues[i]=TraduireMemoire('<<Tous>>'))) then
    begin
      TabChamp[i]:='';
      TabValues[i]:='';
    end;
  end;

  THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(GetControlText('NBVALAFF'),2))).Values.Clear;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(GetControlText('NBVALAFF'),2))).Items.Clear;
  THMultiValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(GetControlText('NBVALAFF'),2))).Text:='';
  THValComboBox(GetControl('COMBOBOXAXE'+MetZero(GetControlText('NBVALAFF'),2))).Values.Clear;
  THValComboBox(GetControl('COMBOBOXAXE'+MetZero(GetControlText('NBVALAFF'),2))).Items.Clear;
  THValComboBox(GetControl('COMBOBOXAXE'+MetZero(GetControlText('NBVALAFF'),2))).Text:='';

  if StrToInt(GetControlText('NBVALAFF'))=1 then
  begin
    NameCurrentAxe:='MCOMBOBOXAXE01';
    TabValuesAxe[1]:='';
    Axe:=1;
  end
  else
  begin
    if ((StrToInt(GetControlText('NBVALAFF'))-1=AxeDate) OR (StrToInt(GetControlText('NBVALAFF'))-1=AxeDate-1))
    then NameCurrentAxe:='MCOMBOBOXAXE'+MetZero(IntToStr(StrToInt(GetControlText('NBVALAFF'))-1),2)
    else NameCurrentAxe:='COMBOBOXAXE'+MetZero(IntToStr(StrToInt(GetControlText('NBVALAFF'))-1),2);
    TabValuesAxe[StrToInt(GetControlText('NBVALAFF'))-1]:='';
    Axe:=StrToInt(GetControlText('NBVALAFF'))-1;
  end;

  TobArbre.ParcoursTraitement(['QBR_NIVEAU',TabChamp[1],TabChamp[2],TabChamp[3],TabChamp[4],TabChamp[5]
  ,TabChamp[6],TabChamp[7],TabChamp[8],TabChamp[9],TabChamp[10],TabChamp[11]],[Axe,TabValues[1]
  ,TabValues[2],TabValues[3],TabValues[4],TabValues[5],TabValues[6],TabValues[7],TabValues[8]
  ,TabValues[9],TabValues[10],TabValues[11]],false, AffecteMenus);

  if StrToInt(GetControlText('NBVALAFF'))<>1 then
  begin
    if ((StrToInt(GetControlText('NBVALAFF'))-1<>AxeDate) AND (StrToInt(GetControlText('NBVALAFF'))-1<>AxeDate-1)) then
    begin
      THValComboBox(GetControl('COMBOBOXAXE'+MetZero(IntToStr(StrToInt(GetControlText('NBVALAFF'))-1),2))).ItemIndex:=0;
      ComboBox_OnChange(THValComboBox(GetControl('COMBOBOXAXE'+MetZero(IntToStr(StrToInt(GetControlText('NBVALAFF'))-1),2))));
    end
    else
    begin
      THValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntToStr(StrToInt(GetControlText('NBVALAFF'))-1),2))).SelectAll;
      ComboBox_OnChange(THValComboBox(GetControl('MCOMBOBOXAXE'+MetZero(IntToStr(StrToInt(GetControlText('NBVALAFF'))-1),2))));
    end
  end;
end;

procedure TOF_QUFVBPEVOLUTION2.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPEVOLUTION2.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPEVOLUTION2.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_QUFVBPEVOLUTION2 ] ) ;
end.

