{***********UNITE*************************************************
Auteur  ...... : EV5
Créé le ...... : xx/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFVBPNEWVAL ()
Mots clefs ... : TOF;QUFVBPNEWVAL
*****************************************************************}
Unit QUFVBPNEWVAL_TOF ;

Interface

Uses UTOF;

Type
  TOF_QUFVBPNEWVAL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    procedure BCalculOnClick(Sender: TObject);
  end ;


var AxeDate,NbDates,numnoeud,NodeNiveau:integer;
    LastLevel:boolean;
    FonctionOrig,CodeSession,CodeAxe,LibAxe,ValeurCodeAxe,ValeurLibAxe : string;
    CodeAxePrec,LibAxePrec,ValeurLibAxePrec:string;
    DatesSession,NumDateOK:string ;
    TabDatesNewVal : array of string;


Implementation

Uses  Controls,Classes,HEnt1,
      {$IFNDEF EAGLCLIENT}
      {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
      {$else}eMul,uTob,
      {$ENDIF}
      sysutils,ComCtrls,HCtrls,HMsgBox,vierge,Uutil,HTB97,BPBasic,BPUtil,BPFctSession;

procedure TOF_QUFVBPNEWVAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPNEWVAL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPNEWVAL.OnUpdate ;
var DatesCheckedNum,DatesChecked,ValuesAxe:string;
NbDateSelected,i,rep:integer;
begin
  Inherited ;
  NbDateSelected:=0;

  for i := 1 to 10 do
  begin
    if i=AxeDate then
    begin
      if ValuesAxe='' then ValuesAxe:='DATES' else ValuesAxe := ValuesAxe + '|' + 'DATES'
    end
    else
    begin
      if THValComboBox(GetControl('COMBOBOXAXE'+IntTostr(i))).value <>''
      then
      begin
        if ValuesAxe='' then ValuesAxe:=THValComboBox(GetControl('COMBOBOXAXE'+IntTostr(i))).value
        else ValuesAxe := ValuesAxe + '|' + THValComboBox(GetControl('COMBOBOXAXE'+IntTostr(i))).value
      end
      else
      begin
        if i<AxeDate then
        begin
          if ValuesAxe='' then ValuesAxe:='SALARIE' else ValuesAxe := ValuesAxe + '|' + 'SALARIE'
        end
        else
        begin
          if ValuesAxe='' then ValuesAxe:='VIDE' else ValuesAxe := ValuesAxe + '|' + 'VIDE'
        end
      end
    end
  end;


  DatesCheckedNum:=THMultiValComboBox(GetControl('MCOMBOBOXAXE'+IntToStr(AxeDate))).Value;

  if ((DatesCheckedNum='') AND (THMultiValComboBox(GetControl('MCOMBOBOXAXE'+IntToStr(AxeDate))).Text=TraduireMemoire('<<Tous>>'))) then
  begin
    For i:=0 to Length(TabDatesNewVal)-1 do
    begin
      if i=0 then DatesChecked:=TabDatesNewVal[i]
      else DatesChecked:=DatesChecked + '|' + TabDatesNewVal[i];
    end;
    NbDateSelected:=Length(TabDatesNewVal);
  end
  else
  begin
    While DatesCheckedNum<>'' do
    begin
      if DatesChecked='' then DatesChecked:=TabDatesNewVal[StrToInt(ReadTokenSt(DatesCheckedNum))-1]
      else DatesChecked:=DatesChecked + '|' + TabDatesNewVal[StrToInt(ReadTokenSt(DatesCheckedNum))-1];
      NbDateSelected:=NbDateSelected+1;
    end;
  end;

  if NbDateSelected = 0 then
  begin
    HShowmessage('1;Création de nouveaux salariés;Aucune Date sélectionnée.;I;O;O;O', '', '');
    SetControlText('EdtFerme','NON');
    TFVierge(ecran).Retour := '';
  end
  else if THValComboBox(GetControl('COMBOTYPECONTRAT')).ItemIndex=-1 then
  begin
    HShowmessage('1;Création de nouveaux salariés;Veuillez sélectionner le type de contrat.;I;O;O;O', '', '');
    SetControlText('EdtFerme','NON');
    TFVierge(ecran).Retour := '';
  end
  else
  begin
    SetControlText('EdtFerme','OUI');
    rep := HShowmessage('1;Création de nouveaux salariés;Vous allez créer '
                          +GetControlText('QBS_NBVALAFF')+ ' salariés.'+
                          #13#10 +' Etes-vous sûr ?;Q;YN;N;N', '', '');

    if rep = mrYes
    then TFVierge(ecran).Retour:='DATESCHECKED='+DatesChecked+';SAISIE1='+GetControlText('EDTSAISIEEVOL1')+
       ';SAISIE2='+GetControlText('EDTSAISIEEVOL2')+';SAISIE3='+GetControlText('EDTSAISIEEVOL3')+
       ';SAISIE4='+GetControlText('EDTSAISIEEVOL4')+';SAISIE5='+GetControlText('EDTSAISIEEVOL5')+
       ';SAISIE6='+GetControlText('EDTSAISIEEVOL6')+';SAISIEQTE='+GetControlText('EDTSAISIQTE')+
       ';NBNEWVAL='+GetControlText('QBS_NBVALAFF')+';VALEURSAXES='+ValuesAxe+';TYPESAL='+THValComboBox(GetControl('COMBOTYPECONTRAT')).value
    else
    begin
      TFVierge(ecran).Retour := '';
      SetControlText('EdtFerme','NON');
    end;
  end
end ;

procedure TOF_QUFVBPNEWVAL.OnLoad ;
var Q:TQuery;
i,j:integer;
DateCur,DateDeb,DateFin:TDateTime;
begin
  Inherited ;
  dateDebFinSesssion(codeSession,DateDeb,DateFin);
  SetLength(TabDatesNewVal,DonneNbIntervalleMailleDateDebDateFin('4',DateDeb,DateFin));
  Q:=OpenSQL('SELECT QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4,QBS_CODEAXES5,QBS_CODEAXES6,'
             +'QBS_CODEAXES7,QBS_CODEAXES8,QBS_CODEAXES9,QBS_CODEAXES10 FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+CodeSession+'"',true);
  for i := 1 to 10 do
  begin
    if Q.Fields[i-1].AsString<> '' then
    begin
      SetControlText('THE_CODEAXES'+IntToStr(i),DonneLibelleAxe(Q.Fields[i-1].AsString));
      if Q.Fields[i-1].AsString = '011' then SetControlEnabled('COMBOBOXAXE'+IntToStr(i),FALSE)
      else SetControlProperty('COMBOBOXAXE'+IntToStr(i),'DATATYPE',ChercheTabletteValeurAxe(Q.Fields[i-1].AsString));
      THValComboBox(GetControl('COMBOBOXAXE'+IntToStr(i))).ItemIndex :=0;
    end
    else
    begin
      SetControlText('THE_CODEAXES'+IntToStr(i),TraduireMemoire('Dates'));
      SetControlVisible('COMBOBOXAXE'+IntToStr(i),FALSE);
      SetControlVisible('MCOMBOBOXAXE'+IntToStr(i),TRUE);
      DateCur := DateDeb;
      NbDates:=1;
      While DateCur<DateFin do
      begin
        THMultiValComboBox(GetControl('MCOMBOBOXAXE'+IntToStr(i))).Items.Add(DateTimeToStr(DateCur));
        THMultiValComboBox(GetControl('MCOMBOBOXAXE'+IntToStr(i))).Values.Add(MetZero(IntToStr(NbDates),3));
        TabDatesNewVal[NbDates-1]:=DateTimeToStr(DateCur);
        AxeDate:=i;
        DateCur:=DateCur+32;
        DateCur:=DEBUTDEMOIS(DateCur);
        NbDates:=NbDates+1;
      end;
      break;
    end;
  end;

  if i<>10 then
  begin
    for j := i+1 to 11 do
    begin
      SetControlVisible('THL_CODEAXES'+IntToStr(j),FALSE);
      SetControlVisible('THE_CODEAXES'+IntToStr(j),FALSE);
      SetControlVisible('COMBOBOXAXE'+IntToStr(j),FALSE);
    end;
  end
  else
  begin
    SetControlText('THE_CODEAXES'+IntToStr(i+1),TraduireMemoire('Dates'));
    SetControlProperty('COMBOBOXAXE'+IntToStr(i+1),'DATATYPE',ChercheTabletteValeurAxe(Q.Fields[i-1].AsString));
    SetControlVisible('COMBOBOXAXE'+IntToStr(i+1),FALSE);
    SetControlVisible('MCOMBOBOXAXE'+IntToStr(i+1),TRUE);
    DateCur := DateDeb;
    NbDates:=1;
    While DateCur<DateFin do
    begin
      THMultiValComboBox(GetControl('MCOMBOBOXAXE'+IntToStr(i+1))).Items.Add(DateTimeToStr(DateCur));
      THMultiValComboBox(GetControl('MCOMBOBOXAXE'+IntToStr(i+1))).Values.Add(MetZero(IntToStr(NbDates),3));
      TabDatesNewVal[NbDates-1]:=DateTimeToStr(DateCur);
      AxeDate:=i+1;
      DateCur:=DateCur+32;
      DateCur:=DEBUTDEMOIS(DateCur);
      NbDates:=NbDates+1;
    end;
  end;

  Ferme(Q);

end ;

procedure TOF_QUFVBPNEWVAL.OnArgument (S : String ) ;
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

  TToolBarButton97(GetControl('BCALCUL')).OnClick := BCalculOnClick;



//  numnoeud:=StrToInt(TrouveArgument(S,'NUMNOEUD',''));
//  NodeNiveau:=StrToInt(TrouveArgument(S,'NIVEAU',''));
//  FonctionOrig:=TrouveArgument(S,'FONCTION','');
//  CodeAxe := TrouveArgument(S,'CODEAXE','');
//  ValeurCodeAxe := TrouveArgument(S,'VALEURCODEAXE','');
//  ValeurLibAxe := DonneLibelleValeurAxe(CodeAxe,ValeurCodeAxe);
//  CodeAxePrec := TrouveArgument(S,'CODEAXEPREC','');

end ;

procedure TOF_QUFVBPNEWVAL.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPNEWVAL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPNEWVAL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPNEWVAL.BCalculOnClick(Sender: TObject);
var CodeSQLAxe:string;
Q:TQuery;
i:integer;
begin
  for i:=1 to AxeDate-2 do
  begin
    if CodeSQLAxe='' then CodeSQLAxe:='QBR_VALAXENIV'+IntToStr(i)+'="'+THValComboBox(GetControl('COMBOBOXAXE'+IntToStr(i))).Value+'"'
    else CodeSQLAxe:=CodeSQLAxe+ ' AND QBR_VALAXENIV'+IntToStr(i)+'="'+THValComboBox(GetControl('COMBOBOXAXE'+IntToStr(i))).Value+'"';
  end;

  Q:=OpenSQL('SELECT AVG(QBR_QTEC),AVG(QBR_OP1),AVG(QBR_OP2),AVG(QBR_OP3),AVG(QBR_OP4),AVG(QBR_OP5),AVG(QBR_OP6) '+
             'FROM QBPARBRE WHERE QBR_CODESESSION="'+CodeSession+'" AND QBR_NIVEAU="'+IntToStr(AxeDate)+'" AND '+CodeSQLAxe,true);
  if not Q.Eof then
  begin
    THEdit(GetControl('EDTSAISIQTE')).Text := Q.Fields[0].AsString;
    THEdit(GetControl('EDTSAISIEEVOL1')).Text := Q.Fields[1].AsString;
    THEdit(GetControl('EDTSAISIEEVOL2')).Text := Q.Fields[2].AsString;
    THEdit(GetControl('EDTSAISIEEVOL3')).Text := Q.Fields[3].AsString;
    THEdit(GetControl('EDTSAISIEEVOL4')).Text := Q.Fields[4].AsString;
    THEdit(GetControl('EDTSAISIEEVOL5')).Text := Q.Fields[5].AsString;
    THEdit(GetControl('EDTSAISIEEVOL6')).Text := Q.Fields[6].AsString;
  end;
  Ferme(Q);
end;




Initialization
  registerclasses ( [ TOF_QUFVBPNEWVAL ] ) ;
end.

