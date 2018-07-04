{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 20/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFVBPARBREMODIF ()
Mots clefs ... : TOF;QUFVBPARBREMODIF
*****************************************************************}
Unit QUFVBPARBREMODIF_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$else}
     uTob,
{$ENDIF}
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     UTOF,
     Vierge;

Type
  TOF_QUFVBPARBREMODIF = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
     codeSession,OkCoeff,OkPrev,retourS:hString;
     nivMax,ValeurAffiche:integer;
     TabNoeud:array [0..16] of integer;
     procedure ClickRadio(Sender: TObject);
     procedure ChangeValeurAxe(Sender: TObject);
  end ;

Implementation

uses UUtil,UDatamem,CstCommun,
     BPFctSession,BPBasic,BPFctArbre,BPUtil ;

procedure TOF_QUFVBPARBREMODIF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPARBREMODIF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPARBREMODIF.OnUpdate ;
begin
  Inherited ;
  retourS:='EVOLPRCT1='+FloatToStr(THNumEdit(GetControl('EDTEVOLPRCT1')).value)+
           ';EVOLVAL1='+GetControlText('EDTEVOLVAL1')+
           ';SAISI1='+GetControlText('EDTSAISI1')+
           ';EVOLPRCT2='+FloatToStr(THNumEdit(GetControl('EDTEVOLPRCT2')).value)+
           ';EVOLVAL2='+GetControlText('EDTEVOLVAL2')+
           ';SAISI2='+GetControlText('EDTSAISI2')+           
           ';EVOLPRCT3='+FloatToStr(THNumEdit(GetControl('EDTEVOLPRCT3')).value)+
           ';EVOLVAL3='+GetControlText('EDTEVOLVAL3')+
           ';SAISI3='+GetControlText('EDTSAISI3')+
           ';EVOLPRCT4='+FloatToStr(THNumEdit(GetControl('EDTEVOLPRCT4')).value)+
           ';EVOLVAL4='+GetControlText('EDTEVOLVAL4')+
           ';SAISI4='+GetControlText('EDTSAISI4')+
           ';EVOLPRCT5='+FloatToStr(THNumEdit(GetControl('EDTEVOLPRCT5')).value)+
           ';EVOLVAL5='+GetControlText('EDTEVOLVAL5')+
           ';SAISI5='+GetControlText('EDTSAISI5')+
           ';EVOLPRCT6='+FloatToStr(THNumEdit(GetControl('EDTEVOLPRCT6')).value)+
           ';EVOLVAL6='+GetControlText('EDTEVOLVAL6')+
           ';SAISI6='+GetControlText('EDTSAISI6')+
           ';EVOLPRCTQTE='+FloatToStr(THNumEdit(GetControl('EDTEVOLQTEPRCT')).value)+
           ';EVOLVALQTE='+GetControlText('EDTEVOLVALQTE')+
           ';SAISIQTE='+GetControlText('EDTSAISIQTE')+
           ';SOMMEQTE='+FloatToStr(VALEUR(GetControlText('EDTT1'))+
           VALEUR(GetControlText('EDTT2'))+VALEUR(GetControlText('EDTT3'))+
           VALEUR(GetControlText('EDTT4'))+VALEUR(GetControlText('EDTT5'))+
           VALEUR(GetControlText('EDTT6'))+VALEUR(GetControlText('EDTT7'))+
           VALEUR(GetControlText('EDTT8'))+VALEUR(GetControlText('EDTT9'))+
           VALEUR(GetControlText('EDTT10'))+VALEUR(GetControlText('EDTT11'))+
           VALEUR(GetControlText('EDTT12'))+VALEUR(GetControlText('EDTT13'))+
           VALEUR(GetControlText('EDTT14'))+VALEUR(GetControlText('EDTT15'))+
           VALEUR(GetControlText('EDTT16'))+VALEUR(GetControlText('EDTT17'))+
           VALEUR(GetControlText('EDTT18'))+VALEUR(GetControlText('EDTT19'))+
           VALEUR(GetControlText('EDTT20')))+';SQLQTE=QBR_QTET1="'+GetControlText('EDTT1')+
           '",QBR_QTET2="'+GetControlText('EDTT2')+'",QBR_QTET3="'+GetControlText('EDTT3')+
           '",QBR_QTET4="'+GetControlText('EDTT4')+'",QBR_QTET5="'+GetControlText('EDTT5')+
           '",QBR_QTET6="'+GetControlText('EDTT6')+'",QBR_QTET7="'+GetControlText('EDTT7')+
           '",QBR_QTET8="'+GetControlText('EDTT8')+'",QBR_QTET9="'+GetControlText('EDTT9')+
           '",QBR_QTET10="'+GetControlText('EDTT10')+'",QBR_QTET11="'+GetControlText('EDTT11')+
           '",QBR_QTET12="'+GetControlText('EDTT12')+'",QBR_QTET13="'+GetControlText('EDTT13')+
           '",QBR_QTET14="'+GetControlText('EDTT14')+'",QBR_QTET15="'+GetControlText('EDTT15')+
           '",QBR_QTET16="'+GetControlText('EDTT16')+'",QBR_QTET17="'+GetControlText('EDTT17')+
           '",QBR_QTET18="'+GetControlText('EDTT18')+'",QBR_QTET19="'+GetControlText('EDTT19')+
           '",QBR_QTET20="'+GetControlText('EDTT20')+'" ';
end ;

procedure TOF_QUFVBPARBREMODIF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPARBREMODIF.OnArgument (S : String ) ;
var Noeud:hString;
    i,NbValAff:integer;
    codeSql,devise:hString;
    Q:TQuery;
    TabCodeAxe:array [0..11] of hString;
    TabValAff,TabLibValAff,TabLibelle:array [0..7] of hString;
    LibelleAxe : hString;
    Enr:TEnreg;
begin
  Inherited ;

  codeSession:=TrouveArgument(S,'SESSION','');
  Noeud:=TrouveArgument(S,'NOEUD','');
  OkCoeff:=TrouveArgument(S,'OKCOEFF','');
  OkPrev:=TrouveArgument(S,'OKPREV','');

  //Cherche niveau Max
  NivMax:=ChercheNivMax(codeSession);
  ChercheCodeAxeSession(codeSession,TabCodeAxe);
  ValeurAffiche:=DonneValeurAffiche(codesession);

  if BPOkOrli then
  //-----------------> ORLI
  begin
    RechercheTabAff(TabValAff,TabLibValAff);
    ChangeListeLibelleColonnes(OkSessionObjectif(codeSession),codeSession,TabValAff,TabLibValAff);

    //onglets visibles
    for i:=1 to 9 do THTabSheet(GetControl('TAB'+IntToStr(i))).tabVisible:=false;
    for i:=0 to 7 do
    if TabValAff[i]<>'' then
    begin
      THTabSheet(GetControl(TabValAff[i])).tabVisible:=true;
      THTabSheet(GetControl(TabValAff[i])).caption:=TabLibValAff[i];
    end;
  end
  //ORLI <-----------------
  else
  begin
    //onglets visibles
    for i:=1 to 8 do THTabSheet(GetControl('TAB'+IntToStr(i))).tabVisible:=false;
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
          end

    end; //CASE
  end;

  if SessionInitPrev(codeSession) then
  begin
    TabCodeAxe[nivMax+1]:='ARTICLE';
    //niv coloris
    if GestionBPColoris then TabCodeAxe[nivMax+2]:='COLORIS';
    //niv fs
    if GestionBPFS then TabCodeAxe[nivMax+3]:='FS';
    //niv mag
    if GestionBPMagasin then TabCodeAxe[nivMax+4]:='MAGASIN';

    NivMax:=ChercheNivMaxSession(codeSession);
    //onglets visibles
    for i:=1 to 9 do THTabSheet(GetControl('TAB'+IntToStr(i))).tabVisible:=false;
    THTabSheet(GetControl('TAB2')).tabVisible:=true;
    THRadioGroup(GetControl('GRP1')).enabled:=false;
  end;

  if SessionDelai(codeSession) then
  begin
    NivMax:=NivMax+1;
    TabCodeAxe[NivMax]:='DELAI';
  end;

  codesql:='';
  for i:=1 to NivMax-1 do
  begin
    if codeSql='' then codesql:=' QBR_VALAXENIV'+IntToStr(i)
    else codesql:=codesql+',QBR_VALAXENIV'+IntToStr(i);
  end;

  if NivMax=1 then codesql:=' QBR_VALAXENIV1';
  codesql:=codesql+',QBR_VALEURAXE ';

  //Cherche code càd le chemin complet codeaxe, valeuraxe
  Q:=MOPenSql('SELECT '+codeSql+',QBR_DEVISE FROM QBPARBRE '+
              'WHERE QBR_CODESESSION="'+codeSession+
              '" AND QBR_NUMNOEUD="'+Noeud+'"','QUFVBPARBREMODIF_TOF (OnArgument).',true);
  if not Q.eof then
  begin
    devise:=Q.fields[NivMax].asString;
    for i:=0 to NivMax-1 do
    begin
      if TabCodeAxe[i+1]='ARTICLE' then
      begin
        if dm_trouveEnr(idm_tmpBPArticleTai,[Q.fields[i].asString],Enr)=0
        then SetControlText('GRILLETAILLE',Enr.ch(BPARTICLETAI_GRILLETAI));
      end;
      if not BPOkOrli then
      begin
        LibelleAxe := DonneLibelleAxe(TabCodeAxe[i+1]);
        if LibelleAxe = '' then LibelleAxe := TabCodeAxe[i+1];
        THRadioGroup(getcontrol('GRP1')).Items.Add(LibelleAxe+'='+Q.fields[i].asString+' ('+DonneLibelleValeurAxe(TabCodeAxe[i+1],Q.fields[i].asString)+')');
      end  //ORLI
      else THRadioGroup(getcontrol('GRP1')).Items.Add(TabCodeAxe[i+1]+'='+Q.fields[i].asString+' ('+DonneLibelleValeurAxe(TabCodeAxe[i+1],Q.fields[i].asString)+')');

      THRadioGroup(getcontrol('GRP1')).Values.add(intToStr(i));
    end;
  end;
  ferme(Q);
  THRadioGroup(getcontrol('GRP1')).ItemIndex:=NivMax-1;

  { EVI / Affichage des devise en fonction de l'axe choisi }
  { THGroupBox(getcontrol('BOX1')).caption:=TraduireMemoire('Montant en "')+devise+'"';
  THGroupBox(getcontrol('BOX3')).caption:=TraduireMemoire('Montant en "')+devise+'"';
  THGroupBox(getcontrol('BOX4')).caption:=TraduireMemoire('Montant en "')+devise+'"';
  THGroupBox(getcontrol('BOX5')).caption:=TraduireMemoire('Montant en "')+devise+'"';
  THGroupBox(getcontrol('BOX6')).caption:=TraduireMemoire('Montant en "')+devise+'"';
  THGroupBox(getcontrol('BOX7')).caption:=TraduireMemoire('Montant en "')+devise+'"'; }

  THRadioGroup(getcontrol('GRP1')).OnClick:=ClickRadio;
  THEdit(GetControl('EDTVALEURAXE')).OnChange:=ChangeValeurAxe;

  for i:=0 to 16 do TabNoeud[i]:=0;
  //remplit tableau num noeud
  TabNoeud[NivMax]:=VALEURI(Noeud);
  for i:=NivMax-1 downto 1 do TabNoeud[i]:=ChercheNoeudPere(codeSession,TabNoeud[i+1]);

  ClickRadio(self);
  retourS:='';

  if SessionInitPrev(codeSession) then
  begin
    THLabel(GetControl('TQBR_QTEREFPRCT')).caption:=TraduireMemoire('Histo');
    THLabel(GetControl('TQBR_QTEREF')).caption:=TraduireMemoire('Réalisé');
    THLabel(GetControl('TQBR_QTEC')).caption:=TraduireMemoire('Prévu');
    THLabel(GetControl('TQBI_SAISIVAL1')).caption:=TraduireMemoire('Qté saisie');
    THLabel(GetControl('LBLQTERETENUE')).visible:=true;
    THEdit(GetControl('EDTQTERETENUE')).visible:=true;
  end
end;

procedure TOF_QUFVBPARBREMODIF.OnClose ;
begin
  Inherited ;
  TFVierge(ECRAN).Retour:=retourS;
end ;

procedure TOF_QUFVBPARBREMODIF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPARBREMODIF.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPARBREMODIF.ClickRadio(Sender: TObject);
var index:integer;
    Q:TQuery;
    OkTaille:boolean;
begin
 Index:=THRadioGroup(getcontrol('GRP1')).ItemIndex;

 if (OkCoeff='X') and (Index=NivMax-1)
  then THTabSheet(GetControl('TAB8')).tabVisible:=true
  else THTabSheet(GetControl('TAB8')).tabVisible:=false;
 OkTaille:=false;
 if (SessionCalculParTaille(codeSession)) or (SessionEclateeParTaille(codeSession))
  then OkTaille:=true;
 if (OkPrev='X') and (Index=NivMax-1) and OkTaille 
  then THTabSheet(GetControl('TAB9')).tabVisible:=true
  else THTabSheet(GetControl('TAB9')).tabVisible:=false;

 Q:=MOPenSql('SELECT QBR_REF1,QBR_REFPRCT1,QBR_OP1,QBR_EVOLPRCT1,QBR_EVOLVAL1,QBR_SAISI1,'+
             'QBR_REF2,QBR_REFPRCT2,QBR_OP2,QBR_EVOLPRCT2,QBR_EVOLVAL2,QBR_SAISI2,'+
             'QBR_REF3,QBR_REFPRCT3,QBR_OP3,QBR_EVOLPRCT3,QBR_EVOLVAL3,QBR_SAISI3,'+
             'QBR_REF4,QBR_REFPRCT4,QBR_OP4,QBR_EVOLPRCT4,QBR_EVOLVAL4,QBR_SAISI4,'+
             'QBR_REF5,QBR_REFPRCT5,QBR_OP5,QBR_EVOLPRCT5,QBR_EVOLVAL5,QBR_SAISI5,'+
             'QBR_REF6,QBR_REFPRCT6,QBR_OP6,QBR_EVOLPRCT6,QBR_EVOLVAL6,QBR_SAISI6,'+
             'QBR_QTEREF,QBR_QTEREFPRCT,QBR_QTEC,QBR_EVOLQTEPRCT,QBR_EVOLQTE,QBR_SAISIQTE,'+
             'QBR_HISTO,QBR_PREVU,QBR_COEFFCALCUL,QBR_COEFFRETENU,'+
             'QBR_COMMENTAIREBP,QBR_VALEURAXE,QBR_VALBLOQUE,QBR_CODEAXE, '+
             'QBR_QTERETENUE,QBR_REALISE,'+
             'QBR_QTET1,QBR_QTET2,QBR_QTET3,QBR_QTET4,QBR_QTET5,'+
             'QBR_QTET6,QBR_QTET7,QBR_QTET8,QBR_QTET9,QBR_QTET10,'+
             'QBR_QTET11,QBR_QTET12,QBR_QTET13,QBR_QTET14,QBR_QTET15,'+
             'QBR_QTET16,QBR_QTET17,QBR_QTET18,QBR_QTET19,QBR_QTET20 '+
             ' FROM QBPARBRE WHERE QBR_CODESESSION="'+codeSession+
             '" AND QBR_NUMNOEUD="'+STRFPOINT(TabNoeud[Index+1])+'"',
             'QUFVBPARBREMODIF_TOF (ClickRadio).',true);
 if not Q.eof
  then
   begin
    //CA1
    THNumEdit(GetControl('EDTREF1')).value:=Q.fields[0].asFloat;
    THNumEdit(GetControl('EDTREFPRCT1')).value:=Q.fields[1].asFloat;
    THNumEdit(GetControl('EDTOP1')).value:=Q.fields[2].asFloat;
    THNumEdit(GetControl('EDTEVOLPRCT1')).value:=Q.fields[3].asFloat;
    SetControlText('EDTEVOLVAL1',Q.fields[4].asString);
    SetControlText('EDTSAISI1',Q.fields[5].asString);
    if Q.fields[5].asString='0'
     then SetControlText('EDTSAISI1','');
    //CA2
    THNumEdit(GetControl('EDTREF2')).value:=Q.fields[6].asFloat;
    THNumEdit(GetControl('EDTREFPRCT2')).value:=Q.fields[7].asFloat;
    THNumEdit(GetControl('EDTOP2')).value:=Q.fields[8].asFloat;
    THNumEdit(GetControl('EDTEVOLPRCT2')).value:=Q.fields[9].asFloat;
    SetControlText('EDTEVOLVAL2',Q.fields[10].asString);
    SetControlText('EDTSAISI2',Q.fields[11].asString);  
    if Q.fields[11].asString='0'
     then SetControlText('EDTSAISI2','');
    //CA3
    THNumEdit(GetControl('EDTREF3')).value:=Q.fields[12].asFloat;
    THNumEdit(GetControl('EDTREFPRCT3')).value:=Q.fields[13].asFloat;
    THNumEdit(GetControl('EDTOP3')).value:=Q.fields[14].asFloat;
    THNumEdit(GetControl('EDTEVOLPRCT3')).value:=Q.fields[15].asFloat;
    SetControlText('EDTEVOLVAL3',Q.fields[16].asString);
    SetControlText('EDTSAISI3',Q.fields[17].asString);   
    if Q.fields[17].asString='0'
     then SetControlText('EDTSAISI3','');
    //CA4
    THNumEdit(GetControl('EDTREF4')).value:=Q.fields[18].asFloat;
    THNumEdit(GetControl('EDTREFPRCT4')).value:=Q.fields[19].asFloat;
    THNumEdit(GetControl('EDTOP4')).value:=Q.fields[20].asFloat;
    THNumEdit(GetControl('EDTEVOLPRCT4')).value:=Q.fields[21].asFloat;
    SetControlText('EDTEVOLVAL4',Q.fields[22].asString);
    SetControlText('EDTSAISI4',Q.fields[23].asString);
    if Q.fields[23].asString='0'
     then SetControlText('EDTSAISI4','');
    //CA5
    THNumEdit(GetControl('EDTREF5')).value:=Q.fields[24].asFloat;
    THNumEdit(GetControl('EDTREFPRCT5')).value:=Q.fields[25].asFloat;
    THNumEdit(GetControl('EDTOP5')).value:=Q.fields[26].asFloat;
    THNumEdit(GetControl('EDTEVOLPRCT5')).value:=Q.fields[27].asFloat;
    SetControlText('EDTEVOLVAL5',Q.fields[28].asString);
    SetControlText('EDTSAISI5',Q.fields[29].asString);   
    if Q.fields[29].asString='0'
     then SetControlText('EDTSAISI5','');
    //CA6
    THNumEdit(GetControl('EDTREF6')).value:=Q.fields[30].asFloat;
    THNumEdit(GetControl('EDTREFPRCT6')).value:=Q.fields[31].asFloat;
    THNumEdit(GetControl('EDTOP6')).value:=Q.fields[32].asFloat;
    THNumEdit(GetControl('EDTEVOLPRCT6')).value:=Q.fields[33].asFloat;
    SetControlText('EDTEVOLVAL6',Q.fields[34].asString);
    SetControlText('EDTSAISI6',Q.fields[35].asString);  
    if Q.fields[35].asString='0'
     then SetControlText('EDTSAISI6','');
    //QTE
    THNumEdit(GetControl('EDTQTEREF')).value:=Q.fields[36].asFloat;
    THNumEdit(GetControl('EDTQTEREFPRCT')).value:=Q.fields[37].asFloat;
    THNumEdit(GetControl('EDTQTEC')).value:=Q.fields[38].asFloat;
    THNumEdit(GetControl('EDTEVOLQTEPRCT')).value:=Q.fields[39].asFloat;
    SetControlText('EDTEVOLVALQTE',Q.fields[40].asString);
    SetControlText('EDTSAISIQTE',Q.fields[41].asString);
    if Q.fields[41].asString='0'
     then SetControlText('EDTSAISIQTE','');
    if SessionInitPrev(codeSession)
     then
      begin                
       THNumEdit(GetControl('EDTQTEREFPRCT')).value:=Q.fields[42].asFloat;
       THNumEdit(GetControl('EDTQTEREF')).value:=Q.fields[51].asFloat;
       THNumEdit(GetControl('EDTQTEC')).value:=Q.fields[43].asFloat;  
       SetControlText('EDTQTERETENUE',Q.fields[50].asString);

       SetControlText('EDTT1',Q.fields[52].asString);
       SetControlText('EDTT2',Q.fields[53].asString);
       SetControlText('EDTT3',Q.fields[54].asString);
       SetControlText('EDTT4',Q.fields[55].asString);
       SetControlText('EDTT5',Q.fields[56].asString);
       SetControlText('EDTT6',Q.fields[57].asString);
       SetControlText('EDTT7',Q.fields[58].asString);
       SetControlText('EDTT8',Q.fields[59].asString);
       SetControlText('EDTT9',Q.fields[60].asString);
       SetControlText('EDTT10',Q.fields[61].asString);
       SetControlText('EDTT11',Q.fields[62].asString);
       SetControlText('EDTT12',Q.fields[63].asString);
       SetControlText('EDTT13',Q.fields[64].asString);
       SetControlText('EDTT14',Q.fields[65].asString);
       SetControlText('EDTT15',Q.fields[66].asString);
       SetControlText('EDTT16',Q.fields[67].asString);
       SetControlText('EDTT17',Q.fields[68].asString);
       SetControlText('EDTT18',Q.fields[69].asString);
       SetControlText('EDTT19',Q.fields[70].asString);
       SetControlText('EDTT20',Q.fields[71].asString);

      end;
    //coeff
    THNumEdit(GetControl('EDTHISTO')).value:=Q.fields[42].asFloat;
    THNumEdit(GetControl('EDTPREVU')).value:=Q.fields[43].asFloat;
    THNumEdit(GetControl('EDTCOEFFCALCUL')).value:=Q.fields[44].asFloat;
    THNumEdit(GetControl('EDTCOEFFRETENU')).value:=Q.fields[45].asFloat;


    SetControlText('EDTCOMMENTAIREBP',Q.fields[46].asString);
    SetControlText('EDTVALEURAXE',Q.fields[47].asString);
    if Q.fields[48].asString='X'
     then THCheckBox(GetControl('CHECKVALBLOQUE')).State:=cbChecked
     else THCheckBox(GetControl('CHECKVALBLOQUE')).State:=cbunChecked;
    THLabel(GetControl('LBLVALEURAXE')).caption:=Q.fields[49].asString;
   end;
 ferme(Q);
 ChangeValeurAxe(self);
end;

procedure TOF_QUFVBPARBREMODIF.ChangeValeurAxe(Sender: TObject);
begin
 THLabel(GetControl('LBLLIBVALEURAXE')).caption:=DonneLibelleValeurAxe(THLabel(GetControl('LBLVALEURAXE')).caption,GetControlText('EDTVALEURAXE'));
end;

Initialization
  registerclasses ( [ TOF_QUFVBPARBREMODIF ] ) ; 
end.
