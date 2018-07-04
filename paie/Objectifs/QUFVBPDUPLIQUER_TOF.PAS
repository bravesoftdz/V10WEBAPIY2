{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/05/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFVBPDUPLIQUER ()
Mots clefs ... : TOF;QUFVBPDUPLIQUER
*****************************************************************}
Unit QUFVBPDUPLIQUER_TOF ;

Interface

Uses UTOF,HEnt1;

Type
  TOF_QUFVBPDUPLIQUER = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
     codeSession:hString;
     procedure Click(Sender: TObject);
  end ;

Implementation

uses Classes,Forms,Sysutils,Vierge,HCtrls,
     {$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     {$else} uTob,
     {$ENDIF}
     BPFctSession,BPBasic;

procedure TOF_QUFVBPDUPLIQUER.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDUPLIQUER.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDUPLIQUER.OnUpdate ;
var niv,dup,session:hString;
begin
  Inherited ;
  dup:=IntToStr(THRadioGroup(getcontrol('RADIODUPLI')).ItemIndex);
  niv:=IntToStr(THRadioGroup(getcontrol('GRP1')).ItemIndex);
  session:=GetControlText('EDTSESSION');
  TFVierge(ECRAN).Retour:='NIV='+niv+';DUP='+dup+';SESSION='+session;
end ;

procedure TOF_QUFVBPDUPLIQUER.OnLoad ;
begin
  Inherited;
  TForm(ecran).Height:=200;
  THRadioGroup(getcontrol('GRP1')).visible:=false;
end ;

procedure TOF_QUFVBPDUPLIQUER.OnArgument (S : String ) ;
begin
  Inherited ;
  THRadioGroup(getcontrol('RADIODUPLI')).OnClick:=Click;
  THEdit(getcontrol('EDTSESSION')).OnChange:=Click;
  if S='PREV'
   then THEdit(getcontrol('EDTSESSION')).plus:=' QBS_TYPENATURE="2" '
   else THEdit(getcontrol('EDTSESSION')).plus:=' QBS_TYPENATURE="1" ';
end ;

procedure TOF_QUFVBPDUPLIQUER.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDUPLIQUER.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDUPLIQUER.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPDUPLIQUER.Click(Sender: TObject);
var Q:TQuery;
    max:integer;
begin
  Inherited ;
  max:=1;
  if THRadioGroup(getcontrol('RADIODUPLI')).itemindex=2
  then
   begin
    TForm(ecran).Height:=370;
    THRadioGroup(getcontrol('GRP1')).visible:=true;

    THRadioGroup(getcontrol('GRP1')).Items.Clear;
    codeSession:=GetControlText('EDTSESSION');
    Q:=OPenSql('SELECT QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,'+
                'QBS_CODEAXES4,QBS_CODEAXES5,QBS_CODEAXES6,QBS_CODEAXES7,'+
                'QBS_CODEAXES8,QBS_CODEAXES9,QBS_CODEAXES10 '+
                'FROM QBPSESSIONBP WHERE QBS_CODESESSION="'+codeSession+'"',true);
    if not Q.eof
     then
      begin
       if Q.fields[0].asString<>''
        then
         begin
          if not BPOkOrli then
            THRadioGroup(getcontrol('GRP1')).Items.Add(DonneLibelleAxe(Q.fields[0].asString))
          else
          //-----------------> ORLI
          THRadioGroup(getcontrol('GRP1')).Items.Add(Q.fields[0].asString);
          THRadioGroup(getcontrol('GRP1')).Values.add('1');
          //ORLI <-----------------
          max:=1;
         end;
       if Q.fields[1].asString<>''
        then
         begin
          if not BPOkOrli then
            THRadioGroup(getcontrol('GRP1')).Items.Add(DonneLibelleAxe(Q.fields[1].asString))
          else
          //-----------------> ORLI
          THRadioGroup(getcontrol('GRP1')).Items.Add(Q.fields[1].asString);
          THRadioGroup(getcontrol('GRP1')).Values.add('2');
          //ORLI <-----------------
          max:=2;
         end;
       if Q.fields[2].asString<>''
        then
         begin
          if not BPOkOrli then
            THRadioGroup(getcontrol('GRP1')).Items.Add(DonneLibelleAxe(Q.fields[2].asString))
          else
          //-----------------> ORLI
          THRadioGroup(getcontrol('GRP1')).Items.Add(Q.fields[2].asString);
          THRadioGroup(getcontrol('GRP1')).Values.add('3');
          //ORLI <-----------------
          max:=3;
         end;
       if Q.fields[3].asString<>''
        then
         begin
          if not BPOkOrli then
            THRadioGroup(getcontrol('GRP1')).Items.Add(DonneLibelleAxe(Q.fields[3].asString))
          else
          //-----------------> ORLI
          THRadioGroup(getcontrol('GRP1')).Items.Add(Q.fields[3].asString);
          THRadioGroup(getcontrol('GRP1')).Values.add('4');
          //ORLI <-----------------
          max:=4;
         end;
       if Q.fields[4].asString<>''
        then
         begin
          if not BPOkOrli then
            THRadioGroup(getcontrol('GRP1')).Items.Add(DonneLibelleAxe(Q.fields[4].asString))
          else
          //-----------------> ORLI
          THRadioGroup(getcontrol('GRP1')).Items.Add(Q.fields[4].asString);
          THRadioGroup(getcontrol('GRP1')).Values.add('5');
          //ORLI <-----------------
          max:=5;
         end;
       if Q.fields[5].asString<>''
        then
         begin
          if not BPOkOrli then
            THRadioGroup(getcontrol('GRP1')).Items.Add(DonneLibelleAxe(Q.fields[5].asString))
          else
          //-----------------> ORLI
          THRadioGroup(getcontrol('GRP1')).Items.Add(Q.fields[5].asString);
          THRadioGroup(getcontrol('GRP1')).Values.add('6');
          //ORLI <-----------------
          max:=6;
         end;
       if Q.fields[6].asString<>''
        then
         begin
          if not BPOkOrli then
            THRadioGroup(getcontrol('GRP1')).Items.Add(DonneLibelleAxe(Q.fields[6].asString))
          else
          //-----------------> ORLI
          THRadioGroup(getcontrol('GRP1')).Items.Add(Q.fields[6].asString);
          THRadioGroup(getcontrol('GRP1')).Values.add('7');
          //ORLI <-----------------
          max:=7;
         end;
       if Q.fields[7].asString<>''
        then
         begin
          if not BPOkOrli then
            THRadioGroup(getcontrol('GRP1')).Items.Add(DonneLibelleAxe(Q.fields[7].asString))
          else
          //-----------------> ORLI
          THRadioGroup(getcontrol('GRP1')).Items.Add(Q.fields[7].asString);
          THRadioGroup(getcontrol('GRP1')).Values.add('8');
          //ORLI <-----------------
          max:=8;
         end;
       if Q.fields[8].asString<>''
        then
         begin
          if not BPOkOrli then
            THRadioGroup(getcontrol('GRP1')).Items.Add(DonneLibelleAxe(Q.fields[8].asString))
          else
          //-----------------> ORLI
          THRadioGroup(getcontrol('GRP1')).Items.Add(Q.fields[8].asString);
          THRadioGroup(getcontrol('GRP1')).Values.add('9');
          //ORLI <-----------------
          max:=9;
         end;
       if Q.fields[9].asString<>''
        then
         begin
          if not BPOkOrli then
            THRadioGroup(getcontrol('GRP1')).Items.Add(DonneLibelleAxe(Q.fields[9].asString))
          else
          //-----------------> ORLI
          THRadioGroup(getcontrol('GRP1')).Items.Add(Q.fields[9].asString);
          THRadioGroup(getcontrol('GRP1')).Values.add('10');
          //ORLI <-----------------
          max:=10;
         end;
      end;
    ferme(Q);
   if SessionEclateeParDelai(codeSession)
    then
     begin
      if not BPOkOrli then
        THRadioGroup(getcontrol('GRP1')).Items.Add(TraduireMemoire('Délai'))
      else
      //-----------------> ORLI
      THRadioGroup(getcontrol('GRP1')).Items.Add('DELAI');
      //ORLI <-----------------
      max:=max+1;
      THRadioGroup(getcontrol('GRP1')).Values.add(IntToStr(max));
     end;
   THRadioGroup(getcontrol('GRP1')).ItemIndex:=max;

   end
  else
   begin
    TForm(ecran).height:=200;
    THRadioGroup(getcontrol('GRP1')).visible:=false;
   end;
end;

Initialization
  registerclasses ( [ TOF_QUFVBPDUPLIQUER ] ) ;
end.
