{***********UNITE*************************************************
Auteur  ...... : EV5
Créé le ...... : xx/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QUFVBPMULTISOC ()
Mots clefs ... : TOF;QUFVBPMULTISOC
*****************************************************************}
Unit QUFVBPMULTISOC_TOF ;

Interface

Uses UTOF,Graphics,HRichOle,HEnt1,Classes;

Type
  TOF_QUFVBPMULTISOC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    private
    procedure BTraitement_OnClick(Sender: TObject);
    procedure MultiDossier1_OnChange(Sender: TObject);
    procedure TextColor( RapMulSoc:THRichEditOle ; Text : hstring; couleur : Tcolor );

  end ;

var CodeSession,LibelleSession:hString;
RapportMultiSoc,RapportMultiSoc1 : THRichEditOle;
lstBase:TStringList;

Implementation

Uses  Controls,
      {$IFNDEF EAGLCLIENT}
      {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
      {$else}eMul,uTob,
      {$ENDIF}
      sysutils,ComCtrls,HCtrls,
      HMsgBox,vierge,Uutil,SynScriptBP,HTB97,UtilPGI,BPUtil;

procedure TOF_QUFVBPMULTISOC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPMULTISOC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPMULTISOC.OnUpdate ;
begin
end ;

procedure TOF_QUFVBPMULTISOC.OnLoad ;
begin
  Inherited ;
  THValComboBox(GetControl('MULTIDOSSIER')).ItemIndex := 1;
  THValComboBox(GetControl('MULTIDOSSIER1')).ItemIndex := 1;
  THValComboBox(GetControl('MULTIDOSSIER1')).OnChange(self);
end ;

procedure TOF_QUFVBPMULTISOC.OnArgument (S : String ) ;
begin
  Inherited ;
  CodeSession:=TrouveArgument(S,'SESSION','');
  LibelleSession:=FuncSynBPDonneLibelleSession([CodeSession],0);
  SetControlText('TH_CODESESSION',CodeSession);
  SetControlText('TH_CODESESSION1',CodeSession);
  SetControlText('TH_LIBSESSION',LibelleSession);
  SetControlText('TH_LIBSESSION1',LibelleSession);
  RapportMultiSoc := THRichEditOle( GetControl( 'THRE_MULTISOC' ) );
  RapportMultiSoc1 := THRichEditOle( GetControl( 'THRE_MULTISOC1' ) );

  TToolBarButton97(GetControl('BTRAITEMENT')).OnClick := BTraitement_OnClick;
  THValComboBox(GetControl('MULTIDOSSIER1')).OnChange := MultiDossier1_OnChange;
  lstBase := TStringList.Create;
end ;


procedure TOF_QUFVBPMULTISOC.MultiDossier1_OnChange(Sender: TObject);
var sSocietes,sBase:string;
i:integer;
begin
  lstBase.Clear;
  for i:=0 to THMultiValComboBox(GetControl('THM_SOCIETES')).Items.count-1 do
    THMultiValComboBox(GetControl('THM_SOCIETES')).Items.delete(0);
  THMultiValComboBox(GetControl('THM_SOCIETES')).clear;
  sSocietes := GetBasesMS(THValComboBox(GetControl('MULTIDOSSIER1')).Value, true);
  if sSocietes <> '' then
  begin
    i:=1;
    sBase := ReadTokenSt(sSocietes) ;
    repeat
      if sBase <> V_PGI.SchemaName then
      begin
        THMultiValComboBox(GetControl('THM_SOCIETES')).Items.Add(sBase);
        THMultiValComboBox(GetControl('THM_SOCIETES')).Values.Add(MetZero(intTostr(i),2));
        lstBase.Add(sBase);
        inc(i);
      end;
    sBase  := ReadTokenSt(sSocietes) ;
    until (sBase = '');
  end
end;

procedure TOF_QUFVBPMULTISOC.BTraitement_OnClick;
var BaseName,BaseValue,BaseValues,sSocietes:string;
rep,i,nbValue,DuplicOK:integer;
lstMsg,lstMsgState:TStringList;
begin
  if TPageControl(GetControl('MULTISOCSESPC')).ActivePageIndex = 0 then
  begin
    lstMsg:=TStringList.Create;
    lstMsgState:=TStringList.Create;
    RapportMultiSoc.Clear;
    //Cohérence Session
    sSocietes := GetBasesMS(THValComboBox(GetControl('MULTIDOSSIER')).Value, true);
    if sSocietes <> '' then
    begin
      VerifCoherenceSessions(sSocietes,codeSession,lstMsg,lstMsgState);
      for i:=0 to lstMsg.Count-1 do
      begin
        TextColor(RapportMultiSoc,lstMsg[i], StringToCOlor(lstMsgState[i]) );
      end
    end else HShowMessage('1;Erreur;Ce regroupement ne contient aucune base.;W;O;O;O;','','');
    lstMsg.free;
    lstMsgState.free;
  end
  else
  begin
    //Copie Session
    if THMultiValComboBox(GetControl('THM_SOCIETES')).Text<>'' then
    begin
      if THMultiValComboBox(GetControl('THM_SOCIETES')).Text=TraduireMemoire('<<Tous>>') then
      begin
        nbValue:=THMultiValComboBox(GetControl('THM_SOCIETES')).Items.count;
        if nbValue = 1
        then rep:=HShowmessage('1;Attention;Vous allez copier la session courante sur 1 base.'+#13#10+
                               ' La session existante sur cette base sera écrasée.'+#13#10+
                               ' Etes-vous sûr ?;Q;YN;N;N', '', '')
        else rep:=HShowmessage('1;Attention;Vous allez copier la session courante sur '+IntToStr(nbValue)+' bases.'+#13#10+
                               ' Les sessions existantes sur ces bases seront écrasées.'+#13#10+
                               ' Etes-vous sûr ?;Q;YN;N;N', '', '');
        if rep=MrYes then
        begin
          RapportMultiSoc1.Clear;
          for i:=0 to nbValue-1 do
          begin
            BaseName:=lstBase[i];
            //copier la session en la mettant en non initialisée
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'"');
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPARBREDETAIL WHERE QBH_CODESESSION="'+codeSession+'"');
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPLOI WHERE QBO_CODESESSION="'+codeSession+'"');
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPCUBETMP WHERE QBQ_CODESESSION="'+codeSession+'"');
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPSESSIONBP WHERE QBS_CODESESSION="'+codeSession+'"');
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPDETCALENDREP WHERE QBE_CALENDREP="'+codeSession+'"');
            DuplicOK:=ExecuteSql('INSERT INTO '+BaseName+'.DBO.QBPSESSIONBP(QBS_CODESESSION,QBS_LIBSESSION,QBS_DATEDEBC,QBS_DATEFINC,'+
                       'QBS_DATEDEBREF,QBS_DATEFINREF,QBS_SAISONCMDC,QBS_SAISONCMDREF,QBS_SESIONOBJECTIF,'+
                       'QBS_SESSIONBRUTE,QBS_DATEOUVERT,QBS_DATEVALIDATION,QBS_TYPENATURE,QBS_CODESTRUCT,'+
                       'QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4,QBS_CODEAXES5,QBS_CODEAXES6,'+
                       'QBS_CODEAXES7,QBS_CODEAXES8,QBS_CODEAXES9,QBS_CODEAXES10,QBS_METHODE,QBS_CODEAXE1,'+
                       'QBS_VALEURAXE1,QBS_CODEAXE2,QBS_VALEURAXE2,QBS_CODEAXE3,QBS_VALEURAXE3,QBS_CODEAXE4,'+
                       'QBS_VALEURAXE4,QBS_NATURECMD,QBS_MODEGESTPREV,QBS_EXTRAPOLABLE,QBS_SESSIONINIT,QBS_SESSIONECLAT,'+
                       'QBS_SESSIONECLATAI,QBS_BPINITIALISE,QBS_SESSIONNETTE,QBS_SESSIONFERMEE,QBS_VUEARBRE,QBS_INITDELAI,'+
                       'QBS_INITCOEFF,QBS_INITPREVISION,QBS_DATECREATION,QBS_DATEMODIF,QBS_CREATEUR,QBS_UTILISATEUR,'+
                       'QBS_OKMODIFSESSION,QBS_OKMODIFVUE,QBS_NUMSESSION,QBS_VALRESTRICTION,QBS_SESSIONVALIDE,'+
                       'QBS_METHODETAILLE,QBS_DATEEXPORT,QBS_OKOBJPREV,QBS_OBJPREV,QBS_NATURE2,QBS_NBVALAFF,QBS_VALAFFH1,'+
                       'QBS_VALAFFH1A,QBS_VALAFFH1B,QBS_VALAFFR1,QBS_VALAFFR1A,QBS_VALAFFR1B,QBS_VALAFFH2,QBS_VALAFFH2A,'+
                       'QBS_VALAFFH2B,QBS_VALAFFR2,QBS_VALAFFR2A,QBS_VALAFFR2B,QBS_VALAFFH3,QBS_VALAFFH3A,QBS_VALAFFH3B,'+
                       'QBS_VALAFFR3,QBS_VALAFFR3A,QBS_VALAFFR3B,QBS_VALAFFH4,QBS_VALAFFH4A,QBS_VALAFFH4B,QBS_VALAFFR4,'+
                       'QBS_VALAFFR4A,QBS_VALAFFR4B,QBS_VALAFFH5,QBS_VALAFFH5A,QBS_VALAFFH5B,QBS_VALAFFR5,QBS_VALAFFR5A,'+
                       'QBS_VALAFFR5B,QBS_VALAFFH6,QBS_VALAFFH6A,QBS_VALAFFH6B,QBS_VALAFFR6,QBS_VALAFFR6A,QBS_VALAFFR6B,'+
                       'QBS_VALAFFH7,QBS_VALAFFH7A,QBS_VALAFFH7B,QBS_VALAFFR7,QBS_VALAFFR7A,QBS_VALAFFR7B,QBS_VALAFFLIB1,'+
                       'QBS_VALAFFLIB2,QBS_VALAFFLIB3,QBS_VALAFFLIB4,QBS_VALAFFLIB5,QBS_VALAFFLIB6,QBS_VALAFFLIB7) '+
                       'SELECT QBS_CODESESSION,QBS_LIBSESSION,QBS_DATEDEBC,QBS_DATEFINC,'+
                       'QBS_DATEDEBREF,QBS_DATEFINREF,QBS_SAISONCMDC,QBS_SAISONCMDREF,QBS_SESIONOBJECTIF,'+
                       'QBS_SESSIONBRUTE,QBS_DATEOUVERT,QBS_DATEVALIDATION,QBS_TYPENATURE,QBS_CODESTRUCT,'+
                       'QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4,QBS_CODEAXES5,QBS_CODEAXES6,'+
                       'QBS_CODEAXES7,QBS_CODEAXES8,QBS_CODEAXES9,QBS_CODEAXES10,QBS_METHODE,QBS_CODEAXE1,'+
                       'QBS_VALEURAXE1,QBS_CODEAXE2,QBS_VALEURAXE2,QBS_CODEAXE3,QBS_VALEURAXE3,QBS_CODEAXE4,'+
                       'QBS_VALEURAXE4,QBS_NATURECMD,QBS_MODEGESTPREV,QBS_EXTRAPOLABLE,"-","-",'+
                       '"-",QBS_BPINITIALISE,QBS_SESSIONNETTE,QBS_SESSIONFERMEE,QBS_VUEARBRE,QBS_INITDELAI,'+
                       'QBS_INITCOEFF,QBS_INITPREVISION,"'+UsDateTime(now)+'","'+UsDateTime(now)+'","'+V_PGI.User+'","'+V_PGI.User+'",'+
                       'QBS_OKMODIFSESSION,QBS_OKMODIFVUE,QBS_NUMSESSION,QBS_VALRESTRICTION,QBS_SESSIONVALIDE,'+
                       'QBS_METHODETAILLE,QBS_DATEEXPORT,QBS_OKOBJPREV,QBS_OBJPREV,QBS_NATURE2,QBS_NBVALAFF,QBS_VALAFFH1,'+
                       'QBS_VALAFFH1A,QBS_VALAFFH1B,QBS_VALAFFR1,QBS_VALAFFR1A,QBS_VALAFFR1B,QBS_VALAFFH2,QBS_VALAFFH2A,'+
                       'QBS_VALAFFH2B,QBS_VALAFFR2,QBS_VALAFFR2A,QBS_VALAFFR2B,QBS_VALAFFH3,QBS_VALAFFH3A,QBS_VALAFFH3B,'+
                       'QBS_VALAFFR3,QBS_VALAFFR3A,QBS_VALAFFR3B,QBS_VALAFFH4,QBS_VALAFFH4A,QBS_VALAFFH4B,QBS_VALAFFR4,'+
                       'QBS_VALAFFR4A,QBS_VALAFFR4B,QBS_VALAFFH5,QBS_VALAFFH5A,QBS_VALAFFH5B,QBS_VALAFFR5,QBS_VALAFFR5A,'+
                       'QBS_VALAFFR5B,QBS_VALAFFH6,QBS_VALAFFH6A,QBS_VALAFFH6B,QBS_VALAFFR6,QBS_VALAFFR6A,QBS_VALAFFR6B,'+
                       'QBS_VALAFFH7,QBS_VALAFFH7A,QBS_VALAFFH7B,QBS_VALAFFR7,QBS_VALAFFR7A,QBS_VALAFFR7B,QBS_VALAFFLIB1,'+
                       'QBS_VALAFFLIB2,QBS_VALAFFLIB3,QBS_VALAFFLIB4,QBS_VALAFFLIB5,QBS_VALAFFLIB6,QBS_VALAFFLIB7 FROM QBPSESSIONBP '+
                       'WHERE QBS_CODESESSION="'+codeSession+'"');
            if DuplicOK=1
            then TextColor(RapportMultiSoc1,TraduireMemoire( 'La session '+codeSession+' a été dupliquée de la base '+V_PGI.SchemaName+' à la base '+BaseName), clGreen )
            else TextColor(RapportMultiSoc1,TraduireMemoire( 'Erreur rencontrée lors de la duplication de la session '+codeSession+' de la base '+V_PGI.SchemaName+' à la base '+BaseName), clRed )
          end;
        end
      end
      else
      begin
        nbValue:=0;
        BaseValues:=THMultiValComboBox(GetControl('THM_SOCIETES')).Text;
        BaseValue := ReadTokenSt(BaseValues) ;
        repeat
          inc(nbValue);
          BaseValue := ReadTokenSt(BaseValues) ;
        until (BaseValue = '');
        if nbValue = 1
        then rep:=HShowmessage('1;Attention;Vous allez copier la session courante sur 1 base.'+#13#10+
                               ' La session existante sur cette base sera écrasée.'+#13#10+
                               ' Etes-vous sûr ?;Q;YN;N;N', '', '')
        else rep:=HShowmessage('1;Attention;Vous allez copier la session courante sur '+IntToStr(nbValue)+' bases.'+#13#10+
                               ' Les sessions existantes sur ces bases seront écrasées.'+#13#10+
                               ' Etes-vous sûr ?;Q;YN;N;N', '', '');
        if rep=MrYes then
        begin
          RapportMultiSoc1.Clear;
          BaseValues:=THMultiValComboBox(GetControl('THM_SOCIETES')).Text;
          BaseValue := ReadTokenSt(BaseValues) ;
          repeat
            BaseName:=lstBase[StrToInt(BaseValue)-1];
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPARBRE WHERE QBR_CODESESSION="'+codeSession+'"');
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPARBREDETAIL WHERE QBH_CODESESSION="'+codeSession+'"');
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPLOI WHERE QBO_CODESESSION="'+codeSession+'"');
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPCUBETMP WHERE QBQ_CODESESSION="'+codeSession+'"');
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPSESSIONBP WHERE QBS_CODESESSION="'+codeSession+'"');
            ExecuteSQL('DELETE FROM '+BaseName+'.DBO.QBPDETCALENDREP WHERE QBE_CALENDREP="'+codeSession+'"');
            DuplicOK:=ExecuteSql('INSERT INTO '+BaseName+'.DBO.QBPSESSIONBP(QBS_CODESESSION,QBS_LIBSESSION,QBS_DATEDEBC,QBS_DATEFINC,'+
                       'QBS_DATEDEBREF,QBS_DATEFINREF,QBS_SAISONCMDC,QBS_SAISONCMDREF,QBS_SESIONOBJECTIF,'+
                       'QBS_SESSIONBRUTE,QBS_DATEOUVERT,QBS_DATEVALIDATION,QBS_TYPENATURE,QBS_CODESTRUCT,'+
                       'QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4,QBS_CODEAXES5,QBS_CODEAXES6,'+
                       'QBS_CODEAXES7,QBS_CODEAXES8,QBS_CODEAXES9,QBS_CODEAXES10,QBS_METHODE,QBS_CODEAXE1,'+
                       'QBS_VALEURAXE1,QBS_CODEAXE2,QBS_VALEURAXE2,QBS_CODEAXE3,QBS_VALEURAXE3,QBS_CODEAXE4,'+
                       'QBS_VALEURAXE4,QBS_NATURECMD,QBS_MODEGESTPREV,QBS_EXTRAPOLABLE,QBS_SESSIONINIT,QBS_SESSIONECLAT,'+
                       'QBS_SESSIONECLATAI,QBS_BPINITIALISE,QBS_SESSIONNETTE,QBS_SESSIONFERMEE,QBS_VUEARBRE,QBS_INITDELAI,'+
                       'QBS_INITCOEFF,QBS_INITPREVISION,QBS_DATECREATION,QBS_DATEMODIF,QBS_CREATEUR,QBS_UTILISATEUR,'+
                       'QBS_OKMODIFSESSION,QBS_OKMODIFVUE,QBS_NUMSESSION,QBS_VALRESTRICTION,QBS_SESSIONVALIDE,'+
                       'QBS_METHODETAILLE,QBS_DATEEXPORT,QBS_OKOBJPREV,QBS_OBJPREV,QBS_NATURE2,QBS_NBVALAFF,QBS_VALAFFH1,'+
                       'QBS_VALAFFH1A,QBS_VALAFFH1B,QBS_VALAFFR1,QBS_VALAFFR1A,QBS_VALAFFR1B,QBS_VALAFFH2,QBS_VALAFFH2A,'+
                       'QBS_VALAFFH2B,QBS_VALAFFR2,QBS_VALAFFR2A,QBS_VALAFFR2B,QBS_VALAFFH3,QBS_VALAFFH3A,QBS_VALAFFH3B,'+
                       'QBS_VALAFFR3,QBS_VALAFFR3A,QBS_VALAFFR3B,QBS_VALAFFH4,QBS_VALAFFH4A,QBS_VALAFFH4B,QBS_VALAFFR4,'+
                       'QBS_VALAFFR4A,QBS_VALAFFR4B,QBS_VALAFFH5,QBS_VALAFFH5A,QBS_VALAFFH5B,QBS_VALAFFR5,QBS_VALAFFR5A,'+
                       'QBS_VALAFFR5B,QBS_VALAFFH6,QBS_VALAFFH6A,QBS_VALAFFH6B,QBS_VALAFFR6,QBS_VALAFFR6A,QBS_VALAFFR6B,'+
                       'QBS_VALAFFH7,QBS_VALAFFH7A,QBS_VALAFFH7B,QBS_VALAFFR7,QBS_VALAFFR7A,QBS_VALAFFR7B,QBS_VALAFFLIB1,'+
                       'QBS_VALAFFLIB2,QBS_VALAFFLIB3,QBS_VALAFFLIB4,QBS_VALAFFLIB5,QBS_VALAFFLIB6,QBS_VALAFFLIB7) '+
                       'SELECT QBS_CODESESSION,QBS_LIBSESSION,QBS_DATEDEBC,QBS_DATEFINC,'+
                       'QBS_DATEDEBREF,QBS_DATEFINREF,QBS_SAISONCMDC,QBS_SAISONCMDREF,QBS_SESIONOBJECTIF,'+
                       'QBS_SESSIONBRUTE,QBS_DATEOUVERT,QBS_DATEVALIDATION,QBS_TYPENATURE,QBS_CODESTRUCT,'+
                       'QBS_CODEAXES1,QBS_CODEAXES2,QBS_CODEAXES3,QBS_CODEAXES4,QBS_CODEAXES5,QBS_CODEAXES6,'+
                       'QBS_CODEAXES7,QBS_CODEAXES8,QBS_CODEAXES9,QBS_CODEAXES10,QBS_METHODE,QBS_CODEAXE1,'+
                       'QBS_VALEURAXE1,QBS_CODEAXE2,QBS_VALEURAXE2,QBS_CODEAXE3,QBS_VALEURAXE3,QBS_CODEAXE4,'+
                       'QBS_VALEURAXE4,QBS_NATURECMD,QBS_MODEGESTPREV,QBS_EXTRAPOLABLE,"-","-",'+
                       '"-",QBS_BPINITIALISE,QBS_SESSIONNETTE,QBS_SESSIONFERMEE,QBS_VUEARBRE,QBS_INITDELAI,'+
                       'QBS_INITCOEFF,QBS_INITPREVISION,"'+UsDateTime(now)+'","'+UsDateTime(now)+'","'+V_PGI.User+'","'+V_PGI.User+'",'+
                       'QBS_OKMODIFSESSION,QBS_OKMODIFVUE,QBS_NUMSESSION,QBS_VALRESTRICTION,QBS_SESSIONVALIDE,'+
                       'QBS_METHODETAILLE,QBS_DATEEXPORT,QBS_OKOBJPREV,QBS_OBJPREV,QBS_NATURE2,QBS_NBVALAFF,QBS_VALAFFH1,'+
                       'QBS_VALAFFH1A,QBS_VALAFFH1B,QBS_VALAFFR1,QBS_VALAFFR1A,QBS_VALAFFR1B,QBS_VALAFFH2,QBS_VALAFFH2A,'+
                       'QBS_VALAFFH2B,QBS_VALAFFR2,QBS_VALAFFR2A,QBS_VALAFFR2B,QBS_VALAFFH3,QBS_VALAFFH3A,QBS_VALAFFH3B,'+
                       'QBS_VALAFFR3,QBS_VALAFFR3A,QBS_VALAFFR3B,QBS_VALAFFH4,QBS_VALAFFH4A,QBS_VALAFFH4B,QBS_VALAFFR4,'+
                       'QBS_VALAFFR4A,QBS_VALAFFR4B,QBS_VALAFFH5,QBS_VALAFFH5A,QBS_VALAFFH5B,QBS_VALAFFR5,QBS_VALAFFR5A,'+
                       'QBS_VALAFFR5B,QBS_VALAFFH6,QBS_VALAFFH6A,QBS_VALAFFH6B,QBS_VALAFFR6,QBS_VALAFFR6A,QBS_VALAFFR6B,'+
                       'QBS_VALAFFH7,QBS_VALAFFH7A,QBS_VALAFFH7B,QBS_VALAFFR7,QBS_VALAFFR7A,QBS_VALAFFR7B,QBS_VALAFFLIB1,'+
                       'QBS_VALAFFLIB2,QBS_VALAFFLIB3,QBS_VALAFFLIB4,QBS_VALAFFLIB5,QBS_VALAFFLIB6,QBS_VALAFFLIB7 FROM QBPSESSIONBP '+
                       'WHERE QBS_CODESESSION="'+codeSession+'"');
            if DuplicOK=1
            then TextColor(RapportMultiSoc1,TraduireMemoire( 'La session '+codeSession+' a été dupliquée de la base '+V_PGI.SchemaName+' à la base '+BaseName), clGreen )
            else TextColor(RapportMultiSoc1,TraduireMemoire( 'Erreur rencontrée lors de la duplication de la session '+codeSession+' de la base '+V_PGI.SchemaName+' à la base '+BaseName), clRed );
            BaseValue := ReadTokenSt(BaseValues) ;
          until (BaseValue = '');
        end
      end
    end
    else HShowMessage('1;Erreur;Aucune base sélectionnée.;W;O;O;O;','','');
  end;
end;

procedure TOF_QUFVBPMULTISOC.OnClose ;
begin
  Inherited ;
  lstBase.Free;
end ;

procedure TOF_QUFVBPMULTISOC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPMULTISOC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_QUFVBPMULTISOC.TextColor( RapMulSoc:THRichEditOle ; Text : hstring; couleur : Tcolor );
begin
  if Assigned( RapMulSoc ) then
  begin
    RapMulSoc.SelAttributes.color := couleur;
    RapMulSoc.Lines.Add( Text );
  end;
end;

{procedure TOF_MBOGENMINMAX.PosCursor( Richedit : THRichEditOle );
var
  i, LastVisibleLine : integer;
begin
  SendMessage( Richedit.Handle, WM_VSCROLL, SB_BOTTOM, 0 );
  LastVisibleLine := Richedit.Perform( EM_EXLINEFROMCHAR, 0, Richedit.selstart );
  SendMessage( Richedit.Handle, WM_VSCROLL, SB_TOP, 0 );
  for i := 0 to ( LastVisibleLine - 7 ) do
    SendMessage( Richedit.Handle, WM_VSCROLL, SB_LINEDOWN, 0 );
end;
PosCursor( RecapMaj );  }

Initialization
  registerclasses ( [ TOF_QUFVBPMULTISOC ] ) ;
end.

