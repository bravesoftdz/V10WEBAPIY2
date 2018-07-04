{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/08/2017
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : BSVSERVER (BSVSERVER)
Mots clefs ... : TOM;BSVSERVER
*****************************************************************}
Unit BSVSERVER_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
{$else}
     eFiche,
     eFichList,
{$ENDIF}
     HDb, 
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob,
     HTB97,
     UdefServices,
     UconnectBSV,
     LicUtil,
     UCryptage,
     Graphics,
     ENTGC ;

Type
  TOM_BSVSERVER = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnLoadAlerte               ; override ;
    procedure OnAfterCommit              ; override ;
    procedure OnAfterRollBack            ; override ;
    procedure OnAfterBeginTrans          ; override ;
  private
    XX : TconnectBSV;
    TOBPARAMBSV : TOB;
    CurrentUser,Currentpasswd : string;
    procedure SetServerConnected;
    procedure ConnecteEtRemplit;
    procedure ChangeParamsConnexion (sender : TObject);
    procedure ChangeArchive (Sender : TObject);
    procedure ChangeWFVALIDBAST (Sender : TObject);
    procedure ChangeChpsEtat (Sender : TObject);
    procedure ArchiveTest (Sender : TObject);
    procedure ChangeCodeEtatReg (Sender : TObject);
    end ;

Implementation

procedure TOM_BSVSERVER.OnNewRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVSERVER.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVSERVER.OnUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVSERVER.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVSERVER.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVSERVER.OnLoadRecord ;
begin
  Inherited ;
  if (GetControlText('BP2_SERVERNAME') <> '') and  (GetControlText('BP2_PORT')<> '0') then
  begin
    if XX.connected then XX.Disconnect;
    ConnecteEtRemplit;
  end;
end ;

procedure TOM_BSVSERVER.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_BSVSERVER.OnArgument ( S: String ) ;
begin
  Inherited ;
  TOBPARAMBSV := TOB.Create ('LES PARAMS BSV',nil,-1);
  CurrentUser := GetEmail(V_PGI.UserLogin);
  Currentpasswd := AnsiLowerCase_(MD5(DeCryptageSt(V_PGI.Password)));
  Currentpasswd := FindEtReplace(Currentpasswd,'-','',true);
  THValComboBox (GetControl('TBASENAME')).Enabled := false;
  XX := TconnectBSV.create;
  GetParamsUploadBSV ( TOBPARAMBSV,XX);
  THEdit(GetControl('BP2_SERVERNAME')).OnExit := ChangeParamsConnexion;
  THEdit(GetControl('BP2_PORT')).OnExit := ChangeParamsConnexion;
  THValComboBox (GetControl('TBASENAME')).OnChange := ChangeArchive;
  THValComboBox (GetControl('WFVALIDBAST')).OnChange := ChangeWFVALIDBAST;
  THValComboBox (GetControl('CHPSETAT')).OnChange := ChangeChpsEtat;
  THValComboBox (GetControl('VALETATREG')).OnChange := ChangeCodeEtatReg;
  TToolbarButton97 (GetControl('BCONNECTARCHIVE')).OnClick := ArchiveTest;
end ;

procedure TOM_BSVSERVER.OnClose ;
begin
  if XX.connected then XX.Disconnect;
  XX.Free;
  TOBPARAMBSV.free;
  Inherited ;
end ;

procedure TOM_BSVSERVER.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_BSVSERVER.OnLoadAlerte;
begin
  Inherited ;
end ;

procedure TOM_BSVSERVER.OnAfterBeginTrans;
begin
  Inherited ;
end ;

procedure TOM_BSVSERVER.OnAfterCommit;
begin
  Inherited ;
end ;

procedure TOM_BSVSERVER.OnAfterRollBack;
begin
  Inherited ;
end ;

procedure TOM_BSVSERVER.ConnecteEtRemplit;
begin
  //
  if XX.connected then Exit;
  //
  XX.BSVServer := GetControlText('BP2_SERVERNAME');
  XX.BSVPORT := GetControlText('BP2_PORT');
  TRY
    if XX.ConnectToServer(CurrentUser,CurrentPasswd) then
    begin
      THLabel(GetControl('SERVERSTATUS')).Caption := 'Connecté';
      THLabel(GetControl('SERVERSTATUS')).Font.color := clGreen;
      SetServerConnected;
    end else
    begin
      THLabel(GetControl('SERVERSTATUS')).Caption := 'Non Connecté';
      THLabel(GetControl('SERVERSTATUS')).Font.color := clRed;
    end;
  EXCEPT
    on E : Exception do
    begin
      PGIInfo(e.Message);
    end;
  END;
end;

procedure TOM_BSVSERVER.SetServerConnected;
var II : Integer;
    AA : TArchive;
    TT,TPERE  : TOB;
begin
  THValComboBox (GetControl('TBASENAME')).Clear;
  THValComboBox (GetControl('TBASENAME')).Items.insert(0,'<<Aucune>>');
  THValComboBox (GetControl('TBASENAME')).values.insert(0,'0');
  THValComboBox (GetControl('TBASENAME')).ItemIndex := 0;
  //
  THValComboBox (GetControl('WFVALIDBAST')).Clear;
  THValComboBox (GetControl('WFVALIDBAST')).Items.insert(0,'<<Aucun>>');
  THValComboBox (GetControl('WFVALIDBAST')).values.insert(0,'0');
  THValComboBox (GetControl('WFVALIDBAST')).ItemIndex := 0;
  //
  THValComboBox (GetControl('CHPSETAT')).clear;
  THValComboBox (GetControl('CHPSETAT')).Items.insert(0,'<<Aucun>>');
  THValComboBox (GetControl('CHPSETAT')).values.insert(0,'0');
  THValComboBox (GetControl('CHPSETAT')).ItemIndex := 0;
  //
  THValComboBox (GetControl('VALETATREG')).clear;
  THValComboBox (GetControl('VALETATREG')).Items.insert(0,'<<Aucun>>');
  THValComboBox (GetControl('VALETATREG')).values.insert(0,'0');
  THValComboBox (GetControl('VALETATREG')).ItemIndex := 0;
  //
  XX.GetArchivesList;
  if XX.LesArchives.Count > 0 then
  begin
    for II := 0 to XX.LesArchives.Count - 1 do
    begin
      AA := XX.LesArchives.items[II];
      THValComboBox (GetControl('TBASENAME')).Items.Insert(II+1,AA.Name);
      THValComboBox (GetControl('TBASENAME')).Values.Insert(II+1,IntToStr(AA.Id));
      if XX.LesArchives.items[II].Id = THDbSpinEdit(GetControl('BP2_ARCHIVE')).value then THValComboBox (GetControl('TBASENAME')).ItemIndex := II+1;
    end;
  end;
  THValComboBox (GetControl('TBASENAME')).Enabled := True;
  if THDbSpinEdit(GetControl('BP2_ARCHIVE')).value <> 0 then
  begin
    XX.GetWorkFlowList(THDbSpinEdit(GetControl('BP2_ARCHIVE')).value);
    // gestion du workflow validation des BAST
    if XX.LesWorkflows.Count > 0 then
    begin
      for II := 0 to XX.LesWorkflows.Count - 1 do
      begin
        AA := XX.LesWorkflows.items[II];
        THValComboBox (GetControl('WFVALIDBAST')).Items.Insert(II+1,AA.Description);
        THValComboBox (GetControl('WFVALIDBAST')).Values.Insert(II+1,IntToStr(AA.Id));
        if XX.LesWorkflows.items[II].Id = THDbSpinEdit(GetControl('BP2_WFVALIDBAST')).value then THValComboBox (GetControl('WFVALIDBAST')).ItemIndex := II+1;
      end;
    end;
    // Gestion mise à jour état des Factures dans la GED
    if TOBPARAMBSV.detail.Count > 0 then
    begin
      TPERE := nil;
      for II := 0 to TOBPARAMBSV.Detail.Count - 1 do
      begin
        TT := TOBPARAMBSV.detail[II];
        THValComboBox (GetControl('CHPSETAT')).Items.Insert(II+1,TT.GetString('LABEL'));
        THValComboBox (GetControl('CHPSETAT')).Values.Insert(II+1,TT.GetString('NAME'));
        if TT.GetString('NAME') = Thedit(GetControl('BP2_CHPSETAT')).Text then
        begin
          THValComboBox (GetControl('CHPSETAT')).ItemIndex := II+1;
          TPERE := TT;
        end;
      end;
      if (Thedit(GetControl('BP2_CHPSETAT')).Text <> '') then
      begin
        if TPERE <> nil then
        begin
          for II := 0 to TPERE.Detail.Count - 1 do
          begin
            TT := TPERE.detail[II];
            THValComboBox (GetControl('VALETATREG')).Items.Insert(II+1,TT.GetString('VALEUR'));
            THValComboBox (GetControl('VALETATREG')).Values.Insert(II+1,TT.GetString('VALEUR'));
            if TT.GetString('VALEUR') = Thedit(GetControl('BP2_VALETATREG')).Text then
            begin
              THValComboBox (GetControl('VALETATREG')).ItemIndex := II+1;
            end;
          end;
        end;
      end;
    end;
  end;
end;


procedure TOM_BSVSERVER.ChangeParamsConnexion(sender: TObject);
begin
  if (GetControlText('BP2_SERVERNAME') <> '') and  (GetControlText('BP2_PORT')<> '0') then
  begin
    if XX.connected then XX.Disconnect;
    ConnecteEtRemplit;
  end;
end;

procedure TOM_BSVSERVER.ChangeArchive(Sender: TObject);
begin
  SetField ('BP2_ARCHIVE', StrToInt(THValComboBox (GetControl('TBASENAME')).Value) );
end;

procedure TOM_BSVSERVER.ArchiveTest(Sender: TObject);
begin
  if GetField('BP2_ARCHIVE')=0 then exit;
  if XX.connected then if XX.ConnectToArchive(GetField('BP2_ARCHIVE'),VH_GC.BSVUploadOK,VH_GC.BSVOpenDoc) then PGIInfo('Connection Ok') else PGIInfo('Pas de connexion');
end;

procedure TOM_BSVSERVER.ChangeWFVALIDBAST(Sender: TObject);
begin
  SetField ('BP2_WFVALIDBAST', StrToInt(THValComboBox (GetControl('WFVALIDBAST')).Value) );
end;

procedure TOM_BSVSERVER.ChangeChpsEtat(Sender: TObject);
var II : Integer;
    TPere,TT : TOB;
begin
  SetField ('BP2_CHPSETAT', THValComboBox (GetControl('CHPSETAT')).Value );
  //
  THValComboBox (GetControl('VALETATREG')).clear;
  THValComboBox (GetControl('VALETATREG')).Items.insert(0,'<<Aucun>>');
  THValComboBox (GetControl('VALETATREG')).values.insert(0,'0');
  THValComboBox (GetControl('VALETATREG')).ItemIndex := 0;
  //
  if (Thedit(GetControl('BP2_CHPSETAT')).Text <> '') then
  begin
    TPere := TOBPARAMBSV.Findfirst(['NAME'],[Thedit(GetControl('BP2_CHPSETAT')).Text],True);
    if TPERE <> nil then
    begin
      for II := 0 to TPERE.Detail.Count - 1 do
      begin
        TT := TPERE.detail[II];
        THValComboBox (GetControl('VALETATREG')).Items.Insert(II+1,TT.GetString('VALEUR'));
        THValComboBox (GetControl('VALETATREG')).Values.Insert(II+1,TT.GetString('VALEUR'));
        if TT.GetString('VALEUR') = Thedit(GetControl('BP2_VALETATREG')).Text then
        begin
          THValComboBox (GetControl('VALETATREG')).ItemIndex := II+1;
        end;
      end;
    end;
  end;

end;

procedure TOM_BSVSERVER.ChangeCodeEtatReg(Sender: TObject);
begin
  SetField ('BP2_VALETATREG', THValComboBox (GetControl('VALETATREG')).Value );
end;
Initialization
  registerclasses ( [ TOM_BSVSERVER ] ) ;
end.
