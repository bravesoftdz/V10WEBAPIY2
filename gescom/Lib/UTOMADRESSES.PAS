unit uTomAdresses;

interface

uses
  StdCtrls
  ,Controls
  ,Classes
  ,forms
  ,sysutils
  ,ComCtrls
{$IFDEF EAGLCLIENT}
  ,Maineagl,UtileAGL,Spin
{$ELSE}
  ,db,dbctrls,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,hDB
{$ENDIF}
  ,hCtrls
  ,hEnt1
  ,hMsgBox
  ,hTB97
  ,uTom
  ,uTob
  ,wCommuns
  ;

type
  TOM_ADRESSES = Class (TOM)
  private
    LibTitre : String;
    MonoAdresse : Boolean;
    client_particulier :Boolean;
    code_client :String;
    code_tiers :String;
    code_affaire :String;
    StateEdit : Boolean;
    NatureAuxi: String;
    LastNumAdresse: Integer;
    TypeAdresse: String;
    TobZonelibre,TOBADRINTB,TOBCONTACTB : TOB;
    ADR_NumeroContactEnter: integer;
    ValuesContact : MyArrayValue;
    CContact : string;
    procedure ADR_Email_OnDblclick(Sender : tObject);
    procedure SetLastError (Num : integer; ou : string );
    procedure AfficheCivilite;
    procedure BtDupAdresseOnClick(Sender: Tobject);
    procedure CopieAdresse(TobAdresse: Tob);
    procedure FLISTE_OnDblClick( Sender: TObject );
    procedure FindContactBTP(CodeClient: string; NumContact: integer);
    procedure ActiveEvents (Etat : boolean);
    procedure ChangeInfoSUp(Sender: Tobject);
    procedure FindAdrBTP(NumAdresse: integer);
    procedure DeleteAdrBTP(NumAdresse: integer);
  public
    procedure OnChangeField(F : TField); override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnArgument (Arguments : String ); override;
    procedure OnClose   ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnDeleteRecord; override;
    procedure ADR_NUMEROCONTACTOnEnter(Sender: TObject);
    procedure ADR_NUMEROCONTACTOnExit(Sender: TObject);
    procedure CHKLIENADRESSETIERSOnClick(Sender: TObject);
    function  SelectAdresseTiers:boolean;
    procedure AfficheContact;
 end;

const
  { libellés des messages }
  TexteMessage: array[1..5] of string = (
          {1}     'Le destinataire de l''adresse est obligatoire'
          {2}    ,'Le code postal de l''adresse est obligatoire'
          {3}    ,'La ville de l''adresse est obligatoire'
          {4}    ,'La région n''existe pas. Vérifiez dans la liste.'
          {5}    ,'Le code NIF de l''adresse est obligatoire'
                 );

implementation

uses
  SaisieList
  ,uTableFiltre
  ,ParamSoc
  ,FactUtil
  ,Dicobtp
  ,entGC
  ,TiersUtil
  ,M3FP
  ,LookUp
  ,MailOl
  ,FactAdresse
  ;

procedure TOM_ADRESSES.SetLastError(Num : integer; ou : string );
begin
  if ou<>'' then SetFocusControl(ou);
  LastError := Num;
  LastErrorMsg := TexteMessage[LastError];
end ;

procedure TOM_ADRESSES.OnArgument(Arguments : String );
var
  Critere: string;
  ChampMul, ValMul: string;
  x: integer;
  {$IFDEF EAGLCLIENT}
  ListLocked : TStrings;
  {$ENDIF}
begin
  inherited ;
  TOBADRINTB := TOB.Create ('BADRESSES',nil,-1);
  TOBCONTACTB := TOB.Create('BCONTACT',nil,-1);
  {$IFDEF EAGLCLIENT}
  if ecran is TFSaisielist then
    begin
    ListLocked := TStringList.Create;
    ListLocked.add( 'bTree' );
    TFSaisieList( ecran ).HMTrad.LockedCtrls := ListLocked;
    ListLocked.Free;
    end;
  {$ENDIF}
  TobZonelibre:=TOB.Create ('TIERSCOMPL', Nil, -1);
  LibTitre := '';
  MonoAdresse := False;
  client_particulier := False;
  LastNumAdresse := 0;
  { Récup des arguments }
  repeat
    Critere:=Trim(ReadTokenSt(Arguments));
    if Critere<>'' then
    begin
      x:=pos('=',Critere);
      if x<>0 then
      begin
        ChampMul := UpperCase(copy(Critere,1,x-1));
        ValMul := Copy(Critere,x+1,length(Critere));
        if (ChampMul = 'YTC_TIERSLIVRE') then code_tiers := ValMul;
        if (ChampMul = 'TITRE') then LibTitre := ValMul;

        if (ChampMul = 'MONOADRESSE') and (ValMul='OUI') then MonoAdresse := True;
        if (ChampMul = 'PART') and (ValMul = 'X') then Client_particulier := True;
        if (ChampMul = 'CLI') then code_client := ValMul;
        if (ChampMul = 'TIERS') or (ChampMul = 'ADR_REFCODE') then SetControlText('TIERS', ValMul);
        if (ChampMul = 'NATUREAUXI') then NatureAuxi := ValMul;
        if (ChampMul = 'TYPEADRESSE') or (ChampMul = 'ADR_TYPEADRESSE') then TypeAdresse := ValMul;
        if (ChampMul = 'AFFAIRE') then code_affaire := ValMul;
        if (ChampMul = 'CONTACT') then CContact := ValMul;
      end;
    end;
  until Critere = '';
  AfficheCivilite;
  { Bouton d'import de l'adresse principale }
  if Assigned(GetControl('BtDupAdresse')) then
    TToolBarButton97(GetControl('BtDupAdresse')).OnClick := BtDupAdresseOnClick;
// SetControlVisible('BADRESSETIERS', (ctxAffaire in V_PGI.PGIContexte) or (VH_GC.GASeria));

if Assigned(GetControl('ADR_NUMEROCONTACT')) then
begin
   ThEdit(GetControl('ADR_NUMEROCONTACT')).OnEnter := ADR_NUMEROCONTACTOnEnter;
   ThEdit(GetControl('ADR_NUMEROCONTACT')).OnExit := ADR_NUMEROCONTACTOnExit;
end;

Thlabel(GetControl('ADR_EMAIL')).OnDblclick := ADR_EMAIL_OnDblclick;

if Assigned(GetControl('FLISTE')) then THGrid(GetControl('FLISTE' )).OnDblClick := FLISTE_OnDblClick;

SetControlVisible('ADR_NIF',  not Client_Particulier);
SetControlVisible('TADR_NIF', not Client_Particulier);

// cas adresse associé affaire.. des zones doivent être cachée
// 2 champs l'un sur l'autre pour formejuridique. il faut force
//à clt entreprise dans le cas affaire  pour pointer sur TTFormeJuridique
If (typeAdresse='AFA') or (typeadresse='INT') then
  begin // cas adresse associé affaire.. des zones doivent être cachée
  SetControlVisible ('ADR_LIVR',false);
  SetControlVisible ('ADR_FACT',false);
  SetControlVisible ('ADR_REGL',false);
  SetControlVisible ('BA0_INT',false);
  SetControlVisible ('BDUPLIQUER',false);
  Client_particulier := False;  // 2 champs l'un sur l'autre pour formejuridique. il faut force à clt entreprise dans le cas affaire  pour pointer sur TTFormeJuridique
  end;

//pas d'export, ni de livraison
If (CtxScot in V_PGI.PGIContexte) or (CtxTempo in V_PGI.PGIContexte) then
  begin
  SetControlVisible ('ADR_LIVR',false);
  SetControlVisible ('ADR_INCOTERM',false);
  SetControlVisible ('ADR_MODEEXP',false);
  SetControlVisible ('ADR_LIEUDISPO',false);
  SetControlVisible ('TADR_INCOTERM',false);
  SetControlVisible ('TADR_MODEEXP',false);
  SetControlVisible ('TADR_LIEUDISPO',false);
  SetControlVisible ('GBLOGISTIQUE',false);
  end;
  if Assigned(Ecran) and (LibTitre<>'') then
    Ecran.Caption:= 'Adresse : ' + LibTitre;

  // mcd 25/05/04 10998 il faut le code auxi pour limiter les contacts
  if (Code_Client ='') and (GetControlText('TIERS') <>'') then
  	Code_Client :=TiersAuxiliaire (GetControlText('TIERS'), False);
  // ------
  // Laisser le PLUS dans le On argument car si pas d'adresses pour le client, le OnLoad n'est pas activé.
  SetControlProperty('ADR_NUMEROCONTACT', 'PLUS', Code_Client);

end;

procedure TOM_ADRESSES.OnNewRecord;
var
  Q: TQuery;
  Str, Ref: String;
  Numero, Max: Integer;
  Duplique: Boolean;
begin
  inherited;
  Duplique := False;
  if Assigned(Ecran) then
  begin
    Duplique := (TFSaisieList(Ecran).LeFiltre.FiltreState = fsDupliq);
    AfficheCivilite; 
    if (ctxAffaire in V_PGI.PGIContexte) then
    begin
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).OnClick := nil;
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).Checked := false;
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).OnClick := CHKLIENADRESSETIERSOnClick;
    end;
    if duplique then
      { Sauve le N° d'adresse de référence lors d'une duplication }
      SetField('ADR_NUMADRESSEREF', LastNumAdresse)
    else if (TFSaisieList(Ecran).LeFiltre.FiltreState = fsCreat) then
      SetControlEnabled('BTDUPADRESSE', True);
  end;
  
  SetField('ADR_TYPEADRESSE', TypeAdresse);

  { Init de la clé }
  Numero := GetSetNumAdresse;
  if Numero > 0 then
    SetField ('ADR_NUMEROADRESSE', Numero)
  else
    SetField ('ADR_NUMEROADRESSE', 1);

  if (not duplique) then
  begin
    { Init des données }
    if (Getfield ('ADR_TYPEADRESSE') = 'AFA') or (Getfield ('ADR_TYPEADRESSE') = 'INT') then
      SetField('ADR_REFCODE', code_affaire)
    else
      SetField('ADR_REFCODE', GetControlText('TIERS'));

    { Données qui ne sont pas dupliquées }
    //mcd 02/10/2003 ne peut pas marcher pour les type  d'adresses autres TIE  !!!!
    if GetField('ADR_TYPEADRESSE')='TIE' then
      SetField('ADR_NADRESSE', 1 + GetTiersLastAdresseNum(GetControlText('TIERS'), NatureAuxi))
    else
    begin
      Ref := GetField('ADR_REFCODE');
      Q := OpenSQL ('SELECT MAX(ADR_NADRESSE) FROM ADRESSES WHERE ADR_REFCODE="'+REF+'" AND ADR_TYPEADRESSE="'
        + GetField('ADR_TYPEADRESSE')+'"', TRUE) ;
      if Q.EOF then Max := 0 else Max := Q.Fields[0].AsInteger  ;
      Ferme(Q) ;
      SetField('ADR_NADRESSE', 1 + MAX)
    end;

    SetField('ADR_PAYS', string(GetParamsoc('SO_GCTIERSPAYS')));
    SetField('ADR_NATUREAUXI', NatureAuxi);
    SetField('ADR_LIVR', 'X');
    SetField('ADR_FACT', 'X');
    SetField('ADR_REGL', 'X');
	  SetField('BA0_INT',  'X');

    If GetField('ADR_TYPEADRESSE')='AFA' then
    begin  //mcd 02/10/2003 change initialisation pour AFA et INT
      SetField('ADR_LIVR', '-');
      SetField('ADR_FACT', 'X');
      SetField('ADR_REGL', '-');
		  SetField('BA0_INT',  '-');
    end
    else If GetField ('ADR_TYPEADRESSE')='INT' then
    begin
      SetField('ADR_LIVR', '-');
      SetField('ADR_FACT', '-');
      SetField('ADR_REGL', '-');
		  SetField('BA0_INT',  'X');
    end
    else
    begin  // si TIE ou PIE
      SetField('ADR_LIVR', 'X');
      SetField('ADR_FACT', 'X');
      SetField('ADR_REGL', 'X');
		  SetField('BA0_INT',  'X');
    end;
    SetField('ADR_NUMEROCONTACT', wGetSqlFieldValue('C_NUMEROCONTACT', 'CONTACT', 'C_AUXILIAIRE = "' + Code_Client + '" AND C_PRINCIPAL = "' + wTrue + '"'));

    if (Client_Particulier = True) then
    begin
      Q := OpenSql ('SELECT T_JURIDIQUE, T_LIBELLE, T_PRENOM FROM TIERS WHERE T_AUXILIAIRE="'+code_client+'"', True);
      if not Q.Eof then
      begin
        str := Q.Findfield('T_JURIDIQUE').AsString;
        SetField ('ADR_JURIDIQUE', str);
        str := Q.Findfield('T_LIBELLE').AsString;
        SetField ('ADR_LIBELLE', str);
        str := Q.Findfield('T_PRENOM').AsString;
        SetField ('ADR_LIBELLE2', str);
      end;
      Ferme (Q);
    end;
  end;
  { }
  if GetControlVisible('PNJURIDIQUE') then
    SetFocusControl('ADR_JURIDIQUE')
  else
    SetFocusControl('ADR_JURIDIQUE_');
end;

procedure TOM_ADRESSES.FindContactBTP (CodeClient: string; NumContact : integer);
var QQ : TQuery;
begin
  TOBCONTACTB.InitValeurs(false);
  TOBCONTACTB.SetString('BC1_TYPECONTACT','T');
  TOBCONTACTB.SetString('BC1_AUXILIAIRE',CodeClient);
  if NumContact <> 0 then
  begin
    QQ := OpenSql ('SELECT * FROM BCONTACT WHERE BC1_TYPECONTACT="T" AND BC1_AUXILIAIRE="'+Code_Client+'" AND BC1_NUMEROCONTACT='+InttoStr(NumContact),true,1,'',true);
    if not QQ.eof then
    begin
      TOBCONTACTB.SelectDB('',QQ);
    end;
    ferme (QQ);
  end;
end;

procedure TOM_ADRESSES.DeleteAdrBTP (NumAdresse : integer);
begin
  ExecuteSql ('DELETE FROM BADRESSES WHERE BA0_TYPEADRESSE="'+GeTField('ADR_TYPEADRESSE')+'" AND BA0_NUMEROADRESSE='+InttoStr(NumAdresse));
end;

procedure TOM_ADRESSES.FindAdrBTP (NumAdresse : integer);
var QQ : Tquery;
begin
  TOBADRINTB.InitValeurs(false);
  TOBADRINTB.SetString('BA0_TYPEADRESSE',Getfield('ADR_TYPEADRESSE'));
  TOBADRINTB.SetInteger('BA0_NUMEROADRESSE',NumAdresse);

  if NumAdresse <> 0 then
  begin
    QQ := OpenSql ('SELECT * FROM BADRESSES WHERE BA0_TYPEADRESSE="'+GeTField('ADR_TYPEADRESSE')+'" AND BA0_NUMEROADRESSE='+InttoStr(NumAdresse),true,1,'',true);
    if not QQ.eof then
    begin
      TOBADRINTB.SelectDB('',QQ);
    end;
    ferme (QQ);
  end;
end;

procedure TOM_ADRESSES.OnLoadRecord;
begin
  inherited;
  ActiveEvents (false);

  SetControlProperty('ADR_REGION','PLUS','RG_PAYS="'+GetField('ADR_PAYS')+'"');
  SetControlProperty('ADR_NUMEROCONTACT', 'PLUS', Code_Client);
  if (not TobZonelibre.SelectDB('"' + Code_Client + '"',Nil )) then
    TobZonelibre.InitValeurs;

  FindContactBTP (Code_Client,GetField('ADR_NUMEROCONTACT'));
  ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], Code_Client, GetField('ADR_NUMEROCONTACT'));
  AfficheContact;

  ADR_NumeroContactEnter := GetField('ADR_NUMEROCONTACT');

  if not AfterInserting then
  begin
    SetControlEnabled('BTDUPADRESSE', False);
    LastNumAdresse := GetField('ADR_NUMEROADRESSE');
  end;

  if Assigned(Ecran) and (LibTitre<>'') then Ecran.Caption:= 'Adresse : ' + LibTitre;

  AfficheCivilite;

  if MonoAdresse then SetControlVisible('BInsert',False);

  { Spécificité adresse de facturation et d'intervention sur affaire }
  if (Getfield ('ADR_TYPEADRESSE') = 'AFA') or (Getfield ('ADR_TYPEADRESSE') = 'INT') then
  begin
    if Assigned(Ecran) and (LibTitre <> '') then
      Ecran.Caption:= LibTitre;
    if not VH_GC.GCIfDefCEGID then
    begin
      SetControlVisible('GB_ADRESSETIERS', True);
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).OnClick := nil;
      if (Getfield('ADR_NUMADRESSEREF') <> 0) then
        TCheckBox(GetControl('CHKLIENADRESSETIERS')).Checked:=true;
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).OnClick := CHKLIENADRESSETIERSOnClick;
    end;
  end;
  FindAdrBTP (GetField('ADR_NUMEROADRESSE'));
  TOBADRINTB.PutEcran(ecran);
  ActiveEvents (true);
end;

procedure TOM_ADRESSES.OnUpdateRecord;
var
  numero: Integer;
  NoTel : String;
  CleTel: String;
  Car	  : String;
  Indice: Integer;
begin
  Inherited;

  if Assigned(Ecran) then
    StateEdit := (TFSaisieList(Ecran).LeFiltre.FiltreState = fsModif);

  { Tests de la saisie }
  if (GetField ('ADR_LIBELLE')= '') then
  begin
    SetLastError(1, 'ADR_LIBELLE');
    EXIT;
  end;

  if (GetField ('ADR_CODEPOSTAL')= '') then
  begin
    SetLastError(2, 'ADR_CODEPOSTAL');
    EXIT;
  end;

  if (GetField ('ADR_VILLE')= '') then
  begin
    SetLastError(3, 'ADR_VILLE');
    EXIT;
  end;

  If (GetField ('ADR_PAYS') <> '') and (GetControlText('ADR_REGION') <> '') and (not existesql('SELECT RG_PAYS FROM REGION WHERE RG_PAYS = "'+string(GetField ('ADR_PAYS'))+'" AND RG_REGION = "'+string(GetField ('ADR_REGION'))+'"')) then
  Begin
    SetLastError(4, 'ADR_REGION');
    exit ;
  end ;

  { Code NIF non obligatoire pour l'instant mais peut-être un jour ...}
{
  if (NATUREAUXI='CLI') and (GetField ('ADR_NIF')= '') and (not Client_Particulier) then
  begin
    SetLastError(5, 'ADR_NIF');
    EXIT;
  end;
}
  if Assigned(Ecran) then
  begin
    if TFSaisieList(Ecran).LeFiltre.FiltreState in [fsCreat, fsDupliq] then
    begin
      { Numéro de l'adresse pour le tiers (saisie concurrentielle) }
      If GetField ('ADR_TYPEADRESSE') <> 'TIE' then exit; //mcd 02/10/2003
      Numero := 1 + GetTiersLastAdresseNum(GetControlText('TIERS'), NatureAuxi);
      if GetField('ADR_NADRESSE') <> Numero then
      begin
        PGIInfo(TraduireMemoire('Le numéro de l''adresse sera') + ' : ' + IntToStr(Numero));
        SetField('ADR_NADRESSE', Numero);
      end;
    end;
  end;

 	NoTel := GetField('ADR_TELEPHONE');

  for indice:=1 to length(NoTel) do
	  begin
    Car:=copy(NoTel,indice,1);
    if (IsNumeric(car)) and (car <>'.') and (car<>',') and (car<>'-') and (car<>' ') then
       CleTel:=CleTel+car;
  	end;

	if GetParamSoc('SO_RTSUPPZERO') then CleTel := '0'+ CleTel;
  
  TOBADRINTB.GetEcran(ecran);
  TOBADRINTB.SetString('BA0_TYPEADRESSE',Getfield('ADR_TYPEADRESSE'));
  TOBADRINTB.SetInteger('BA0_NUMEROADRESSE',Getfield('ADR_NUMEROADRESSE'));
  TOBADRINTB.PutValue('BA0_CLETELEPHONE',Cletel);
  If GetField('ADR_TYPEADRESSE')='AFA' then TOBADRINTB.PutValue('BA0_INT','-') else TOBADRINTB.PutValue('BA0_INT','X'); 
  TOBADRINTB.SetAllModifie(true);
  TOBADRINTB.InsertOrUpdateDB(false);

end;

procedure TOM_ADRESSES.OnAfterUpdateRecord;
var
  Adr, AdrRef: integer;
begin
  inherited;
  if not VH_GC.GCIfDefCEGID then
      if (ctxAffaire in V_PGI.PGIContexte) or (VH_GC.GASeria) then
      begin
        if StateEdit then
        begin
          if (GetControl('GB_ADRESSETIERS').Visible = True) then
          begin
            // Sur les adresses Affaires ou propositions
            AdrRef:= GetField('ADR_NUMADRESSEREF');
            if (AdrRef <> 0) then
            begin
              if (PGIAsk(TraduitGA('Vous avez validé des modifications sur une adresse liée à l''adresse d''un tiers.' + chr(13) + 'Voulez-vous conserver ce lien et répercuter les modifications sur toutes les adresses liées ?'), TFSaisieList(ecran).Caption) = mrYes) then
              begin
                Adr:= GetField('ADR_NUMEROADRESSE');
                // req de répercution des modifs de l'adresse de l'affaire sur l'adresse du tiers
                if Not ModifierAdrMereParFille( inttostr(Adr) ) then exit;
                // Req de répercution des modifs de l'adresse du tiers sur toutes les adresses liées
                ModifierAdrFillesParMere( inttostr(AdrRef) );
                TFSaisieList(Ecran).Close;
              end
              else
              begin
                DS.Edit;  // pour passer DS.state en mode dsEdit
                SetField('ADR_NUMADRESSEREF', 0);
                TFSaisieList(Ecran).LeFiltre.Post;
              end;
            end
            else
            begin
              if (Getfield ('ADR_TYPEADRESSE') = 'TIE') then
              begin
                // Cas de modif de l'adresse du tiers
                Adr := GetField('ADR_NUMEROADRESSE');
                // Req de répercution des modifs de l'adresse du tiers sur toutes les adresses liées
                ModifierAdrFillesParMere( inttostr(Adr) );
                TFSaisieList(Ecran).Close;
              end;
            end;
          end;
        end;
      end;
end;

procedure TOM_ADRESSES.OnChangeField(F: TField);
begin
  Inherited;
  if (DS.State<>dsBrowse) and (F.FieldName = 'ADR_PAYS') and (GetField('ADR_PAYS')<>'') then
  begin
    SetControlProperty('ADR_REGION','PLUS','RG_PAYS="'+GetField('ADR_PAYS')+'"');
  end;
end;

procedure TOM_ADRESSES.ADR_Email_OnDblclick(Sender : tObject);
begin
  SendMail('',GetControlText('ADR_EMAIL'),'',nil,'',False);
end;

procedure TOM_ADRESSES.CHKLIENADRESSETIERSOnClick(Sender: TObject);
begin
  if not(DS.State in [dsInsert,dsEdit]) then
    DS.edit; // pour passer DS.state en mode dsEdit

  if not TCheckBox(GetControl('CHKLIENADRESSETIERS')).Checked then
  begin
    if (PGIAsk( TraduitGA( 'Attention, vous allez perdre le lien entre l''adresse en cours et l''adresse du tiers d''origine.' + chr(13) + 'Voulez-vous confirmer ?') ,TFSaisieList(ecran).Caption) = mrYes) then
    begin
      SetField('ADR_NUMADRESSEREF', 0);
    end
    else
    begin
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).OnClick := nil;
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).Checked := true;
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).OnClick := CHKLIENADRESSETIERSOnClick;
    end;
  end
  else
  begin
    if Not SelectAdresseTiers then
    begin
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).OnClick := nil;
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).Checked := false;
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).OnClick := CHKLIENADRESSETIERSOnClick;
    end;
  end;
end;

function TOM_ADRESSES.SelectAdresseTiers:boolean;
var
  rep: string;
  TobAdresses: TOB;
  Q: TQuery;
begin
  Result := False;
  if not(DS.State in [dsInsert,dsEdit]) then
    DS.edit; // pour passer DS.state en mode dsEdit
  Rep := AGLLanceFiche('AFF','AFADRTIERS_MUL','','','TIERS=' + Code_tiers + ';ACTION=CONSULTATION');
  if (Rep <> '') then
	begin
    TobAdresses := Tob.Create('ADRESSES',nil,-1);
    Q := Nil;
    try
      Q:=OpenSQL('SELECT * FROM ADRESSES WHERE ADR_NUMEROADRESSE="'+rep+'"',True) ;
      if Not TobAdresses.SelectDB('', Q) then exit;
    finally
      if Q <> Nil then ferme(Q) ;
    end;

    SetField('ADR_NUMADRESSEREF', TobAdresses.GetValue('ADR_NUMEROADRESSE'));
    CopieAdresse(TobAdresses);
	  SetField('ADR_JURIDIQUE', TobAdresses.GetValue('ADR_JURIDIQUE'));
	  SetField('ADR_LIBELLE', TobAdresses.GetValue('ADR_LIBELLE'));
	  SetField('ADR_LIBELLE2', TobAdresses.GetValue('ADR_LIBELLE2'));
	  SetField('ADR_ADRESSE1', TobAdresses.GetValue('ADR_ADRESSE1'));
	  SetField('ADR_ADRESSE2', TobAdresses.GetValue('ADR_ADRESSE2'));
	  SetField('ADR_ADRESSE3', TobAdresses.GetValue('ADR_ADRESSE3'));
	  SetField('ADR_CODEPOSTAL', TobAdresses.GetValue('ADR_CODEPOSTAL'));
	  SetField('ADR_VILLE', TobAdresses.GetValue('ADR_VILLE'));
	  SetField('ADR_PAYS', TobAdresses.GetValue('ADR_PAYS'));
	  SetField('ADR_TELEPHONE', TobAdresses.GetValue('ADR_TELEPHONE'));
	  SetField('ADR_EMAIL', TobAdresses.GetValue('ADR_EMAIL'));
	  SetField('ADR_BLOCNOTE', TobAdresses.GetValue('ADR_BLOCNOTE'));

    TCheckBox(GetControl('CHKLIENADRESSETIERS')).OnClick := nil;
    if not VH_GC.GCIfDefCEGID then
      TCheckBox(GetControl('CHKLIENADRESSETIERS')).Checked := true;
    TCheckBox(GetControl('CHKLIENADRESSETIERS')).OnClick := CHKLIENADRESSETIERSOnClick;
    Result:=true;
    TobAdresses.Free;
	end;
end;

procedure TOM_ADRESSES.AfficheCivilite;
begin
  if (Client_Particulier = True) then
  begin
    SetControlText('TADR_LIBELLE', '&Nom');
//    SetControlText('TADR_LIBELLE2', '&Prénom'); gm pas sur la fiche ?
    SetControlText('TADR_JURIDIQUE', '&Civilité');
  end;
  SetControlVisible('PNCIVILITE', Client_Particulier);
  SetControlVisible('PNJURIDIQUE', not Client_Particulier);
end;

procedure TOM_ADRESSES.BtDupAdresseOnClick(Sender: Tobject);
var
  TobAdresse: Tob;
begin
  if PgiAsk(TraduireMemoire('Importer l''adresse principale du Tiers ?')) = mrYes then
  begin
    TobAdresse := Tob.Create('ADRESSES', nil, -1);
    try
      GetAdrFromCode(TobAdresse, GetControlText('TIERS'), True);   { Charge la TobAdresse }
      CopieAdresse(TobAdresse);
      SetControlEnabled('BTDUPADRESSE', False);
    finally
      TobAdresse.Free;
    end;
  end;
end;

procedure TOM_ADRESSES.CopieAdresse(TobAdresse: Tob);
begin
  SetField('ADR_JURIDIQUE'    , TobAdresse.GetValue('ADR_JURIDIQUE'));
  SetField('ADR_LIBELLE'      , TobAdresse.GetValue('ADR_LIBELLE'));
  SetField('ADR_LIBELLE2'     , TobAdresse.GetValue('ADR_LIBELLE2'));
  SetField('ADR_ADRESSE1'     , TobAdresse.GetValue('ADR_ADRESSE1'));
  SetField('ADR_ADRESSE2'     , TobAdresse.GetValue('ADR_ADRESSE2'));
  SetField('ADR_ADRESSE3'     , TobAdresse.GetValue('ADR_ADRESSE3'));
  SetField('ADR_CODEPOSTAL'   , TobAdresse.GetValue('ADR_CODEPOSTAL'));
  SetField('ADR_VILLE'        , TobAdresse.GetValue('ADR_VILLE'));
  SetField('ADR_PAYS'         , TobAdresse.GetValue('ADR_PAYS'));
  SetField('ADR_TELEPHONE'    , TobAdresse.GetValue('ADR_TELEPHONE'));
  SetField('ADR_EMAIL'        , TobAdresse.GetValue('ADR_EMAIL'));
  SetField('ADR_BLOCNOTE'     , TobAdresse.GetValue('ADR_BLOCNOTE'));
  SetField('ADR_INCOTERM'     , TobAdresse.GetValue('ADR_INCOTERM'));   
  SetField('ADR_MODEEXP'      , TobAdresse.GetValue('ADR_MODEEXP'));    
  SetField('ADR_LIEUDISPO'    , TobAdresse.GetValue('ADR_LIEUDISPO'));
  SetField('ADR_EAN'          , TobAdresse.GetValue('ADR_EAN'));
  SetField('ADR_NIF'          , TobAdresse.GetValue('ADR_NIF'));
  SetField('ADR_REGION'       , TobAdresse.GetValue('ADR_REGION'));
end;

procedure AGLSelectAdresseTiers(parms: array of variant; nb: integer);
var
  F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFSaisieList) then 
    OM := TFSaisieList(F).OM
  else
    exit;

  if (OM is TOM_ADRESSES) then
    TOM_ADRESSES(OM).SelectAdresseTiers()
  else
    exit;
end;


procedure TOM_ADRESSES.OnClose;
begin
  inherited;
  TobZonelibre.free;
  TOBADRINTB.free;
  TOBCONTACTB.free;
end;

procedure TOM_ADRESSES.ADR_NUMEROCONTACTOnEnter(Sender: TObject);
begin
  { Sauve le contact en entrée }
   ADR_NumeroContactEnter := ValeurI(GetControlText('ADR_NUMEROCONTACT'));
end;

procedure TOM_ADRESSES.ADR_NUMEROCONTACTOnExit(Sender: TObject);
begin
  { Si le n° du contact à changé  }
  if ADR_NumeroContactEnter <> ValeurI(GetField('ADR_NUMEROCONTACT')) then     
  begin
    if (LookupValueExist(GetControl('ADR_NUMEROCONTACT'))) then
    begin
      FindContactBTP (Code_Client,GetField('ADR_NUMEROCONTACT'));
      //
      ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], Code_Client, ValeurI(GetField('ADR_NUMEROCONTACT')));
      //
      SetControlText('C_CONTACT',ValuesContact[0]);
      SetControlText('C_CONTACTPRENOM', ValuesContact[1]);

      SetControlText('BC1_BATIMENT', TOBCONTACTB.GetString('BC1_BATIMENT'));
      SetControlText('BC1_ETAGE', TOBCONTACTB.GetString('BC1_ETAGE'));
      SetControlText('BC1_ESCALIER', TOBCONTACTB.GetString('BC1_ESCALIER'));
      SetControlText('BC1_PORTE', TOBCONTACTB.GetString('BC1_PORTE'));
      //
    end else
    begin
      SetField('ADR_NUMEROCONTACT', 0);
      SetControlText('C_CONTACT', '');
      SetControlText('C_CONTACTPRENOM', '');
      SetControlText('BC1_BATIMENT', '');
      SetControlText('BC1_ETAGE', '');
      SetControlText('BC1_ESCALIER', '');
      SetControlText('BC1_PORTE', '');
    end;
  end
end;

procedure TOM_ADRESSES.FLISTE_OnDblClick(Sender: TObject);
begin
  TFSaisieList(Ecran).retour := GetField('ADR_NUMEROADRESSE');
  TFSaisieList(Ecran).Close;
end;

procedure TOM_ADRESSES.OnCancelRecord;
begin
  inherited;
  FindContactBTP (Code_Client,Adr_NumeroContactEnter);
  ValuesContact:= GetNomsFromContact(['C_NOM','C_PRENOM'], Code_Client, Adr_NumeroContactEnter);
  AfficheContact;
  ActiveEvents (false);
  FindAdrBTP (LastNumAdresse);
  TOBADRINTB.PutEcran(ecran);
  ActiveEvents (true);
end;

procedure TOM_ADRESSES.AfficheContact;
begin
  if ValuesContact <> nil then
  begin
    SetControlText('C_CONTACT',ValuesContact[0]);
    SetControlText('C_CONTACTPRENOM', ValuesContact[1]);

    SetControlText('BC1_BATIMENT', TOBCONTACTB.GetString('BC1_BATIMENT'));
    SetControlText('BC1_ETAGE', TOBCONTACTB.GetString('BC1_ETAGE'));
    SetControlText('BC1_ESCALIER', TOBCONTACTB.GetString('BC1_ESCALIER'));
    SetControlText('BC1_PORTE', TOBCONTACTB.GetString('BC1_PORTE'));
  end else
  begin
    SetControlText('C_CONTACT', '');
    SetControlText('C_CONTACTPRENOM', '');
    SetControlText('BC1_BATIMENT', '');
    SetControlText('BC1_ETAGE', '');
    SetControlText('BC1_ESCALIER', '');
    SetControlText('BC1_PORTE', '');
  end;
end;

procedure TOM_ADRESSES.ChangeInfoSUp (Sender : Tobject);
begin
  if not (DS.State in [dsInsert, dsEdit]) then
  begin
    DS.edit;
    SetControlEnabled('bPost',true);
    SetControlEnabled('bRechercher',false);
    SetControlEnabled('bFirst',false);
    SetControlEnabled('bPrev',false);
    SetControlEnabled('bNext',false);
    SetControlEnabled('bLast',false);
    SetControlEnabled('bDefaire',true);
    SetControlEnabled('bInsert',false);
    SetControlEnabled('bDelete',false);
    SetControlEnabled('bDupliquer',false);
  end;
end;

procedure TOM_ADRESSES.ActiveEvents(Etat: boolean);
begin

  if Etat then
  begin
    if assigned(GetControl ('BA0_GARDIEN'))     then THEdit(GetControl ('BA0_GARDIEN')).OnChange := ChangeInfoSUp;
    if assigned(GetControl ('BA0_CODEENTREE1')) Then THEdit(GetControl ('BA0_CODEENTREE1')).OnChange := ChangeInfoSUp;
    if assigned(GetControl ('BA0_CODEENTREE1')) Then THEdit(GetControl ('BA0_CODEENTREE2')).OnChange := ChangeInfoSUp;
  end
  else
  begin
    if assigned(GetControl ('BA0_GARDIEN'))     then THEdit(GetControl ('BA0_GARDIEN')).OnChange := nil;
    if assigned(GetControl ('BA0_CODEENTREE1')) then THEdit(GetControl ('BA0_CODEENTREE1')).OnChange := nil;
    if assigned(GetControl ('BA0_CODEENTREE2')) then THEdit(GetControl ('BA0_CODEENTREE2')).OnChange := nil;
  end;

end;

procedure TOM_ADRESSES.OnDeleteRecord;
begin
  inherited;
  DeleteAdrBTP (GetField('ADR_NUMEROADRESSE'));
end;

Initialization
  RegisterClasses([TOM_ADRESSES]);
  RegisterAglProc( 'SelectAdresseTiers', TRUE , 0, AGLSelectAdresseTiers);
end.

