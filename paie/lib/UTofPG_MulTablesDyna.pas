{***********UNITE*************************************************
Auteur  ...... : GGU
Créé le ...... : 07/02/2007
Modifié le ... :
Description .. : Unit de gestion du multi critère des tables dynamiques
Suite ........ :
Mots clefs ... : PAIE
*****************************************************************}
{
PT1 13/12/2007 V8 GGU FQ 15058 Gestion des sens pour les critères des tables dynamiques
PT5 09/10/2008 FC FQ 15826 changement de sens sur tous les dossiers la première fois qu'on va dans les tables
               dynamiques d'un dossier
}
Unit UTofPG_MulTablesDyna;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
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
     UTOF ;

Type
  TOF_PGMulTablesDyna = Class (TOF)
  private
    procedure OnChangeNIVSAIS(Sender : TObject);
    procedure OnChangeNATURETABLE(Sender : TObject);
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation

Uses
  ParamSoc, Ent1; //PT1


procedure TOF_PGMulTablesDyna.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_PGMulTablesDyna.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_PGMulTablesDyna.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_PGMulTablesDyna.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMulTablesDyna.OnArgument (S : String ) ;
var
//PT5  ValNewSens : Boolean; //PT1
  ValNewSens,AffMsg : String; //PT5
  Q : TQuery;  //PT1
begin
  Inherited ;
  //Debut PT1
  { Gestion du changement de gestion des sens lors du passage V08.02.000.008
    à V08.02.000.010 }
  { On regarde le paramsoc pour savoir si le changement à eu lieu }
  //DEB PT5
  //En PCL on regarde le paramsoc de la DB000000 pour le faire le changement qu'une seule fois
  //sinon on regarde le paramsoc de la base
  if V_PGI.ModePCL='1' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM DB000000.DBO.PARAMSOC WHERE SOC_NOM = "SO_TDYCHGTSENS"', TRUE);
    if not Q.EOF then ValNewSens := Q.FindField('SOC_DATA').AsString;
    Ferme(Q);
  end
  else
  begin
    if GetParamSocSecur('SO_TDYCHGTSENS', True) then
      ValNewSens := 'X'
    else
      ValNewSens := '-';
  end;
  //FIN PT5
  if ValNewSens = '-' then  //PT5
  begin
    { Si il n'a pas eu lieu, on regarde si un paramétrage des tables a été fait }
    Q := OpenSQL('SELECT COUNT(*) FROM TABLEDIMENT', True, 1);
    if Q.Fields[0].AsInteger > 0 then
    begin
      { Si des tables dyna sont paramétrées, on lance les requetes pour changer les sens }
      ExecuteSQL('Update tablediment set PTE_SENS1 = "I" WHERE PTE_SENS1 = "<"');
      ExecuteSQL('Update tablediment set PTE_SENS1 = "IE" WHERE PTE_SENS1 = "<="');
      ExecuteSQL('Update tablediment set PTE_SENS1 = "S" WHERE PTE_SENS1 = ">"');
      ExecuteSQL('Update tablediment set PTE_SENS1 = "SE" WHERE PTE_SENS1 = ">="');
      ExecuteSQL('Update tablediment set PTE_SENS1 = ">" WHERE PTE_SENS1 = "I"');
      ExecuteSQL('Update tablediment set PTE_SENS1 = ">=" WHERE PTE_SENS1 = "IE"');
      ExecuteSQL('Update tablediment set PTE_SENS1 = "<" WHERE PTE_SENS1 = "S"');
      ExecuteSQL('Update tablediment set PTE_SENS1 = "<=" WHERE PTE_SENS1 = "SE"');
      ExecuteSQL('Update tablediment set PTE_SENS2 = "I" WHERE PTE_SENS2 = "<"');
      ExecuteSQL('Update tablediment set PTE_SENS2 = "IE" WHERE PTE_SENS2 = "<="');
      ExecuteSQL('Update tablediment set PTE_SENS2 = "S" WHERE PTE_SENS2 = ">"');
      ExecuteSQL('Update tablediment set PTE_SENS2 = "SE" WHERE PTE_SENS2 = ">="');
      ExecuteSQL('Update tablediment set PTE_SENS2 = ">" WHERE PTE_SENS2 = "I"');
      ExecuteSQL('Update tablediment set PTE_SENS2 = ">=" WHERE PTE_SENS2 = "IE"');
      ExecuteSQL('Update tablediment set PTE_SENS2 = "<" WHERE PTE_SENS2 = "S"');
      ExecuteSQL('Update tablediment set PTE_SENS2 = "<=" WHERE PTE_SENS2 = "SE"');
      ExecuteSQL('Update tablediment set PTE_SENSINT = "I" WHERE PTE_SENSINT = "<"');
      ExecuteSQL('Update tablediment set PTE_SENSINT = "IE" WHERE PTE_SENSINT = "<="');
      ExecuteSQL('Update tablediment set PTE_SENSINT = "S" WHERE PTE_SENSINT = ">"');
      ExecuteSQL('Update tablediment set PTE_SENSINT = "SE" WHERE PTE_SENSINT = ">="');
      ExecuteSQL('Update tablediment set PTE_SENSINT = ">" WHERE PTE_SENSINT = "I"');
      ExecuteSQL('Update tablediment set PTE_SENSINT = ">=" WHERE PTE_SENSINT = "IE"');
      ExecuteSQL('Update tablediment set PTE_SENSINT = "<" WHERE PTE_SENSINT = "S"');
      ExecuteSQL('Update tablediment set PTE_SENSINT = "<=" WHERE PTE_SENSINT = "SE"');
      { et on met à jour le paramsoc pour désactiver la Maj et afficher des messages}
      //DEB PT5 On met à jour les paramsoc de la DB000000 en mode PCL
      if V_PGI.ModePCL='1' then
      begin
        ExecuteSQL('UPDATE DB000000.DBO.PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM = "SO_TDYCHGTSENS"');
        ExecuteSQL('UPDATE DB000000.DBO.PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM = "SO_TDYAFFMSGCHGTSENS"');
      end
      else
      begin
        SetParamSoc('SO_TDYCHGTSENS', True);
        SetParamSoc('SO_TDYAFFMSGCHGTSENS', True);
      end;
      //FIN PT5
    end else begin
      { Sinon, on met à jour le paramsoc pour désactiver la Maj et ne pas mettre de messages : (paramsoc à 11) }
      //DEB PT5 On met à jour les paramsoc de la DB000000 en mode PCL
      if V_PGI.ModePCL='1' then
      begin
        ExecuteSQL('UPDATE DB000000.DBO.PARAMSOC SET SOC_DATA="X" WHERE SOC_NOM = "SO_TDYCHGTSENS"');
        ExecuteSQL('UPDATE DB000000.DBO.PARAMSOC SET SOC_DATA="-" WHERE SOC_NOM = "SO_TDYAFFMSGCHGTSENS"');
      end
      else
      begin
        SetParamSoc('SO_TDYCHGTSENS', True);
        SetParamSoc('SO_TDYAFFMSGCHGTSENS', False);
      end;
      //FIN PT5
    end;
{$IFDEF EAGLCLIENT}
    AvertirCacheServer('PARAMSOC');
{$ENDIF}
    RechargeParamSoc;
    Ferme(Q);
  end;
  { On regarde à nouveau le paramsoc pour savoir si le changement à eu lieu }
  //DEB PT5
  if V_PGI.ModePCL='1' then
  begin
    Q := OpenSql('SELECT SOC_DATA FROM DB000000.DBO.PARAMSOC WHERE SOC_NOM = "SO_TDYCHGTSENS"', TRUE);
    if not Q.EOF then ValNewSens := Q.FindField('SOC_DATA').AsString;
    Ferme(Q);
  end
  else
  begin
    if GetParamSocSecur('SO_TDYCHGTSENS', True) then
      ValNewSens := 'X'
    else
      ValNewSens := '-';
  end;
  if ValNewSens = 'X' then
  begin
    if V_PGI.ModePCL='1' then
    begin
      Q := OpenSql('SELECT SOC_DATA FROM DB000000.DBO.PARAMSOC WHERE SOC_NOM = "SO_TDYAFFMSGCHGTSENS"', TRUE);
      if not Q.EOF then AffMsg := Q.FindField('SOC_DATA').AsString;
      Ferme(Q);
    end
    else
    begin
      if GetParamSocSecur('SO_TDYAFFMSGCHGTSENS', False) then
        AffMsg := 'X'
      else
        AffMsg := '-';
    end;
    if AffMsg = 'X' then
    //FIN PT5
    begin
      if (PGIAsk('la convention des sens de tables dynamiques a changé. Nous vous invitons à consulter l''aide détaillée.#10#13#10#13#10#13Ne plus afficher ce message ?', 'Tables dynamiques') = mrYes) then
      begin
        //DEB PT5
        if V_PGI.ModePCL='1' then
          ExecuteSQL('UPDATE DB000000.DBO.PARAMSOC SET SOC_DATA="-" WHERE SOC_NOM = "SO_TDYAFFMSGCHGTSENS"')
        else
          SetParamSoc('SO_TDYAFFMSGCHGTSENS', False);
        //FIN PT5
{$IFDEF EAGLCLIENT}
        AvertirCacheServer('PARAMSOC');
{$ENDIF}
        RechargeParamSoc;
      end;
    end;
  end;
  //Fin PT1
  if GetControl('PTE_NIVSAIS') is THValComboBox then
    (GetControl('PTE_NIVSAIS') as THValComboBox).OnChange := OnChangeNIVSAIS;
  if GetControl('PTE_NATURETABLE') is THValComboBox then
    (GetControl('PTE_NATURETABLE') as THValComboBox).OnChange := OnChangeNATURETABLE;
end ;

procedure TOF_PGMulTablesDyna.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_PGMulTablesDyna.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_PGMulTablesDyna.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_PGMulTablesDyna.OnChangeNIVSAIS(Sender: TObject);
var
  st : string;
begin
  st := GetControlText('PTE_NIVSAIS');
  if GetControl('PTE_NIVSAIS') is THValComboBox then
  begin
    if st = 'ETB' then
    begin
      setControlProperty('PTE_VALNIV','DataType','TTETABLISSEMENT');
      setControlProperty('PTE_VALNIV','Enabled',True);
      setControlProperty('TPTE_VALNIV','Enabled',True);
    end else if st = 'CON' then
    begin
      setControlProperty('PTE_VALNIV','DataType','PGCONVENTIONS');
      setControlProperty('PTE_VALNIV','Enabled',True);
      setControlProperty('TPTE_VALNIV','Enabled',True);
    end else begin
      setControlText('PTE_VALNIV','');
      // L'affectation de '' ne fonctionne pas, la dernière tablette associé reste
      setControlProperty('PTE_VALNIV','DataType','');
      setControlProperty('PTE_VALNIV','Enabled',False);
      setControlProperty('TPTE_VALNIV','Enabled',False);
    end;
    (GetControl('PTE_VALNIV') as THValComboBox).Reload ;
  end;
end;

procedure TOF_PGMulTablesDyna.OnChangeNATURETABLE(Sender: TObject);
var
  st : string;
begin
  st := GetControlText('PTE_NATURETABLE');
  if GetControl('PTE_NATURETABLE') is THValComboBox then
  begin
    if st = 'DSA' then
    begin
      setControlProperty('PTE_NIVSAIS','Enabled',True);
      setControlProperty('TPTE_NIVSAIS','Enabled',True);
      setControlProperty('PTE_VALNIV','Enabled',False);
      setControlProperty('TPTE_VALNIV','Enabled',False);
      setControlText('PTE_VALNIV','');
    end else begin
      setControlProperty('PTE_NIVSAIS','Enabled',False);
      setControlProperty('TPTE_NIVSAIS','Enabled',False);
      setControlText('PTE_NIVSAIS','');
      setControlProperty('PTE_VALNIV','Enabled',False);
      setControlProperty('TPTE_VALNIV','Enabled',False);
      setControlText('PTE_VALNIV','');
    end;
  end;
end;


Initialization
  registerclasses ( [ TOF_PGMulTablesDyna ] ) ;
end.
