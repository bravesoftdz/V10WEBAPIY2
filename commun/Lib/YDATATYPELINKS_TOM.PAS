{***********UNITE*************************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 13/01/2005
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : YDATATYPELINKS (YDATATYPELINKS)
Mots clefs ... : TOM;YDATATYPELINKS
*****************************************************************}
{
PT1   : 14/06/2007 GGU Adaptation à la paie, ajout d'un filtre qui permet de n'afficher
                       qu'une partie des liens (par exemple ceux de la ged paie en mettant
                       un filtre égal à 'PGLIBGED%')
}
Unit YDATATYPELINKS_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} 
     Fiche, 
     FichList, 
{$else}
     eFiche, 
     eFichList, 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM, 
     UTob ,LicUtil
{$ifdef GIGI}
     ,DicoAf
{$endif}
     ,ParamSoc,
{$IFNDEF PAIEGRH} //PT1
     UtilGc,
{$ENDIF}
     UtilPgi, EntGC ;

Type
  TOM_YDATATYPELINKS = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( StArgument: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
  public
{$IFDEF PAIEGRH} //PT1
//    boGedPaieOnly : Boolean;
    FiltreLinks : String;
{$ENDIF}
  private
    Racine : String; // stocke la racine que doit avoir CodeHdtLink
    TobDEt : Tob;
    procedure ChargeTablette;
  end ;

Implementation

uses StrUtils;
const
  TexteMsg: array[1..10] of string = (
    {1}'Le code liaison doit être renseigné.',
    {2}'La tablette maître doit être renseignée.',
    {3}'La tablette esclave doit être renseignée.',
    {4}'Le type de liaison doit être renseigné.',
    {5}'La tablette esclave doit être différente de la table maître.',
    {6}'Le code liaison doit commencer par ',
    {7}'Le nom de la tablette maître n''existe pas. ',
    {8}'Le nom de la tablette esclave n''existe pas. ',
    {9}'Vous avez 2 fois le même esclave dans votre paramétrage.Cela peut poser des problèmes. ',
   {10}'La table esclave ne fait pas partie de la même nature que la table maître. '
    );

procedure TOM_YDATATYPELINKS.OnNewRecord ;
begin
  Inherited ;
  SetField('YDL_TYPELINK', 'NOR');
  SetField('YDL_PREDEFINI', 'DOS');
  if Racine <>'' then SetField ('YDL_CODEHDTLINK',Racine);
end ;

procedure TOM_YDATATYPELINKS.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_YDATATYPELINKS.OnUpdateRecord ;
var vtob : tob;
  StAbrege, Critere,StAb2 : string;
  BoOk : Boolean;
begin
  Inherited ;
    //mise nom liaison en majuscule
  SetField('YDL_CODEHDTLINK',Uppercase (GetField('YDL_CODEHDTLINK')));
  if GetField('YDL_CODEHDTLINK') = '' then
  begin //le code liaison doit être renseigné
    LastError := 1;
    LastErrorMsg := TexteMsg[LastError];
    exit;
  end;

  if GetField('YDL_MDATATYPE') = '' then
  begin //le code maitre doit exister
    LastError := 2;
    LastErrorMsg := TexteMsg[LastError];
    exit;
  end;

  if GetField('YDL_SDATATYPE') = '' then
  begin   //esclave doit exister
    LastError := 3;
    LastErrorMsg := TexteMsg[LastError];
    exit;
  end;

  if GetField('YDL_TYPELINK') = '' then
  begin   //l etype de liaision doit être renseigner
    LastError := 4;
    LastErrorMsg := TexteMsg[LastError];
    exit;
  end;

  if GetField('YDL_MDATATYPE') = GetField('YDL_SDATATYPE') then
  begin  //on verifie que maitre # esclave
    LastError := 5;
    LastErrorMsg := TexteMsg[LastError];
    exit;
  end;

  if Racine <>'' then
    begin //on verifie que le nom de la liaison à bien la bonne racine
    if Copy(GetField('YDL_CODEHDTLINK'),1,Strlen(Pchar(RAcine))) <> racine then
      begin
      LastError := 6;
      LastErrorMsg := TexteMsg[LastError]+ Racine;
      exit;
      end;
    end;
  if Not ExisteSql ('SELECT DO_COMBO FROM DECOMBOS where DO_COMBO="'+ GetField('YDL_MDATATYPE')+
     '" AND DO_NOMCHAMP LIKE "%HIERARCHIQUE%"')
     then begin //si saisie directe nom table, elle doit exister
     LastError := 7;
     LastErrorMsg := TexteMsg[LastError];
     exit;
     end;
  if Not ExisteSql ('SELECT DO_COMBO FROM DECOMBOS where DO_COMBO="'+ GetField('YDL_SDATATYPE')+
     '" AND DO_NOMCHAMP LIKE "%HIERARCHIQUE%"')
     then begin //si saisie directe nom table, elle doit exister
     LastError := 8;
     LastErrorMsg := TexteMsg[LastError];
     exit;
     end;
  if ExisteSql ('SELECT YDL_CODEHDTLINK FROM YDATATYPELINKS where YDL_SDATATYPE="'+ GetField('YDL_SDATATYPE')
         +'" AND YDL_CODEHDTLINK <> "'+ GetField('YDL_CODEHDTLINK') +'"')
     then begin //message si 2 fois même esclave dans la base
     PgiInfo(TexteMsg[9],'Administration hiérarchie');
     end;
    //message si "table" associée au maitre incohérent avec "table" associée esclave
  If Tobdet <> Nil then
    begin
    vTob := TobDet.FindFirst(['DO_COMBO'],[GetField('YDL_MDATATYPE')], false);
    StAbrege :='';
    if (vtob<>nil) then StAbrege:=vtob.getValue('DO_NOMCHAMP');
    BoOk :=False;
    Critere:=(Trim(ReadTokenSt(stAbrege)));
    While (Critere <>'') do
     begin
      if Copy(Critere,1,6)='@TABLE' then
        begin   //on regarde si le critère existe dans l'autre table
        vTob := TobDet.FindFirst(['DO_COMBO'],[GetField('YDL_SDATATYPE')], false);
        if (vtob<>nil) then StAb2:=vtob.getValue('DO_NOMCHAMP');
        If pos(Critere,StAb2 ) <>0 then BoOk :=true;
        end;
      Critere := (Trim(ReadTokenSt(stAbrege)));
     end;
    If Not BoOK  then PgiInfo(TexteMsg[10],'Administration hiérarchie');
    end;
end ;

procedure TOM_YDATATYPELINKS.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_YDATATYPELINKS.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_YDATATYPELINKS.OnLoadRecord ;
begin
  Inherited ;
  SetControlEnabled ('YDL_MDATATYPE',true);
  SetControlEnabled ('YDL_SDATATYPE',true);
  if not (DS.State in [dsInsert]) then
    begin //mcd 15/09/2006 AGL 12811
    If ExisteSql ('Select ydt_codehdtlink from ydataTypetrees where ydt_codehdtlink="'+ GetField('ydl_codehdtlink')+'"') then
      begin
      SetControlEnabled ('YDL_MDATATYPE',false);
      SetControlEnabled ('YDL_SDATATYPE',false);
      if ctxscot in v_pgi.pgicontexte then SetControlEnabled ('YDL_TYPELINK',false);
      end;
    end
end ;

procedure TOM_YDATATYPELINKS.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_YDATATYPELINKS.OnArgument ( StArgument: String ) ;
var critere ,champ,val, st, st1: string;
    X ,II: integer;
    CC ,CC1: THValComboBox ;
    Cegid : boolean;
//{$IFDEF PAIEGRH} //PT1
//    PaieArgument : String;
//    intLength : Integer;
//{$ENDIF}
begin
  Inherited ;
 //PT1  Gestion d'un parametre GedPaieOnly pour n'afficher que les tablettes
 // de la Ged Paie
{$IFDEF PAIEGRH}
  FiltreLinks := '';
//  boGedPaieOnly := False;
//  if pos('GedPaieOnly;',StArgument) > 0 then
//  begin
//    boGedPaieOnly := True;
//    PaieArgument := LeftStr(StArgument,pos('GedPaieOnly;',StArgument)-1);
//    intLength := length(StArgument)-pos('GedPaieOnly;',StArgument)-11;
//    if intLength > 0 then PaieArgument := PaieArgument + RightStr(StArgument,intLength);
//    StArgument := PaieArgument;
//  end;
{$ENDIF}

 Cegid := false;
 SetControlEnabled('YDL_PREDEFINI',false); //en règle générale non accessible
 SetControlVisible('YDL_LCOMMUNE',false); //en règle générale non accessible
 Critere:=(Trim(ReadTokenSt(stArgument)));
 While (Critere <>'') do
 begin
  if Critere<>'' then
    begin
    X:=pos('=',Critere);
    if x<>0 then
    begin
      Champ:=copy(Critere,1,X-1);
      Val:=Copy (Critere,X+1,length(Critere)-X);
    end
    else begin   //pour paramètre CEGID
      Champ :=Critere;
      Val:='';
      end;
    if Champ = 'PREDEFINI' then
      begin
      X := pos ('+', val);
      if X <>0 then
        begin  //on passe plusieurs valeurs,il faut donc
               //donner accès à la tablette mais restreint aux valeurs passées
        st := 'AND (';
        val := StringReplace (val, '+', ';', [rfReplaceAll]);
        Critere:=(Trim(ReadTokenSt(Val)));
        SetControlEnabled('YDL_PREDEFINI',true); 
        St1 :='';
        While (Critere <>'') do
         begin
          st := st +st1 + ' CO_CODE="'+critere+'"';
          Critere := (Trim(ReadTokenSt(Val)));
          st1 :=' OR ';
         end;
         st := st +')';
         ThDBValCOmboBox(GetCOntrol('YDL_PREDEFINI')).plus :=St;
        end;
      end
    else if Champ ='CEGID' then Cegid :=true
    else if Champ = 'RACINEAPPLI' then Racine:=val;
{$IFDEF PAIEGRH}
    if Champ = 'FILTRE' then FiltreLinks:=val;
{$ENDIF}
    end;
  Critere := (Trim(ReadTokenSt(stArgument)));
 end;
 if (V_PGI.PassWord = CryptageSt(DayPass(V_PGI.DateEntree))) or  (Cegid) then
   begin // dans ce cas , tout est permis. ne tient pas compte des paramètres passé
   SetControlEnabled('YDL_PREDEFINI',true); //on accède à la zone
   Racine :=''; // on permet tout...
   SetControlVisible('YDL_LCOMMUNE',True); //on accède à la zone
   end
  else begin  //on charge les valeurs possible dans la saisie des tablettes
    ChargeTablette;
    SetControlProperty('YDL_MDATATYPE','DataType',' ');
    SetControlProperty('YDL_SDATATYPE','DataType',' ');
    CC := THValComboBox(GetControl('YDL_MDATATYPE')) ;
    CC1 := THValComboBox(GetControl('YDL_SDATATYPE')) ;
    if (CC = nil) or (CC1 = nil) then exit ;
    For ii:=0 to TobDet.detail.count-1 do
      begin   //on alimente items et values en fct des info de la tob
      CC.Items.Add (TobDet.detail[ii].getvalue('DO_LIBELLE')) ;
      CC.Values.Add(TobDet.detail[ii].getvalue('DO_COMBO')) ;
      CC1.Items.Add (TobDet.detail[ii].getvalue('DO_LIBELLE')) ;
      CC1.Values.Add(TobDet.detail[ii].getvalue('DO_COMBO')) ;
      end;
   end;
end;

procedure TOM_YDATATYPELINKS.OnClose ;
begin
  Inherited ;
tobdet.free;
end ;

procedure TOM_YDATATYPELINKS.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_YDATATYPELINKS.ChargeTablette;
var Libelle,Nom,VstType, VstSql,Separ: string;
    II: integer;
begin
  //on va charger une tob, avec les tablettes voulues par l'applicatif
TobDet  :=Tob.create ('les tablettes',nil,-1);
Vstsql := 'SELECT DO_COMBO,DO_LIBELLE,DO_NOMCHAMP from DECOMBOS where DO_NOMCHAMP like "%HIERARCHIQUE%" AND (';
Separ :='';
  // ajouter ici , les contextes non gérer
If CtxGescom In V_PGI.PGIContexte then
  begin
  VStSql := VstSQl + separ +' DO_NOMCHAMP LIKE "%HIERAR_GC%"';
  Separ :=' OR ';
  end;
{$ifdef STK}
  VStSql := VstSQl + separ +' DO_NOMCHAMP LIKE "%HIERAR_STK%"';
  Separ :=' OR ';
{$endif}
If CtxGcAff In V_PGI.PGIContexte then
  begin
  VStSql := VstSQl + separ +' DO_NOMCHAMP LIKE "%HIERAR_GA%" OR DO_NOMCHAMP LIKE "%HIERAR_RES%"';
  Separ :=' OR ';
  end;
If CtxAffaire In V_PGI.PGIContexte then
  begin
  VStSql := VstSQl + separ +' DO_NOMCHAMP LIKE "%HIERAR_GA%" OR DO_NOMCHAMP LIKE "%HIERAR_RES%"';
  Separ :=' OR ';
  end;
If CtxGRC In V_PGI.PGIContexte then
  begin
  VStSql := VstSQl + separ +' DO_NOMCHAMP LIKE "%HIERAR_GRC%"';
  Separ :=' OR ';
  end;
If CtxPaie In V_PGI.PGIContexte then
  begin
  VStSql := VstSQl + separ +' DO_NOMCHAMP LIKE "%HIERAR_PAIE%" OR DO_NOMCHAMP LIKE "%HIERAR_RES%"';
{$IFDEF PAIEGRH} //PT1
//  if boGedPaieOnly then
//    VStSql := VStSql +' AND DO_COMBO like "PGLIBGED%" ';
  if FiltreLinks <> '' then
    VStSql := VStSql +' AND DO_COMBO like "'+FiltreLinks+'" ';
{$ENDIF}
  Separ :=' OR ';
  end;
If GetParamSocSecur ('SO_RTGESTIONGED', false) then
  begin
  VStSql := VstSQl + separ +' DO_NOMCHAMP LIKE "%HIERAR_GED%"';
  Separ :=' OR ';
  end;
VStSql := VstSql +')';
if ctxscot in V_PGI.PGIContexte then
  begin
  VStSQl := VStSql + ' and DO_COMBO <>"GCCOLLECTION"'; //mcd 17/08/2005 12307 existe en GC, mais pas en GI
  end;

TobDEt.LoadDetailFromSql (VstSql);
For ii:=(TobDEt.detail.count -1) downto 0 do
 begin
 Libelle := TobDet.detail[ii].GetValue ('DO_LIBELLE');
{$ifdef GIGI}
 Libelle := traduitGa (libelle);      //mcd 17/08/2005
{$endif}
 Nom := TobDet.detail[ii].GetValue ('DO_COMBO');
  if (pos('LIBRE',Nom) >0) then
   begin
   {$IFNDEF PAIEGRH} //PT1 GGU
   vStType:=CodeTabletteLibre (Nom);
   Libelle := RechDomZoneLibre (vstType, false);
   {$ENDIF}
  end
    // si tablette sur famille article on change le libellé
  else if (pos('FAMILLENIV',Nom) >0)  then
   begin
   vStType:='LF' + Copy (Nom,13,1);
   Libelle := RechDom ('GCLIBFAMILLE', vstType,false)
   end
  else if (pos('RTLIBGED',Nom) >0)  then
   begin
   vStType:='RD' + Copy (Nom,9,1);
   Libelle := RechDom ('RTLIBGED', vstType,false)
   end
  else if nom = 'AFFAIREPART1' then
     begin //cas partie du code à affaire, à traduire
     if GetParamSocSecur ('SO_AFFCO1TYPE','')='LIS' then Libelle :=GetParamSocSecur ('SO_AFFCO1LIB','')
        else LIbelle:= '.-';
     end
  else if nom = 'AFFAIREPART2' then
     begin
     if (GetParamSocSecur ('SO_AFFCO2TYPE','')='LIS') and (GetParamSocSecur ('SO_AFFCO2VISIBLE',false)) then Libelle :=GetParamSocSecur ('SO_AFFCO2LIB','')
        else LIbelle:= '.-';
     end
  else if nom = 'AFFAIREPART3' then
     begin
     if (GetParamSocSecur ('SO_AFFCO3TYPE','')='LIS') and (GetParamSocSecur ('SO_AFFCO3VISIBLE',false)) then Libelle :=GetParamSocSecur ('SO_AFFCO3LIB','')
        else LIbelle:= '.-';
     end ;
  if (copy(Libelle,1,2) <> '.-') then
    begin  //on peut garder l'info
    TobDet.detail[ii].PutValue ('DO_LIBELLE', Libelle );
    end
    else begin // il faut détruire enrgt
    TobDet.detail[ii].free;
    end
 end;
end;


Initialization
  registerclasses ( [ TOM_YDATATYPELINKS ] ) ;
end.
