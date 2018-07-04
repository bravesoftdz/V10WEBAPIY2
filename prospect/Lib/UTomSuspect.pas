{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 19/10/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : SUSPECTS (SUSPECTS)
Mots clefs ... : TOM;SUSPECTS
*****************************************************************}
Unit UTomSuspect ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, 
     HCtrls, HEnt1, HMsgBox, UTOM, UTob ,M3FP,Graphics,extctrls,utilGC,lookup,
{$IFDEF EAGLCLIENT}
      eFiche,spin,
{$ELSE}
      Fiche,db,dbctrls,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}HDB,
{$ENDIF}
      ParamSoc,UtilRT, ent1,UtilPGI,EntRT,Web
 ;
Type
  TOM_SUSPECTS = Class (TOM)
  private
    TobZonelibre : TOB;
    ModifLot : boolean;
    StSQL : string;
    TParticulier :TRadioGroup;
    Procedure ZLSuspectsIsModified ;
    procedure SetArguments(StSQL : string);
    procedure SetLastError (Num : integer; ou : string );
    procedure ModifParticulier(Sender: TObject);
    procedure RSU_RVA_OnDblclick(Sender : tObject);        
    procedure RSU_CONTACTRVA_OnDblclick(Sender : tObject);
    function ValidDateNaissance : Boolean;
    procedure ThEnseigneOnElipsisClick (Sender : Tobject);
  public
{$IFDEF EAGLCLIENT}
    THEnseigne : THEdit;
{$ELSE}
    THEnseigne : THDBEdit;
{$ENDIF}
    procedure OnArgument ( S: String )   ; override ;
    procedure OnNewRecord                ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnCancelRecord             ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnClose                    ; override ;
    procedure OnChangeField (F : TField) ; override ;
    procedure FermeSuspectOnClick(Sender: TObject);
    end ;

Function GetLibelleSuspect (champ:string;Particulier:boolean):string;

const
// libellés des messages
    TexteMessage: array[1..8] of string 	= (
      {1}        'La raison sociale doit être renseignée'
      {2}       ,'Le numéro de Siret ou Siren n''est pas correct '
      {3}       ,'Ce numéro de Siret existe déjà sur la fiche suspect : '// GRC
      {4}       ,'Cette région n''appartient pas à ce pays'
      {5}       ,'Double-cliquez pour envoyer un message'
      {6}       ,'Double-cliquez pour accéder au site Web'
      {7}       ,'Le jour de naissance n''est pas correct'
      {8}       ,'L''année de naissance doit être saisie sur 4 caractères'
      );

LibSuspectParticulier : array[1..14,1..3] of String 	= (
   ('RSU_LIBELLE'                   , 'Raison Sociale'      , 'Nom'           ),
   ('RSU_TELEPHONE'                 , 'Téléphone'           , 'Tél domicile'  ),
   ('RSU_FAX'                       , 'Fax'                 , 'Tél bureau'    ),
   ('RSU_TELEX'                     , 'Tél portable'        , 'Tél portable'  ),
   ('RSU_RVA'                       , 'Site Web'            , 'E-mail'        ),
   ('RSU_JURIDIQUE'                 , 'Abréviat. postale'   , 'Civilité'      ),
   ('RSU_SUSPECT'                   , 'Code suspect'        ,'Code suspect'   ),
   ('RSU_EAN'                       , 'Code Client EDI'     ,''               ),
   ('RSU_NIF'                       , 'Code client UE'      ,''               ),
   ('RSU_APE'                       , 'Code NAF'            ,''               ),
   ('RSU_CONTACTCIVILITE'           , 'Civilité Contact'    ,'Civilité Contact'),
   ('RSU_CONTACTPRENOM'             , 'Prénom Contact'      ,'Prénom Contact'  ),
   ('RSU_CONTACTNOM'                , 'Nom Contact'         ,'Nom Contact'     ),
   ('RSU_CONTACTRVA'                , 'E-mail Contact'      ,'E-mail Contact'  )
   );

Implementation

uses
  TiersUtil,BTPUtil,HrichOle;

procedure TOM_SUSPECTS.OnUpdateRecord ;
var CodeSuspect : string;
    Qr : TQuery;
begin
  Inherited ;
  if (ds<>nil) then
    begin
    if (GetField('RSU_LIBELLE')='') then
      begin
      SetLastError(1,'RSU_LIBELLE');  exit;
      end;
    if (ds.state = dsinsert) then
      begin
      CodeSuspect:=RTAttribNewCodeSuspect(GetParamsocSecur('SO_RTCOMPTEURSUSP','0'));
      SetField ('RSU_SUSPECT', CodeSuspect);
      {$IFDEF EAGLCLIENT}
      SetControlText('RSU_SUSPECT',CodeSuspect);
      {$ENDIF}
      end;

    if (GetField('RSU_PARTICULIER')<>'X') and (GetField('RSU_SIRET') <> '') then
      begin
      if not VerifSiret( GetField('RSU_SIRET')) then
        begin
        SetLastError(2, 'RSU_SIRET'); exit ;
        end;

      Qr := OpenSql('SELECT RSU_SUSPECT FROM SUSPECTS WHERE RSU_PARTICULIER<>"X" AND RSU_FERME<>"X" AND '+
      'RSU_SIRET LIKE "'+GetField('RSU_SIRET')+'%" AND RSU_SUSPECT<>"'+GetField('RSU_SUSPECT')+'"', False);
      if not Qr.Eof then
        begin
        LastError:=3;
        LastErrorMsg:=TexteMessage[LastError]+Qr.Findfield('RSU_SUSPECT').AsString;
        Ferme(Qr);
        exit;
        end
      else Ferme(Qr);
      end ;
    If (GetField ('RSU_PAYS') <> '') and (GetField ('RSU_REGION') <> '') and (not existesql('SELECT RG_PAYS FROM REGION WHERE RG_PAYS = "'+string(GetField ('RSU_PAYS'))+'" AND RG_REGION = "'+string(GetField ('RSU_REGION'))+'"')) then
    Begin
      SetLastError(4, 'RSU_REGION');
      exit ;
    end ;
    if (GetField('RSU_PARTICULIER') = 'X') and (ValidDateNaissance = False) then exit;
    if (ds.state = dsinsert) then SetParamSoc('SO_RTCOMPTEURSUSP',CodeSuspect) ;

    end;
    TFFiche(Ecran).Retour :=  GetField('RSU_SUSPECT');

end;

procedure TOM_SUSPECTS.OnAfterUpdateRecord;
begin
  inherited;
  TobZoneLibre.GetEcran (TFfiche(Ecran),Nil);
  TobZoneLibre.PutValue ('RSC_SUSPECT', GetField('RSU_SUSPECT')) ;
  TobZoneLibre.InsertOrUpdateDB (FALSE);
  TobZoneLibre.SetAllModifie(False);
  SetControlVisible('BPROSPECT', (ds.state <> dsinsert));
  if ModifLot then TFFiche(ecran).BFermeClick(nil);
end;

procedure TOM_SUSPECTS.OnLoadRecord ;
begin
  Inherited ;
  if (ds.state = dsinsert ) or (GetField('RSU_FERME')='X') then
    SetControlEnabled('BPROSPECT', False);
  if (GetField('RSU_DATESUSPRO') = iDate1900) then
     begin
     SetControlVisible('TRSU_DATESUSPRO',False);
     SetControlVisible('RSU_DATESUSPRO',False);
     end else
     begin
     SetControlEnabled ('RSU_FERME', False);
     SetControlVisible('BPROSPECT',False);
     end;
  if not TobZonelibre.SelectDB('"'+GetField('RSU_SUSPECT')+'"',Nil ) then TobZonelibre.InitValeurs;
  TobZonelibre.PutEcran(TFfiche(Ecran));
  SetControlEnabled ('RSU_SUSPECT', False);
  SetFocusControl ('RSU_JURIDIQUE');
  TCheckBox(GetControl('RSU_FERME')).OnClick := FermeSuspectOnClick;
  if ModifLot then SetArguments(StSQL);
  if (GetField('RSU_PARTICULIER')='X') then
  begin
    TParticulier.ITemIndex := 0;
    ModifParticulier(Nil);
  end else TParticulier.ITemIndex := 1;
  SetControlProperty('RSU_REGION','PLUS','RG_PAYS="'+GetControlText('RSU_PAYS')+'"');
end ;

procedure TOM_SUSPECTS.SetLastError (Num : integer; ou : string );
begin
  if ou<>'' then SetFocusControl(ou);
  LastError:=Num;
  LastErrorMsg:=TexteMessage[LastError];
end ;

procedure TOM_SUSPECTS.SetArguments(StSQL : string);
var Critere,ChampMul,ValMul : string ;
    x,y : integer ;
    Ctrl : TControl;
    Fiche : TFFiche;
begin
SetControlVisible('BSTOP',TRUE);
DS.Edit;
Fiche := TFFiche(ecran);
Repeat
    Critere:=Trim(ReadTokenPipe(StSQL,'|')) ;
    if Critere<>'' then
        begin
        x:=pos('=',Critere);
        if x<>0 then
           begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           y := pos(',',ValMul);
           if y<>0 then ValMul:=copy(ValMul,1,length(ValMul)-1);
           if copy(ValMul,1,1)='"' then ValMul:=copy(ValMul,2,length(ValMul));
           if copy(ValMul,length(ValMul),1)='"' then ValMul:=copy(ValMul,1,length(ValMul)-1);
           SetField(ChampMul,ValMul);
           Ctrl:=TControl(Fiche.FindComponent(ChampMul));
           if Ctrl=nil then exit;
{$IFDEF EAGLCLIENT}
           if (Ctrl is TCustomCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is TCustomEdit) then
           begin
            TEdit(Ctrl).Font.Color:=clRed;
            SetControlText(ChampMul,ValMul);
           end
           else if Ctrl is TSpinEdit then TSpinEdit(Ctrl).Font.Color:=clRed
              //mcd 25/01/01 ajout test TH pour zones YTC de tierscompl
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              TSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ELSE}
           if (Ctrl is TDBCheckBox) or (Ctrl is THDBValComboBox) Or (Ctrl is THDBEdit) then TEdit(Ctrl).Font.Color:=clRed
           else if Ctrl is THDBSpinEdit then THDBSpinEdit(Ctrl).Font.Color:=clRed
              //mcd 25/01/01 ajout test TH pour zones YTC de tierscompl
           else if (Ctrl is TCheckBox) or (Ctrl is THValComboBox) Or (Ctrl is THEdit)Or (Ctrl is THNumEdit)then
              begin
              THDBSpinEdit(Ctrl).Font.Color:=clRed;
              SetControlText(ChampMul,ValMul);
              end;
{$ENDIF}
           end;
        end;
until  Critere='';
end;

procedure TOM_SUSPECTS.OnArgument ( S: String ) ;
var x : integer ;
    Ctrl: TControl;
begin
  Inherited ;
  AppliqueFontDefaut (THRichEditOle(GetControl('RSU_BLOCNOTE')));
  TobZonelibre:=TOB.Create ('SUSPECTSCOMPL', Nil, -1);
  x := pos('MODIFLOT',S);
  ModifLot := x<>0;
  if ModifLot then
    begin
    TFfiche(Ecran).MonoFiche:=true;
    StSQL := copy(S,x+9,length(S));
    end
    else if ( Not AutoriseCreationTiers ('PRO') ) or
            ( (VH_RT.RTExisteConfident=false) and (GetParamsocSecur('SO_RTCONFIDENTIALITE',False) = true) ) then
              SetControlVisible('BPROSPECT',False);

  TParticulier := TRadioGroup(GetControl('RSU_PARTICULIER'));
  TParticulier.ItemIndex := 1;
  TParticulier.OnClick := Modifparticulier;
//{$ifdef GIGI}
// SetControlVisible('RSU_ZONECOM',false);
// SetControlVisible('TRSU_ZONECOM',false);
// SetControlVisible('RSU_REPRESENTANT',false);
// SetControlVisible('TRSU_REPRESENTANT',false);
//{$endif}
{$IFDEF EAGLCLIENT}
  Ctrl := GetControl('RSU_CONTACTRVA');
  if Assigned(Ctrl) and (Ctrl is thEdit) then ThEdit(Ctrl).OnDblclick := RSU_CONTACTRVA_OnDblclick;
  Ctrl := GetControl('RSU_RVA');
  if assigned(Ctrl) and (Ctrl is ThEdit) then ThEdit(Ctrl).OnDblclick := RSU_RVA_OnDblclick;
{$ELSE}
  Ctrl := GetControl('RSU_CONTACTRVA');
  if Assigned(Ctrl) and (Ctrl is thDBEdit) then ThDBEdit(Ctrl).OnDblclick := RSU_CONTACTRVA_OnDblclick;
  Ctrl := GetControl('RSU_RVA');
  if Assigned(Ctrl) and (Ctrl is ThDBEdit) then ThDBEdit(Ctrl).OnDblclick := RSU_RVA_OnDblclick;
{$ENDIF}

  //FQ10467 gestion des commerciaux
  if GereCommercial then
  begin
    SetControlVisible ('GB_COMMERCIAUX', True);
    SetControlVisible('RSU_ZONECOM', True);
    SetControlVisible('TRSU_ZONECOM', True);
  end
  else
  begin
    SetControlVisible ('GB_COMMERCIAUX', False);
    SetControlVisible('RSU_ZONECOM', False);
    SetControlVisible('TRSU_ZONECOM', False);
  end;

  
{$IFDEF EAGLCLIENT}
  THEnseigne := THEdit(GetControl('RSU_ENSEIGNE'));
{$ELSE}
  THEnseigne := THDBEdit(GetControl('RSU_ENSEIGNE'));
{$ENDIF}
  if THEnseigne <> nil then
  begin
    if GetParamSocSecur('SO_GCENSEIGNETAB', False, True) then
    begin
      THEnseigne.DataType := 'RTENSEIGNE';
      THEnseigne.ElipsisButton := True;
      THEnseigne.OnElipsisClick := ThEnseigneOnElipsisClick;
    end
    else
    begin
      THEnseigne.DataType := '';
      THEnseigne.ElipsisButton := False;
    end;
  end;
end ;

procedure TOM_SUSPECTS.OnClose ;
begin
  Inherited ;
TobZonelibre.free;
end ;

procedure TOM_SUSPECTS.OnCancelRecord ;
begin
  Inherited ;
if TFfiche(Ecran)<>nil then TobZonelibre.PutEcran(TFfiche(Ecran));
end ;

Procedure TOM_SUSPECTS.ZLSuspectsIsModified ;
BEGIN
TobZoneLibre.GetEcran (TFfiche(Ecran),Nil);
if TobZoneLibre.IsOneModifie and  not(DS.State in [dsInsert,dsEdit])then
    begin
    DS.edit; // pour passer DS.state en mode dsEdit
{$IFDEF EAGLCLIENT}
    TFFiche(Ecran).QFiche.CurrentFille.Modifie:=true;
{$ELSE}
    SetField ('RSU_SUSPECT', GetControlText ('RSU_SUSPECT'));
{$ENDIF}
    end;
END;

procedure TOM_SUSPECTS.FermeSuspectOnClick(Sender: TObject);
begin
  if TCheckBox(GetControl('RSU_FERME')).Checked then
  begin
    SetControlVisible ('RSU_DATEFERMETURE', TRUE);
    SetControlVisible ('TRSU_DATEFERMETURE', TRUE);
    SetControlVisible ('RSU_MOTIFFERME', TRUE);
    SetControlVisible ('TRSU_MOTIFFERME', TRUE);
    SetControlEnabled ('BPROSPECT',FALSE);
    SetField('RSU_FERME','X') ;
    SetField('RSU_DATEFERMETURE', V_PGI.DateEntree);
  end else
  begin
    SetControlVisible ('RSU_DATEFERMETURE', FALSE);
    SetControlVisible ('TRSU_DATEFERMETURE', FALSE);
    SetControlVisible ('RSU_MOTIFFERME', FALSE);
    SetControlVisible ('TRSU_MOTIFFERME', FALSE);
    SetControlEnabled ('BPROSPECT',TRUE);
    SetField('RSU_FERME','-') ;
    Setfield('RSU_DATEFERMETURE',iDate1900) ;
    SetField('RSU_MOTIFFERME','') ;
  end;
end;

procedure TOM_SUSPECTS.OnDeleteRecord ;
begin
inherited ;
ExecuteSQL('DELETE FROM CIBLAGEELEMENT WHERE RVB_SUSPECT="'+string(GetField('RSU_SUSPECT'))+'"') ;
ExecuteSQL('DELETE FROM SUSPECTSCOMPL WHERE RSC_SUSPECT="'+string(GetField('RSU_SUSPECT'))+'"') ;
end;

procedure TOM_SUSPECTS.RSU_CONTACTRVA_OnDblclick(Sender : tObject);
var stChamp : string;
begin
  if TControl(Sender).name = 'RSU_RVA' then stChamp := GetControlText('RSU_RVA')
  else stChamp := GetControlText('RSU_CONTACTRVA');
  if stChamp <> '' then
     PGIEnvoiMail('',stChamp,'',nil,'',False);
end;

procedure TOM_SUSPECTS.RSU_RVA_OnDblclick(Sender : tObject);
var
  sHttp: String;
begin
  sHttp := GetControlText('RSU_RVA');
  if Pos('HTTP://', UpperCase(sHttp)) = 0 then
    sHttp := 'http://'+sHttp;
  LanceWeb(sHttp,True);
end;

function TOM_SUSPECTS.ValidDateNaissance : Boolean;
var mois : integer;
begin
Result := False;
mois := GetField('RSU_MOISNAISSANCE');
if (GetField('RSU_JOURNAISSANCE') > 30) and ((mois = 2) or (mois = 4) or (mois = 6) or (mois = 9) or (mois = 11)) then begin SetLastError(7, 'RSU_JOURNAISSANCE'); exit; end;
if (GetField('RSU_ANNEENAISSANCE') <> 0) and (GetField('RSU_ANNEENAISSANCE') < 1900) then begin SetLastError(8, 'RSU_ANNEENAISSANCE'); exit; end;
Result := True;
end;

// <<<<<<<<<<<<<< Fonctions pour script AGL >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

procedure AGLZLSuspectsIsModified( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else exit;
if (OM is TOM_SUSPECTS)
    then TOM_SUSPECTS(OM).ZLSuspectsIsModified
    else exit;
end;


procedure TOM_SUSPECTS.ModifParticulier(Sender: TObject);
var Particulier :Boolean;
    Ctrl : TControl;
begin
  Particulier := (TParticulier.ItemIndex = 0);
  SetControlVisible ('RSU_JOURNAISSANCE', Particulier);
  SetControlVisible ('RSU_MOISNAISSANCE', Particulier);
  SetControlVisible ('RSU_ANNEENAISSANCE', Particulier);
  SetControlVisible ('RSU_DATNAISSANCE', Particulier);
  SetControlVisible ('SEP_MOISNAISSANCE', Particulier);
  SetControlVisible ('SEP_ANNEENAISSANCE', Particulier);
  SetControlVisible ('TRSU_PRENOM',Particulier);
  SetControlVisible ('RSU_SEXE', Particulier);
  SetControlVisible ('TRSU_SEXE', Particulier);

  SetControlVisible ('RSU_SIRET', not Particulier);
  SetControlVisible ('TRSU_SIRET', not Particulier);
  SetControlVisible ('RSU_NIF', not Particulier);
  SetControlVisible ('TRSU_NIF', not Particulier);
  SetControlVisible ('RSU_EAN', not Particulier);
  SetControlVisible ('TRSU_EAN', not Particulier);
  SetControlVisible ('RSU_APE', not Particulier);
  SetControlVisible ('TRSU_APE', not Particulier);
  SetControlVisible ('RSU_FORMEJURIDIQUE', not Particulier);
  SetControlVisible ('TRSU_FORMEJURIDIQUE', not Particulier);
  SetControlVisible ('RSU_ENSEIGNE', not Particulier);
  SetControlVisible ('TRSU_ENSEIGNE', not Particulier);

  if (Particulier) then
    begin
    SetField('RSU_PARTICULIER','X');
    SetControlProperty ('RSU_JURIDIQUE', 'DataType', 'TTCIVILITE');
    Ctrl := GetControl('RSU_RVA');
{$IFDEF EAGLCLIENT}
    if assigned(Ctrl) and (Ctrl is ThEdit) then ThEdit(Ctrl).OnDblclick := RSU_CONTACTRVA_OnDblclick;
{$ELSE}
    if Assigned(Ctrl) and (Ctrl is ThDBEdit) then ThDBEdit(Ctrl).OnDblclick := RSU_CONTACTRVA_OnDblclick;
{$ENDIF}
    Ctrl.Hint := TexteMessage[5];
    end else
    begin
    SetField('RSU_PARTICULIER','-');
    SetControlProperty ('RSU_JURIDIQUE', 'DataType', 'TTFORMEJURIDIQUE');
    Ctrl := GetControl('RSU_RVA');
{$IFDEF EAGLCLIENT}
    if assigned(Ctrl) and (Ctrl is ThEdit) then ThEdit(Ctrl).OnDblclick := RSU_RVA_OnDblclick;
{$ELSE}
    if Assigned(Ctrl) and (Ctrl is ThDBEdit) then ThDBEdit(Ctrl).OnDblclick := RSU_RVA_OnDblclick;
{$ENDIF}
    Ctrl.Hint := TexteMessage[6];
    end;

  SetControlText ('TRSU_LIBELLE', '&'+GetLibelleSuspect('RSU_LIBELLE',Particulier));
  SetControlText ('TRSU_TELEPHONE', '&'+GetLibelleSuspect('RSU_TELEPHONE',Particulier));
  SetControlText ('TRSU_FAX', '&'+GetLibelleSuspect('RSU_FAX',Particulier));
  SetControlText ('TRSU_TELEX', '&'+GetLibelleSuspect('RSU_TELEX',Particulier));
  SetControlText ('TRSU_RVA', '&'+GetLibelleSuspect('RSU_RVA',Particulier));
  SetControlText ('TRSU_JURIDIQUE', '&'+GetLibelleSuspect('RSU_JURIDIQUE',Particulier));

  TFFiche(Ecran).Refresh;
end;

procedure TOM_SUSPECTS.OnNewRecord;
begin
  inherited;
  TParticulier.ItemIndex := 1;
  SetField('RSU_PARTICULIER','-');
  SetFocusControl ('RSU_JURIDIQUE');
  SetControlChecked('RSU_PUBLIPOSTAGE',True);
  SetField('RSU_PUBLIPOSTAGE','X');
end;

procedure TOM_SUSPECTS.OnChangeField(F: TField);
begin
Inherited;
  if (F.FieldName = 'RSU_PAYS') and (GetField('RSU_PAYS')<>'') then
    SetControlProperty('RSU_REGION','PLUS','RG_PAYS="'+GetControlText('RSU_PAYS')+'"');
end;

Function GetLibelleSuspect (champ:string;Particulier:boolean):string;
var Icpt,Iparticulier :integer;
begin
  if Particulier then Iparticulier:=3 else Iparticulier:=2;
  for Icpt:=Low(LibSuspectParticulier) to High(LibSuspectParticulier) do
    if (LibSuspectParticulier[Icpt, 1]=champ) then result := TraduireMemoire(LibSuspectParticulier [Icpt,Iparticulier]);
end;


procedure TOM_SUSPECTS.ThEnseigneOnElipsisClick(Sender: Tobject);
var
  Numdisp : integer;

begin
  Numdisp := 0;
  if GetParamSocSecur('SO_GCENSEIGNECREAT', False, True) then
    Numdisp := 43;

  // on retourne le libelle et non pas le code
  LookupList(THEnseigne, 'Enseignes', 'CHOIXCOD', 'CC_LIBELLE', 'CC_LIBELLE', 'CC_TYPE="REN"', 'CC_LIBELLE',true, NumDisp,'',tlDefault ) ;

end;

Initialization
  registerclasses ( [ TOM_SUSPECTS ] ) ;
  RegisterAglProc( 'ZLSuspectsIsModified', TRUE , 0, AGLZLSuspectsIsModified);
end.

