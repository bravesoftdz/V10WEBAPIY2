{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 31/10/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PARACTIONS (PARACTIONS)
Mots clefs ... : TOM;PARACTIONS
*****************************************************************}
Unit UtomParActions ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db, FichList,Fiche,
     EdtREtat,
{$ELSE}
     eFichList, UtileAGL,eFiche,
{$ENDIF}
     forms,
     sysutils,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     UTob,
     uTableFiltre,
     SaisieList,
     EntRT, M3FP,Paramsoc,Yplanning,variants ;

const
//GP_20071123_TP_GP14529
  TexteMessage: array[1..11] of string 	= (
    {1}  'Le type d''action doit être renseigné'
    {2} ,'Vous ne pouvez pas effacer ce type d''action, il est actuellement utilisé dans les actions'
    {3} ,'Vous ne pouvez pas effacer ou déplacer cette action, elle est actuellement utilisée dans les chaînages'
    {4} ,'Le nombre de jours ne peut pas être négatif'
    {5} ,'Vous ne pouvez pas effacer ce type d''action, il est actuellement utilisé dans les types de chaînage'
    {6} ,'L''heure de l''action doit être renseignée dans ce cas'
    {7} ,'Le responsable n''existe pas'
    {8} ,'Le libellé de l''action doit être renseigné'
    {9} ,'Ce type d''action existe déjà'
    {10} ,'Vous devez renseigner la nature de travail'
//GP_20071123_TP_GP14529
    {11} ,'Vous devez renseigner la famille de l''action'    {Qualité}
    );

  nonRealisable : string = 'NRE';
  prevue : string = 'PRE';

Type
  TOM_PARACTIONS = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnCancelRecord             ; override ;
  Private
       TF : TTableFiltre;
       Old_Typeaction,stProduitpgi      : string;
       soRtactrappel: boolean;
       soRtactchoixrappel:string;
       soRtchsuppact:boolean;
       soRtactgestech:boolean;
       soRtactgestcout:boolean;
       soRtactgestchrono:boolean;
    procedure LookEcran(Valeur:integer);
    procedure RTChChgtTypeAction;
    procedure RTChaineFlecheHaut;
    procedure RTChaineFlecheBas;
    procedure RappelClick(Sender: TObject);
    procedure EnvoiMailClick(Sender: TObject);
    procedure MailEtatValClick(Sender: TObject);
    procedure MailEtatReaClick(Sender: TObject);
    procedure MailEtatPreClick(Sender: TObject);
    procedure MailEtatAnuClick(Sender: TObject);
    procedure GestEchClick(Sender: TObject);
{$IFDEF QUALITE}
    procedure QNatureAction_OnChange(Sender: TObject);
{$ENDIF QUALITE}
    procedure DUREEACT_OnClick(Sender: TObject);
    end ;

//procedure AGLRTChChgtTypeAction( parms: array of variant; nb: integer ) ;
procedure AGLRTValeursDef1( parms: array of variant; nb: integer ) ;
procedure AGLRTValeursDef2( parms: array of variant; nb: integer ) ;
procedure AGLRTChaineFlecheBas( parms: array of variant; nb: integer ) ;
procedure AGLRTChaineFlecheHaut( parms: array of variant; nb: integer ) ;
Implementation
uses UtilAction;

procedure TOM_PARACTIONS.OnNewRecord ;
var nb : integer;
begin
  Inherited ;
  SetField( 'RPA_TYPEACTION', '' );
  if TF<>nil then
    begin
    nb:= TF.Reccount;
    SetControlEnabled('RPA_MAILRESPVAL',False);
    if (nb=1) then { =1 sur la première ligne que l'on crée }
      begin
      SetField( 'RPA_NUMLIGNE', 1 );
      SetControlEnabled('RPA_MAILETATVAL',True);
      SetControlEnabled('RPA_MAILVALAUTO',True);
      SetField( 'RPA_ETATACTION', prevue );
      SetField('RPA_MAILRESPVAL','RSP') ;
      end
    else
      begin
      SetField( 'RPA_NUMLIGNE', TF.MaxValue('RPA_NUMLIGNE')+1 );
      SetField( 'RPA_ETATACTION', nonRealisable );
      SetControlEnabled('RPA_MAILETATVAL',False);
      SetControlEnabled('RPA_MAILVALAUTO',False);
      end;
    SetField( 'RPA_CHAINAGE',TF.TOBFiltre.GetValue('RPG_CHAINAGE'));
    SetField('RPA_INTERVENANT',VH_RT.RTResponsable);
    Old_Typeaction := '';
    end
  else
    begin
      SetField( 'RPA_CHAINAGE','---');
      if soRtactrappel = true then
      begin
        SetControlChecked('RPA_GESTRAPPEL',true) ;
        SetField('RPA_GESTRAPPEL','X');
        SetField('RPA_DELAIRAPPEL',soRtactchoixrappel);
      end;
      SetField( 'RPA_ETATACTION', prevue );
    end ;
  SetField( 'RPA_PRODUITPGI', stProduitpgi );
  if soRtactgestech= true then SetField('RPA_GESTDATECH','X');
  if soRtactgestcout= true then SetField('RPA_GESTCOUT','X');
  if soRtactgestchrono= true then SetField('RPA_GESTCHRONO','X');
  SetField ('RPA_DUREEACTION',30);
  if Assigned(GetControl('DUREEACT')) then
    SetControlText('DUREEACT','030');
end ;

procedure TOM_PARACTIONS.OnDeleteRecord ;
var TobTypeEncours : tob;
begin
  Inherited ;

  VH_RT.TobTypesAction.Load;

  if TF=nil then
  begin
    { Param type action : on vérifie que le type n'est pas utilisé dans les actions }
    if ExisteSQL( 'SELECT RAC_NUMACTION FROM ACTIONS WHERE RAC_TYPEACTION="'+GetField('RPA_TYPEACTION')+'"') then
    begin
      LastError:=2;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit ;
    end;
    { Param type action : on vérifie que le type n'est pas utilisé dans les types de chainage }

    TobTypeEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION'],[GetField('RPA_TYPEACTION')],TRUE) ;
    while (TobTypeEncours <> Nil) do
    begin
      if (TobTypeEncours.GetValue('RPA_CHAINAGE')<>'---') and
         (TobTypeEncours.GetValue('RPA_NUMLIGNE')<>0) then
      begin
          LastError:=5;
          LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
          exit ;
      end;
      TobTypeEncours:=VH_RT.TobTypesAction.FindNext(['RPA_TYPEACTION'],[GetField('RPA_TYPEACTION')],TRUE) ;
    end;
  end
  else
  begin
    if ExisteSQL( 'SELECT RAC_NUMACTION FROM ACTIONS WHERE RAC_TYPEACTION="'+GetField('RPA_TYPEACTION')+
       '" AND RAC_CHAINAGE="'+GetField('RPA_CHAINAGE')+
       '" AND RAC_NUMLIGNE='+IntToStr(GetField('RPA_NUMLIGNE'))) then
    begin
      if soRtchsuppact = false then
      begin
        LastError:=3;
        LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
        exit ;
      end ;
    end;
  end;
  TobTypeEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField ('RPA_TYPEACTION'),GetField ('RPA_CHAINAGE'),GetField ('RPA_NUMLIGNE')],TRUE) ;
  if (TobTypeEncours <> Nil) then
     TobTypeEncours.free;
end ;

procedure TOM_PARACTIONS.OnUpdateRecord ;
var TobTypeEncours : tob;
//GP_20071123_TP_GP14529 >>>

  function VerifTypeAction: Boolean;
  begin
    if GetField ('RPA_TYPEACTION') = '' then
    begin
      Lasterror := 1;
      LastErrorMsg := TraduireMemoire(TexteMessage[LastError]);
      SetFocusControl('RPA_TYPEACTION') ;
      Result := False;
    end
    else
      Result := True;
  end;

//GP_20071123_TP_GP14529 <<<
begin
  Inherited ;

  VH_RT.TobTypesAction.Load;

  if TF<>nil then
  begin
    if GetField ('RPA_TYPEACTION') = '' then
    begin
      SetFocusControl('RPA_TYPEACTION') ;
      Lasterror:=1;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit;
    end;
//GP_20071123_TP_GP14529 >>>
    if not VerifTypeAction then
      Exit;
//GP_20071123_TP_GP14529 <<<
    if GetField ('RPA_LIBELLE') = '' then
    begin
      SetFocusControl('RPA_LIBELLE') ;
      Lasterror:=8;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit;
    end;
	  if ((GetField ('RPA_INTERVENANT') <> '') and  (Not ExisteSQL ('SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE="'+GetField('RPA_INTERVENANT')+'"'))) then
  	begin
    	SetFocusControl('RPA_INTERVENANT');
      LastError := 7;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit;
    end;
    if GetField ('RPA_NBJOURS') < 0 then
    begin
      SetFocusControl('RPA_NBJOURS') ;
      Lasterror:=4;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit;
    end;
    if  (GetField ('RPA_GESTRAPPEL') = 'X') and (GetField ('RPA_HEUREACTION') = iDate1900) then
    begin
      SetFocusControl('RPA_HEUREACTION') ;
      Lasterror:=6;
      LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
      exit;
    end;
  end ;

  if GetField ('RPA_DELAIDATECH') < 0 then
  begin
    SetFocusControl('RPA_DELAIDATECH') ;
    Lasterror:=4;
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
    exit;
  end;
//GP_20071123_TP_GP14529 >>>

  {$IFDEF QUALITE}
    if stProduitPGI='QNC' then
    begin
      if (GetControlText('RPA_QNATUREACTION')='WPR') and (GetControlText('RPA_NATURETRAVAIL')='') then
      begin
        LastError := 10;
        SetFocusControl('RPA_NATURETRAVAIL');
        LastErrorMsg := TraduireMemoire(TexteMessage[LastError]);
        Exit;
      end;
      if not VerifTypeAction then
        Exit;
      if GetField ('RPA_QNATUREACTION') = '' then
      begin
        LastError := 11;
        LastErrorMsg := TraduireMemoire(TexteMessage[LastError]);
        SetFocusControl('RPA_QNATUREACTION');
        Exit;
      end;
    end;
  {$ENDIF QUALITE}
//GP_20071123_TP_GP14529 <<<

  { mise à jour de la Tob Mémoire des types d'actions }
  if ( ds.State = dsEdit ) then
  begin
    TobTypeEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField ('RPA_TYPEACTION'),GetField ('RPA_CHAINAGE'),GetField ('RPA_NUMLIGNE')],TRUE) ;
    if (TobTypeEncours <> Nil) then
       TobTypeEncours.GetEcran(TForm(Ecran),nil);
  end
  else


  TobTypeEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField ('RPA_TYPEACTION'),GetField ('RPA_CHAINAGE'),GetField ('RPA_NUMLIGNE')],TRUE) ;
  if (TobTypeEncours <> Nil) then
     if ( ds.State = dsEdit ) then
       TobTypeEncours.GetEcran(TForm(Ecran),nil)
     else
       begin
       SetFocusControl('RPA_TYPEACTION') ;
       Lasterror:=9;
       LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
       exit;
       end;

    if ( ds.State = dsInsert ) then
    begin
      TobTypeEncours:=TOB.create ('PARACTIONS',VH_RT.TobTypesAction,-1);
      if (TobTypeEncours <> Nil) then
      begin
       TobTypeEncours.GetEcran(TForm(Ecran),nil);
       TobTypeEncours.PutValue('RPA_CHAINAGE',GetField ('RPA_CHAINAGE'));
       TobTypeEncours.PutValue('RPA_NUMLIGNE',GetField ('RPA_NUMLIGNE'));
       TobTypeEncours.PutValue('RPA_PRODUITPGI',stProduitpgi);
      end;
    end;
end ;

procedure TOM_PARACTIONS.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARACTIONS.OnLoadRecord ;
begin
  Inherited ;
  if TF<>nil then
    begin
    SetControlEnabled('RPA_MAILRESPVAL',False);
    if GetField( 'RPA_NUMLIGNE') = 1 then
      begin
      SetControlEnabled('RPA_MAILETATVAL',True);
      SetControlEnabled('RPA_MAILVALAUTO',True);
      SetControlEnabled('BFLECHEHAUT',False);
      end
    else
      begin
      SetControlEnabled('RPA_MAILETATVAL',False);
      SetControlEnabled('RPA_MAILVALAUTO',False);
      SetControlEnabled('BFLECHEHAUT',True);
      end;
    if Getfield ('RPA_TYPEACTION') <> '' then
       begin
       Old_Typeaction := Getfield ('RPA_TYPEACTION');
       //LookEcran();
       end;
    if Getfield('RPA_NUMLIGNE')=TF.MaxValue('RPA_NUMLIGNE') then
       SetControlEnabled('BFLECHEBAS',False)
    else
       SetControlEnabled('BFLECHEBAS',True);

    end
  else
    begin
      if GetField( 'RPA_NUMLIGNE') <> 0 then
        ModifAutorisee:=False;
      EnvoiMailClick(Ecran);
    end;
  if GetField('RPA_GESTRAPPEL') = 'X' then
    SetControlEnabled('RPA_DELAIRAPPEL',true)
  else
    SetControlEnabled('RPA_DELAIRAPPEL',false) ;
  if GetField('RPA_GESTDATECH') = 'X' then
     begin
     SetControlEnabled('RPA_DELAIDATECH',true);
     SetControlEnabled('TRPA_DELAIDATECH',true);
     if GetField('RPA_DELAIDATECH') = 0 then SetControlEnabled('RPA_WEEKEND',False)
     else SetControlEnabled('RPA_WEEKEND',true);
     end
  else
     begin
     SetControlEnabled('RPA_DELAIDATECH',false) ;
     SetControlEnabled('TRPA_DELAIDATECH',false);
     SetControlEnabled('RPA_WEEKEND',false);
     end;
  if Assigned(GetControl('DUREEACT')) then
    if not VarIsNull (Getfield('RPA_DUREEACTION')) then
      SetControlText ('DUREEACT',PCSDureeCombo(Getfield('RPA_DUREEACTION')))
    else
      SetControlText ('DUREEACT','');
  MailEtatValClick(Ecran);
  MailEtatPreClick(Ecran);
  MailEtatReaClick(Ecran);
  MailEtatAnuClick(Ecran);

  SetControlEnabled('BVALIDER',ModifAutorisee);
{$IFDEF QUALITE}
	{ Qualité }
  if stProduitpgi = 'QNC' then
  begin
		SetControlProperty('RPA_MAILRESP','PLUS', ' AND CO_CODE="RSP"');
	  SetControlProperty('RPA_NATURETRAVAIL', 'ENABLED', GetControlText('RPA_QNATUREACTION')='WPR');
  end;
{$ENDIF QUALITE}
end ;

procedure TOM_PARACTIONS.OnChangeField ( F: TField ) ;
begin
  Inherited ;
{$IFDEF QUALITE}
  if (F.FieldName = 'RPA_TYPEACTION') then
  begin
  	if ( Ecran is TFSaisieList) and (DS.State in [DsInsert]) then
    	RTChChgtTypeAction();
  end
  else if (F.FieldName = 'RPA_DELAIDATECH') then
	begin
  	if (DS.State <> dsBrowse) then
  	begin
    	if GetField ('RPA_DELAIDATECH') > 0 then SetControlEnabled('RPA_WEEKEND',True)
    	else begin SetControlEnabled('RPA_WEEKEND',False); SetField ('RPA_WEEKEND','-'); end;
  	end;
  end
  ;
  if (F.FieldName = 'RPA_ETATACTION') and ( Ecran is TFSaisieList) then
    if GetField('RPA_ETATACTION') <> nonRealisable then
      begin
      SetControlEnabled('RPA_MAILETATVAL',true);
      SetControlEnabled('RPA_MAILRESPVAL',true);
      end
    else
      begin
      SetControlEnabled('RPA_MAILETATVAL',False);
      SetControlEnabled('RPA_MAILRESPVAL',False);
      end;
  
{$ELSE QUALITE}
  if (F.FieldName = 'RPA_TYPEACTION') and ( Ecran is TFSaisieList)
                  and (DS.State in [DsInsert]) then
     RTChChgtTypeAction();
  if (F.FieldName = 'RPA_DELAIDATECH') and (DS.State <> dsBrowse) then
     begin
     if GetField ('RPA_DELAIDATECH') > 0 then SetControlEnabled('RPA_WEEKEND',True)
     else begin SetControlEnabled('RPA_WEEKEND',False); SetField ('RPA_WEEKEND','-'); end;
     end;
{$ENDIF QUALITE}
end ;

procedure TOM_PARACTIONS.OnArgument ( S: String ) ;
var Tcb,Tcb1,Tcb2,Tcb3,Tcb4,Tcb5:TCheckBox;
begin
  Inherited ;
  stProduitpgi:=S;
  if (Ecran<>nil) and (Ecran is TFSaisieList ) then
    begin
    TF := TFSaisieList(Ecran).LeFiltre;
    //stProduitpgi:=TF.TOBFiltre.GetValue('RPG_PRODUITPGI');
    end;

  if TF = Nil then
  begin
    Tcb1:=TCheckBox(GetControl('RPA_ENVOIMAIL'));
    Tcb1.OnClick:=EnvoiMailClick;
    if (Not WriteYPLanning ('RAI')) or (stProduitpgi = 'QNC') then SetControlVisible ('RPA_PLANIFIABLE',False);
  end;  

  Tcb5:=TCheckBox(GetControl('RPA_MAILETATVAL'));
  Tcb5.OnClick:=MailEtatValClick;
  Tcb2:=TCheckBox(GetControl('RPA_MAILETATPRE'));
  Tcb2.OnClick:=MailEtatPreClick;
  Tcb3:=TCheckBox(GetControl('RPA_MAILETATREA'));
  Tcb3.OnClick:=MailEtatReaClick;
  Tcb4:=TCheckBox(GetControl('RPA_MAILETATANU'));
  Tcb4.OnClick:=MailEtatAnuClick;
{$IFDEF CCS3}
  SetControlVisible ('LNOTIFICATION1',False);
  SetControlVisible ('TRAIT3',False);
  SetControlVisible ('RPA_MODIFRESP',False);
{$ENDIF}
  if stProduitpgi='GRC' then
  begin
    SetControlVisible ('TRPA_TABLELIBREF1',False); SetControlVisible ('RPA_TABLELIBREF1',False);
    SetControlVisible ('TRPA_TABLELIBREF2',False); SetControlVisible ('RPA_TABLELIBREF2',False);
    SetControlVisible ('TRPA_TABLELIBREF3',False); SetControlVisible ('RPA_TABLELIBREF3',False);
    soRtactrappel:=GetParamsocSecur('SO_RTACTRAPPEL',True);
    soRtactchoixrappel:=GetParamSocSecur('SO_RTACTCHOIXRAPPEL','001');
    soRtchsuppact:=GetParamSocSecur('SO_RTCHSUPPACT',False);
    soRtactgestech:=GetParamSocSecur('SO_RTACTGESTECH',False);
    soRtactgestcout:=GetParamSocSecur('SO_RTACTGESTCOUT',False);
    soRtactgestchrono:=GetParamSocSecur('SO_RTACTGESTCHRONO',False);
    if ( GetParamSocSecur('SO_RTGESTINFOS001',False)=false ) then SetControlVisible ('BZOOM',False) ;
    if TF <> Nil then
    begin
      SetControlVisible ('TRPG_TABLELIBREF1',False); SetControlVisible ('RPG_TABLELIBREF1',False);
      SetControlVisible ('TRPG_TABLELIBREF2',False); SetControlVisible ('RPG_TABLELIBREF2',False);
      SetControlVisible ('TRPG_TABLELIBREF3',False); SetControlVisible ('RPG_TABLELIBREF3',False);
    end;
  end
{$IFDEF QUALITE}
  else if stProduitPGI='QNC' then
  begin
 	  SetControlVisible ('PNPRODUCTION'		, True);
    if not (ctxGPAO in V_PGI.PGIContexte) then
    begin
    	SetControlVisible ('RPA_NATURETRAVAIL' , False);
    	SetControlVisible ('TRPA_NATURETRAVAIL', False);
    	SetControlProperty('RPA_QNATUREACTION' , 'PLUS', ' AND CO_CODE<>"WPR"');
    end;
    SetControlVisible ('TRPA_TABLELIBREF1',False); SetControlVisible ('RPA_TABLELIBREF1',False);
    SetControlVisible ('TRPA_TABLELIBREF2',False); SetControlVisible ('RPA_TABLELIBREF2',False);
    SetControlVisible ('TRPA_TABLELIBREF3',False); SetControlVisible ('RPA_TABLELIBREF3',False);
    soRtactrappel			:=GetParamsocSecur('SO_RTACTRAPPEL',True);
    soRtactchoixrappel:=GetParamSocSecur('SO_RTACTCHOIXRAPPEL','001');
    soRtchsuppact			:=GetParamSocSecur('SO_RTCHSUPPACT',False);
    soRtactgestech		:=GetParamSocSecur('SO_RTACTGESTECH',False);
    soRtactgestcout		:=GetParamSocSecur('SO_RTACTGESTCOUT',False);
    soRtactgestchrono	:= False;
    if ( GetParamSocSecur('SO_RTGESTINFOS001',False)=false ) then SetControlVisible ('BZOOM',False) ;
    if TF <> Nil then
    begin
      SetControlVisible ('TRPG_TABLELIBREF1',False); SetControlVisible ('RPG_TABLELIBREF1',False);
      SetControlVisible ('TRPG_TABLELIBREF2',False); SetControlVisible ('RPG_TABLELIBREF2',False);
      SetControlVisible ('TRPG_TABLELIBREF3',False); SetControlVisible ('RPG_TABLELIBREF3',False);
    end;
		SetControlProperty('PAGE2', 'TABVISIBLE', False);
  end
{$ENDIF QUALITE}
  else
    begin
    SetControlVisible ('TRPA_TABLELIBRE1',False); SetControlVisible ('RPA_TABLELIBRE1',False);
    SetControlVisible ('TRPA_TABLELIBRE2',False); SetControlVisible ('RPA_TABLELIBRE2',False);
    SetControlVisible ('TRPA_TABLELIBRE3',False); SetControlVisible ('RPA_TABLELIBRE3',False);
    soRtactrappel:=GetParamsocSecur('SO_RFACTRAPPEL',True);
    soRtactchoixrappel:=GetParamSocSecur('SO_RFACTCHOIXRAPPEL','001');
    soRtchsuppact:=GetParamSocSecur('SO_RFCHSUPPACT',False);
    soRtactgestech:=GetParamSocSecur('SO_RFACTGESTECH',false);
    soRtactgestcout:=GetParamSocSecur('SO_RFACTGESTCOUT',False);
    soRtactgestchrono:=GetParamSocSecur('SO_RFACTGESTCHRONO',False);
    SetControlVisible ('BZOOM',False) ;
    if TF <> Nil then
        begin
        SetControlProperty ('RPA_TYPEACTION','Datatype','RTTYPEACTIONFOU');
        SetControlVisible ('TRPG_TABLELIBRE1',False); SetControlVisible ('RPG_TABLELIBRE1',False);
        SetControlVisible ('TRPG_TABLELIBRE2',False); SetControlVisible ('RPG_TABLELIBRE2',False);
        SetControlVisible ('TRPG_TABLELIBRE3',False); SetControlVisible ('RPG_TABLELIBRE3',False);
        end;
    end;
  if soRtactrappel = FALSE then
     begin
     SetControlEnabled('RPA_GESTRAPPEL',FALSE) ;
     SetControlEnabled('RPA_DELAIRAPPEL',FALSE) ;
     end
  else
     begin
     Tcb:=TCheckBox(GetControl('RPA_GESTRAPPEL'));
     Tcb.OnClick:=RappelClick;
     end;
  if soRtactgestech= False then
     begin
     SetControlEnabled('RPA_GESTDATECH',False);
     SetControlEnabled('RPA_DELAIDATECH',False);
     SetControlEnabled('TRPA_DELAIDATECH',False);
     SetControlEnabled('RPA_WEEKEND',False);
     end
  else
     begin
     Tcb:=TCheckBox(GetControl('RPA_GESTDATECH'));
     Tcb.OnClick:=GestEchClick;
     end;
  if soRtactgestcout= False then SetControlEnabled('RPA_GESTCOUT',False);
  if soRtactgestchrono= False then SetControlEnabled('RPA_GESTCHRONO',False);

{$IFDEF QUALITE}
  if Assigned(GetControl('RPA_QNATUREACTION')) then
    thValComboBox(GetControl('RPA_QNATUREACTION')).OnChange:= QNatureAction_OnChange;
{$ENDIF QUALITE}

  if Assigned(GetControl('DUREEACT')) then
    thValComboBox(GetControl('DUREEACT')).OnClick:= DUREEACT_OnClick;

end ;

procedure TOM_PARACTIONS.OnCancelRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARACTIONS.LookEcran (valeur : integer);
var TobTypActEncours : Tob;
begin
   VH_RT.TobTypesAction.Load;

   TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_TYPEACTION','RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RPA_TYPEACTION'),'---',0],TRUE) ;
   if TobTypActEncours = Nil then exit;
   if Valeur <> 1 then
       begin
       if (Valeur = 2) and (not(DS.State in [dsInsert,dsEdit])) and (TF<>Nil) then
            TF.Edit;
       SetField('RPA_GESTDATECH', TobTypActEncours.GetValue('RPA_GESTDATECH')  );
       SetField('RPA_DELAIDATECH',TobTypActEncours.GetValue('RPA_DELAIDATECH')) ;
       SetField('RPA_WEEKEND',TobTypActEncours.GetValue('RPA_WEEKEND')) ;
       SetField('RPA_GESTCOUT', TobTypActEncours.GetValue('RPA_GESTCOUT') ) ;
       SetField('RPA_GESTCHRONO', TobTypActEncours.GetValue('RPA_GESTCHRONO') ) ;
       SetField ('RPA_TABLELIBRE1',TobTypActEncours.GetValue('RPA_TABLELIBRE1'));
       SetField ('RPA_TABLELIBRE2',TobTypActEncours.GetValue('RPA_TABLELIBRE2'));
       SetField ('RPA_TABLELIBRE3',TobTypActEncours.GetValue('RPA_TABLELIBRE3'));
       SetField ('RPA_TABLELIBREF1',TobTypActEncours.GetValue('RPA_TABLELIBREF1'));
       SetField ('RPA_TABLELIBREF2',TobTypActEncours.GetValue('RPA_TABLELIBREF2'));
       SetField ('RPA_TABLELIBREF3',TobTypActEncours.GetValue('RPA_TABLELIBREF3'));

       SetField('RPA_GESTRAPPEL', TobTypActEncours.GetValue('RPA_GESTRAPPEL') ) ;
       SetField ('RPA_DELAIRAPPEL',TobTypActEncours.GetValue('RPA_DELAIRAPPEL'));
       SetField('RPA_OUTLOOK', TobTypActEncours.GetValue('RPA_OUTLOOK'));
       SetField('RPA_MODIFRESP', TobTypActEncours.GetValue('RPA_MODIFRESP'));
       if not VarIsNull (TobTypActEncours.GetValue('RPA_DUREEACTION')) then
         begin
         SetField('RPA_DUREEACTION', TobTypActEncours.GetValue('RPA_DUREEACTION'));
         SetControlText ('DUREEACT',PCSDureeCombo(Getfield('RPA_DUREEACTION')));
         end;
       if GetField('RPA_GESTDATECH') = 'X' then
           begin
           SetControlEnabled('RPA_DELAIDATECH',true);
           SetControlEnabled('TRPA_DELAIDATECH',true);
           if GetField('RPA_DELAIDATECH') = 0 then SetControlEnabled('RPA_WEEKEND',False)
           else SetControlEnabled('RPA_WEEKEND',true);
           end
       else
           begin
           SetControlEnabled('RPA_DELAIDATECH',false) ;
           SetControlEnabled('TRPA_DELAIDATECH',false);
           SetControlEnabled('RPA_WEEKEND',false);
           end;
       end;
   if Valeur <> 2 then
   begin
     if (Valeur = 1) and (not(DS.State in [dsInsert,dsEdit])) and (TF<>Nil) then
          TF.Edit;

     SetField('RPA_MAILETATVAL', TobTypActEncours.GetValue('RPA_MAILETATVAL') );
     SetField('RPA_MAILRESPVAL', TobTypActEncours.GetValue('RPA_MAILRESPVAL') );
     SetField('RPA_MAILVALAUTO', TobTypActEncours.GetValue('RPA_MAILVALAUTO') );

     SetField('RPA_MAILETATPRE', TobTypActEncours.GetValue('RPA_MAILETATPRE') );
     SetField('RPA_MAILRESPPRE', TobTypActEncours.GetValue('RPA_MAILRESPPRE') );
     SetField('RPA_MAILPREAUTO', TobTypActEncours.GetValue('RPA_MAILPREAUTO') );

     SetField('RPA_MAILETATREA', TobTypActEncours.GetValue('RPA_MAILETATREA') );
     SetField('RPA_MAILRESPREA', TobTypActEncours.GetValue('RPA_MAILRESPREA') );
     SetField('RPA_MAILREAAUTO', TobTypActEncours.GetValue('RPA_MAILREAAUTO') );

     SetField('RPA_MAILETATANU', TobTypActEncours.GetValue('RPA_MAILETATANU') );
     SetField('RPA_MAILRESPANU', TobTypActEncours.GetValue('RPA_MAILRESPANU') );
     SetField('RPA_MAILANUAUTO', TobTypActEncours.GetValue('RPA_MAILANUAUTO') );
     MailEtatValClick(Ecran);
     MailEtatPreClick(Ecran);
     MailEtatReaClick(Ecran);
     MailEtatAnuClick(Ecran);
     if (ds.State=dsInsert) and (TobTypActEncours.GetValue('RPA_REPLICLIB') = 'X' )
        and (Valeur <> 1) and (Valeur <> 2) then
       SetField ('RPA_LIBELLE',THValComboBox(GetControl('RPA_TYPEACTION')).Text);
   end;
end;

procedure TOM_PARACTIONS.RTChChgtTypeAction;
begin
if (GetField ('RPA_TYPEACTION') <> '') and (GetField ('RPA_TYPEACTION') <> Old_Typeaction) then
   begin
   Old_Typeaction := GetField ('RPA_TYPEACTION');
   LookEcran(0);
   end;
end;

procedure TOM_PARACTIONS.RTChaineFlecheBas;
var INum,MaxNum : integer;
    TobTypActEncours,TobInter : Tob;
    StEtat1,StEtat2 : String;
    Index1,Index2 : Integer;
begin
  VH_RT.TobTypesAction.Load;

  //if Getfield('RPA_NUMLIGNE')=TF.MaxValue('RPA_NUMLIGNE') then exit;
  if TF.Recno = TF.Reccount then exit;
{ le déplacement d'une ligne modifie son numéro. Si ce type de chainage a déjà été utilisé,
  l'action correspondante perdra le comportement prévu du fait de la modification des lignes
  du type de chainage (idem suppression d'une ligne ds type chainage)
  le déplacement est donc possible sur choix ParamSoc
  }
  if ExisteSQL( 'SELECT RAC_NUMACTION FROM ACTIONS WHERE RAC_TYPEACTION="'+GetField('RPA_TYPEACTION')+
     '" AND RAC_CHAINAGE="'+GetField('RPA_CHAINAGE')+
     '" AND RAC_NUMLIGNE='+IntToStr(GetField('RPA_NUMLIGNE'))) then
  begin
    if soRtchsuppact = false then
    begin
      LastError:=3;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),ecran.Caption);
      exit ;
    end ;
  end;
  //PGIInfo(TraduireMemoire('Vous modifiez l''ordre des actions, vérifiez les paramètres relatifs à la chronologie'),ecran.caption);
  { controle également de la ligne suivante }
  TF.Next;
  if ExisteSQL( 'SELECT RAC_NUMACTION FROM ACTIONS WHERE RAC_TYPEACTION="'+GetField('RPA_TYPEACTION')+
     '" AND RAC_CHAINAGE="'+GetField('RPA_CHAINAGE')+
     '" AND RAC_NUMLIGNE='+IntToStr(GetField('RPA_NUMLIGNE'))) then
  begin
    if soRtchsuppact = false then
    begin
      LastError:=3;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),ecran.Caption);
      exit ;
    end ;
  end;
  TF.SelectRecord(TF.Recno-1);
  { fin controle de la ligne suivante }

  INum:=Getfield('RPA_NUMLIGNE');
  MaxNum:=TF.MaxValue('RPA_NUMLIGNE');
  StEtat1:=Getfield('RPA_ETATACTION');

  if DS.State = dsEdit then
      TF.Post
  else  if DS.State = dsInsert then TF.Insert;
  // ligne en cours : numligne=max+1

  TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RPA_CHAINAGE'),INum],TRUE) ;
  if (TobTypActEncours <> Nil) then
      TobTypActEncours.PutValue('RPA_NUMLIGNE',MaxNum+1)
  else exit;
  Index1:=TobTypActEncours.GetIndex;
  ExecuteSql ('update paractions set rpa_numligne='+IntToStr(TF.MaxValue('RPA_NUMLIGNE')+1)+' where rpa_chainage="'+
         GetField('RPA_CHAINAGE')+'" and rpa_numligne='+IntToStr(INum));

  // ligne suivante : numligne= numligne en cours
  TF.Next;
  Index2:=0;
  TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RPA_CHAINAGE'),GetField('RPA_NUMLIGNE')],TRUE) ;
  if (TobTypActEncours <> Nil) then
  begin
    Index2:=TobTypActEncours.GetIndex;
    TobTypActEncours.PutValue('RPA_NUMLIGNE',INum);
    StEtat2:=Getfield('RPA_ETATACTION');
    TobTypActEncours.PutValue('RPA_ETATACTION',StEtat1);
  end;
  ExecuteSql ('update paractions set rpa_numligne='+IntToStr(INum)+', rpa_etataction="'+StEtat1+
    '" where rpa_chainage="'+GetField('RPA_CHAINAGE')+'" and rpa_numligne='+IntToStr(GetField('RPA_NUMLIGNE')));

  // ligne max+1 : numligne= numligne ligne suivante
  TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RPA_CHAINAGE'),MaxNum+1],TRUE) ;
  if (TobTypActEncours <> Nil) then
  begin
    TobTypActEncours.PutValue('RPA_NUMLIGNE',GetField('RPA_NUMLIGNE'));
    TobTypActEncours.PutValue('RPA_ETATACTION',StEtat2);
  end;
  ExecuteSql ('update paractions set rpa_numligne='+IntToStr(GetField('RPA_NUMLIGNE'))+', rpa_etataction="'+StEtat2+
    '" where rpa_chainage="'+GetField('RPA_CHAINAGE')+'" and rpa_numligne='+IntToStr(TF.MaxValue('RPA_NUMLIGNE')+1));
  { on interverti les tob en mémoire car elle est triée sur chainage/num ligne }
  TobInter:=TOB.Create('je ne sais pas a priori',Nil,-1) ;
  TobInter.Dupliquer(VH_RT.TobTypesAction.Detail[Index1],FALSE,TRUE,False) ;
  VH_RT.TobTypesAction.Detail[Index1].Dupliquer(VH_RT.TobTypesAction.Detail[Index2],FALSE,TRUE,False) ;
  VH_RT.TobTypesAction.Detail[Index2].Dupliquer(TobInter,FALSE,TRUE,False) ;
  TobInter.Free;
  INum:=TF.Recno;
  TF.RefreshLignes;
  TF.SelectRecord(INum);
end;


procedure TOM_PARACTIONS.RTChaineFlecheHaut;
var INum,MaxNum : integer;
    TobTypActEncours,TobInter : Tob;
    StEtat1,StEtat2 : String;
    Index1,Index2 : Integer;
begin
  VH_RT.TobTypesAction.Load;

  if TF.Recno = 1 then exit;
  INum:=GetField('RPA_NUMLIGNE');
  MaxNum:=TF.MaxValue('RPA_NUMLIGNE');
  StEtat1:=Getfield('RPA_ETATACTION');
  { le déplacement d'une ligne modifie son numéro. Si ce type de chainage a déjà été utilisé,
  l'action correspondante perdra le comportement prévu du fait de la modification des lignes
  du type de chainage (idem suppression d'une ligne ds type chainage)
  le déplacement est donc possible sur choix ParamSoc
  }
  if ExisteSQL( 'SELECT RAC_NUMACTION FROM ACTIONS WHERE RAC_TYPEACTION="'+GetField('RPA_TYPEACTION')+
     '" AND RAC_CHAINAGE="'+GetField('RPA_CHAINAGE')+
     '" AND RAC_NUMLIGNE='+IntToStr(GetField('RPA_NUMLIGNE'))) then
  begin
    if soRtchsuppact = false then
    begin
      LastError:=3;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),ecran.Caption);
      exit ;
    end ;
  end;
  { controle également de la ligne précédente }
  TF.SelectRecord(TF.Recno-1);
  if ExisteSQL( 'SELECT RAC_NUMACTION FROM ACTIONS WHERE RAC_TYPEACTION="'+GetField('RPA_TYPEACTION')+
     '" AND RAC_CHAINAGE="'+GetField('RPA_CHAINAGE')+
     '" AND RAC_NUMLIGNE='+IntToStr(GetField('RPA_NUMLIGNE'))) then
  begin
    if soRtchsuppact = false then
    begin
      LastError:=3;
      PGIBox(TraduireMemoire(TexteMessage[LastError]),ecran.Caption);
      exit ;
    end ;
  end;
  TF.Next;
  { fin controle de la ligne suivante }  
  //PGIInfo(TraduireMemoire('Vous modifiez l''ordre des actions, vérifiez les paramètres relatifs à la chronologie'),ecran.caption);
  if DS.State = dsEdit then
      TF.Post
  else  if DS.State = dsInsert then TF.Insert;
  // ligne en cours : numligne=max+1
  TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RPA_CHAINAGE'),INum],TRUE) ;
  if (TobTypActEncours <> Nil) then
      TobTypActEncours.PutValue('RPA_NUMLIGNE',MaxNum+1)
  else exit;
  Index1:=TobTypActEncours.GetIndex;
  ExecuteSql ('update paractions set rpa_numligne='+IntToStr(MaxNum+1)+' where rpa_chainage="'+
          GetField('RPA_CHAINAGE')+'" and rpa_numligne='+IntToStr(INum));

  // ligne précédente
  //TF.Prior;
  Index2:=0;
  TF.SelectRecord(TF.Recno-1);
  TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RPA_CHAINAGE'),GetField('RPA_NUMLIGNE')],TRUE) ;
  if (TobTypActEncours <> Nil) then
  begin
    Index2:=TobTypActEncours.GetIndex;
    TobTypActEncours.PutValue('RPA_NUMLIGNE',INum);
    StEtat2:=Getfield('RPA_ETATACTION');
    TobTypActEncours.PutValue('RPA_ETATACTION',StEtat1);
  end;
  ExecuteSql ('update paractions set rpa_numligne='+IntToStr(INum)+', rpa_etataction="'+StEtat1+
    '" where rpa_chainage="'+GetField('RPA_CHAINAGE')+'" and rpa_numligne='+IntToStr(GetField('RPA_NUMLIGNE')));

  // ligne max+1 : numligne= numligne ligne précédente
  TobTypActEncours:=VH_RT.TobTypesAction.FindFirst(['RPA_CHAINAGE','RPA_NUMLIGNE'],[GetField('RPA_CHAINAGE'),MaxNum+1],TRUE) ;
  if (TobTypActEncours <> Nil) then
  begin
    TobTypActEncours.PutValue('RPA_NUMLIGNE',GetField('RPA_NUMLIGNE'));
    TobTypActEncours.PutValue('RPA_ETATACTION',StEtat2);
  end;
  ExecuteSql ('update paractions set rpa_numligne='+IntToStr(GetField('RPA_NUMLIGNE'))+', rpa_etataction="'+StEtat2+
    '" where rpa_chainage="'+GetField('RPA_CHAINAGE')+'" and rpa_numligne='+IntToStr(MaxNum+1));
  { on interverti les tob en mémoire car elle est triée sur chainage/num ligne }
  TobInter:=TOB.Create('je ne sais pas a priori',Nil,-1) ;
  TobInter.Dupliquer(VH_RT.TobTypesAction.Detail[Index1],FALSE,TRUE,False) ;
  VH_RT.TobTypesAction.Detail[Index1].Dupliquer(VH_RT.TobTypesAction.Detail[Index2],FALSE,TRUE,False) ;
  VH_RT.TobTypesAction.Detail[Index2].Dupliquer(TobInter,FALSE,TRUE,False) ;
  TobInter.Free;
  INum:=TF.Recno;
  TF.RefreshLignes;
  TF.SelectRecord(INum);
end;

procedure TOM_PARACTIONS.RappelClick(Sender: TObject);
begin
{if (DS.State in [dsEdit,dsInsert]) then
    begin}
      if GetCheckBoxState('RPA_GESTRAPPEL') = cbChecked then
        begin
        SetControlEnabled('RPA_DELAIRAPPEL',true) ;
        SetField('RPA_DELAIRAPPEL',soRtactchoixrappel);
        end
      else
        begin
        SetControlEnabled('RPA_DELAIRAPPEL',FALSE) ;
        SetField('RPA_DELAIRAPPEL','');
        end;
    {end;}
end;

procedure TOM_PARACTIONS.EnvoiMailClick(Sender: TObject);
begin
  if GetCheckBoxState('RPA_ENVOIMAIL') = cbChecked then
    begin
    SetControlEnabled('RPA_MAILRESP',true) ;
    SetControlEnabled('RPA_MAILRESPAUTO',true) ;
{$IFDEF QUALITE}
		{ Module Qualité }
  	if (stProduitPGI = 'QNC') and (DS.State<>dsBrowse) then
  	begin
			SetField('RPA_MAILRESP','RSP');
			SetControlText('RPA_MAILRESP','RSP');
  	end;
{$ENDIF QUALITE}
    end
  else
    begin
    SetControlEnabled('RPA_MAILRESP',false) ;
    SetControlEnabled('RPA_MAILRESPAUTO',false) ;
    SetField('RPA_MAILRESP','') ;
    SetField('RPA_MAILRESPAUTO','-') ;
    end;
end;
procedure TOM_PARACTIONS.MailEtatValClick(Sender: TObject);
begin
  if TCheckBox(GetControl('RPA_MAILETATVAL')).Enabled = true then
  begin
    if GetCheckBoxState('RPA_MAILETATVAL') = cbChecked then
      begin
      SetControlEnabled('RPA_MAILVALAUTO',true) ;
      SetField('RPA_MAILRESPVAL','RSP') ;
      end
    else
      begin
      SetControlEnabled('RPA_MAILVALAUTO',false) ;
      if DS.State <> dsBrowse then
        begin
        SetField('RPA_MAILRESPVAL','') ;
        SetField('RPA_MAILVALAUTO','-') ;
        end;
      end;
  end;
end;
procedure TOM_PARACTIONS.MailEtatPreClick(Sender: TObject);
begin
  if GetCheckBoxState('RPA_MAILETATPRE') = cbChecked then
    begin
    SetControlEnabled('RPA_MAILRESPPRE',true) ;
    SetControlEnabled('RPA_MAILPREAUTO',true) ;
    end
  else
    begin
    SetControlEnabled('RPA_MAILRESPPRE',false) ;
    SetControlEnabled('RPA_MAILPREAUTO',false) ;
    if DS.State <> dsBrowse then
      begin
      SetField('RPA_MAILRESPPRE','') ;
      SetField('RPA_MAILPREAUTO','-') ;
      end;
    end;
end;
procedure TOM_PARACTIONS.MailEtatReaClick(Sender: TObject);
begin
  if GetCheckBoxState('RPA_MAILETATREA') = cbChecked then
    begin
    SetControlEnabled('RPA_MAILRESPREA',true) ;
    SetControlEnabled('RPA_MAILREAAUTO',true) ;
    end
  else
    begin
    SetControlEnabled('RPA_MAILRESPREA',false) ;
    SetControlEnabled('RPA_MAILREAAUTO',false) ;
    if DS.State <> dsBrowse then
      begin
      SetField('RPA_MAILRESPREA','') ;
      SetField('RPA_MAILREAAUTO','-') ;
      end;
    end;
end;
procedure TOM_PARACTIONS.MailEtatAnuClick(Sender: TObject);
begin
  if GetCheckBoxState('RPA_MAILETATANU') = cbChecked then
    begin
    SetControlEnabled('RPA_MAILRESPANU',true) ;
    SetControlEnabled('RPA_MAILANUAUTO',true) ;
    end
  else
    begin
    SetControlEnabled('RPA_MAILRESPANU',false) ;
    SetControlEnabled('RPA_MAILANUAUTO',false) ;
    if DS.State <> dsBrowse then
      begin
      SetField('RPA_MAILRESPANU','') ;
      SetField('RPA_MAILANUAUTO','-') ;
      end;
    end;
end;

procedure TOM_PARACTIONS.GestEchClick(Sender: TObject);
begin
  if GetCheckBoxState('RPA_GESTDATECH') = cbChecked then
    begin
    SetControlEnabled('RPA_DELAIDATECH',true) ;
    SetControlEnabled('TRPA_DELAIDATECH',true) ;
    if (GetField('RPA_DELAIDATECH') = 0) then SetControlEnabled('RPA_WEEKEND',False)
    else SetControlEnabled('RPA_WEEKEND',true) ;
    end
  else
    begin
    SetControlEnabled('RPA_DELAIDATECH',false) ;
    SetControlEnabled('TRPA_DELAIDATECH',false) ;
    SetControlEnabled('RPA_WEEKEND',false) ;
    if DS.State <> dsBrowse then
      begin
      SetField('RPA_DELAIDATECH',0) ;
      SetField('RPA_WEEKEND','-') ;
      end;
    end;
end;

procedure TOM_PARACTIONS.DUREEACT_OnClick(Sender: TObject);
begin
  SetField('RPA_DUREEACTION',Valeur(GetControlText('DUREEACT')));
  //ForceUpdate;
  if DS.State = dsBrowse then DS.Edit;
  if Assigned (TF) then
    UpdateChamp('RPA_DUREEACTION');
end      ;

{procedure AGLRTChChgtTypeAction( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_PARACTIONS) then TOM_PARACTIONS(OM).RTChChgtTypeAction else exit;
end;
}
procedure AGLRTValeursDef1( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_PARACTIONS) then TOM_PARACTIONS(OM).LookEcran(1) else exit;
end;

procedure AGLRTValeursDef2( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_PARACTIONS) then TOM_PARACTIONS(OM).LookEcran(2) else exit;
end;

procedure AGLRTChaineFlecheBas( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_PARACTIONS) then TOM_PARACTIONS(OM).RTChaineFlecheBas else exit;
end;

procedure AGLRTChaineFlecheHaut( parms: array of variant; nb: integer ) ;
var  F : TForm ;
     OM : TOM ;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFFiche) then OM:=TFFiche(F).OM else
   if (F is TFSaisieList) then OM:=TFSaisieList(F).LeFiltre.OM else exit;
if (OM is TOM_PARACTIONS) then TOM_PARACTIONS(OM).RTChaineFlecheHaut else exit;
end;

procedure AGLRTActionsTypeImprime( parms: array of variant; nb: integer ) ;
var StChainage,stWhere,stProduit : string;
begin
StChainage:=string(Parms[1]);
stProduit:=string(Parms[2]);
stWhere:=' RPA_CHAINAGE="'+StChainage+'" and RPA_PRODUITPGI="'+stProduit+'"';
if stProduit='GRC' then
  LanceEtat('E','RTA','RTA',True,False,False,Nil,stWhere,'',False,0,'')
else
  LanceEtat('E','RTA','RTF',True,False,False,Nil,stWhere,'',False,0,'');
end;


{$IFDEF QUALITE}
procedure TOM_PARACTIONS.QNatureAction_OnChange(Sender: TObject);
begin
  SetControlProperty('RPA_NATURETRAVAIL', 'ENABLED', GetControlText('RPA_QNATUREACTION')='WPR');
	if (GetControlText('RPA_QNATUREACTION')<>'WPR') then  // Lié à la production
    SetControlText('RPA_NATURETRAVAIL', '');
end;
{$ENDIF QUALITE}
Initialization
  registerclasses ( [ TOM_PARACTIONS ] ) ;
//  RegisterAglProc( 'RTChChgtTypeAction', TRUE , 0, AGLRTChChgtTypeAction);
RegisterAglProc( 'RTValeursDef1', TRUE , 0, AGLRTValeursDef1);
RegisterAglProc( 'RTValeursDef2', TRUE , 0, AGLRTValeursDef2);
RegisterAglProc( 'RTChaineFlecheBas', TRUE , 0, AGLRTChaineFlecheBas);
RegisterAglProc( 'RTChaineFlecheHaut', TRUE , 0, AGLRTChaineFlecheHaut);
RegisterAglProc( 'RTActionsTypeImprime', TRUE , 2, AGLRTActionsTypeImprime);
end.

