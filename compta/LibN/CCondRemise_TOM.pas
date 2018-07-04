{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 04/02/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CCONDREMISE (CCONDREMISE)
Mots clefs ... : TOM;CCONDREMISE
*****************************************************************}
Unit CCONDREMISE_TOM ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
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
     UTob ;

Procedure FicheCondRemiseESPCP ( CondRemise : String ; action : TActionFiche ) ;

Type
  TOM_CCONDREMISE = Class (TOM)
   private
    Procedure BFraisClick ( Sender : TObject) ;
   protected
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
   public
    end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}

  AGLInit, HTB97, CFraisRemise_TOM, Ent1, Lookup, UlibEncaDecaESP, HDB,
{$IFDEF eAGLClient}
    MaineAGL         // AGLLanceFiche (et non pas AGLMaine !!)
{$ELSE}
    FE_Main
{$ENDIF eAGLClient}
    ;
//==================================================
// Definition des variable
//==================================================
var
    MESS : array [0..11] of string = (
     {0}   'Ces conditions ont des frais assocciés #13 Désirez-vous le supprimer aussi?' ,
     {1}   'Il faut supprimer avant les frais de remise!',
     {2}   'Vous devez renseigner une libellé!',
     {3}   'L''écart entre le nombre minimum et maximum est incohérent!',
     {4}   'Vous devez renseigner un compte général de risque!',
     {5}   'Le compte de risque renseigné n''existe pas!',
     {6}   'Vous devez renseigner un compte général pour l''intérêt!',
     {7}   'Le compte pour l''intérêt renseigné n''existe pas!',
     {8}   'Vous devez renseigner un compte général pour la banque!',
     {9}   'Le compte pour la banque renseigné n''existe pas!',
    {10}   'Le journal choisi n''appartiant pas au compte général de la banque choisi!',
    {11}   'On n''a pas renseigner le journal de génération!'
    ) ;

Procedure FicheCondRemiseESPCP ( CondRemise : String ; action : TActionFiche ) ;
Begin
  if (Action<>taConsult) and (Trim(CondRemise)<>'') and (not ExisteSQL('select CCB_CONDREMISE from CCONDREMISE where CCB_CONDREMISE="'+CondRemise+'"')) then //XMG 05/04/04
     Action:=taCreat ;
  AGLLanceFiche('CP','CPESPCONDREMISE','',CondRemise,actiontostring(Action)+';') ;
End ;

procedure TOM_CCONDREMISE.OnNewRecord ;
begin
  Inherited ;
  SetField('CCB_TYPEREMISE','ENC') ;
end ;

procedure TOM_CCONDREMISE.OnDeleteRecord ;
begin
  Inherited ;
  if ExisteSQL('select CFR_CONDREMISE from CFRAISREMISE where CFR_CONDREMISE="'+GetField('CCB_CONDREMISE')+'"') then begin
     if PGIAsk(Mess[0])=mryes then
        ExecuteSQL('delete from CFRAISREEMISE where CFR_CONDREMISE="'+GetField('CCB_CONDREMISE')+'"')
     else
        LastError:=1 ;
  end ;
  if LastError>0 then
     LastErrorMsg:=Mess[LastError] ;
end ;

procedure TOM_CCONDREMISE.OnUpdateRecord ;
begin
  Inherited ;
  if trim(GetField('CCB_LIBELLE'))='' then
     Lasterror:=2
  else
  if GetField('CCB_GENERAL')='' then
     LastError:=4
  else
  if (not LookUpValueExist( GetControl('CCB_GENERAL'))) then
     LastError:=5
  else
  if GetField('CCB_NBREJMAX')<GetField('CCB_NBREJMIN') then
     LastError:=3
  else
  if GetField('CCB_TYPEREMISE')='ESC' then begin
     if (Arrondi(GetField('CCB_PLAFOND'),V_PGI.OkDecV)>0) then Begin
        if GetField('CCB_GENERALRISQUE')='' then
           LastError:=4
        else
        if (not LookUpValueExist( GetControl('CCB_GENERALRISQUE'))) then
           LastError:=5
     End else
     if (Arrondi(GetField('CCB_TAUXINT'),CCB_NbrDecTauxInteret)>0) then Begin
        if GetField('CCB_GENERALINTERET')='' then
           LastError:=6
        else
        if (not LookUpValueExist( GetControl('CCB_GENERALINTERET'))) then
           LastError:=7
     End ;
     if trim(GetField('CCB_JOURNAL'))='' then
        lastError:=11 else
     if not ExisteSQL('select J_JOURNAL from JOURNAL where J_JOURNAL="'+GetField('CCB_JOURNAL')+'" and J_NATUREJAL="BQE" and J_CONTREPARTIE="'+GetField('CCB_GENERAL')+'"') then
        LastError:=10
  end ;
  if LastError>0 then
     LastErrorMsg:=Mess[LastError] ;
end ;

procedure TOM_CCONDREMISE.OnAfterUpdateRecord ;
begin
  Inherited ;
  if assigned(Ecran) then
     SetControlEnabled('BFraisRemise',TRUE) ;
end ;

procedure TOM_CCONDREMISE.OnLoadRecord ;
begin
  Inherited ;
  if assigned(Ecran) then
     SetControlEnabled('BFraisRemise',ds.state<>dsInsert) ;
end ;

procedure TOM_CCONDREMISE.OnChangeField ( F: TField ) ;
var ii : Integer ;
    St : String ;
begin
  Inherited ;
  if F.FieldName='CCB_TYPEREMISE' then begin
     if assigned(Ecran) then
        SetControlVisible('TSESCOMPTE',F.AsString='ESC') ;
     if F.AsString<>'ESC' then begin
        SetField('CCB_PLAFOND',0) ;
        SetField('CCB_TAUXINT',0) ;
        SetField('CCB_MODECALCULINT','NAT') ;
        SetField('CCB_GENERALRISQUE','') ;
        SetField('CCB_GENERALINTERET','') ;
        For ii:=1 to 3 do begin
            SetField('CCB_REFINTERNE'+inttostr(ii),'') ;
            SetField('CCB_LIBELLE'+inttostr(ii),'') ;
            SetField('CCB_REFEXTERNE'+inttostr(ii),'') ;
            SetField('CCB_REFLIBRE'+inttostr(ii),'') ;
        end ;
     End ;
  End else
  if (pos(';'+F.FieldName+';',';CCB_GENERAL;CCB_GENERALRISQUE;CCB_GENERALINTERET;')>0) then begin
     if (Ds.State in [dsEdit,dsInsert]) and
        (F.AsString<>'')                                                         then begin
        St:=Bourreladonc(F.AsString,fbGene) ;
        if St<>F.AsString then
           F.AsString:=St ;
     End ;
     if F.FieldName='CCB_GENERAL' then
        SetControlText('COPIEGENERAL',F.AsString) ;
  End ;
end ;

procedure TOM_CCONDREMISE.OnArgument ( S: String ) ;
begin
  Inherited ;
  if Assigned(Ecran) then begin
    TFFicheListe(Ecran).TypeAction:=stringtoaction(ReadTokenSt(s)) ;
    TToolBarButton97(GetControl('BFRAISREMISE')).OnClick:=BFraisClick ;
    THDBEdit(GetControl('CCB_TAUXINT')).DisplayFormat:=strfMask(CCB_NbrDecTauxInteret,'',TRUE,FALSE) ;
  End ;
end ;

procedure TOM_CCONDREMISE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CCONDREMISE.OnCancelRecord ;
begin
  Inherited ;
end ;

Procedure TOM_CCONDREMISE.BFraisClick ( Sender : TObject) ;
Begin
  FicheFraisRemiseESPCP(GetField('CCB_CONDREMISE'),'',TFFicheListe(Ecran).TypeAction,'Banque='+GetField('CCB_GENERAL')+';') ;
End ;

Initialization
  registerclasses ( [ TOM_CCONDREMISE ] ) ;
end.
