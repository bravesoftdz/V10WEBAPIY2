{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 15/03/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CONTABON (CONTABON)
Mots clefs ... : TOM;CONTABON
*****************************************************************}
Unit CONTABON_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     Fiche, 
     FichList,
     FE_Main,
{$else}
     eFiche, 
     eFichList,
     MaineAGL,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     UTOM,
     HTB97,
     Graphics,
     Ent1,
     UTob ;

Type
  TOM_CONTABON = Class (TOM)
   private
    FOkRes     : Boolean ;
    Function  VerifiSiExiste : Boolean ;
    procedure BGuideClick( Sender : TObject ) ;
   public
    procedure OnNewRecord                ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    end ;


Procedure ParamAbonnement(EstCompta : Boolean ; Lequel : String ; Mode : TActionFiche) ;

Implementation

uses
 UtilPGI , CPGUIDE_TOM ;


Procedure ParamAbonnement(EstCompta : Boolean ; Lequel : String ; Mode : TActionFiche) ;
var
 lBoOkRes : Boolean ;
 lStRange : string ;
 lStArg : string ;
BEGIN
 lBoOkRes:=True ;
 if Mode=taConsult then lBoOkRes:=False ;
 if ((Mode=taModif) and (Lequel<>'')) then lBoOkRes:=False ; {Appel depuis Génère abo}
 if lBoOkRes then if _Blocage(['nrCloture','nrBatch'],True,'nrBatch') then Exit ;

 if Mode= taConsult then lStArg := 'ACTION=TACONSULT;' else lStArg := 'ACTION=TAMODIF;' ;
 if EstCompta then lStRange := 'X' else lStRange := '-' ;
 if lBoOkRes then lStArg := lStArg + 'X;' else lStArg := lStArg +'-;' ;

 if Lequel <> '' then
   lStRange := lStRange + ';' + ReadTokenSt(Lequel) ;
 
 AGLLanceFiche('CP','CPCONTABON', lStRange, '', lStArg ) ;

END ;

procedure TOM_CONTABON.OnNewRecord ;
begin
 Inherited ;

 SetField('CB_RECONDUCTION'  , 'TAC') ;
 SetField('CB_SEPAREPAR'     , '1M') ;
 SetField('CB_ARRONDI'       , 'PAS') ;
 SetField('CB_QUALIFPIECE'   , 'N') ;

 SetField('CB_DATECREATION'  , Date) ;

end ;


function TOM_CONTABON.VerifiSiExiste : Boolean ;
var
 lQ : TQuery ;
begin
 lQ     := OpenSql('Select CB_COMPTABLE,CB_CONTRAT from CONTABON Where CB_COMPTABLE="'+GetField('CB_COMPTABLE')+'" '+
                   'And CB_CONTRAT="'+GetField('CB_CONTRAT')+'"',True) ;
 Result := lQ.EOF ;
 Ferme(lQ) ;
END ;

procedure TOM_CONTABON.OnUpdateRecord ;
begin

 Inherited ;

 if DS.State in [dsEdit,dsInsert] then
  begin
   if GetField('CB_CONTRAT')='' then
     begin
     LastErrorMsg := 'Vous devez renseigner un code' ;
     SetFocusControl('CB_CONTRAT') ;
     LastError := 1 ;
     exit ;
     end ;
   if GetField('CB_GUIDE')='' then
     begin
     LastErrorMsg := 'Vous devez renseigner un guide' ;
     SetFocusControl('CB_GUIDE') ;
     LastError := 1 ;
     exit ;
     end ;
  end ;

 if DS.State in [dsInsert] then
  begin
   if not VerifisiExiste then
     begin
     LastErrorMsg := 'Le code que vous avez saisi existe déjà. Vous devez le modifier' ;
     SetFocusControl('CB_CONTRAT')  ;
     LastError := 1 ;
     exit ;
     end ;

   if GetField('CB_DATECONTRAT') < VH^.Encours.Deb then
     begin
     LastErrorMsg := 'Vous ne pouvez pas définir des contrats d''abonnement sur des exercices non ouverts' ;
     LastError := 1 ;
     exit ;
     end ;
   SetField('CB_DATEDERNGENERE',iDate1900) ;
  end
 else if GetField('CB_NBREPETITION') < GetField('CB_DEJAGENERE') then
   begin
   SetField('CB_NBREPETITION', GetField('CB_DEJAGENERE') ) ;
   LastErrorMsg := 'Le nombre de répétition ne peut pas être inférieure au nombre de contrat généré.' ;
   SetFocusControl('CB_NBREPETITION')  ;
   LastError := 1 ;
   Exit ;
   end ;

 if ( DS.State in [dsInsert] ) and ( GetField('CB_DEJAGENERE') <= 0 ) then
  SetField('CB_DATEDERNGENERE',GetField('CB_DATECONTRAT')) ;

 SetField('CB_DATEMODIF',Date) ;

end ;

procedure TOM_CONTABON.OnChangeField ( F: TField ) ;
begin

 Inherited ;

 if F.FieldName = 'CB_SEPAREPAR' then
   begin
   if ( GetField('CB_SEPAREPAR') = 'SEM' ) or ( GetField('CB_SEPAREPAR') = 'QUI' ) then
     begin
     if not (DS.state in [dsBrowse] ) then
       SetField('CB_ARRONDI','PAS') ;
     SetControlEnabled('CB_ARRONDI',False) ;
     end
   else
     SetControlEnabled('CB_ARRONDI',True) ;
   end ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 05/09/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CONTABON.OnArgument ( S : String ) ;
begin

 READTOKENST(S) ;
 FOkRes := READTOKENST(S) = 'X' ;

 if FOkRes then _Bloqueur('nrBatch',false) ;

 Inherited ;

 TToolbarButton97(Getcontrol('BGUIDE',True)).OnClick := BGuideClick ;

 TFFicheListe(Ecran).HelpContext := 7337000 ;

 Ecran.Caption := Ecran.Caption + ' :' ; // pour afficher el titre de la fenetre ds les msg d'erreur
 UpdateCaption(ecran);

end ;

procedure TOM_CONTABON.OnClose ;
begin
 if FOkRes then _Bloqueur('nrBatch',false) ;
 Inherited ;
end ;


procedure TOM_CONTABON.BGuideClick(Sender: TObject);
var St : String;
    lAction : TActionFiche ;
begin
 NextPrevControl(Ecran) ; // on remettre a jour le GetField qd on change le code guide
 St := GetField('CB_GUIDE') ;

 if ( DS.State = dsInsert ) and ( St = '' )
   then lAction := taCreat
   else lAction := TFFicheListe(Ecran).TypeAction ;

 St := ParamGuide(St,'ABO', lAction ) ;

 if TFFicheListe(Ecran).TypeAction = taConsult then Exit ;
{$IFDEF EAGLCLIENT}
 THValComboBox(Getcontrol('CB_GUIDE',True)).Reload ;
{$ELSE}
 THDBValComboBox(Getcontrol('CB_GUIDE',True)).Reload ;
{$ENDIF}

 if lAction = taCreat then
   begin
   SetField( 'CB_GUIDE', St ) ;
//   SetControlText('CB_GUIDE', St ) ;
   end ;

end;

procedure TOM_CONTABON.OnLoadRecord;
begin

  inherited;

  if GetField('CB_DEJAGENERE') <= 0 then
    begin
    SetControlProperty('CB_DATECONTRAT', 'Color', clWindow ) ;
    SetControlProperty('CB_QUALIFPIECE', 'Color', clWindow ) ;
    SetControlEnabled('CB_DATECONTRAT',True) ;
    SetControlEnabled('CB_QUALIFPIECE',True) ;
    end
  else
    begin
    SetControlProperty('CB_DATECONTRAT', 'Color', clBtnFace ) ;
    SetControlProperty('CB_QUALIFPIECE', 'Color', clBtnFace ) ;
    SetControlEnabled('CB_DATECONTRAT',False) ;
    SetControlEnabled('CB_QUALIFPIECE',False) ;
    end ;

end;


Initialization
  registerclasses ( [ TOM_CONTABON ] ) ; 
end.
