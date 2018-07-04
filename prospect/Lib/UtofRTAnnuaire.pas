{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 17/10/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTANNUAIRE ()
Mots clefs ... : TOF;RTANNUAIRE
*****************************************************************}
Unit UtofRTAnnuaire ;

Interface

Uses 
     Classes,
{$IFNDEF EAGLCLIENT}
     mul,
{$else}
     eMul, 
{$ENDIF}
     forms, 
     HEnt1,
     UTOF,
{$ifdef GIGI}
     UtofAftraducChampLibre,
{$endif}
{$ifdef GRC}
     UtilRT,
{$endif}
      ParamSoc,UtilSelection,SysUtils,Hctrls,UtilPgi ;

Type
{$ifdef GIGI}
  TOF_RTANNUAIRE = Class (TOF_AFTRADUCCHAMPLIBRE)   //mcd 03/07/06 pour valeur par defaut ressource
{$else}
  TOF_RTANNUAIRE = Class (TOF)
{$endif}
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    ContactAffaire: Boolean;
     stProduitpgi : string;
     Telephone : THEdit;
    procedure TelephoneChange(Sender: Tobject);
  end ;

Implementation

procedure TOF_RTANNUAIRE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTANNUAIRE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RTANNUAIRE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_RTANNUAIRE.OnLoad ;
var xx_where : String;
begin
  Inherited ;
{$IFDEF GRC}
  SetControlText('XX_WHERE','');
  // BDU - 30/11/06, Ajout du type "Contact d'un autre tiers" pour les intervenants de l'affaire
  if ContactAffaire then
    SetControlText('XX_WHERE', 'C_TYPECONTACT = "T" AND T_NATUREAUXI = "CLI"');
//  if (V_PGI.menucourant = 92) then
  if stProduitpgi = 'GRC' then
  begin
  {$ifdef GIGI}
    SetControlText('XX_WHERE','AND (T_NATUREAUXI = "CLI" or T_NATUREAUXI = "PRO" or T_NATUREAUXI = "NCP") ');
  {$else}
    SetControlText('XX_WHERE','AND (T_NATUREAUXI = "CLI" or T_NATUREAUXI = "PRO") ');
  {$endif}
    // mng 15/03/04 direct dans la fiche SetControlEnabled('T_NATUREAUXI',false);
  end;
  if (ctxGRC in V_PGI.PGIContexte) and
//     ( (V_PGI.menucourant=92) or ( (GetControlText('T_NATUREAUXI') = 'CLI') or
     ( (stProduitpgi = 'GRC') or ( (GetControlText('T_NATUREAUXI') = 'CLI') or
        (GetControlText('T_NATUREAUXI') = 'PRO') ) ) then
   begin
      if (GetControlText('XX_WHERE') = '') then
          SetControlText('XX_WHERE',RTXXWhereConfident('CON'))
      else
          begin
          xx_where := GetControlText('XX_WHERE');
          xx_where := xx_where + RTXXWhereConfident('CON');
          SetControlText('XX_WHERE',xx_where) ;
          end;
   end
   else
//     if (V_PGI.menucourant=93) or ( GetControlText('T_NATUREAUXI') = 'FOU') then SetControlText('XX_WHERE',RTXXWhereConfident('CONF'));
     if ( GetControlText('T_NATUREAUXI') = 'FOU') then SetControlText('XX_WHERE',RTXXWhereConfident('CONF'));
{$ENDIF}
end ;

procedure TOF_RTANNUAIRE.OnArgument (S : String ) ;
var F : TForm;
    ChampMul,ValMul,Critere : string ;
    x: integer ;
begin
  Inherited ;
			//mcd 21/02/2006 si appeler depuis ANNUAIRE ou JURIDIQUE , on cache les champs du tiers
			//car le lien tiers n'est pas obligatoire
  ContactAffaire := False;
  Repeat
   Critere:=uppercase(Trim(ReadTokenSt(S))) ;
   if Critere<>'' then
      begin
      x:=pos('=',Critere);
      if x<>0 then
         begin
         ChampMul:=copy(Critere,1,x-1);
         ValMul:=copy(Critere,x+1,length(Critere));
         end else
         begin
         ChampMul:=Critere; ValMul:='';
         end;

       if (ChampMul='ANNUAIRE') then 
          begin
					SetControlVisible('T_LIBELLE',False);
					SetControlVisible('TT_LIBELLE',False);
					SetControlVisible('TELEPHONE',False);
					SetControlVisible('TC_TELEPHONE',False);
					SetControlVisible('RAISONSOCIALE',False);
					SetControlVisible('LRAISONSOCIALE',False);
					SetControlVisible('T_VILLE',False);
					SetControlVisible('TT_VILLE',False);
					SetControlVisible('T_CODEPOSTAL',False);
					SetControlVisible('TT_CODEPOSTAL',False);
					SetControlVisible('PCOMPLEMENT',False);
					SetControlVisible('PZLIBRE',False);
          TfMUl(Ecran).DBListe := 'DPCONTACT' ;
          End;
       // BDU - 30/11/06, Ajout du type "Contact d'un autre tiers" pour les intervenants de l'affaire
       if ChampMul = 'CONTACT_AFFAIRE' then
         ContactAffaire := True;
       if (ChampMul = 'GRC') then stProduitpgi := 'GRC';
      end;
until  Critere='';
      //fin mcd 21/02/2006
{$IFDEF GRC}
  RTMajChampsLibresContact(TFMul(Ecran));
  if (ctxGRC in V_PGI.PGIContexte) then
  begin
      SetControlVisible('BCOURRIER',True);
      if (pos('CLI',TFMul(Ecran).FRange) <> 0) or ( TFMul(Ecran).FRange='' ) then
        begin
//        if TfMul(Ecran).Q <> NIL then TfMul(Ecran).Q.Liste  := 'RTANNUAIRE';
        if TfMul(Ecran).Q <> NIL then TfMul(Ecran).SetDBListe('RTANNUAIRE');
        if GetParamSocSecur('SO_RTGESTINFOS006',False) = True then
          begin
          F := TForm (Ecran);
          MulCreerPagesCL(F,'NOMFIC=YYCONTACT');
          end;
        end;
  end;
  if (pos('FOU',TFMul(Ecran).FRange) <> 0) then
//     if TfMul(Ecran).Q <> NIL then TfMul(Ecran).Q.Liste:='RTANNUAIREFOU' ;
     if TfMul(Ecran).Q <> NIL then TfMul(Ecran).SetDBListe('RTANNUAIREFOU') ;
  Telephone := THEdit(Ecran.FindComponent('TELEPHONE'));
  Telephone.OnChange := TelephoneChange;

{$ENDIF}

end ;

procedure TOF_RTANNUAIRE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RTANNUAIRE.OnDisplay () ;
begin
  Inherited ;
  if GetField('C_RVA') <> '' then SetControlEnabled ('BNEWMAIL',True)
  else SetControlEnabled ('BNEWMAIL',False);
end ;

procedure TOF_RTANNUAIRE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_RTANNUAIRE.TelephoneChange(Sender: Tobject);
begin
  ThEdit(getCOntrol('C_CLETELEPHONE')).Text := CleTelForFind (ThEdit(getCOntrol('TELEPHONE')).Text);
end;

Initialization
  registerclasses ( [ TOF_RTANNUAIRE ] ) ;
end.

