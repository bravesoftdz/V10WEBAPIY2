unit UtilRT;

interface

uses  ComCtrls,Ent1, Sysutils, HCtrls, Forms,StdCtrls, HEnt1, Controls,UtilTOM,
{$IFDEF EAGLCLIENT}
     eFiche,MaineAGL,eMul,
{$ELSE}
     Fiche,{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}Mul,NumConv,
{$ENDIF}
     Menus, M3FP,Paramsoc,HMsgBox,utob,Classes, HStatus,EntRT, UtilGC,cube,EntGC,TiersUtil;

Procedure RTMajChampsLibres( FF : Tform ) ;
Procedure RTMajChampsLibresFou( FF : Tform ) ;
Procedure RTMajChampsLibresPersp( FF : Tform ) ;
Procedure RTMajChampsLibresAction( FF : Tform ) ;
Procedure RTMajChampsLibresActionF( FF : Tform ) ;
Procedure RTMajChampsLibresChainage( FF : Tform ) ;
Procedure RTMajChampsLibresChainageF( FF : Tform ) ;
Procedure RTMajChampsLibresSuspect( FF : Tform ) ;

function  RTXXWhereConfident (typeconf:string) :string;
function  RTDroitModifTiers (CodeTiers : string ): Boolean;
function  RTDroitModifFou (CodeTiers : string ): Boolean;
Procedure RTMajPopMenu( FF : Tform; MenuPop : String; NoLigne : Integer; AConserver : boolean) ;
function RTControleModifTiers (FF : Tform ;TobProspect: TOB; ActionModif :Boolean): Boolean;
function RTValeurChampLibre(NomChamp,ValeurChamp: string):string;
Procedure CommercialToRessource ; // pour test CEGID charge la table des commerciaux dans la table Ressource
function RTDroitModifActions (CodeTiers,typeAction,Responsable :string): Boolean;
function RTDroitModifActionsF (CodeTiers,typeAction,Responsable :string): Boolean;
Procedure RTCreerProspectPourClient ;  // créer l'enregistrement prospect pour les tiers clients
Procedure RTCorrespondSuspectProspect(CleChoixCode : string; FromTOB, ToTOB : TOB);
Function RTRechResponsable (Tiers: string) : string;
Function RTFormateDateHeure (Date: TDateTime; Heure : String) : String;
Function RTChangeLibre2 ( NomChamp : string ; FF : Tform ) : boolean;

function RTAttribNewCodeSuspect (DernierChrono : string):string;
Procedure RTCreerLiensConcurrents ;  // déplace les concurrents des propositions dans fichier lien
function RTDroitModifTypeAction : Boolean;
Procedure RTAlimCleTelephonTiers ;
Function RTFormateCleTelephone(Notel:String) : String;
Function RTFormatTelephonContact(Notel:String) : String;
Procedure RTAlimTelephonContact ;
Procedure RTMajNatureContacts(StAuxi,StNat : String);
Procedure RTMajNatureAdresses(StAuxi,StNat : String);
Procedure RTMajNatureFrais(StAuxi,StNat : String);
Function RTAccesMenu(NoMenu,NoTag,NoGrp:integer) : Boolean;
procedure MajProspectClient ( TOBGenere : Tob ; TOBTiers : Tob);
procedure TransformeProspectClient ( TOBGenere : Tob ; NomChamp : String);

implementation


Procedure RTMajChampsLibres( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 10 do if i<10 then ChangeLibre2('TYTC_TABLELIBRETIERS'+intToStr(i),FF)else ChangeLibre2('TYTC_TABLELIBRETIERSA',FF);
For i:=1 to 3 do ChangeBoolLibre('YTC_BOOLLIBRE'+intToStr(i),FF);
For i:=1 to 3 do ChangeLibre2('TYTC_VALLIBRE'+intToStr(i),FF);
For i:=1 to 3 do ChangeLibre2('TYTC_DATELIBRE'+intToStr(i),FF);
end;

Procedure RTMajChampsLibresFou( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TYTC_TABLELIBREFOU'+intToStr(i),FF);
For i:=1 to 3 do ChangeBoolLibre('YTC_BOOLLIBREFOU'+intToStr(i),FF);
For i:=1 to 3 do ChangeLibre2('TYTC_VALLIBREFOU'+intToStr(i),FF);
For i:=1 to 3 do ChangeLibre2('TYTC_DATELIBREFOU'+intToStr(i),FF);
end;


Procedure RTMajChampsLibresContact( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
//For i:=1 to 3 do ChangeLibre2('TC_LIBRECONTACT'+intToStr(i),FF);
For i:=1 to 10 do if i<10 then ChangeLibre2('TC_LIBRECONTACT'+intToStr(i),FF) else ChangeLibre2('TC_LIBRECONTACTA',FF);
For i:=1 to 3 do  ChangeBoolLibre('C_BOOLLIBRE'+intToStr(i),FF);
For i:=1 to 3 do ChangeLibre2('TC_VALLIBRE'+intToStr(i),FF);
For i:=1 to 3 do ChangeLibre2('TC_DATELIBRE'+intToStr(i),FF);
end;

Procedure RTMajChampsLibresContactSSelect( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
//For i:=1 to 3 do RTChangeLibre2('TC_LIBRECONTACT'+intToStr(i),FF);
For i:=1 to 10 do if i<10 then RTChangeLibre2('TC_LIBRECONTACT'+intToStr(i),FF)else RTChangeLibre2('TC_LIBRECONTACTA',FF);
//For i:=1 to 3 do  RTChangeBoolLibre('C_BOOLLIBRE'+intToStr(i),FF);
//For i:=1 to 3 do RTChangeLibre2('TC_VALLIBRE'+intToStr(i),FF);
//For i:=1 to 3 do RTChangeLibre2('TC_DATELIBRE'+intToStr(i),FF);
end;

Procedure RTMajChampsLibresPersp( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TRPE_TABLELIBREPER'+intToStr(i),FF);
end;

Procedure RTMajChampsLibresAction( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TRAC_TABLELIBRE'+intToStr(i),FF);
end;

Procedure RTMajChampsLibresActionF( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TRAC_TABLELIBREF'+intToStr(i),FF);
end;

Procedure RTMajChampsLibresChainage( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TRCH_TABLELIBRECH'+intToStr(i),FF);
end;

Procedure RTMajChampsLibresChainageF( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 3 do ChangeLibre2('TRCH_TABLELIBRECHF'+intToStr(i),FF);
end;

Procedure RTMajChampsTypeAction( FF : Tform ) ;
var i: integer;
    CC : THLabel;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
  begin
  CC:=THLabel(TomTofGetControl(FF,'TRPA_TABLELIBRE'+intToStr(i)));
  if CC=Nil then Exit;
  if CC.FocusControl = nil then Exit;
  CC.Caption:=ChampToLibelle('RAC_TABLELIBRE'+intToStr(i));
  if (Length(CC.Caption)>0) and (copy(CC.Caption,1,2)='.-') then
     begin
     CC.visible:=false;
     TomTofSetControlVisible(FF,CC.FocusControl.Name, False);
     end;
  end;
end;

Procedure RTMajChampsTypeActionF( FF : Tform ) ;
var i: integer;
    CC : THLabel;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
  begin
  CC:=THLabel(TomTofGetControl(FF,'TRPA_TABLELIBREF'+intToStr(i)));
  if CC=Nil then Exit;
  if CC.FocusControl = nil then Exit;
  CC.Caption:=ChampToLibelle('RAC_TABLELIBREF'+intToStr(i));
  if (Length(CC.Caption)>0) and (copy(CC.Caption,1,2)='.-') then
     begin
     CC.visible:=false;
     TomTofSetControlVisible(FF,CC.FocusControl.Name, False);
     end;
  end;
end;

Procedure RTMajChampsTypeChainage( FF : Tform ) ;
var i: integer;
    CC : THLabel;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
  begin
  CC:=THLabel(TomTofGetControl(FF,'TRPG_TABLELIBRECH'+intToStr(i)));
  if CC=Nil then Exit;
  if CC.FocusControl = nil then Exit;
  CC.Caption:=ChampToLibelle('RCH_TABLELIBRECH'+intToStr(i));
  if (Length(CC.Caption)>0) and (copy(CC.Caption,1,2)='.-') then
     begin
     CC.visible:=false;
     TomTofSetControlVisible(FF,CC.FocusControl.Name, False);
     end;
  end;

end;
Procedure RTMajChampsTypeChainageF( FF : Tform ) ;
var i: integer;
    CC : THLabel;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
  begin
  CC:=THLabel(TomTofGetControl(FF,'TRPG_TABLELIBRECHF'+intToStr(i)));
  if CC=Nil then Exit;
  if CC.FocusControl = nil then Exit;
  CC.Caption:=ChampToLibelle('RCH_TABLELIBRECHF'+intToStr(i));
  if (Length(CC.Caption)>0) and (copy(CC.Caption,1,2)='.-') then
     begin
     CC.visible:=false;
     TomTofSetControlVisible(FF,CC.FocusControl.Name, False);
     end;
  end;

end;

Procedure RTMajChampsLibresSuspect( FF : Tform ) ;
var i: integer;
    AuMoinsUn,AuMoinsUnZL : boolean;
begin
if FF=Nil then  exit;
AuMoinsUn:=False;
For i:=0 to 9 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBTABLE'+intToStr(i),FF) or AuMoinsUn);
if FF is TFMul then TTabSheet(FF.FindComponent('PTABLELIBRE')).TabVisible := AuMoinsUn;
AuMoinsUnZL:=False;
For i:=0 to 2 do AuMoinsUnZL:=(ChangeBoolLibre('RSC_RSCLIBBOOL'+intToStr(i),FF) or AuMoinsUnZL);
For i:=0 to 2 do AuMoinsUnZL:=(ChangeLibre2('TRSC_RSCLIBVAL'+intToStr(i),FF) or AuMoinsUnZL);
For i:=0 to 2 do AuMoinsUnZL:=(ChangeLibre2('TRSC_RSCLIBDATE'+intToStr(i),FF) or AuMoinsUnZL);
if FF is TFMul then TTabSheet(FF.FindComponent('PZONELIBRE')).TabVisible := AuMoinsUnZL;
if FF is TFCube then TTabSheet(FF.FindComponent('PCOMPLEMENT')).TabVisible := (AuMoinsUn or AuMoinsUnZL);
if FF is TFFiche then
  begin
  AuMoinsUn:=False;
  For i:=0 to 2 do AuMoinsUn:=(ChangeLibre2('TRSC_RSCLIBTEXTE'+intToStr(i),FF) or AuMoinsUn);
  end
end;

Procedure RTMasqueChampsLibresTiers( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
    begin
    TOMTOFSetControlVisible(FF,'YTC_DATELIBRE'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'TYTC_DATELIBRE'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'YTC_DATELIBRE'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'TYTC_DATELIBRE'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'YTC_BOOLLIBRE'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'YTC_VALLIBRE'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'TYTC_VALLIBRE'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'YTC_VALLIBRE'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'TYTC_VALLIBRE'+intToStr(i)+'_',false);
    end;
end;

Procedure RTMasqueChampsLibresFou( FF : Tform ) ;
var i: integer;
begin
if FF=Nil then  exit;
For i:=1 to 3 do
    begin
    TOMTOFSetControlVisible(FF,'YTC_DATELIBREFOU'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'TYTC_DATELIBREFOU'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'YTC_DATELIBREFOU'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'TYTC_DATELIBREFOU'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'YTC_VALLIBREFOU'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'TYTC_VALLIBREFOU'+intToStr(i),false);
    TOMTOFSetControlVisible(FF,'YTC_VALLIBREFOU'+intToStr(i)+'_',false);
    TOMTOFSetControlVisible(FF,'TYTC_VALLIBREFOU'+intToStr(i)+'_',false);
    end;
end;

function RTChangeLibre2 ( NomChamp : string ; FF : Tform ) : boolean;
var CC : THLabel;
begin
result:=True; {Visible}
CC:=THLabel(TomTofGetControl(FF,NomChamp));
if CC=Nil then Exit;
if CC.FocusControl = nil then Exit;
CC.Caption:=ChampToLibelle(copy(NomChamp,2,length(NomChamp)));
if (Length(CC.Caption)>0) and (copy(CC.Caption,1,2)='.-') then
   begin
   CC.visible:=false; TomTofSetControlVisible(FF,CC.Name+'_', False);
   TomTofSetControlVisible(FF,CC.FocusControl.Name, False);  TomTofSetControlVisible(FF,CC.FocusControl.Name+'_', False);
   Result:=False;
   end;
end;

/////// Pour Script AGL ///////////////////////
procedure AGLRTMajChampsLibres( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibres(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresFou( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresFou(TForm(Longint(Parms[0]))) ;
end;

procedure AGLRTMajChampsLibresContact( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresContact(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresContactSSelect( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresContactSSelect(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresPersp( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresPersp(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresAction( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresAction(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresActionF( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresActionF(TForm(Longint(Parms[0]))) ;
end;

procedure AGLRTMajChampsLibresChainage( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresChainage(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsLibresChainageF( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresChainageF(TForm(Longint(Parms[0]))) ;
end;

procedure AGLRTMajChampsTypeAction( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsTypeAction(TForm(Longint(Parms[0]))) ;
end;
procedure AGLRTMajChampsTypeActionF( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsTypeActionF(TForm(Longint(Parms[0]))) ;
end;

procedure AGLRTMajChampsTypeChainage( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsTypeChainage(TForm(Longint(Parms[0]))) ;
end;

procedure AGLRTMajChampsTypeChainageF( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsTypeChainageF(TForm(Longint(Parms[0]))) ;
end;

Function AGLRTXXWhereConfident( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTXXWhereConfident(Parms[0]) ;
end;

Function AGLRTDroitModifTiers( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTDroitModifTiers(string(Parms[0])) ;
end;

Function AGLRTDroitModifFou( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTDroitModifFou(string(Parms[0])) ;
end;

Function AGLRTDroitModifActions( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTDroitModifActions(string(Parms[0]),string(parms[1]),string(parms[2])) ;
end;
Function AGLRTDroitModifActionsF( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTDroitModifActionsF(string(Parms[0]),string(parms[1]),string(parms[2])) ;
end;


procedure AGLRTMajChampsLibresSuspect( parms: array of variant; nb: integer ) ;
begin
  RTMajChampsLibresSuspect(TForm(Longint(Parms[0]))) ;
end;

Function AGLRTRechResponsable( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTRechResponsable(Parms[0]) ;
end;

Function  AGLRTRecupResponsable( parms: array of variant; nb: integer ): variant ;
begin
if Parms[0] = 'C' then
   result:=VH_RT.RTResponsable
else
   result:=VH_RT.RTNomResponsable;
end;


Function AGLRTFormateDateHeure( parms: array of variant; nb: integer ): variant ;
begin
  result:=RTFormateDateHeure(Parms[0],Parms[1]) ;
end;

Function AGLRTRecupTime( parms: array of variant; nb: integer ): variant ;
begin
//{$IFNDEF EAGLCLIENT}
 result:= FormatDateTime ('hh:nn',Time);
//{$ENDIF}
end;

procedure AGLRTMasqueChampsLibresTiers( parms: array of variant; nb: integer ) ;
begin
  RTMasqueChampsLibresTiers(TForm(Longint(Parms[0]))) ;
end;

procedure AGLRTMasqueChampsLibresFou( parms: array of variant; nb: integer ) ;
begin
  RTMasqueChampsLibresFou(TForm(Longint(Parms[0]))) ;
end;

function RTXXWhereConfident (typeconf:string) :string;
begin
  result:='';
  if typeconf ='CON' then result := VH_RT.RTConfWhereConsult
  else if typeconf ='MOD' then result := VH_RT.RTConfWhereModif
       else if typeconf ='CONF' then result := VH_RT.RFConfWhereConsult
            else if typeconf ='MODF' then result := VH_RT.RFConfWhereModif
  ;
end;

function RTDroitModifTiers (CodeTiers : string): Boolean;
var Sqlconf : string;
begin
Result:=true;
if (GetParamsoc('SO_RTCONFIDENTIALITE') = False) then exit;
Sqlconf := RTXXWhereConfident('MOD');
Result :=  VH_RT.DroitModifTiers ;
if VH_RT.DroitModifTiers  and (trim(Sqlconf) <> '') then result:=ExisteSQL('SELECT T_TIERS from RTTIERS WHERE T_TIERS ="'+CodeTiers+'"'+ Sqlconf );
end;

function RTDroitModifFou (CodeTiers : string): Boolean;
var Sqlconf : string;
begin
Result:=true;
if (GetParamsoc('SO_RFCONFIDENTIALITE') = False) then exit;
Sqlconf := RTXXWhereConfident('MODF');
Result :=  VH_RT.RFDroitModifTiers ;
if VH_RT.RFDroitModifTiers  and (trim(Sqlconf) <> '') then result:=ExisteSQL('SELECT T_TIERS from RFFOURNISSEURS WHERE T_TIERS ="'+CodeTiers+'"'+ Sqlconf );
end;


function RTDroitModifActions (CodeTiers,typeAction,Responsable :string): Boolean;
var    QQ :tquery;
       TobAcces : TOB;
       ListeActions,Operateur1,Operateur2 : string;
       Trouve, Autorise, Createur ,vide:boolean;
begin
  result := true;
  if (GetParamsoc('SO_RTCONFIDENTIALITE') = False) then exit;
{  if ((CodeTiers <>'') and not RTDroitModifTiers(CodeTiers)) then
  begin
    result := False; exit;
  end;    }
  TobAcces:=TOB.create ('PROSPECTCONF',NIL,-1);
  QQ := OpenSQL('Select * from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'" and RTC_TYPECONF="ACT"  and RTC_PRODUITPGI="GRC" ',True) ;
  if Not QQ.EOF then
  begin
    TobAcces.SelectDB('',QQ);
    ListeActions := TobAcces.GetValue('RTC_SQLCONF');
    Operateur1 :=  TobAcces.GetValue('RTC_OPER1');
    Operateur2 :=  TobAcces.GetValue('RTC_OPER2');
    if (ListeActions = '') then Vide := true else Vide := False;
    if (pos(typeAction,ListeActions ) > 0 ) then Trouve := true else  Trouve := False;
    If (operateur1 = '=') then Autorise:= true else  Autorise:=False;
    If  (operateur2 = '=') then Createur := true else Createur := False;
    if ((Responsable = '') or (VH_RT.RTResponsable = '')) then Createur := False;
    if ( Createur and (VH_RT.RTResponsable=Responsable) ) or (not Createur) then Createur := true else Createur :=False;
    if (( (Vide) or (Trouve and Autorise) or (not Trouve and not Autorise)) and Createur) then Result := true else Result := False ;
  end else
  begin
    if ((Responsable <> '') and  (VH_RT.RTResponsable <> '') and (VH_RT.RTResponsable=Responsable)) then
      Result := true else Result := False;;
  end;
  Ferme(QQ);
  TobAcces.free;
end;

function RTDroitModifActionsF (CodeTiers,typeAction,Responsable :string): Boolean;
var    QQ :tquery;
       TobAcces : TOB;
       ListeActions,Operateur1,Operateur2 : string;
       Trouve, Autorise, Createur ,vide:boolean;
begin
  result := true;
  if (GetParamsoc('SO_RFCONFIDENTIALITE') = False) then exit;
  TobAcces:=TOB.create ('PROSPECTCONF',NIL,-1);
  QQ := OpenSQL('Select * from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'" and RTC_TYPECONF="ACT"  and RTC_PRODUITPGI="GRF" ',True) ;
  if Not QQ.EOF then
  begin
    TobAcces.SelectDB('',QQ);
    ListeActions := TobAcces.GetValue('RTC_SQLCONF');
    Operateur1 :=  TobAcces.GetValue('RTC_OPER1');
    Operateur2 :=  TobAcces.GetValue('RTC_OPER2');
    if (ListeActions = '') then Vide := true else Vide := False;
    if (pos(typeAction,ListeActions ) > 0 ) then Trouve := true else  Trouve := False;
    If (operateur1 = '=') then Autorise:= true else  Autorise:=False;
    If  (operateur2 = '=') then Createur := true else Createur := False;
    if ((Responsable = '') or (VH_RT.RTResponsable = '')) then Createur := False;
    if ( Createur and (VH_RT.RTResponsable=Responsable) ) or (not Createur) then Createur := true else Createur :=False;
    if (( (Vide) or (Trouve and Autorise) or (not Trouve and not Autorise)) and Createur) then Result := true else Result := False ;
  end else
  begin
    if ((Responsable <> '') and  (VH_RT.RTResponsable <> '') and (VH_RT.RTResponsable=Responsable)) then
      Result := true else Result := False;;
  end;
  Ferme(QQ);
  TobAcces.free;
end;

function RTControleModifTiers (FF : Tform ;TobProspect: TOB; ActionModif :Boolean): Boolean;
var TobModAcces,TobTiers,TobZonelibre : TOB;
    QQ :TQUERY;
    InfoAcces : string;
    ind :integer;

    Function CompareInfosFiche(TobAcces :TOB):Boolean;
    var Resultat : Array[0..2] of boolean;
        ValeurConf,ValeurConf2,LibelleChamp,LibelleValeur,LibelleValeur2,Operateur,LibOperateur,Lien,Prefixe : Array[0..2] of String;
        NomChamp,ValeurFiche,ValeurConf0 :string;
        ListeChamp: TStrings;
        ResulCompare,i,ResulCompare1,ResulCompare2 :integer;

        Function ResulConf (Resul1, Resul2:Boolean;lien :string):boolean;
        begin
          if (lien = 'Ou') then result := (Resul1 or Resul2)
          else  result := (Resul1 and Resul2);
        end;
    begin
      ResulCompare1:=0; ResulCompare2:=0;
      Result := false;
      if (TobAcces.GetValue('RTC_SQLCONF')='') then exit;
      ListeChamp := TStringList.Create;
      for i:=1 to 3 do
      begin
        NomChamp := TobAcces.getvalue('RTC_CHAMP'+intToStr(i));
        Prefixe [i-1] := Copy(NomChamp,1,Pos('_',NomChamp)-1);
        if (NomChamp <> '') then
        begin
          ListeChamp.add (NomChamp);
          LibelleChamp [i-1] := ChampToLibelle (NomChamp);
          Operateur [i-1] := TobAcces.getvalue('RTC_OPER'+intToStr(i));
          LibOperateur [i-1] := RechDom ('TTCOMPARE', Operateur[i-1],False);
          if (Operateur [i-1]='E') or (Operateur [i-1]='G') then
          begin
            ValeurConf0 := TobAcces.getvalue('RTC_VAL'+intToStr(i));
            ValeurConf    [i-1] := ReadTokenSt (ValeurConf0);
            ValeurConf2   [i-1] := ReadTokenSt (ValeurConf0);
            LibelleValeur [i-1] := RTValeurChampLibre(NomChamp,ValeurConf [i-1]);
            LibelleValeur2[i-1] := RTValeurChampLibre(NomChamp,ValeurConf2 [i-1]);
          end else
          begin
            ValeurConf [i-1] := TobAcces.getvalue('RTC_VAL'+intToStr(i));
            if (Operateur[i-1]='=') or (Operateur[i-1]='<>') or (Operateur[i-1]='>') or (Operateur[i-1]='>=') or (Operateur [i-1]='<') or (Operateur [i-1]='<=') then
              LibelleValeur[i-1]:= RTValeurChampLibre(NomChamp,ValeurConf [i-1])
            else  LibelleValeur[i-1]:= ValeurConf [i-1];
            ValeurConf2 [i-1] := '';   LibelleValeur2 [i-1]:= '';
          end;
        end else
        begin
          ValeurConf [i-1] := '';    Operateur [i-1] := '';
          LibelleChamp [i-1] := '';  LibelleValeur [i-1]:= '';
          ValeurConf2 [i-1] := '';   LibelleValeur2 [i-1]:= '';
        end;
      end;
      if ListeChamp.Count = 0 then Resultat[0] := true;

      Lien[0] := TobAcces.getvalue('RTC_LIEN1');
      Lien[1] := TobAcces.getvalue('RTC_LIEN2');
      Lien[2] := '';

      for i := 0 to ListeChamp.Count - 1 do
      begin
          ValeurFiche := '';
          NomChamp := ListeChamp [i];
          if ( (Operateur[i] = 'E' ) or (Operateur[i] = 'G' ) ) then
          InfoAcces := InfoAcces +'('+LibelleChamp[i]+' '+LibOperateur[i]+' '+LibelleValeur[i]+' et '+LibelleValeur2[i]+') '+Lien[i]+' '
          else
          InfoAcces := InfoAcces + '(' + LibelleChamp[i] + ' ' + LibOperateur[i]+ ' '  + LibelleValeur[i] +') '+Lien[i]+' ';

          if (not ActionModif) and (((FF.name = 'GCTIERS') and (Prefixe [i] = 'RPR'))
          or ((FF.name = 'RTPARAMCL') and (Prefixe [i] <> 'RPR'))) then
          begin  //En création controle séparé des Infos Fiche tiers et des Infos complémentaires
            Resultat[i] := true;
            Continue;
          end;
          if (Prefixe [i] = 'T')   then ValeurFiche := TobTiers.Getvalue(NomChamp)
          else if (Prefixe [i] = 'YTC') then  ValeurFiche := TobZoneLibre.Getvalue(NomChamp)
          else if (Prefixe [i] = 'RPR') then  ValeurFiche := TobProspect.Getvalue(NomChamp);
          ResulCompare := AnsiCompareText(ValeurFiche,ValeurConf[i]);
          if ( (Operateur[i] = 'E' ) or (Operateur[i] = 'G' ) ) then
          begin
             ResulCompare1 := AnsiCompareText(ValeurFiche,ValeurConf[i]);
             ResulCompare2 := AnsiCompareText(ValeurFiche,ValeurConf2[i]);
          end;
          if ( (Operateur[i] = '=' ) and (ResulCompare  = 0 ) )
          or ( (Operateur[i] = '<>') and (ResulCompare <> 0 ) )
          or ( (Operateur[i] = '>' ) and (ResulCompare  > 0 ) )
          or ( (Operateur[i] = '>=') and (ResulCompare >= 0 ) )
          or ( (Operateur[i] = '<' ) and (ResulCompare  < 0 ) )
          or ( (Operateur[i] = '<=') and (ResulCompare <= 0 ) )
          or ( (Operateur[i] = 'C' ) and (pos (ValeurConf[i],ValeurFiche )  =  1  ) ) // commence par
          or ( (Operateur[i] = 'D' ) and (pos (ValeurConf[i],ValeurFiche ) <>  1  ) )// ne commence pas par
          or ( (Operateur[i] = 'L' ) and (pos (ValeurConf[i],ValeurFiche )  >  0  ) ) // contient
          or ( (Operateur[i] = 'M' ) and (pos (ValeurConf[i],ValeurFiche )  =  0  ) ) // ne contient pas
          or ( (Operateur[i] = 'I' ) and (pos (ValeurFiche,ValeurConf[i] )  >  0  ) ) // Est dans
          or ( (Operateur[i] = 'J' ) and (pos (ValeurFiche,ValeurConf[i] )  =  0  ) ) // N'est pas dans
          or ( (Operateur[i] = 'E' ) and (ResulCompare1 >= 0 ) and (ResulCompare2 <= 0 ) ) // Est entre
          or ( (Operateur[i] = 'G' ) and ((ResulCompare1 < 0 ) or (ResulCompare2 > 0 ))  ) // N'est pas entre
          then Resultat[i] := true;
      end;
      Result := Resultat[0];
      for i := 1 to ListeChamp.Count - 1 do
      begin
         Result := ResulConf (Result,Resultat[i],Lien[i-1]);
      end;
      ListeChamp.free;
    end;

begin
  Result := true;  InfoAcces := '';
  TobModAcces:=Nil;
  if (GetParamsoc('SO_RTCONFIDENTIALITE') = False) or ((FF.name <> 'GCTIERS') and (FF.name <> 'RTPARAMCL'))
  or (VH_RT.RTConfWhereModif = '') then exit;
  Result := true;
  try
  TobTiers:=TOB.Create ('TIERS', Nil, -1);
  TobZonelibre:=TOB.Create ('TIERSCOMPL', Nil, -1);
  if (FF.name = 'GCTIERS') then
  begin
    TobTiers.GetEcran (FF,Nil);
    TobZoneLibre.GetEcran (FF,Nil);
    Tobprospect :=TOB.Create ('PROSPECTS', Nil, -1);
    Tobprospect.SelectDB('"'+TOBTiers.GetValue('T_AUXILIAIRE')+'"',Nil );
  end
  else if (FF.name = 'RTPARAMCL') then
  begin
    TobTiers.SelectDB('"'+Tobprospect.GetValue('RPR_AUXILIAIRE')+'"',Nil );
    TobZoneLibre.SelectDB('"'+Tobprospect.GetValue('RPR_AUXILIAIRE')+'"',Nil );
  end;
  TobModAcces:=TOB.create ('GRC Confidentialité',NIL,-1);
  QQ := OpenSQL('Select * from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'" and RTC_TYPECONF LIKE "MO%" and RTC_PRODUITPGI="GRC" ORDER BY RTC_TYPECONF DESC',True) ;
  if Not QQ.EOF then
  begin
    TobModAcces.LoadDetailDB('PROSPECTCONF','','',QQ, False) ;
    Result := False;
    for ind:=0 to TobModAcces.detail.count-1 do
    begin
     if (TobModAcces.detail[ind].GetValue('RTC_SQLCONF')<>'') then
       InfoAcces := InfoAcces + '#13 Critère '+IntToStr(ind+1)+' : ';
     Result := Result or CompareInfosFiche (TobModAcces.detail[ind]);
    end;
    if not Result then
    begin
     if ActionModif then  PGIBox(InfoAcces,'Accès en modification non autorisé')
     else PGIBox(InfoAcces,'Accès en création non autorisé');
    end;
  end;
  Ferme(QQ);

  Finally
    TobModAcces.Free;
    TobTiers.Free;
    TobZoneLibre.Free;
    if (FF.name = 'GCTIERS') then Tobprospect.Free;
  end;
end;

function RTValeurChampLibre(NomChamp,ValeurChamp: string):string;
var datatype :string;
begin
   result := ValeurChamp;
   if copy(NomChamp,1,19)='YTC_TABLELIBRETIERS' then    //YTC_TABLELIBRETIERS0 -> GCLIBRETIERS0
      datatype := FindEtReplace(NomChamp,'YTC_TABLE','GC',false)
   else if copy(NomChamp,1,15)='RPR_RPRLIBTABLE' then   //RPR_RPRLIBTABLE0 -> RT RPRLIBTABLE0
      datatype := FindEtReplace(NomChamp,'RPR_','RT',false)
   else
      datatype := Get_Join(NomChamp );
   if datatype <> '' then
     Result :=RechDom(datatype,ValeurChamp,false)
end;

{***********A.G.L.***********************************************
Auteur  ...... : Garnier Marie-Noëlle
Créé le ...... : 21/09/2001
Modifié le ... :   /  /
Description .. : Suppression d'un ligne dans un PopUpMenu
Mots clefs ... : POPUPMENU
*****************************************************************}
procedure AGLRTMajPopMenu( parms: array of variant; nb: integer ) ;
begin
  RTMajPopMenu(TForm(Longint(Parms[0])), String(Parms[1]), Integer(Parms[2]), Boolean(Parms[3])) ;
end;

Procedure RTMajPopMenu( FF : Tform; MenuPop : String; NoLigne : Integer; AConserver : boolean ) ;
var pop : TPopupMenu ;
begin
    pop :=TPopupMenu(TomTofGetControl(FF,MenuPop) ) ;
    pop.items[NoLigne].visible:=AConserver;
end;

Procedure CommercialToRessource ;
var T,TR,TC : TOB;
    Q : TQuery;
    i : integer;
begin
T:=Tob.create('Les ressources',Nil,-1);
TC:=Tob.create('Les commerciaux',Nil,-1);
Q := OpenSQL( 'SELECT * FROM COMMERCIAL', True );
if Not Q.EOF then TC.loadDetailDB('COMMERCIAL','','',Q,False );
Ferme(Q);
if TC.detail.count>0 then
   begin
   For i:=0 to TC.detail.count-1 do
       begin
       TR:=Tob.create('RESSOURCE',T,-1);
       TR.initValeurs (False);
       TR.putValue('ARS_TYPERESSOURCE','SAL');
       TR.putValue('ARS_RESSOURCE',TC.detail[i].getvalue('GCL_COMMERCIAL'));
       TR.putValue('ARS_LIBELLE',TC.detail[i].getvalue('GCL_LIBELLE'));
       end;
   T.InsertDBByNivel(True);
   end;
T.free; TC.Free;
end;

Procedure RTCreerProspectPourClient ;
var TOBProspects: TOB;
    Q : TQuery;
begin
if PGIAsk('Voulez-vous initialiser les enregistrements complémentaires pour les tiers clients ?','')<>mrYes then exit ;
TOBProspects:=Tob.create('PROSPECTS',Nil,-1);
Q:=OPENSQL('select t_auxiliaire from tiers where (t_natureauxi="CLI" or t_natureauxi="PRO") and NOT EXISTS (select rpr_auxiliaire from prospects where rpr_auxiliaire=tiers.t_auxiliaire)',true);
InitMove(Q.recordcount, '');
while not Q.EOF do
  BEGIN
  TOBProspects.InitValeurs(False);
  TOBProspects.PutValue('RPR_AUXILIAIRE',Q.Fields[0].AsString);
  if not TOBProspects.InsertOrUpdateDB(False) then V_PGI.IoError:=oeUnknown ;
  Q.Next ;
  MoveCur(false);  
  END ;
  
FiniMove();
Ferme(Q);
TOBProspects.free;
end;

Function AGLRTExisteConfident ( parms: array of variant; nb: integer ): variant ;
begin
  result := 'MOD';
  if (GetParamsoc('SO_RTCONFIDENTIALITE') = False) then exit;
  //QQ := OpenSQL('Select * from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'" and RTC_TYPECONF="MOD"  ',True) ;
  //if QQ.EOF then
  { mng 03-12-2003 }
  if VH_RT.RTExisteConfident = false then
     result := '';
  //Ferme(QQ);
end;

Function AGLRTEagl ( parms: array of variant; nb: integer ): Variant ;
begin
result:=false;
{$IFDEF EAGLCLIENT}
result:=true;
{$ENDIF}
end;

Procedure RTCorrespondSuspectProspect(CleChoixCode : string; FromTOB, ToTOB : TOB);
var TobCorrespond: TOB;
    PrefixeTo,ChampTo,FieldNameFrom,FieldNameTo :string;
    ind : integer;
begin
PrefixeTo := TableToPrefixe (ToTOB.NomTable);
if (PrefixeTo='YTC') then ChampTo := 'RSP_CHTIERS'
else if (PrefixeTo='RPR') then ChampTo := 'RSP_CHCOMPL';
TobCorrespond:=tob.create('Correspondance',Nil,-1) ;
TobCorrespond.LoadDetailDB('PARSUSPECTCOR','','',nil,false,true);
//RSP_CHSUSPECT:champ libre suspect,RSP_CHTIERS:champ libre tiers,RSP_CHCOMPL:champ libre prospect
if TobCorrespond.detail.count>0 then
begin  
  For ind:=0 to TobCorrespond.detail.count-1 do
    begin
      FieldNameFrom := TobCorrespond.detail[ind].getvalue('RSP_CHSUSPECT');
      FieldNameTo := TobCorrespond.detail[ind].getvalue(ChampTo);
      if (trim (FieldNameFrom)<>'') and (copy (FieldNameTo ,1,3) = PrefixeTo) then
        ToTOB.PutValue (FieldNameTo, FromTOB.GetValue (FieldNameFrom));
    end;
end;
TobCorrespond.cleardetail;
TobCorrespond.free;
end;

Function RTRechResponsable (Tiers: string) : string;
var Q : TQuery;
    Select,Utilisateur,Commercial : string;
begin
result:='';
Select := 'SELECT T_REPRESENTANT FROM TIERS WHERE T_TIERS = "'+ tiers +'"';
Q:=OpenSQL(Select, True);
if not Q.Eof then
   Commercial := Q.Fields[0].AsString;
Ferme(Q) ;
if commercial <> '' then
    begin
    Select := 'SELECT GCL_UTILASSOCIE FROM COMMERCIAL WHERE GCL_COMMERCIAL = "'+ commercial+'"';
    Q:=OpenSQL(Select, True);
    if not Q.Eof then
       Utilisateur := Q.Fields[0].AsString;
    Ferme(Q) ;
    if Utilisateur <> '' then
        begin
        Select := 'SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_UTILASSOCIE = "'+ Utilisateur+'"';
        Q:=OpenSQL(Select, True);
        if not Q.Eof then
            Result := Q.Fields[0].AsString;
        Ferme(Q) ;
        end;
    end;
end;
Function RTFormateDateHeure (Date: TDateTime; Heure : String) : String;
var hh : TDateTime;
begin
//hh:=strToDateTime(Heure);
hh:=Date+strToDateTime(Heure);
//{$IFNDEF EAGLCLIENT}
result:=USDATETIME(hh);
//{$ENDIF}
end;

function RTAttribNewCodeSuspect (DernierChrono : string):string;
var ChronoCode :string;
    i_Chrono,i_Reste: Int64;
begin
  ChronoCode := DernierChrono ;
  if not isnumeric(ChronoCode) then ChronoCode := '0';
  repeat
    i_Chrono := StrToInt64(ChronoCode);
    i_Chrono := i_Chrono+1;
    i_Reste:=i_Chrono Mod 10 ;
    if i_Reste = 0 then i_Chrono := i_Chrono+1;
    ChronoCode := IntToStr(i_Chrono);
    if (Length(ChronoCode) > 17) then ChronoCode := '1';
  until (not ExisteSQL('SELECT RSU_SUSPECT FROM SUSPECTS WHERE RSU_SUSPECT = "'+ChronoCode+'"'));
  Result := ChronoCode;
end;
// déplacement des Concurrents des propositions dans un fichier liens
Procedure RTCreerLiensConcurrents  ;
var Q: TQuery;
    StListe,StConc,Sql : String;
    i : integer;
begin
{if GetParamsoc('SO_RTDEPLACECON') = true then
  PGIBox('Le déplacement des concurrents a déjà été effectué','Déplacement des concurrents')
else
   begin}
   if PGIAsk('Voulez-vous déplacer les concurrents ?', 'Déplacement des concurrents') <> mrYes then exit
   else
      begin

Q := OpenSQL ('SELECT RPE_CONCURRENTS,rpe_perspective,rpe_tiers FROM PERSPECTIVES where rpe_concurrents <> ""',True) ;
while Not Q.EOF do
   begin
   StListe := Q.FindField('RPE_CONCURRENTS').AsString;
   if StListe <> '' then
      Repeat
      StConc:=Trim(ReadTokenSt(StListe)) ;
      // recherche du code à traiter dans les codes restants pour éliminer les doublons
      if StListe <> '' then
      begin
        i:=1;
        While i<>0 do
        begin
          i:=Pos(StConc, StListe);
          if i <> 0 then
          begin
            Delete(StListe,i,length(StConc));
            if Copy(StListe,i,1)=';' then
               Delete(StListe,i,1);                
          end;
        end;
      end;
      // fin recherche du code à traiter dans les codes restants pour éliminer les doublons
      if (StConc<>'') and (length(StConc) < 18) then
          begin
              Sql :='INSERT INTO PERSPECTIVESTIERS (RPT_TIERS,RPT_PERSPECTIVE) VALUES ("';
              Sql := Sql+StConc+'",'+IntToStr(Q.FindField('rpe_perspective').AsInteger)+')';
              ExecuteSQL(Sql);
          end;
      until  StConc='' ;
   Q.Next;
   end;
Ferme(Q) ;

{      ExecuteSql ('Update perspectives set rpe_concurrents="" where rpe_concurrents <> ""');
      SetParamSoc ('SO_RTDEPLACECON',True)
      end;}
   end;
end;

Function  AGLRTAffichePlusMesActions( parms: array of variant; nb: integer ): variant ;
begin
  // parms 1 = 'X' : lire, = '-' : Ecrire
  if parms[1] = 'X' then
  begin
     if GetSynRegKey('RTMesActions','X',TRUE)='X' then
        Result:=False
     else
        Result:=True;
  end
  else
     SaveSynRegKey('RTMesActions',parms[0],TRUE);
end;

Function  AGLRTAfficheParamRappel ( parms: array of variant; nb: integer ): variant ;
var Duree,Critere : String;
begin
  // parms 1 = 'X' : lire, = '-' : Ecrire
  if parms[1] = 'X' then
  begin
    if GetSynRegKey('RTRappelAuto','X',TRUE)='X' then
      Result:='X;'
    else
      Result:='-;';

    Duree:=GetSynRegKey('RTRappelMinutes','',true);
    if Duree='' then
      Result:=Result+IntToStr(GetParamSoc('SO_RTACTDUREE'))
    else
      Result:=Result+Duree;
  end
else
  begin
  Critere:=parms[0];
  SaveSynRegKey('RTRappelAuto',ReadToKenSt(Critere),TRUE);
  SaveSynRegKey('RTRappelMinutes',ReadToKenSt(Critere),TRUE);
  end;
end;

function RTDroitModifTypeAction : Boolean;
begin
  result := true;
  //if (GetParamsoc('SO_RTCONFIDENTIALITE') = False) then exit;
  result:=ExisteSQL('Select RTC_VAL1 from PROSPECTCONF Where RTC_INTERVENANT="'+V_PGI.User+'" and RTC_TYPECONF="ACT"  and RTC_PRODUITPGI="GRC" and RTC_VAL1="="');
end;

Procedure RTAlimCleTelephonTiers ;
var Notel,Cletel : string;
    Q : TQuery;
begin
   if PGIAsk('Voulez-vous mettre à jour la clé téléphone ?', 'MAJ Clé téléphone') <> mrYes then exit
   else
      begin
      Q := OpenSQL ('SELECT t_tiers,t_telephone,t_cletelephone FROM tiers where (t_natureauxi="CLI" or t_natureauxi="PRO") and t_telephone<>""',True) ;
      while Not Q.EOF do
         begin
         Cletel:='';
         Notel := Q.FindField('t_telephone').AsString;
         Cletel:=RTFormateCleTelephone(Notel);
         if Cletel <> '' then
            ExecuteSQL('UPDATE tiers SET t_cletelephone="'+Cletel+'" where t_tiers="'+Q.FindField('t_tiers').asString+'"');
         Q.Next;
         end;
      Ferme(Q) ;
   end;
end;

Procedure RTAlimTelephonContact ;
var Notel,Cletel : string;
    Q : TQuery;
begin
if PGIAsk('Voulez-vous formater les no téléphones contacts ?', 'MAJ no téléphone') <> mrYes then exit
else
  begin
  Q := OpenSQL ('SELECT C_AUXILIAIRE,C_NUMEROCONTACT,c_telephone FROM contact where (c_natureauxi="CLI" or c_natureauxi="PRO") and c_telephone<>"" and c_typecontact="T"',True) ;
  while Not Q.EOF do
     begin
     Cletel:='';
     Notel := Q.FindField('c_telephone').AsString;
     Cletel:=RTFormatTelephonContact(Notel);
     if Cletel <> '' then
        ExecuteSQL('UPDATE contact SET c_telephone="'+Cletel+'" where C_TYPECONTACT="T" and C_NUMEROCONTACT='+
           Q.FindField('C_NUMEROCONTACT').asString+' and C_AUXILIAIRE="'+Q.FindField('C_AUXILIAIRE').asString+'"');
     Q.Next;
     end;
  Ferme(Q) ;
  end;
end;

Function RTFormateCleTelephone(Notel:String) : String;
var Cletel,car : String;
    i : integer;
begin
Cletel:='';
for i:=1 to length(Notel) do
   begin
   car:=copy(Notel,i,1);
   if (IsNumeric(car)) and (car <>'.') and (car<>',') and (car<>'-') and (car<>' ') then
       Cletel:=Cletel+car;
   end;
Result:=Cletel;
end;

Function RTFormatTelephonContact(Notel:String) : String;
var Cletel,car : String;
    i : integer;
begin
for i:=1 to length(Notel) do
   begin
   car:=copy(Notel,i,1);
   if (IsNumeric(car)) and (car <>'.') and (car<>',') and (car<>'-') and (car<>' ') then
       Cletel:=Cletel+car;
   end;
if copy(Notel,1,1)='+' then Cletel:='+'+Cletel;
Result:=Cletel;
end;

Procedure RTMajNatureContacts(StAuxi,StNat : String);
begin
ExecuteSQL ('Update contact set c_natureauxi="'+StNat+'" where c_typecontact="T" and c_auxiliaire="'+StAuxi+'"');
end;
Procedure RTMajNatureAdresses(StAuxi,StNat : String);
begin
ExecuteSQL ('Update ADRESSES set adr_natureauxi="'+StNat+'" where ADR_REFCODE="'+StAuxi+'"');
end;
Procedure RTMajNatureFrais(StAuxi,StNat : String);
begin
ExecuteSQL ('Update TIERSFRAIS set gtf_natureauxi="'+StNat+'" where gtf_tiers="'+StAuxi+'"');
end;

Function RTAccesMenu(NoMenu,NoTag,NoGrp:integer) : Boolean;
var Q : TQuery;
begin
  result:=true;
  Q:=OpenSql('SELECT MN_ACCESGRP FROM MENU WHERE MN_1='+IntToStr(NoMenu)+' AND MN_TAG='+IntToStr(NoTag),true);
  if not Q.Eof then
    if Copy(Q.FindField('MN_ACCESGRP').AsString,NoGrp,1) = '-' then result:=false;
  ferme(Q);
end;

Function AGLRTAccesMenu( parms: array of variant; nb: integer ) : variant;
begin
  result:=RTAccesMenu(Integer(Parms[0]),Integer(parms[1]),Integer(parms[2])) ;
end;
{***********A.G.L.***********************************************
Auteur  ...... : Garnier Marie-Noëlle
Créé le ...... : 12/09/2001
Modifié le ... : 17/09/2001
Description .. : Procédure de transformation d'un Tiers de Nature Prospect
Suite ........ : à la Nature Client
Mots clefs ... : NATURE TIERS
*****************************************************************}
procedure MajProspectClient ( TOBGenere : Tob ; TOBTiers : Tob);
begin
if GetInfoParPiece(TOBGenere.GetValue('GP_NATUREPIECEG'),'GPP_PROCLI') = 'X' then
    begin
    if (TOBTiers.GetValue('T_NATUREAUXI')='PRO') then
       begin
       // transformer le tiers de PRO en CLI
       TOBTiers.putValue('T_NATUREAUXI','CLI');
       TOBTiers.putValue('T_DATEPROCLI',V_PGI.DateEntree);
       end;

    TransformeProspectClient(TOBGenere,'GP_TIERSLIVRE');
    TransformeProspectClient(TOBGenere,'GP_TIERSFACTURE');
    TransformeProspectClient(TOBGenere,'GP_TIERSPAYEUR');
    RTMajNatureContacts(TOBTiers.GetValue('T_AUXILIAIRE'),'CLI');
    RTMajNatureAdresses(TOBTiers.GetValue('T_TIERS'),'CLI');
    RTMajNatureFrais(TOBTiers.GetValue('T_TIERS'),'CLI');
    end;

end;

procedure TransformeProspectClient ( TOBGenere : Tob ; NomChamp : String);
var Requete : String;
begin
// transformer le tiers de PRO en CLI
if ( TOBGenere.GetValue(NomChamp) <> '' ) and
   ( TOBGenere.GetValue('GP_TIERS')<>TOBGenere.GetValue(NomChamp) ) then
   begin
   Requete:='UPDATE TIERS SET T_NATUREAUXI="CLI",T_DATEPROCLI="'+USDATETIME (V_PGI.DateEntree)+
     '" WHERE T_TIERS="'+TOBGenere.GetValue(NomChamp)+'" and T_NATUREAUXI="PRO"';
   ExecuteSQL (Requete);
   // mng 27-12-2002 : maj nature contacts
   RTMajNatureContacts(TiersAuxiliaire(TOBGenere.GetValue(NomChamp),false),'CLI');
   RTMajNatureAdresses(TOBGenere.GetValue(NomChamp),'CLI');
   RTMajNatureFrais(TOBGenere.GetValue(NomChamp),'CLI');   
   end;
end;

Initialization
RegisterAglFunc( 'RTeAGL', FALSE , 0, AGLRTEagl);
RegisterAglProc( 'RTMajChampsLibres', TRUE , 0, AGLRTMajChampsLibres);
RegisterAglProc( 'RTMajChampsLibresFou', TRUE , 0, AGLRTMajChampsLibresFou);
RegisterAglProc( 'RTMajChampsLibresContact', TRUE , 0, AGLRTMajChampsLibresContact);
RegisterAglProc( 'RTMajChampsLibresPersp', TRUE , 0, AGLRTMajChampsLibresPersp);
RegisterAglProc( 'RTMajChampsLibresAction', TRUE , 0, AGLRTMajChampsLibresAction);
RegisterAglProc( 'RTMajChampsLibresActionF', TRUE , 0, AGLRTMajChampsLibresActionF);
RegisterAglProc( 'RTMajChampsLibresChainage', TRUE , 0, AGLRTMajChampsLibresChainage);
RegisterAglProc( 'RTMajChampsLibresChainageF', TRUE , 0, AGLRTMajChampsLibresChainageF);
RegisterAglProc( 'RTMajChampsTypeAction', TRUE , 0, AGLRTMajChampsTypeAction);
RegisterAglProc( 'RTMajChampsTypeActionF', TRUE , 0, AGLRTMajChampsTypeActionF);
RegisterAglProc( 'RTMajChampsTypeChainage', TRUE , 0, AGLRTMajChampsTypeChainage);
RegisterAglProc( 'RTMajChampsTypeChainageF', TRUE , 0, AGLRTMajChampsTypeChainageF);
RegisterAglProc( 'RTMajPopMenu', TRUE , 2, AGLRTMajPopMenu);
RegisterAglFunc( 'RTXXWhereConfident', FALSE , 1, AGLRTXXWhereConfident);
RegisterAglFunc( 'RTDroitModifTiers', FALSE , 1, AGLRTDroitModifTiers);
RegisterAglFunc( 'RTDroitModifFou', FALSE , 1, AGLRTDroitModifFou);
RegisterAglFunc( 'RTDroitModifActions', FALSE , 3, AGLRTDroitModifActions);
RegisterAglFunc( 'RTDroitModifActionsF', FALSE , 3, AGLRTDroitModifActionsF);
RegisterAglFunc( 'RTExisteConfident', FALSE , 0, AGLRTExisteConfident);
RegisterAglProc( 'RTMajChampsLibresSuspect', TRUE , 0, AGLRTMajChampsLibresSuspect);
RegisterAglFunc( 'RTRechResponsable', FALSE , 1, AGLRTRechResponsable);
RegisterAglProc( 'RTMasqueChampsLibresTiers', TRUE , 0, AGLRTMasqueChampsLibresTiers);
RegisterAglProc( 'RTMasqueChampsLibresFou', TRUE , 0, AGLRTMasqueChampsLibresFou);
RegisterAglProc( 'RTMajChampsLibresContactSSelect', TRUE , 0, AGLRTMajChampsLibresContactSSelect);
RegisterAglFunc( 'RTFormateDateHeure', FALSE , 2, AGLRTFormateDateHeure);
RegisterAglFunc( 'RTRecupTime', FALSE , 2, AGLRTRecupTime);
RegisterAglFunc( 'RTRecupResponsable', FALSE , 1, AGLRTRecupResponsable);
RegisterAglFunc( 'RTAffichePlusMesActions', FALSE , 2, AGLRTAffichePlusMesActions);
RegisterAglFunc( 'RTAccesMenu', false , 3, AGLRTAccesMenu);
RegisterAglFunc( 'RTAfficheParamRappel', FALSE , 2, AGLRTAfficheParamRappel);
end.



