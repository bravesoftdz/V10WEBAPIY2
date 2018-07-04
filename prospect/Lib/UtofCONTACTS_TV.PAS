{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 08/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CONTACTS_TV ()
Mots clefs ... : TOF;CONTACTS_TV
*****************************************************************}
Unit UtofCONTACTS_TV ;

Interface

Uses StdCtrls, 
     Classes,
     UTobView,  UtilGC,
{$IFNDEF EAGLCLIENT}
     Fe_Main,
{$else}
     MainEAGL,
{$ENDIF}
{$ifdef AFFAIRE}
      UtofAfTraducChampLibre,
{$ENDIF}
     forms,
     sysutils, 
     HEnt1,
     UTOF,
     UtilRT,
{$IFDEF GIGI}
      EntGc,
{$ENDIF}
     UtilSelection,ParamSoc ;

Function RTLanceFiche_CONTACTS_TV(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Type
{$ifdef AFFAIRE}
                //mcd 11/05/2006 12940  pour faire affectation depuis ressource si paramétré
  TOF_CONTACTS_TV = Class (TOF_AFTRADUCCHAMPLIBRE)
 {$else}
  TOF_CONTACTS_TV = Class (TOF)
{$endif}
     private
         TobViewer1: TTobViewer;
         procedure ContTVOnDblClickCell(Sender: TObject ) ;
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

Function RTLanceFiche_CONTACTS_TV(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_CONTACTS_TV.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CONTACTS_TV.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CONTACTS_TV.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CONTACTS_TV.OnLoad ;
begin
Inherited ;
SetControlText('XX_WHERE',RTXXWhereConfident('CON')) ;
end ;

procedure TOF_CONTACTS_TV.OnArgument (S : String ) ;
var F : TForm;
    Memo : TMemo;
begin
Inherited ;
F := TForm (Ecran);
MulCreerPagesCL(F,'NOMFIC=GCTIERS');
if (GetParamSocSecur('SO_RTGESTINFOS006',False) = True) then
  MulCreerPagesCL(F,'NOMFIC=YYCONTACT');
TobViewer1:=TTobViewer(getcontrol('TV'));
TobViewer1.OnDblClick:= ContTVOnDblClickCell ;
memo := TMemo(GetControl('FSQL'));
if (GetControl('YTC_RESSOURCE1') <> nil)  then
  begin
 if not (ctxaffaire in V_PGI.PGICONTEXTE) then SetControlVisible ('PRESSOURCE',false)
  else begin
    GCMAJChampLibre (TForm (Ecran), False, 'EDIT', 'YTC_RESSOURCE', 3, '_');
    Memo.lines[6]:='YTC_RESSOURCE1,YTC_RESSOURCE2,YTC_RESSOURCE3,YTC_TABLELIBRETIERS1,YTC_TABLELIBRETIERS2,';
    if not (ctxscot in V_PGI.PGICOntexte) then
       begin
       SetControlVisible ('T_MOISCLOTURE',false);
       SetControlVisible ('T_MOISCLOTURE_',false);
       SetControlVisible ('TT_MOISCLOTURE',false);
       SetControlVisible ('TT_MOISCLOTURE_',false);
       end;
    end;
  end;
{$Ifdef GIGI}
 if (GetControl('T_REPRESENTANT') <> nil) then  SetControlVisible('T_REPRESENTANT',false);
 if (GetControl('TT_REPRESENTANT') <> nil) then  SetControlVisible('TT_REPRESENTANT',false);
 if (GetControl('YTC_REPRESENTANT2') <> nil) then  SetControlVisible('YTC_REPRESENTANT2',false);
 if (GetControl('YTC_REPRESENTANT2_') <> nil) then  SetControlVisible('YTC_REPRESENTANT2_',false);
 if (GetControl('YTC_REPRESENTANT3') <> nil) then  SetControlVisible('YTC_REPRESENTANT3',false);
 if (GetControl('YTC_REPRESENTANT3_') <> nil) then  SetControlVisible('YTC_REPRESENTANT3_',false);
 SetControlText('T_NatureAuxi','');    //on efface les valeurs CLI et PO, car NCP en plus
 SetControlProperty ('T_NATUREAUXI', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI', 'Plus', VH_GC.AfNatTiersGRCGI);
 SetControlProperty ('T_NATUREAUXI_', 'Complete', true);
 SetControlProperty ('T_NATUREAUXI_', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('T_NATUREAUXI_', 'Plus', VH_GC.AfNatTiersGRCGI);
 Memo.lines[3]:='T_SECTEUR,T_SOCIETEGROUPE,T_PRESCRIPTEUR,';
 Memo.lines[5]:='T_MOISCLOTURE,';
{$endif}
end ;

procedure TOF_CONTACTS_TV.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CONTACTS_TV.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CONTACTS_TV.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CONTACTS_TV.ContTVOnDblClickCell(Sender: TObject );
begin
with TTobViewer(sender) do
  begin
  if copy(ColName[CurrentCol],1,2) = 'C_' then
     V_PGI.DispatchTT (16,taConsult ,'T;'+AsString[ColIndex('T_AUXILIAIRE'), CurrentRow]+';'+IntToStr(AsInteger[ColIndex('C_NUMEROCONTACT'), CurrentRow]), '','')
  else
     V_PGI.DispatchTT (28,taConsult ,AsString[ColIndex('T_AUXILIAIRE'), CurrentRow], '','')
  end;
end;

Initialization
  registerclasses ( [ TOF_CONTACTS_TV ] ) ; 
end.
