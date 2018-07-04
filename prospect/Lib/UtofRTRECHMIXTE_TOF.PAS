{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/07/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RTRECHMIXTE ()
Mots clefs ... : TOF;RTRECHMIXTE
*****************************************************************}
Unit UtofRTRECHMIXTE_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,
{$else}
     eMul,
     MainEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     EntGC,
//     EntRT,     //La variable VH_RT
     utilGC,
     KPMGutil,
     utobdebug,
     HMsgBox,
     HTB97,
     utofMixte,
     UTOF ;

Type
  TOF_RTRECHMIXTE = Class (TOF_MIXTE)
   private
      LaListe : THGrid;

   public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

    procedure LaListeOnDblClick (Sender : Tobject);

  end ;

Function RTLanceFiche_RTRECHMIXTE(Nat,Cod : String ; Range,Lequel,Argument : string) : string;

Implementation

{$IFDEF VER150}
  uses
    Variants
    , uTOFComm;
{$ENDIF VER150}

Function RTLanceFiche_RTRECHMIXTE(Nat,Cod : String ; Range,Lequel,Argument : string) : string;
begin
AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_RTRECHMIXTE.OnArgument (S : String ) ;
begin
  With ParamMixte do
  begin
    Suffixe             := 'CCM';
    Action              := taModif;
    TypeFiche           := 'RECHERCHE';
  end;

  Inherited ;
  LaListe               := THGrid(GetControl('FLISTE'));
  LaListe.OnDblClick    := LaListeOnDblClick;

  SetControlVisible('BSELECTALL', False);

  {$IFDEF GIGI}
    if assigned(GetControl('CCM_REPRESENTANT')) then
     begin
     SetControlVisible ('CCM_REPRESENTANT', False);
     SetControlVisible ('TCCM_REPRESENTANT', False);
     end;
   if (GetControl('CCM_ZONECOM') <> nil) then  SetControlVisible('CCM_ZONECOM',false);
   if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
   if (GetControl('TT_TARIFTIERS') <> nil) then  SetControlVisible('TT_TARIFTIERS',false);
   if (GetControl('CCM_TARIFTIERS') <> nil) then  SetControlVisible('CCM_TARIFTIERS',false);
   if (GetControl('CCM_CONFIDENTIEL') <> nil) then  SetControlVisible('CCM_CONFIDENTIEL',false);
//   SetControlText('CCM_NATURE','CLI;PRO;NCP');
//   SetControlProperty ('CCM_NATURE', 'Complete', true);
//   SetControlProperty ('CCM_NATURE', 'Datatype', 'TTNATTIERS');
//   SetControlProperty ('CCM_NATURE', 'Plus', VH_GC.AfNatTiersGRCGI);
  {$ENDIF GIGI}

end ;

procedure TOF_RTRECHMIXTE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RTRECHMIXTE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RTRECHMIXTE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_RTRECHMIXTE.OnLoad ;
var
  F                     : TFMul;

begin
  Inherited ;
  F                     := TFMul(Ecran);
  ReConstruireListe(F);
end ;


procedure TOF_RTRECHMIXTE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_RTRECHMIXTE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_RTRECHMIXTE.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_RTRECHMIXTE.LaListeOnDblClick(Sender: Tobject);
begin
  BZoomOnClick(Self);
end;

Initialization
  registerclasses ( [ TOF_RTRECHMIXTE ] ) ;
end.
