{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 02/04/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPMULABO ()
Mots clefs ... : TOF;CPMULABO
*****************************************************************}
Unit CPMULABO_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, 
     mul,
     FE_Main,
     HDB,
{$else}
     eMul, 
     uTob,
     MaineAGL,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     AGLInit, // pour TheMulQ
     CONTABON_TOM,
     HTB97,
     HMsgBox,
     UTOF ; 

Type
  TOF_CPMULABO = Class (TOF)
   private
    {$IFDEF EAGLCLIENT}
    FListe   : THGrid ;
    {$ELSE}
    FListe   : THDBGrid ;
    {$ENDIF}
    procedure FListeDblClick(Sender: TObject) ;
    procedure BChercheClick(Sender: TObject);
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

procedure ListeAbonnements ;

Implementation

uses
 UtilPgi,
 SaisUtil ;  // pour le GereSelectionsGrid

procedure ListeAbonnements ;
begin

 if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;

 AGLLanceFiche('CP','CPMULABO','','','') ;

end ;

procedure TOF_CPMULABO.BChercheClick(Sender: TObject);
begin
 TFMul(Ecran).BChercheClick(Sender);  // inherited;
 {$IFDEF EAGLCLIENT}
 TheMulQ := TFMul(Ecran).Q.TQ ;
 GereSelectionsGrid(FListe,TheMulQ) ;
{$ELSE}
 TheMulQ := TFMul(Ecran).Q ;
// GereSelectionsGrid(FListe,THQuery(TheMulQ)) ;
{$ENDIF}

end;

procedure TOF_CPMULABO.FListeDblClick(Sender: TObject);
begin
 inherited;
 {$IFDEF EAGLCLIENT}
 TheMulQ := TFMul(Ecran).Q.TQ ;
{$ELSE}
 TheMulQ := TFMul(Ecran).Q ;
{$ENDIF}
 if TheMulQ.EOF and TheMulQ.Bof then Exit ;
 ParamAbonnement(True,GetField('CB_CONTRAT'),taModif) ;
end;

procedure TOF_CPMULABO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPMULABO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPMULABO.OnUpdate ;
begin
 Inherited ;
end ;

procedure TOF_CPMULABO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPMULABO.OnArgument (S : String ) ;
begin

 Inherited ;

 {$IFDEF EAGLCLIENT}
 TheMulQ := TFMul(Ecran).Q.TQ ;
 FListe  := THGrid (GetControl ('FLISTE')) ;
{$ELSE}
 TheMulQ := TFMul(Ecran).Q ;
 FListe := THDBGrid (GetControl ('FLISTE'));
{$ENDIF}

 FListe.OnDblClick := FListeDblClick ;
 TToolbarButton97(GetControl('BCHERCHE')).OnClick := BChercheClick ;
  {JP 14/06/06 : FQ 17406}
  Ecran.HelpContext := 7346000;
end ;

procedure TOF_CPMULABO.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPMULABO.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPMULABO.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CPMULABO ] ) ; 
end.
