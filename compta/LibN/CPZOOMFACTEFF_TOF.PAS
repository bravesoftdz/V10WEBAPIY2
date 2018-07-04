{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 02/05/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPZOOMFACTEFF ()
Mots clefs ... : TOF;CPZOOMFACTEFF
*****************************************************************}
Unit CPZOOMFACTEFF_TOF;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,    // AGLLanceFiche
     eMul,        // TFMul
{$ELSE}
     db,
     dbtables,
     FE_Main,     // AGLLanceFiche
     Mul,         // TFMul 
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     AGLInit,     // TheData
     SAISUTIL,    // TOBM
     UTOF ; 

Procedure ZoomFactCliEff(LOBM : TList);

Type
  TOF_CPZOOMFACTEFF = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    LOBM : TList ;
    procedure FaitXX_WHERE;
  end ;

Implementation

Procedure ZoomFactCliEff(LOBM : TList) ;
begin
  TheData := LOBM;
  AGLLanceFiche('CP','CPZOOMFACTEFF','','', '');
end;

procedure TOF_CPZOOMFACTEFF.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPZOOMFACTEFF.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPZOOMFACTEFF.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPZOOMFACTEFF.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CPZOOMFACTEFF.OnArgument (S : String ) ;
begin
  Inherited ;

  // Récupères les arguments
  LOBM := TList(TheData);
  TheData := nil;

  // FormShow
  FaitXX_WHERE ;
  TFMul(Ecran).BAgrandirClick(Nil);
  SetControlVisible('BAGRANDIR', False);
  SetControlVisible('BREDUIRE', False);
end ;

procedure TOF_CPZOOMFACTEFF.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPZOOMFACTEFF.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPZOOMFACTEFF.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_CPZOOMFACTEFF.FaitXX_WHERE;
var
  St,St1 : String ;
  O : TOBM ;
  i : Integer ;
begin
  St:='' ;
  for i:=0 To LOBM.Count-1 do begin
    O := TOBM(LOBM[i]);
    St1:='(E_JOURNAL="'+O.GetMvt('E_JOURNAL')+'" AND E_EXERCICE="'+QuelExo(DateToStr(O.GetMvt('E_DATECOMPTABLE')))+'" AND '+
         ' E_DATECOMPTABLE="'+UsDateTime(O.GetMvt('E_DATECOMPTABLE'))+'" AND '+
         ' E_NUMEROPIECE='+IntToStr(O.GetMvt('E_NUMEROPIECE'))+' AND '+
         ' E_NUMLIGNE='+IntToStr(O.GetMvt('E_NUMLIGNE'))+' AND E_NUMECHE='+IntToStr(O.GetMvt('E_NUMECHE'))+') ';
    If St='' then St:=St1
             else St:=St+' OR '+St1;
  end;
  If St='' Then Exit;
  St := '('+St+')';
  SetControlText('XX_WHERE', St);
end;

Initialization
  registerclasses ( [ TOF_CPZOOMFACTEFF ] ) ; 
end.
