{***********UNITE*************************************************
Auteur  ...... : DECOSSE
Créé le ...... : 09/03/2001
Modifié le ... :   /  /
Description .. : Source TOF de la TABLE : CONSULTREVIS ()
Mots clefs ... : TOF;CONSULTREVIS
*****************************************************************}
Unit ConsultRevis_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL,  // AGLLanceFiche
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     FE_Main,   // AGLLanceFiche
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     UTOB,
     AGLInit,
     Stat ;

procedure CPLanceFiche_ConsultRevis(pszArgument : String);

Type
  TOF_CONSULTREVIS = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnClose                  ; override ;
    procedure OnArgument (S : String ) ; override ;
  end ;

Implementation

procedure CPLanceFiche_ConsultRevis(pszArgument : String);
begin
  AGLLanceFiche('CP','CPCONSULTREVIS','','',pszArgument);
end;

procedure TOF_CONSULTREVIS.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_CONSULTREVIS.OnClose ;
begin
TFStat(Ecran).LaTOB.Free ; TFStat(Ecran).LaTOB:=Nil ;
  Inherited ;
end ;

procedure TOF_CONSULTREVIS.OnArgument (S : String ) ;
Var Q : TQuery ;
    TOBEcr : TOB ;
    NomT : String ;
begin
Ecran.HelpContext := 7365000;
TOBEcr:=TOB.Create('',Nil,-1) ;
if Pos('Y_',S)>0 then NomT:='ANALYTIQ' else NomT:='ECRITURE' ;
Q:=OpenSQL('SELECT * FROM '+NomT+' '+S,True) ;
TOBEcr.LoadDetailDB(NomT,'','',Q,False) ;
Ferme(Q) ;
TFStat(Ecran).LaTOB:=TOBEcr ;
TMemo(getcontrol('FSQL')).Lines.add('SELECT * FROM '+NomT+' '+S) ;
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_CONSULTREVIS ] ) ;
end.

