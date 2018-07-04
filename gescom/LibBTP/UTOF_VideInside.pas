{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 26/03/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : VIDEINSIDE ()
Mots clefs ... : TOF;VIDEINSIDE
*****************************************************************}
Unit UTOF_VideInside ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     FE_Main,
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     Maineagl,
     eMul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     AglInit,
     SaisieConsommations,
     BTMODELHEBDO_TOF,
     BTCLOTUREEX_TOF,
     REGENERECONSO_tof
     ,vierge
     ,BTTVA_TOF;

Type
  TOF_VIDEINSIDE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    fretour : string;
  end ;

function SaisirHorsInside(Nomfiche : string; param:string='') : string;

Implementation
uses fichcomm, AssistparamSocBTP;

var ValRetour : string;

function SaisirHorsInside(Nomfiche : string; param:string='') : string;
var S : string;
begin
  Valretour := '';
	S :='FICHE='+NomFiche;
  if Param <> '' then S := S + ';'+ param;
  AGLLanceFiche('BTP','BTVIDEINSIDE','','',S) ;
  if ValRetour <> '' then Result := ValRetour;
end;

procedure TOF_VIDEINSIDE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_VIDEINSIDE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_VIDEINSIDE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_VIDEINSIDE.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_VIDEINSIDE.OnArgument (S : String ) ;
var position : integer;
    laFiche,LeQuel : string;
    Argument : string;
    Arguments : string;
begin
  Inherited ;
  leQuel := '';
  fretour := '';
  Arguments := S;

  Position :=  Pos('FICHE=',Arguments);
  if Position > 0 then
  begin
  	laFiche := READTOKENST (Arguments);
    laFiche := Copy(laFiche,7,255);
  end;

  Position :=  Pos('LEQUEL=',Arguments);
  if Position > 0 then
  begin
  	leQuel := READTOKENST (Arguments);
    LeQuel := Copy(LeQuel,8,255);
  end;

  Argument := Arguments;

  if lafiche = '' then exit;

  if laTob <> nil then TheTOB := LATOB;

  if Argument = '' then Argument := 'ACTION=MODIFICATION';

  ecran.Caption := LaFiche;

  if lafiche='PARAMSOC' then
  begin
  	AppelParamSociete;
  end
  else if lafiche ='MODPAIE' then
  begin
  	FicheModePaie_AGL(' ');
  end
  else if lafiche ='GCUNITEMESURE' then
  begin
  	AGLLanceFiche('GC',LaFiche,'','','');
  end
  else if Lafiche = 'GCCODECPTA' then
  begin
  	AGLLanceFiche('GC',LaFiche,'','','');
  end
  else if Lafiche = 'SAISIECONSOMMATIONS' then SaisieConsommation
  else if Lafiche = 'SAISIEINTERVENTION' then SaisieIntervention
  else if Lafiche = 'TVA' then
  begin
  	ecran.caption := '';
  	ParamTVATPF(true);
  end
  Else if  LaFiche = 'BTREGENERECONSO'  then GestionRegenerationConso
  Else if (LaFiche = 'BTPARSOCCPTAGS1') Or
    			(LaFiche = 'BTPARAMSOCS')     Or
          (LaFiche = 'BTPARAMGENS1')    Or
          (LaFiche = 'BTPARAMSOCPART')  Or
          (Lafiche = 'BTSFAMART')       then
       AGLLanceFiche('BTP',LaFiche,'','','ALONE')
  //else if Lafiche = 'BTENVOIECRS1' then
  // 		 AGLLANCEFICHE('BTP','BTENVOIECRS1','','','MODIFICATION')
  //else if Lafiche = 'BTRECUPDOSS1' then
  // 		 AGLLANCEFICHE('BTP','BTRECUPDOSS1','','','MODIFICATION')
  else if Lafiche = 'BTCPTPIECE_S1' then
	     AGLLanceFiche('BTP',LaFiche,'','','NATURE=DBT;NUMERO=0')
  else if LaFiche = 'BTCLOTUREEX' then BTClotureEx
  else if LaFiche = 'GCSELDOCREA_MUL' then AGLLANCEFICHE ('GC',laFiche,'','','ACTION=MODIFICATION')
  else if LaFiche = 'BTAPPELINT' then AGLLANCEFICHE ('BTP',laFiche,'','',Argument)
  else  if Copy(LaFiche,1,1) = 'Z' then ValRetour  := AGLLANCEFICHE ('Z',laFiche,'',LeQuel,Argument)
  else ValRetour  := AGLLANCEFICHE ('BTP',laFiche,'',LeQuel,Argument);

end ;

procedure TOF_VIDEINSIDE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_VIDEINSIDE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_VIDEINSIDE.OnCancel () ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOF_VIDEINSIDE ] ) ; 
end.

