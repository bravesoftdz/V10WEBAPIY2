{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/12/2001
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CAMORTCST ()
Mots clefs ... : TOF;TOFAMORTCST
*****************************************************************}
Unit TOFEMPCHOIXIMP;
                  
Interface

Uses
     StdCtrls,
     Controls,
     Graphics,
     Classes,
     buttons,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     forms,
     Windows,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Vierge,
     HTB97,
{$IFNDEF EAGLCLIENT}
     FE_Main,
 {$ELSE}
     MaineAGL,
 {$ENDIF}
     UTOF;

Type


  TOF_EMPCHOIXIMP = Class (TOF)
    published
       procedure OnNew                    ; override ;
       procedure OnDelete                 ; override ;
       procedure OnUpdate                 ; override ;
       procedure OnLoad                   ; override ;
       procedure OnArgument (S : String ) ; override ;
       procedure OnClose                  ; override ;
  end ;

  TOF_EMPCHOIXDATE = Class (TOF)
    private

    published
       procedure OnArgument (S : String ) ; override ;
       procedure OnUpdate                 ; override ;
  end ;

Function ChoixDeviseImpression(lStSymbDevEmprunt, lStSymbDevSaisie : String) : String;
Function ChoixDateReference : TDateTime;


Implementation


Function ChoixDeviseImpression(lStSymbDevEmprunt, lStSymbDevSaisie : String) : String;
begin
   Result := AGLLanceFiche('FP','FEMPCHOIXIMP','','',lStSymbDevEmprunt + ';' + lStSymbDevSaisie);
end;

Function ChoixDateReference : TDateTime;
var
   lStDate : String;
begin
   lStDate := AGLLanceFiche('FP','FEMPCHOIXDATE','','','');

   if lStDate = '' then
      Result := -1
   else
      Result := StrToDate(lStDate);
end;


procedure TOF_EMPCHOIXIMP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_EMPCHOIXIMP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_EMPCHOIXIMP.OnUpdate ;
begin
  Inherited ;

  if TRadioButton(GetControl('RBDEVDOSSIER')).Checked then
     TFVierge(Ecran).Retour := 'DOSSIER'
  else
     TFVierge(Ecran).Retour := 'ORIGINE';
end ;

procedure TOF_EMPCHOIXIMP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_EMPCHOIXIMP.OnArgument (S : String ) ;
var
  lStSymbDevEmprunt : String;
  lStSymbDevSaisie  : String;
begin
  Inherited ;

  lStSymbDevEmprunt := ReadTokenSt(S);
  lStSymbDevSaisie  := ReadTokenSt(S);

  if lStSymbDevEmprunt <> '' then
     SetControlCaption('rbDevDossier', 'Edition dans la devise du &dossier (' + lStSymbDevEmprunt + ')');

  if lStSymbDevSaisie <> '' then
     SetControlCaption('rbDevOrig', 'Edition dans la devise d''origine de l''&emprunt (' + lStSymbDevSaisie + ')');
end ;



procedure TOF_EMPCHOIXIMP.OnClose ;
begin
  Inherited ;
end ;
















procedure TOF_EMPCHOIXDATE.OnArgument (S : String ) ;
begin
  Inherited ;

  TFVierge(Ecran).Retour := '';

end ;


procedure TOF_EMPCHOIXDATE.OnUpdate ;
var
  lDtDate : TDateTime;
begin
  Inherited ;

  try
     lDtDate := StrToDate(GetControlText('edDate'));
     TFVierge(Ecran).Retour := DateToStr(lDtDate);
  except
     PGIError('La date saisie n''est pas valide !',TitreHalley);
     Exit;
  end;

  Ecran.Close;
end ;


Initialization
  registerclasses ( [ TOF_EMPCHOIXIMP, TOF_EMPCHOIXDATE ] ) ;
end.

