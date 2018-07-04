{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 16/10/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFPLANNINGGENERER ()
Mots clefs ... : TOF;AFPLANNINGGENERER
*****************************************************************}
Unit UTofAFPLANNINGGENERER;

Interface

Uses StdCtrls,
     Controls,
     Classes, paramsoc,
     vierge, 
{$IFNDEF EAGLCLIENT}
     db,
     dbtables,
     FE_Main,
{$ELSE}
     MaineAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF, AFPlanningGene, utob, DicoAF, UtilTaches,
     utilPlanning, AFPlanningCst;

Type
  TOF_AFPLANNINGGENERER = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

    private
      fGenerateur : TAFGenerateur;
      fStAffaire  : String;
      fTache      : RecordTache;
      fSelection  : Boolean;

  end;

  Function AFLanceFicheAFPlanningGenerer(Lequel, Argument : String) : String;

Implementation

Function AFLanceFicheAFPlanningGenerer(Lequel, Argument : String) : String;
begin
  result := AGLLanceFiche('AFF','AFPLANNINGGENERER','', Lequel, Argument);
end;

procedure TOF_AFPLANNINGGENERER.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGGENERER.OnDelete ;
begin
  Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 15/10/2002
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_AFPLANNINGGENERER.OnUpdate;
var
 vInMode  : Integer;

begin
  Inherited;

  vInMode := THRadioGroup(GetControl('RG_REMPLACER')).ItemIndex;
  if TCheckbox(GetControl('CB_SUPPRIMERMANUEL')).Checked then
    vInMode := vInMode + 1;
  // mode ajout
  if vInMode = 1 then vInMode := cInAjout;
  if TCheckbox(GetControl('CB_VALORISE')).Checked then
    TFVierge(ecran).retour := IntToSTr(vInMode) + ';TRUE'
  else
    TFVierge(ecran).retour := IntToSTr(vInMode) + ';FALSE';

  if not EtatInterdit(vInMode, fStAffaire, fTache) then
  begin
    fGenerateur := TAFGenerateur.create;
    try
      if fGenerateur.GenererPlanning(fStAffaire,
                                     fTache,
                                    vInMode,
                                    TCheckbox(GetControl('CB_VALORISE')).Checked) then
        begin
          Close;
          PGIInfoAF ('La génération du planning est terminée', '');
        end;

    finally
      fGenerateur.Free;
    end;
  end;
end;

procedure TOF_AFPLANNINGGENERER.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFPLANNINGGENERER.OnArgument (S : String ) ;
var
  Tmp     : String;
  champ   : String;
  valeur  : String;

begin

  fSelection := S = 'SELECTION';
  if not fSelection then
  begin
    // traitement des arguments
    fStAffaire := '';
    fTache.StNumeroTache := '';
    Tmp:=(Trim(ReadTokenSt(S)));
    While (Tmp <>'') do
      Begin
        If Tmp<>'' then
          Begin
            DecodeArgument(Tmp, Champ, valeur);
            If Champ='AFFAIRE'        then fStAffaire := valeur else
            If Champ='NUMEROTACHE'    then fTache.StNumeroTache := valeur else
            if Champ='LIBTACHE'       then fTache.StLibTache := valeur else
            if Champ='FCTTACHE'       then fTache.StFctTache := valeur else
            if Champ='ARTICLE'        then fTache.StArticle := valeur else
            if champ='TIERS'          then fTache.StTiers := valeur else
            if Champ='UNITE'          then fTache.StUnite := valeur else
            if Champ='LASTDATEGENE'   then fTache.StLastDateGene := valeur else
            if Champ='COMPTEUR'       then fTache.BoCompteur := ((StrToInt(valeur) <> 0) and getParamsoc('SO_AFREALPLAN')) else
            if Champ='ACTIVITEREPRIS' then fTache.StActiviteRepris := valeur;
          End;
          Tmp:=(Trim(ReadTokenSt(S)));
      End;
  end;
  TCheckBox(GetControl('CB_VALORISE')).Checked := False;
end;
   
procedure TOF_AFPLANNINGGENERER.OnClose ;
begin
  Inherited;
end;

Initialization
  registerclasses ( [ TOF_AFPLANNINGGENERER ] ) ;
end.
