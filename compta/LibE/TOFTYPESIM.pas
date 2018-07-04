{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/12/2001
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CTYPESIM ()
Mots clefs ... : TOF;TOFTYPESIM
*****************************************************************}
Unit TOFTYPESIM;

Interface

Uses
     StdCtrls,
     Controls,
     Classes,
     forms,
     Windows,
     sysutils,
     HCtrls,
     HEnt1,
     Vierge,
  {$IFNDEF EAGLCLIENT}
     FE_Main,
 {$ELSE}
     MaineAGL,
 {$ENDIF}
     UObjetEmprunt,
     UTOF;

Type
  TOF_TYPESIM = Class (TOF)
    private

    published
       procedure OnNew                    ; override ;
       procedure OnDelete                 ; override ;
       procedure OnUpdate                 ; override ;
       procedure OnLoad                   ; override ;
       procedure OnArgument (S : String ) ; override ;
       procedure OnClose                  ; override ;
  end ;


Function ChoixTypeSim(pStTypeTaux,pStTypeVers : String) : TTypeRecherche;

Implementation


// Lancement de la fenêtre de choix du type de simulation
Function ChoixTypeSim(pStTypeTaux,pStTypeVers : String) : TTypeRecherche;
var
   lStTypeSim : String;
begin
   lStTypeSim := AGLLanceFiche('FP','FTYPESIM','','',pStTypeTaux + ';' + pStTypeVers);
   if lStTypeSim = 'C' then
      Result := trCapital
   else
   if lStTypeSim = 'D' then
      Result := trDuree
   else
   if lStTypeSim = 'T' then
      Result := trTaux
   else
   if lStTypeSim = 'V' then
      Result := trVersement
   else
      Result := trAucune;
end;


//
procedure TOF_TYPESIM.OnArgument (S : String ) ;
var
  lStTypeTaux : String;
  lStTypeVers : String;
begin
  Inherited ;

  lStTypeTaux := ReadTokenSt(S);
  lStTypeVers := ReadTokenSt(S);

  // Récup des controls
  SetControlChecked('rbCapital',   True);
  SetControlChecked('rbTaux',      not ((lStTypeVers = tvVariable) and (lStTypeTaux = ttConstant)) );
  SetControlChecked('rbDuree',     (lStTypeTaux = ttConstant) );
  SetControlChecked('rbVersement', not ((lStTypeVers = tvVariable) and (lStTypeTaux = ttConstant)) );
  SetControlChecked('rbPlan',      (lStTypeVers = tvVariable) and (lStTypeTaux = ttConstant) );

  SetFocusControl('rbCapital');

  TFVierge(Ecran).Retour := '';

end ;

// Validation -> valeur de retour
procedure TOF_TYPESIM.OnUpdate ;
var
  lStRetour : string;
begin
  Inherited ;

  lStRetour := '';
  if TRadioButton(GetControl('rbCapital')).Checked then
     lStRetour := 'C'
  else
  if TRadioButton(GetControl('rbTaux')).Checked then
     lStRetour := 'T'
  else
  if TRadioButton(GetControl('rbDuree')).Checked then
     lStRetour := 'D'
  else
  if TRadioButton(GetControl('rbVersement')).Checked then
     lStRetour := 'V'
  else
  if TRadioButton(GetControl('rbPlan')).Checked then
     lStRetour := 'P';

  TFVierge(Ecran).Retour := lStRetour;

  Ecran.Close;
end ;

procedure TOF_TYPESIM.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_TYPESIM.OnDelete ;
begin
  Inherited ;
end ;


procedure TOF_TYPESIM.OnLoad ;
begin
  Inherited ;
end ;


procedure TOF_TYPESIM.OnClose ;
begin
  Inherited ;
end ;




Initialization
  registerclasses ( [ TOF_TYPESIM ] ) ;
end.

