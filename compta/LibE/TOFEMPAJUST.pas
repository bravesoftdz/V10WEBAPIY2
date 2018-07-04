{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/12/2001               
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CAMORTCST ()
Mots clefs ... : TOF;TOFAMORTCST
*****************************************************************}
Unit TOFEMPAJUST;

Interface

Uses
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
     UCreCommun,
     UTOF;

Type
  TOF_EMPAJUST = Class (TOF)
    private
       edValAjust  : THNumEdit;

    published
       procedure OnNew                    ; override ;
       procedure OnDelete                 ; override ;
       procedure OnUpdate                 ; override ;
       procedure OnLoad                   ; override ;
       procedure OnArgument (S : String ) ; override ;
       procedure OnClose                  ; override ;
  end ;


Function SaisieAjustement(pRdMontant : Double) : Double;


Implementation


// ----------------------------------------------------------------------------
// Nom    : SaisieAjustement
// Date   : 06/11/2002
// Auteur : D. ZEDIAR
// Objet  : Lance la saisie du montant de l'ajustement
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function SaisieAjustement(pRdMontant : Double) : Double;
var
   lStMontant : String;
begin
   lStMontant := AGLLanceFiche('FP','FEMPAJUST','','',FloatToStr(pRdMontant));
   if lStMontant = '' then
      Result := -1
   else
      Result := Valeur(lStMontant);
end;


// ----------------------------------------------------------------------------
// Nom    : OnArgument
// Date   : 06/11/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_EMPAJUST.OnArgument (S : String ) ;
begin
  Inherited ;

  // Récup des controls
  edValAjust := THNumEdit(GetControl('edValAjust'));
  edValAjust.CurrencySymbol := gSymbDevEuro;
  if S = '' then
     edValAjust.Value := 0
  else
     edValAjust.Value := Valeur(S);
  SetFocusControl('edValAjust');

  TFVierge(Ecran).Retour := '';

end ;

//
procedure TOF_EMPAJUST.OnNew ;
begin
  Inherited ;
end ;

//
procedure TOF_EMPAJUST.OnDelete ;
begin
  Inherited ;
end ;

// Validation de la saisie -> valeur de retour
procedure TOF_EMPAJUST.OnUpdate ;
begin
  Inherited ;
  TFVierge(Ecran).Retour := FloatToStr(Valeur(edValAjust.Text));
  Ecran.Close;
end ;

//
procedure TOF_EMPAJUST.OnLoad ;
begin
  Inherited ;
end ;


//
procedure TOF_EMPAJUST.OnClose ;
begin
  Inherited ;
end ;





Initialization
  registerclasses ( [ TOF_EMPAJUST ] ) ;
end.

