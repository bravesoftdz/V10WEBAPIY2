{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/12/2001
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CAMORTCST ()
Mots clefs ... : TOF;TOFAMORTCST
*****************************************************************}
Unit TOFEMPCST;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Windows,
     sysutils,
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
     UTOF,
     UCreCommun;


Type
  TOF_EMPCST = Class (TOF)
    private
       fBoAmortissement : Boolean;
       procedure ButtonClick(Sender: TObject);

    published
       procedure OnNew                    ; override ;
       procedure OnDelete                 ; override ;
       procedure OnUpdate                 ; override ;
       procedure OnLoad                   ; override ;
       procedure OnArgument (S : String ) ; override ;
       procedure OnClose                  ; override ;
  end ;


Function SaisieMontantConstant(pBoAmortissement : Boolean; pRdMontant : Double) : Double;


Implementation

// ----------------------------------------------------------------------------
// Nom    : SaisieMontantConstant
// Date   : 06/11/2002
// Auteur : D. ZEDIAR
// Objet  : Lance la saisie de l'élément constant (versment ou amortissement)
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function SaisieMontantConstant(pBoAmortissement : Boolean; pRdMontant : Double) : Double;
var
   lStArgs    : String;
   lStMontant : String;
begin
   if pBoAmortissement then
      lStArgs := 'A'
   else
      lStArgs := 'V';

   lStArgs := lStArgs + ';' + FloatToStr(pRdMontant);

   lStMontant := AGLLanceFiche('FP','FEMPCST','','',lStArgs);

   if lStMontant = '' then
      Result := -1
   else
      result := Valeur(lStMontant);
end;


// ----------------------------------------------------------------------------
// Nom    : OnArgument
// Date   : 06/11/2002
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_EMPCST.OnArgument (S : String ) ;
var
  lStType    : String;
  lStMontant : String;
begin
  Inherited ;

  lStType    := ReadTokenSt(S);
  lStMontant := ReadTokenSt(S);

  fBoAmortissement := (lStType = 'A') or (lStType = 'a');

  SetControlProperty('edValCalc', 'Value', Valeur(lStMontant));
  SetControlProperty('edValCst', 'Value', Valeur(lStMontant));

  if not fBoAmortissement then
  begin
     Ecran.Caption     := 'Saisie du versement constant';
     SetControlCaption('laValCalc', 'Versement calculé');
     SetControlCaption('laValCst', 'Montant du versement');
  end;

  SetControlProperty('edValCalc', 'CurrencySymbol', gSymbDevEmprunt);
  SetControlProperty('edValCst', 'CurrencySymbol', gSymbDevEmprunt);
  SetFocusControl('edValCst');

  TToolbarButton97(GetControl('btDown')).OnClick := ButtonClick;

  TFVierge(Ecran).Retour := '';

end ;

//
procedure TOF_EMPCST.OnNew ;
begin
  Inherited ;
end ;

//
procedure TOF_EMPCST.OnDelete ;
begin
  Inherited ;
end ;

// Validation de la saisie -> valeur de retour
procedure TOF_EMPCST.OnUpdate ;
begin
  Inherited ;

  TFVierge(Ecran).Retour := FloatToStr(Valeur(GetControlText('edValCst')));
  Ecran.Close;

end ;

//
procedure TOF_EMPCST.OnLoad ;
begin
  Inherited ;
end;

//
procedure TOF_EMPCST.OnClose ;
begin
  Inherited ;
end ;

// Copie du montant calculé dans le champ de saisie
procedure TOF_EMPCST.ButtonClick(Sender: TObject);
begin
   SetControlProperty('edValCst', 'Value', Valeur(GetControlText('edValCalc')));
end;



Initialization
  registerclasses ( [ TOF_EMPCST ] ) ;
end.

