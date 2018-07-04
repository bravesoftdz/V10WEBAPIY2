{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/12/2001
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CAMORTCST ()
Mots clefs ... : TOF;TOFAMORTCST
*****************************************************************}
Unit TOFEMPMONTANT;

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
{$IFNDEF EAGLCLIENT}
     FE_Main,
 {$ELSE}
     MaineAGL,
 {$ENDIF}
     UTOF,
     UCreCommun;

Type
  TOF_EMPMONTANT = Class (TOF)
    private
       edMontant    : THNumEdit;
       edConversion : THNumEdit;

       fRdTaux : Double;

       Procedure edMontantChange(Sender : TObject);

    published
       procedure OnNew                    ; override ;
       procedure OnDelete                 ; override ;
       procedure OnUpdate                 ; override ;
       procedure OnLoad                   ; override ;
       procedure OnArgument (S : String ) ; override ;
       procedure OnClose                  ; override ;
  end ;


Function SaisieMontant(pRdMontant : Double) : Double;


Implementation


// ----------------------------------------------------------------------------
// Nom    : SaisieMontant
// Date   : 07/11/2001
// Auteur : D. ZEDIAR
// Objet  : Lancement de la saisie d'un montant
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function SaisieMontant(pRdMontant : Double) : Double;
var
   lStRetour : String;
begin
   lStRetour := AGLLanceFiche('FP','FEMPMONTANT','','',FloatToStr(pRdMontant) );
   if lStRetour = '' then
      result := -1
   else
      Result := Valeur(lStRetour);
end;

// ----------------------------------------------------------------------------
// Nom    : OnArgument
// Date   : 07/11/2001
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_EMPMONTANT.OnArgument (S : String ) ;
begin
  Inherited ;

  // Récup des controls
  edMontant    := THNumEdit(GetControl('edMontant'));
  edConversion := THNumEdit(GetControl('edConversion'));

  edMontant.OnChange := edMontantChange;
  SetFocusControl('edMontant');

  if gCodeDevSaisie = '' then
  begin
    if gCodeDevEmprunt = gCodeDevEuro then
    begin
       SetControlCaption('laDevMnt', gSymbDevFranc);
       SetControlCaption('laDevConv', gSymbDevEuro);
       fRdTaux := 1 / V_PGI.TauxEuro;
    end
    else
    begin
       SetControlCaption('laDevMnt', gSymbDevEuro);
       SetControlCaption('laDevConv', gSymbDevFranc);
       fRdTaux := V_PGI.TauxEuro;
    end;
  end
  else
  begin
     SetControlCaption('laDevMnt', gSymbDevSaisie);
     SetControlCaption('laDevConv', gSymbDevEmprunt);
     if gCodeDevSaisie = gCodeDevEuro then
       fRdTaux := V_PGI.TauxEuro
     else
       fRdTaux := 1 / V_PGI.TauxEuro;
  end;

  if S = '' then
     edConversion.Value := 0
  else
     edConversion.Value := Valeur(S);
  edMontant.Value := edConversion.Value / fRdTaux;
  edMontantChange(Self);

  TFVierge(Ecran).Retour := '';

end ;

//
procedure TOF_EMPMONTANT.OnNew ;
begin
  Inherited ;
end ;

//
procedure TOF_EMPMONTANT.OnDelete ;
begin
  Inherited ;
end ;

//
procedure TOF_EMPMONTANT.OnLoad ;
begin
  Inherited ;
end ;

// Validation de la saisie -> valeur de retour
procedure TOF_EMPMONTANT.OnUpdate ;
begin
  Inherited ;
  TFVierge(Ecran).Retour := FloatTostr(edConversion.Value);
  Ecran.Close;
end ;

//
Procedure TOF_EMPMONTANT.edMontantChange(Sender : TObject);
begin
   // Calcul du montant converti
   try
      edConversion.Value := Arrondi(fRdTaux * Valeur(edMontant.Text),2);
   except
   end;
end;

//
procedure TOF_EMPMONTANT.OnClose ;
begin
  Inherited ;
end ;


Initialization
  registerclasses ( [ TOF_EMPMONTANT ] ) ;
end.

