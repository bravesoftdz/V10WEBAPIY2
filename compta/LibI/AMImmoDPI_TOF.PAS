{***********UNITE*************************************************
Auteur  ...... : TGA
Créé le ...... : 18/07/2006
Modifié le ... :
Description .. : Source TOF de la FICHE : AMImmoDPI ()
               : Saisie des affectations de la DPI pour une immobilisation

Mots clefs ... : TOF;AMImmoDPI

Appel de la fonction
    result := AGLLanceFiche('AM', 'AMIMMODPI','' , '', code);
    code = code immo, montant maxi

Suite : TGA 05/09/2006 Rev5
11/10/2006 - MVG FQ 18956
13/10/2006 - MVG FQ 18972
10/06      - BTY FQ 18949 F6 doit ramener la plus petite des valeurs solde dispo et base éco
10/06      - BTY 10/06 Afficher le solde dispo dans l'écran de saisie
10/06      - BTY 10/06 Oter test de présence d'ancienne affectation immo
10/06      - BTY 10/06 Dates de la mention Solde dispo
11/06      - BTY FQ 19087 On peut saisir un montant à 0
******************************************************************}
Unit AMIMMODPI_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MaineAGL,
{$ENDIF}
     ImEnt,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     HTB97,
     Vierge,
     uTOB,
     windows,
     AmDPI,   // BTY afficher solde dispo
     AGLInit, // BTY TheData
     UTOF ;

Type
  TOF_AMIMMODPI = Class (TOF)
    procedure OnArgument (S : String ) ; override;
    procedure OnAfterFormShowEcran;
    procedure OnUpdate ; override;

    private
    { Variables }
    FBaseFisDec : Double;
    FBaseAmoEco : Double;
    FMontantMax : Double;
    FAncienMontant: Double;
    FDateFin : TDateTime;
    procedure OnbFermeClick(Sender: TObject);
    Function  ControleDuMontant :Boolean;
    procedure OnGrilleKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
  end ;

function AMLanceFiche_MontantImmoDPI(Stcode: string): String ;

Implementation
// 11/10/2006 - MVG FQ 18956

const
  HM: array[0..4] of string =
   ('Le montant affecté doit être positif et différent de 0.',
    'Le montant affecté doit être inférieur à la DPI disponible à cette date soit ',
    'Le montant de la DPI ne peut pas dépasser la base fiscalement déductible soit ',
    'Le montant de la DPI ne peut pas dépasser la base amortissement économique soit ',
    'Voulez-vous enregistrer cette opération ?');


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 18/07/2006
Modifié le ... :   /  /
Description .. : Appel de la fenêtre de saisie
Mots clefs ... :
*****************************************************************}
Function AMLanceFiche_MontantImmoDPI(Stcode: string): String ;
begin
   result:= AGLLanceFiche('AM', 'AMIMMODPI','' , '', Stcode);
End;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 18/07/2006
Modifié le ... :   /  /
Description .. :  Préparation de la fenêtre DPI
Mots clefs ... :
*****************************************************************}
procedure TOF_AMIMMODPI.OnArgument (S: String ) ;
Var
  //Q : Tquery;
  stDPIDispo : string;
begin
  Inherited ;
  // Paramêtres en entrée : Base fiscalement déductible
  //                        Base amortissement économique
  //                        Montant maximum de la DPI
  //                        fin de l'exrcice en cours

  FBaseFisDec := StrToFloat(ReadTokenSt(S));
  FBaseAmoEco := StrToFloat(ReadTokenSt(S));
  FMontantMax := StrToFloat(ReadTokenSt(S));
  FDateFin    := StrToDate(ReadTokenSt(S));

  // Préparation de la fenêtre
  TToolbarButton97(GetControl('BFerme')).Onclick := OnbFermeClick;
  TFVIERGE(ECRAN).retour := IntToStr(MrNo);
  TFVIERGE(ECRAN).OnAfterFormShow := OnAfterFormShowEcran;
  TFVIERGE(ECRAN).OnKeyDown := OnGrilleKeyDown;


  // Lecture de IMMOMVTD pour IZ_DATE = fin de l'exercice en cours
  // BTY 10/06  QUEZAQUO ??
  {Q :=OpenSQL('SELECT * FROM IMMOMVTD WHERE IZ_DATE="'+USDATETIME(FDateFin)+'"  AND ' +
                          'IZ_NATURED="IMM"', FALSE) ;
  If Not Q.EOF Then
     FAncienMontant := Q.FindField('IZ_MONTANT').AsFloat
  Else
     FAncienMontant := 0;
  Ferme(Q);

  // Affichage du montant de l'affectation de la DPI
  SetControlText('DE_DMTDPI',StrFMontant(FAncienMontant,15,V_PGI.OkDecV,'',True));
  }
  FAncienMontant := 0;

  // BTY 10/06 Afficher le solde DPI utilisable à fin exo N en cours
  if (TheData <> nil) then
    if (TheData is TAmDPI) then
       // TheData variable globale pointe depuis l'appelant sur l'objet fAmDPI alimenté
       begin
       StDPIDispo := 'DPI utilisables du '   +
                   DateToStr(TAmDPI(TheData).Exercice[5].Deb)  + ' au ' +
                   DateToStr(TAmDPI(TheData).Exercice[1].Fin)  + '    ';
       SetControlText ('StDPIDispo', StDPIDispo +
                   StrFMontant(fMontantMax, 15, V_PGI.OkDecV, '', true));
       end;
 End ;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 27/04/2006
Modifié le ... :   /  /
Description .. :  Mise à Jour des masques apres affichage
Mots clefs ... :
*****************************************************************}
procedure TOF_AMIMMODPI.OnAfterFormShowEcran;
begin
  SetControlProperty('DE_DMTDPI','DisplayFormat',StrfMask(V_PGI.OkDecV,'', True));;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 27/04/2006
Modifié le ... :   /  /
Description .. : Sortie par validation
Mots clefs ... :
*****************************************************************}
procedure TOF_AMIMMODPI.OnUpdate;

var erreur  : Boolean;
    montant : String;
    mt :Double;
begin
   Inherited ;

   // Récupération du montant de l'opération
   mt := Valeur(GetControlText('DE_DMTDPI'));
   SetControlText('DE_DMTDPI', StrFMontant(mt,15,V_PGI.OkDecV,'',True));

   erreur := ControleDuMontant;
   IF erreur=True then
     Begin
       ecran.ModalResult :=mrNone;
       SetFocusControl('DE_DMTDPI');
       exit;
     end;

   // Mise à jour du montant
   montant := GetControlText('DE_DMTDPI');
   TFVIERGE(ECRAN).retour := montant;

End;


{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 27/04/2006
Modifié le ... :   /  /
Description .. : Contrôle du montant de l'investissement
Mots clefs ... :
*****************************************************************}
Function TOF_AMIMMODPI.ControleDuMontant :Boolean;
var
   montant:Double;
   i : Integer;
   com : double;
begin
   i := 0;
   com :=0;
   result := False;

   // Récupération du montant de la déduction saisi
   montant := Valeur(GetControlText('DE_DMTDPI'));

   // MVG 13/10/2006 FQ 18972
   // Controle de l'affectation de la DPI

   IF montant<=0 then
      // Ctl montant differend de 0
      begin
      // FQ 19087
      i := 0;
      PGIBox(HM[i]);
      result:=True;
      exit;
      end
   Else IF (montant > FBaseAmoEco) then
      // Ctl mt > base amortissement économique
      Begin
         i:=3;
         Com := FBaseAmoEco;
      End
   Else If ((FBaseFisDec<>0) AND (montant > FBaseFisDec)) then
      // Ctl mt < base fiscalement déductible
      Begin
         i:=2;
         Com := FBaseFisDec;
      End
   Else If (montant > FMontantMax) Then
      // Ctl mt < au maximum
      Begin
         i:=1;
         Com := FMontantMax;
      End;

   IF i<>0 Then
     Begin
       PGIBox(HM[i]+' '+StrFMontant(com,15,V_PGI.OkDecV,'',True)+'.');
       result:=True;
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 27/04/2006
Modifié le ... :   /  /
Description .. : Sortie par la X sans validation
Mots clefs ... :
*****************************************************************}
procedure TOF_AMIMMODPI.OnbFermeClick(Sender: TObject);
Var montant : Double;
   question : TmodalResult;
begin

  // Récupération du montant de l'opération
  montant := Valeur(GetControlText('DE_DMTDPI'));
  SetControlText('DE_DMTDPI', StrFMontant(montant,15,V_PGI.OkDecV,'',True));

  // Ctl si la saisie à été modifiée
  IF (montant = FAncienMontant) THEN
    Begin
      // Pas de saisie on sort
      Ecran.ModalResult := mrNo;
      TFVIERGE(ECRAN).retour := 'ECHAP';
    end
  Else
    Begin
     // Action Echap ou Fermer par la croix
     Question := PGIAskCancel(HM[4],Ecran.Caption);
     If Question = mrYes then
        // contrôle validation
        OnUpdate
     else If Question = mrNo then
        Begin
          // Sortir
          Ecran.ModalResult := mrNo ;
          TFVIERGE(ECRAN).retour := 'ECHAP';
        End
     else
        Begin
          // Cancel
          Ecran.ModalResult :=  mrNone;
          TFVIERGE(ECRAN).retour := 'ECHAP';
        End;
    end
end;

{***********A.G.L.***********************************************
Auteur  ...... : TGA
Créé le ...... : 27/04/2006
Modifié le ... :   /  /
Description .. : Test sur Touche F6
Mots clefs ... :
*****************************************************************}
procedure TOF_AMIMMODPI.OnGrilleKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of
    VK_F6 : begin
            // FQ 18949
            //SetControlText('DE_DMTDPI',StrFMontant(FMontantMax,15,V_PGI.OkDecV,'',True));
            if Arrondi (FMontantMax-FBaseAmoEco, V_PGI.OkDecV) > 0 then
               SetControlText('DE_DMTDPI',StrFMontant(FBaseAmoEco,15,V_PGI.OkDecV,'',True))
            else
               SetControlText('DE_DMTDPI',StrFMontant(FMontantMax,15,V_PGI.OkDecV,'',True));
            end;
  end;
end;




Initialization
  registerclasses ( [ TOF_AMIMMODPI ] ) ;
end.


