{***********UNITE*************************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 19/03/2001
Modifié le ... : 29/03/2001
Description .. : Source TOF de la TABLE : ICCPARAMETRE ()
Mots clefs ... : TOF;ICCPARAMETRE
*****************************************************************}
Unit uTOFIccParametre;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     UTOB,
     MaineAGL,             // AGLLanceFICHE
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,              // AGLLanceFICHE
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     Windows,
     Ent1,
     ParamSoc,
     Htb97,                // TtoolBarButton97
     Graphics,
     HFLabel,              // TFlashingLabel
     HPanel,               // THPanel
     HRichOLE,             // THRichEditOLE
     uTOFIccInteret,       // CPLanceFiche_ICCINTERET
     IccGlobale;

procedure CPLanceFiche_ICCPARAMETRE(psz : String);

Type
  TOF_ICCPARAMETRE = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnLoad                   ; override ;
    procedure OnClickBValider          ( Sender : TObject );
    procedure OnKeyDownForm            ( Sender: TObject ; var Key: Word;Shift: TShiftState );
    procedure OnExitZDateDu            ( Sender : TObject );
    procedure OnExitZDateAu            ( Sender : TObject );

    procedure VerificationAvantCalcul  ( vStArgument : String ; D1, D2 : TDateTime  );

  private
    FStParam               : String;
    RichEdit               : THRichEditOLE;
    ZDateDu                : THEdit;
    ZDateAu                : THEdit;

  public

  end ;

Implementation

procedure CPLanceFiche_ICCPARAMETRE(psz : String);
begin
  AGLLanceFiche('CP', 'ICCPARAMETRE', '', '', psz);
end;

procedure TOF_ICCPARAMETRE.OnArgument (S : String ) ;
var stArg : String;
begin
  Inherited;
  { Paramètres passés dans la chaine S
      DATE 1 ( Date de début du calcul )
      DATE 2 ( Date de fin du calcul )
      LISTE DES COMPTES }

  stArg := S;

  { Date de Début du calcul }
  FStParam := ReadTokenSt(stArg);
  SetControlText('ZDATEDU',FStParam);

  {Date de fin du calcul }
  FStParam := ReadTokenSt(stArg);
  SetControlText('ZDATEDU_', FStParam);

  { Listre des comptes sélectionnés }
  FStParam := stArg;

  ZDateDu := THEdit(GetControl('ZDATEDU'));
  ZDateAu := THEdit(GetControl('ZDATEDU_'));

  ZDateDu.OnExit := OnExitZDateDu;
  ZDateAu.OnExit := OnExitZDateAu;

  Ecran.OnKeyDown := OnKeyDownForm;

  TToolBarButton97(GetControl('BVALIDER')).OnClick := OnClickBValider;
  RichEdit := THRichEditOLE(GetControl('RichEdit'));
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCPARAMETRE.OnLoad;
begin
  inherited;
  SetControlText('ZMODE',GetParamSocSecur('SO_ICCNOMBREJOURS',''));

  if GetParamSocSecur('SO_ICCMETHODE','') = 'ICC' then
    THRadioGroup(GetControl('ZMETHODE')).ItemIndex := 0
  else
    THRadioGroup(GetControl('ZMETHODE')).ItemIndex := 1;

  VerificationAvantCalcul( FStParam, StrToDate(ZDateDu.Text), StrToDate(ZDateAu.Text) );

  // OU BIEN
  // Vérification des comptes Icc par rapport à ceux de la compta
  // Date1 := VH^.EnCours.Deb;
  // Date2 := IIF(VH^.Suivant.Code <> '',VH^.Suivant.Fin,VH^.Encours.Fin);
  //  VerificationAvantCalcul( FStParam, Date1, Date2 );

end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCPARAMETRE.OnExitZDateAu(Sender: TObject);
begin
  VerificationAvantCalcul( FStParam, StrToDate(ZDateDu.Text), StrToDate(ZDateAu.Text) );
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCPARAMETRE.OnExitZDateDu(Sender: TObject);
begin
  VerificationAvantCalcul( FStParam, StrToDate(ZDateDu.Text), StrToDate(ZDateAu.Text) );
end;

procedure TOF_ICCPARAMETRE.OnClickBValider(Sender: TObject);
var Methode : Word;
    Mode : String;
    D1,D2 : TDateTime;
begin
  D1 := StrToDate(ZDateDu.Text);
  D2 := StrToDate(ZDateAu.Text);

  if D1 > D2 then
  begin
    PgiInfo('La date de fin du calcul doit être supérieure à la date de début.','Saisie de la date');
    Exit;
  end;

  if not ((D1 >= VH^.Encours.Deb) and (D2 <= VH^.Encours.Fin)) then
    if not ((D1 >= VH^.Suivant.Deb) and (D2 <= VH^.Suivant.Fin)) then
    begin
      PgiInfo('Les dates de calcul doivent être comprises dans l''intervalle de dates de l''un des exercices.','Saisie de la date');
      Exit;
    end;

  Methode := THRadioGRoup(GetControl('ZMETHODE')).ItemIndex;
  Mode := THValComboBox(GetControl('ZMODE')).Text;

  Icc_Data.AvecRecapitulatif := TCheckBox(GetControl('ZRECAPITULATIF')).Checked;

  CPLanceFiche_ICCINTERET(Mode + ';' + IntToStr(Methode) + ';' + DateToStr(D1) + ';' + DateToStr(D2) + ';' + FStParam);
  TToolBarButton97(GetControl('BFERME')).Click;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/10/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCPARAMETRE.OnKeyDownForm(Sender: TObject; var Key: Word;Shift: TShiftState);
begin
  if (Key = VK_F10) or (Key = VK_RETURN) then
    TToolBarButton97(GetControl('BVALIDER')).Click;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/09/2001
Modifié le ... : 21/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCPARAMETRE.OnClose ;
begin
  inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/10/2001
Modifié le ... :   /  /
Description .. : Vérifie les comptes et les éléments nationaux avant les calculs
Mots clefs ... :
*****************************************************************}
procedure TOF_ICCPARAMETRE.VerificationAvantCalcul( vStArgument : String ; D1 , D2 : TDateTime );
var lStCompte : string;
begin
  if vStArgument = '' then Exit;

  Icc_Data.TauxAbsent        := False ;
  Icc_Data.TauxProvisoire    := False ;
  Icc_Data.NouvelleEcriture  := False ;
  Icc_Data.AvecRecapitulatif := False ;

  while vStArgument <> '' do
  begin
    lStCompte := ReadTokenSt( vStArgument );

    { Vérification si de nouvelles écritures ont été saisies dans la COMPTA }
    if not Icc_Data.NouvelleEcriture then
      Icc_Data.NouvelleEcriture := ExisteNouvelleEcriture ( lStCompte , D1 , D2 ) ;
  end;

  { Vérification si les taux sont saisis et de la façon dont ils sont définis }
  CalculTauxMaxLegal( D1 , D2 );

  RichEdit.Clear;
 
  if ICC_Data.SoldeDifferent then
  begin
    RichEdit.lines.Add('Le solde est différent de celui de la comptabilité');
    RichEdit.lines.Add(' ');
  end;

  if ICC_Data.NouvelleEcriture then
  begin
    RichEdit.lines.Add('Des nouvelles écritures n''ont pas été récupérées dans ICC');
    RichEdit.lines.Add(' ');
  end;

  if Icc_Data.TauxAbsent then
  begin
    RichEdit.lines.Add('Les éléments nationaux sont absents pour cette période');
    RichEdit.lines.Add(' ');
  end;

  if Icc_Data.TauxProvisoire then
  begin
    RichEdit.lines.Add('Les éléments nationaux sont définis en "provisoires"');
    RichEdit.lines.Add(' ');
  end;

end;

Initialization
  registerclasses ( [ TOF_ICCPARAMETRE ] ) ;
end.
