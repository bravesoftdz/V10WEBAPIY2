{-------------------------------------------------------------------------------------
    Version    |   Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
                  03/08/03  JP   Migration eAGL
                  13/08/03  JP   Correction de la requête de CompteOnChange
                  14/08/03  JP   Gestion des devises dans MontantOnChange
                  07/05/04  JP   Différentes modifications sur le calcul des soldes
                                 et la génération de champs explicites
 6.50.001.001     24/05/05  JP   Nouvelle gestion du formatage des montant cf. 24/05/05
 7.09.001.001     20/10/06  JP   Filtre sur les comptes en fonctions de dossiers du regroupement
 7.09.001.003     26/12/06  JP   FQ 10387 : Ajout du paramétrage de la date du virement
--------------------------------------------------------------------------------------}
unit TofModifVir ;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, 
  {$ENDIF}
  StdCtrls, Controls, Classes, forms, sysutils, HCtrls, HEnt1, HMsgBox, UTOF, UTOB, HTB97,
  UObjGen;

type
  TOF_MODIFVIR = class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  protected
    Rester     : Boolean; // Gestion d'erreur en sortie
    TypeAction : string;
    CurTob     : TOB;
    edMontant  : THNumEdit;
    SLibelle,
    DLibelle,
    SDevise    : string;
    DDevise    : string;
    Montant,
    SSolde,
    DSolde     : Double; // Soldes d'origine des comptes
    ObjTaux    : TObjDevise;
    PasTouche  : Boolean;
    Nature     : string;
    FDateEqui  : TDateTime; {26/12/06 : FQ 10387}

    procedure CompteOnChange  (Sender : TObject);
    procedure MontantOnChange (Sender : TObject);
    procedure bValiderOnClick (Sender : TObject);
    procedure MajDevises;
    procedure MajCtreValeur;
  end ;

function TRLanceFiche_ModifVir(Dom, Fiche, Range, Lequel, Arguments : string): string;

implementation

uses
  AglInit, Vierge, TofEquilibrage, Commun, UProcSolde, Constantes;


{---------------------------------------------------------------------------------------}
function TRLanceFiche_ModifVir(Dom, Fiche, Range, Lequel, Arguments : string): string;
{---------------------------------------------------------------------------------------}
begin
  Result := AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end ;

{JP 28/11/03 : Nouvelle gestion : on ne gère plus trois devises mais seulement la devise
               du compte source, le solde destinataire est donc converti à la volée.
               En fin d'unité j'ai conservé le source précédent en cas de retour en arrière
{---------------------------------------------------------------------------------------}
procedure TOF_MODIFVIR.OnArgument(S : string ) ;
{---------------------------------------------------------------------------------------}
var
  Compte : THEdit;
  Mnt    : Double;
begin
  inherited;
  Ecran.HelpContext := 150;
  PasTouche := True;
  {Est-on en modification ou en création}
  if ReadTokenSt(S) = 'MOD' then TypeAction := 'M'
                            else TypeAction := 'N';

  {26/12/06 : FQ 10387 : récupération de la date d'équilibrage}
  FDateEqui := StrToDate(ReadTokenSt(S));

  {On récupère le filtre sur la nature de écritures}
  Nature := ReadTokenSt(S);

  //CurTob := TheTOB;
  CurTob := TOB(TheData);
  Compte := THEdit(GetControl('SCOMPTE'));
  edMontant := THNumEdit(GetControl('MONTANT'));

  {Création de l'objet qui va servir à la conversion des soldes dans la devise d'affichage}
  ObjTaux := TObjDevise.Create(FDateEqui);

  {Création d'un nouveau virement}
  if TypeAction = 'N' then begin
    Ecran.Caption := 'Création de virement';
    UpdateCaption(Ecran);
    Compte.OnChange := CompteOnChange;

    Compte := THEdit(GetControl('DCOMPTE'));
    Compte.OnChange := CompteOnChange;
    Montant := 0.00;
    edMontant.Value := Montant;
  end

  {Modification d'un virement existant : on ne peut changer que le montant}
  else begin
    Ecran.Caption := 'Modification de virement';
    UpdateCaption(Ecran);
    {Gestion des zones source}
    Montant := CurTob.GetValue('MONTANT');
    Compte.Text := CurTob.GetValue('SCOMPTE');
    Compte.Enabled := False;
    SSolde := CurTob.GetValue('SSOLDE');
    SDevise := CurTob.GetValue('SDEVISE');
    THNumEdit(GetControl('SSOLDE')).Value := SSolde;
    {20/10/06 : On se limite au comptes des dossiers du regroupement Tréso}
    Compte.Plus := FiltreBanqueCp(Compte.DataType, '', '');

    {Gestion des zones destination}
    Compte := THEdit(GetControl('DCOMPTE'));
    Compte.Text := CurTob.GetValue('DCOMPTE');
    Compte.Enabled := False;
    {20/10/06 : On se limite au comptes des dossiers du regroupement Tréso}
    Compte.Plus := FiltreBanqueCp(Compte.DataType, '', '');

    Mnt := CurTob.GetValue('DSOLDE');
    DDevise := CurTob.GetValue('DDEVISE');
    {Conversion du solde de la destination dans la devise du source}
    ObjTaux.ConvertitMnt(Mnt, DSolde, DDevise, GetControlText('SCOMPTE'));

    {On met le montant au format de la devise source}
    THNumEdit(GetControl('DSOLDE')).Decimals := ObjTaux.GetNbDecimalesFromDev(SDevise);{24/05/05}
    THNumEdit(GetControl('DSOLDE')).Value := DSolde;
    edMontant.Decimals := ObjTaux.GetNbDecimalesFromDev(SDevise);
    edMontant.Value := Montant;
  end;

  {Affichages des drapeaux et des libellés Devise}
  MajDevises;

  {Pour la mise à jour des soldes après changement du montant}
  edMontant.OnChange := MontantOnChange;

  TToolBarButton97(GetControl('BVALIDER')).OnClick := bValiderOnClick;
  (GetControl('SCOMPTE') as THEdit).Plus := FiltreBanqueCp((GetControl('SCOMPTE') as THEdit).DataType, '', '');
  (GetControl('DCOMPTE') as THEdit).Plus := FiltreBanqueCp((GetControl('DCOMPTE') as THEdit).DataType, '', '');

  Rester := False;
  PasTouche := False;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_MODIFVIR.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if Rester then begin
    Rester := False;
    LastError := -1;
  end
  else
    FreeAndNil(ObjTaux);
end ;

{Affichage du solde et la devise du compte choisi (En création de virement)
{---------------------------------------------------------------------------------------}
procedure TOF_MODIFVIR.CompteOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  CptName,
  Compte : string;
  Solde  : Double;
  MD     : Double;
  aTob   : TOB;
begin
  {Remarque : seul les virements en création permettent un changement des comptes }
  CptName := THEdit(Sender).Name;
  {Récupération du compte général}
  Compte := GetControlText(CptName);

  {07/05/04 : pour réinitialiser les soldes}
  edMontant.Value := 0;

  {Si on ne trouve pas le compte, c'est qu'on est en train de changer le compte,
   mais que ce n'est pas fini}
  if ObjTaux.Corresp.IndexOf(Compte) = - 1 then Exit;
  {Pour ne pas exécuter MontantOnChange}
  PasTouche := True;

  aTob := CurTob.FindFirst(['SCOMPTE'], [Compte], False);

  {Si le solde a déjà été calculé, car le compte choisi est un compte source
   dans les propositions de virement}
  if aTob <> nil then
    Solde := aTob.GetValue('SSOLDE')
  else begin
    {Sinon on regarde s'il s'agit d'un compte destination pour récupérer le solde}
    aTob := CurTob.FindFirst(['DCOMPTE'], [Compte], False);
    if aTob <> nil then
      Solde := aTob.GetValue('DSOLDE')
    else
      {Récupération du solde du Sender. +1 Car GetSoldeInit fonctionne avec une inégalité stricte
       26/12/06 : FQ 10387 : paramétrage de la date d'opération}
      Solde := GetSoldeInit(Compte, DateToStr(FDateEqui + 1), Nature, True);
  end;

  Montant := edMontant.Value;
  {On vient de modifier le compte général source}
  if CptName[1] = 'S' then begin
    SLibelle := RechDom('TRBANQUECP', Compte, False);
    SSolde := Solde;
    {On convertit tous les montants dans la même devise}
    MajCtreValeur;
    {On récupère la devise du compte}
    SDevise := ObjTaux.GetDeviseCpt(Compte);
    {On met à jour l'affichage des devises}
    MajDevises;
    Solde := SSolde - Montant;
    {On met le montant au format de la devise source}
    THNumEdit(GetControl('SSOLDE')).Decimals := ObjTaux.GetNbDecimalesFromDev(SDevise);{24/05/05}
    THNumEdit(GetControl('SSOLDE')).Value := Solde;
  end

  {On vient de modifier le compte général destination}
  else begin
    DLibelle := RechDom('TRBANQUECP', Compte, False);
    DDevise := ObjTaux.GetDeviseCpt(Compte);
    MD := Solde;
    {Conversion du solde de la destination dans la devise du source}
    ObjTaux.ConvertitMnt(MD, DSolde, DDevise, FloatToStr(Valeur(GetControlText('SCOMPTE'))));
    Solde := DSolde + Montant;
    {On met le montant au format de la devise source}
    THNumEdit(GetControl('DSOLDE')).Decimals := ObjTaux.GetNbDecimalesFromDev(SDevise);{24/05/05}
    THNumEdit(GetControl('DSOLDE')).Value := Solde;
  end;
  PasTouche := False;
end;

{Ajuste les soldes selon le virement
{---------------------------------------------------------------------------------------}
procedure TOF_MODIFVIR.MontantOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Mnt : Double;
  Sld : Double;
begin
  {Si à True, c'est qu'on est train de convertir les montants => on ne réaffecte pas les soldes}
  if PasTouche then Exit;

  Mnt := Valeur(edMontant.Text);
  Sld := SSolde + Montant - Mnt;
  {On met le montant au format de la devise source}
  THNumEdit(GetControl('SSOLDE')).Decimals := ObjTaux.GetNbDecimalesFromDev(SDevise);{24/05/05}
  THNumEdit(GetControl('SSOLDE')).Value := Sld;

  Sld := DSolde - Montant + Mnt;
  {On met le montant au format de la devise source}
  THNumEdit(GetControl('DSOLDE')).Decimals := ObjTaux.GetNbDecimalesFromDev(SDevise);{24/05/05}
  THNumEdit(GetControl('DSOLDE')).Value := Sld;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_MODIFVIR.bValiderOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  SCompte,
  DCompte : string;
  Tob1Vir : TOB;
  Mnt     : Double;
  MD, ME  : Double;
begin
  Mnt := Valeur(edMontant.Text);

  {Le montant d'un virement doit être une valeur positive}
  if Mnt <= 0 then begin
    TrShowMessage(Ecran.Caption, 10, '', '');
    edMontant.SetFocus;
    Rester := True;
    Exit;
  end;

  SCompte := GetControlText('SCOMPTE');
  DCompte := GetControlText('DCOMPTE');
  if TypeAction = 'N' then begin
    {Le compte d''origine et celui de destination ne peuvent être identiques}
    if SCompte = DCompte then begin
      TrShowMessage(Ecran.Caption, 11, '', '');
      Rester := True;
      Exit;
    end
    else if Trim(SCompte) = '' then begin
      HShowMessage('1;Virements;Veuillez renseigner le compte d''origine !;W;O;O;O;', '', '');
      Rester := True;
      Exit;
    end
    else if Trim(DCompte) = '' then begin
      HShowMessage('0;Virements;Veuillez renseigner le compte de destination !;W;O;O;O;', '', '');
      Rester := True;
      Exit;
    end;

    {Un virement entre ces deux comptes existe déjà}
    if ((CurTob.FindFirst(['SCOMPTE'], [SCompte], False) <> nil) and
        (CurTob.FindFirst(['DCOMPTE'], [DCompte], False) <> nil)) or

       ((CurTob.FindFirst(['SCOMPTE'], [DCompte], False) <> nil) and
        (CurTob.FindFirst(['DCOMPTE'], [SCompte], False) <> nil)) then begin
      TrShowMessage(Ecran.Caption, 12, '', '');
      Rester := True;
      Exit;
    end;

    Tob1Vir := TOB.Create('1', CurTob, -1);
    Tob1Vir.AddChampSupValeur('SCOMPTE', SCompte);
    Tob1Vir.AddChampSupValeur('SDEVISE', SDevise);
    Tob1Vir.AddChampSupValeur('MONTANT', Mnt); // Devise ?
    Tob1Vir.AddChampSupValeur('DCOMPTE', DCompte);
    Tob1Vir.AddChampSupValeur('DDEVISE', DDevise);
    Tob1Vir.AddChampSup('SSOLDE', False); // Les soldes seront écrit pas MajSolde !
    Tob1Vir.AddChampSup('DSOLDE', False);
    Tob1Vir.AddChampSup('SOURCE', False);
    {07/05/04 : Ajout de ces quatre champs pour rendre le TobViewer plus lisible}
    Tob1Vir.AddChampSup('DESTINATION', False);
    Tob1Vir.AddChampSup('VIREMENT', False);
    Tob1Vir.AddChampSupValeur('SLIBELLE', SLibelle);
    Tob1Vir.AddChampSupValeur('DLIBELLE', DLibelle);

    Tob1Vir.AddChampSupValeur('DDOSSIER', GetDossierFromBQCP(DCompte));
    Tob1Vir.AddChampSupValeur('SDOSSIER', GetDossierFromBQCP(SCompte));
    Tob1Vir.AddChampSupValeur('DDOSSLIB', GetLibDossierFromBQ(DCompte));
    Tob1Vir.AddChampSupValeur('SDOSSLIB', GetLibDossierFromBQ(SCompte));
    Tob1Vir.AddChampSupValeur('DATEEQUI', FDateEqui);{26/12/06 : FQ 10387}
  end

  {On est en mofification de virement}
  else begin
    {Les soldes et source seront écrit pas MajSolde !}
    CurTob.PutValue('MONTANT', Mnt);
    {Racine pour MajSolde}
    CurTob := CurTob.Parent;
  end;

  {On reconvertit le montant dans la devise du compte destination}
  ME := Valeur(GetControlText('DSOLDE'));
  ObjTaux.ConvertitMnt(ME, MD, SDevise, DCompte);
  TOF_EQUILIBRAGE.MajSolde(CurTob, DCompte, MD);
  {07/05/04 : Inversion de l'ordre (DSOLDE puis SSOLDE) car MajSolde(SSOLDE) exécute le PutSource,
              if est préférable que DSOLDE soit à jour !!!}
  TOF_EQUILIBRAGE.MajSolde(CurTob, SCompte, Valeur(GetControlText('SSOLDE')));

  TFVierge(Ecran).Retour := 'X'; // MaJ écran à faire
end;

{---------------------------------------------------------------------------------------}
procedure TOF_MODIFVIR.MajDevises;
{---------------------------------------------------------------------------------------}
begin
  {Maj des infos devise
   24/05/05 : Nouvelle gestion de l'affichage des devises}
  MajAffichageDevise(GetControl('IDEV' ), THLabel(GetControl('SDEVISE')), SDevise, sd_Aucun);
  MajAffichageDevise(GetControl('IDEV1'), THLabel(GetControl('DDEVISE')), SDevise, sd_Aucun);
  MajAffichageDevise(GetControl('IDEV2'), THLabel(GetControl('MDEVISE')), SDevise, sd_Aucun);
end;

{Conversion du solde de destination et du montant dans la devise du nouveau compte source
{---------------------------------------------------------------------------------------}
procedure TOF_MODIFVIR.MajCtreValeur;
{---------------------------------------------------------------------------------------}
var
  MD : Double;
begin
  if DSolde <> 0 then begin
    {Conversion du solde de la destination dans la devise du source}
    ObjTaux.ConvertitMnt(DSolde, MD, DDevise, GetControlText('SCOMPTE'));
    {On met le montant au format de la devise source}
    THNumEdit(GetControl('SSOLDE')).Decimals := ObjTaux.GetNbDecimalesFromDev(SDevise);{24/05/05}
    THNumEdit(GetControl('SSOLDE')).Value := MD;
  end;

  if Montant <> 0 then begin
    {Conversion du montant saisie dans la devise du compte source}
    ObjTaux.ConvertitMnt(Montant, MD, SDevise, GetControlText('SCOMPTE'));
    {On met le montant au format de la devise source}
    THNumEdit(GetControl('MONTANT')).Decimals := ObjTaux.GetNbDecimalesFromDev(SDevise); {24/05/05}
    THNumEdit(GetControl('MONTANT')).Value := MD;
    Montant := MD;
  end;
end;

initialization
  registerclasses ( [ TOF_MODIFVIR ] ) ;

end.

