{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/01/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPRESULTATFISCAL ()
Mots clefs ... : TOF;CPRESULTATFISCAL
*****************************************************************}
Unit uTofCPRESULTATFISCAL ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     Fe_Main,       //Agllancefiche
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     MainEAgl,      //AglLancefiche
     eMul,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     hpanel,
     Vierge,
     Graphics,        //[FsBold]...
     Grids,           //propriétés grille
     Windows,         //Formkeydown
     HTB97,           //TToolBarButton97
     HrichOLE,        //THRichEditOle
     M3FP,            // registerAGLproc, registerAGLfunc
     ent1,            // VH^
     OutilsDRF,
     UTOF,
     galOutil;        // GetGuidPer

Type
  TOF_CPRESULTATFISCAL = Class (TOF)
  private
    TobDRF              : Tob;                  //les données de la fiche
    LaTobDRF            : Tob;                  //la table CPRESULTATFISCAL
    TobCalcul           : Tob;                  //La tob des règles de calcul
    TobLien             : Tob;                  //La tob des lien pour accéder aux onglets
    TobSimpl            : Tob;                  //La tob des concordances pour régime simplifié
    bModif              : Boolean;              //TObDrf modifiée ?

    procedure EnregMemoDRF;
    procedure ChargeFicheDRF;
    procedure AffectationBtOutil;
    procedure RechercheCalcul;
    procedure RechercheLien;
    procedure RechercheSimpl;
    procedure TraiteSimpl;
    procedure RecalculFiche;
    function IsConditionGood(LaCondition : String) : boolean;
    function LeCalcul (StrCalcul : String) : Extended;
    procedure GetDetailCalcul(var StrCalcul : String ; var OpeActu : String ; var LeChamp : String);
    Function RechPrecExe (ExoPrec : String) : Tob;
    procedure EventRecalcul (Sender : Tobject);
    procedure RemiseEnForme (Sender : Tobject);
    function AjouteVersTob (LaNature : String; LeComm : String; Mt1 : Extended; LIndice : integer = 0; Mt2 : Extended = 0; Mt3 : Extended = 0; Mt4 : Extended = 0; Mt5 : Extended = 0) : boolean;
    function GetChampResultat( LeChamp : string ; Depart : integer = 0) : string ; // BVE 14.09.07
    procedure EffaceDansTob (LaTob : Tob ; LaNature : String ; QueTotal : boolean = False);
    procedure BvaliderOnClick (Sender : Tobject);
    procedure LeditOnKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure BretourOnClick (Sender : Tobject);
    procedure LeLabelOnClick (Sender : Tobject);
    procedure LeLabelOnMouseEnter (Sender : Tobject);
    procedure LeLabelOnMouseLeave(Sender: Tobject);
    procedure FormOnKeyDown (Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure SortieSansSauvegarde;
    procedure LaSortie (Sender : Tobject);
    procedure RefreshLaGrille;

    function RemunerationExploitant : Extended;
    function TaxeVoitureParticuliere : Extended;
    function ImpotSurSociete : Extended;
    procedure Calcul_2058A_WI_MOA;                                    // on recalcul 2058A_WI_MOA et 2058A_WU_MOA

//    procedure BOutil_Onclick (Sender : Tobject);
//    procedure AppelOutil (LaNature : String ; LaProc : TNotifyEvent ; LeTitre : String = '' ; LaBorne1 : extended = 0 ; LaBorne2 : extended = 0 ; LeTag : integer = 0);
    procedure B2058A_XL_MOA_OnClick (Sender : Tobject);               // suivi des déficits
    procedure B2058A_WC_MOA_OnClick (Sender : Tobject);               // rémunération du conjoint de l'exploitant
    procedure B2058A_XA_MOA_OnClick (Sender : Tobject);               // régime des sociétés mères et des filiales
    procedure B2058B_CA_MOA_OnClick (Sender : Tobject);               // provisions pour risques et charges
    procedure B2058B_CD_MOA_OnClick (Sender : Tobject);               // provisions pour dépréciation
    procedure B2058B_CG_MOA_OnClick (Sender : Tobject);               // Charges à payer
    procedure B2058B_ZV_MOA_OnClick (Sender : Tobject);               // Indemnités congès à payer
    procedure B2058A_WN_MOA_OnClick (Sender : Tobject);               // suivi de PVCT différés
    procedure B2058A_WM_MOA_OnClick (Sender : Tobject);               // suivi des MVLT
    procedure B2058A_VA_MOA_OnClick (Sender : Tobject);               // imposition des PVLT
    procedure B2058AB_AH_MOA_OnClick (Sender : Tobject);              // suivi des déficits (intégration fiscale)
    procedure B2058AB_AG_MOA_OnClick (Sender : Tobject);              // suivi des MVLT (intégration fiscale)
    procedure B2058AB_AF_MOA_OnClick (Sender : Tobject);              // imposition des PVLT (intégration fiscale)
    procedure BJetonPresence_OnClick (Sender : Tobject);              // Jeton de présence
    procedure BCessionImo_OnCLick (Sender : Tobject);                      // Cession d'immobilisation
    procedure BDetail_OnClick (Sender : Tobject);                     // saisie en grille

    procedure BDetail_bgValideOnClick (Sender : Tobject);
    procedure B2058A_XL_MOA_bgValideOnClick (Sender : Tobject);
    procedure B2058A_WC_MOA_bgValideOnClick (Sender : Tobject);
    procedure B2058A_XA_MOA_bgValideOnClick (Sender : Tobject);
    procedure B2058B_CA_MOA_bgValideOnClick (Sender : Tobject);
    procedure B2058B_CD_MOA_bgValideOnClick (Sender : Tobject);
    procedure B2058B_CG_MOA_bgValideOnClick (Sender : Tobject);
    procedure B2058B_ZV_MOA_bgValideOnClick (Sender : Tobject);
    procedure B2058A_WN_MOA_bgValideOnClick (Sender : Tobject);
    procedure B2058A_WM_MOA_bgValideOnClick (Sender : Tobject);
    procedure B2058A_VA_MOA_bgValideOnclick (Sender : Tobject);
    procedure B2058AB_AH_MOA_bgValideOnclick (Sender : Tobject);
    procedure B2058AB_AG_MOA_bgValideOnclick (Sender : Tobject);
    procedure B2058Ab_AF_MOA_bgValideOnclick (Sender : Tobject);
    procedure BCessionImo_bgValideOnclick (Sender : Tobject);
    procedure BJetonPresence_bgValideOnclick (Sender : Tobject);

    procedure BDetail_Valider;
    procedure B2058A_XL_MOA_Valider;
    procedure B2058A_WC_MOA_Valider;
    procedure B2058A_XA_MOA_Valider;
    procedure B2058B_CA_MOA_Valider;
    procedure B2058B_CD_MOA_Valider;
    procedure B2058B_CG_MOA_Valider;
    procedure B2058B_ZV_MOA_Valider;
    procedure B2058A_WN_MOA_Valider;
    procedure B2058A_WM_MOA_Valider;
    procedure B2058A_VA_MOA_Valider;
    procedure B2058AB_AH_MOA_Valider;
    procedure B2058AB_AG_MOA_Valider;
    procedure B2058AB_AF_MOA_Valider;
    procedure BCessionImo_Valider;
    procedure BJetonPresence_Valider;

    procedure HPBOnClose (Sender : Tobject);
    procedure HPBFerme;


  public
    BValider            : TToolbarButton97;     //bouton Valider
    BFerme              : TToolbarButton97;     //bouton Fermer
    Rg                  : THEdit;               //régime fiscal en cours
    IntegFisc           : THEdit;               //nom de l'onglet intégration fiscale
    TSHFisc             : TTabSheet;            //onglet intégration fiscale
    TSH1                : TTabSheet;            //premier onglet
    TSHDEV              : TTabSheet;            //Onglet stockage infos
    TSHSYN              : TTabSheet;            //onglet de synthese
    LaPage              : TPageControl;         //La page control
    HPB                 : TDRFToolBox;          //fenêtre pour les multi saisie et outils
    bHPB                : boolean;              //fenêtre en cours
    tHPB                : String;               //Zone en cours d'utilisation de la fenêtre
    BRetour             : TToolbarButton97;     //Bouton retour à l'onglet de synthèse
    BgValide            : TToolbarButton97;     //bouton valider de la saisie en grille
    BgAnnule            : TToolbarButton97;     //bouton annule de la saisie en grille
    BgAjouteLg          : TToolbarButton97;     //bouton ajoute une ligne de la saisie en grille
    BgSupprimeLg        : TToolbarButton97;     //Bouton supprime une ligne de la saisie en grille
    RemConjointMin      : Extended;             //rémunération mini pour conjoint
    RemConjointMax      : Extended;             //rémunération max pour conjoint
    BSimpl              : boolean;              // Régime simplifié ?

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnArgument (Argument : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  end ;

  procedure LanceFiche_DRF(Argument : String);
  function OkLanceDRF  : String;
  function AGLChercheResultatComptable (parms: array of variant; nb: integer) : Variant;
  function AGLChercheRegimeFiscal (parms: array of variant; nb: integer) : Variant;
  Procedure AGLRendInvisible (parms: array of variant; nb: integer);


Implementation

uses DateUtils, TntWideStrings, TntComCtrls, TntGrids, TntStdCtrls;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/02/2007
Modifié le ... : 22/02/2007
Description .. : Contrôle l'existence de la fiche millésimée DRF
Suite ........ : et retourne l'année de la fiche la plus récente
Mots clefs ... : 
*****************************************************************}
function OkLanceDRF : String;
var
  lstAnnee              : String;
  lstSql                : String;
  ltobFiche             : Tob;
  lTob                  : Tob;
  lstRegleFisc          : string;


begin
  Result                := '';
  lstAnnee              := FormatDateTime('yyyy', VH^.EnCours.Deb);
  lstRegleFisc          := GetColonneSQL('DPFISCAL','DFI_REGLEFISC','DFI_GUIDPER="'+GetGuidPer(V_PGI.NoDossier)+'"');
  if lstRegleFisc = 'BIP' then lstRegleFisc := 'BIC';
  lstSql                := 'SELECT DFM_FORME, Substring(DFM_FORME,13,len(DFM_FORME)) ANNEE  FROM FORMES ' +
                           'WHERE DFM_FORME LIKE "CPRFISCAL'+lStRegleFisc+'%" ORDER BY ANNEE DESC';

  lTobFiche             := Tob.Create ('FICHES', nil, -1);
  try
    lTobFiche.LoaddetailFromSql(lStsql);
    if lTobFiche.Detail.Count = 0 then
      Exit;

    // on test l'année en cours
    lTob                := lTobFiche.findFirst(['ANNEE'], [lStAnnee], False);
    // si pas de résultat on recherche l'année précédente, sinon Rien
    if lTob = nil then
    begin
      lstAnnee          := IntToStr(VALEURI(lstAnnee) -1);
      lTob              := lTobFiche.findFirst(['ANNEE'], [lStAnnee], False);
    end;

    if lTob <> nil then
      Result            := lTob.GetString('ANNEE');

  finally
    FreeAndNil(ltobFiche);
  end;  
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 10/01/2007
Modifié le ... :   /  /
Description .. : procedure d'appel de la fiche
Mots clefs ... :
*****************************************************************}
procedure LanceFiche_DRF(Argument : String);
var
  LaFiche               : String;

begin

  LaFiche               := 'CPRFISCALBIC' + Argument;
  AGLLanceFiche('CP', LaFiche, '', '', '');

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 10/01/2007
Modifié le ... :   /  /
Description .. : paramètrages à l'ouverture de la fiche
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.OnArgument (Argument : String ) ;
var
  StrSql                : String;
  StrCaption            : String;
  Ledit                 : THEdit;
  TobV                  : Tob;
  Q                     : TQuery;




begin
  inherited;
  //les contrôles
  BFerme                := TToolbarButton97(GetControl('BFERME'));
  BValider              := TToolbarButton97(GetControl('BVALIDER'));
  Rg                    := THEdit(GetControl('RGFSC'));                         //contient la valeur du régime
  IntegFisc             := THEdit(GetControl('INTEGFISC'));                     //contient l'onglet pour intégration fiscale
  LaPage                := TPageControl(GetControl('FE_PC'));                   //La page control
  TSH1                  := TTabSheet(GetControl('FE_TSH1'));                    //premier onglet
  TSHDEV                := TTabSheet(GetControl('FE_DEV'));                     //onglet pour dev
  TSHFisc               := TTabSheet(GetControl(IntegFisc.Text));               //l'onglet de l'intégration fiscale
  BgValide              := TToolbarButton97(GetControl('BGVALIDE'));
  BgAnnule              := TToolbarButton97(GetControl('BGANNULE'));
  BgAjouteLg            := TToolbarButton97(GetControl('BGAJOUTELG'));
  BgSupprimeLg          := TToolbarButton97(GetControl('BGSUPPRIMELG'));
  BRetour               := TToolbarButton97(GetControl('BRETOUR'));

  BSimpl                := False;         //par défaut, régime normal (normal)
  bModif                := False;         //par défaut, pas de modif
//
  RechercheCalcul;
  RechercheLien;

  //les propriétés

  Ecran.OnKeyDown  := FormOnKeyDown;

  if not Assigned(TSHDEV) then
  begin
    PGIInfo('La fiche pour détermination de résultat fiscal n''est pas valide');
    TForm(Ecran).Close;
  end;
  TSHDEV.TabVisible     := False;                       //on cache l'onglet pour dev
  if V_PGI.SAV then
    TSHDEV.TabVisible   := True;


//  LaPage.ActivePage     := TSH1;                        // on se positionne sur le premier onglet
  BFerme.OnClick        := LaSortie;
  BValider.OnClick      := BvaliderOnClick;


  if Assigned(BRetour) then
  begin
    Ledit               := THEdit(GetControl('ECRANSYNTHESE'));
    if Assigned(Ledit) then
      TSHSYN            := TTabSheet(GetControl(Ledit.Text));
    BRetour.Visible     := False;                         // on le rend invisible pour le moment, car il ne sert vraiment à grand chose
  end;

  if Assigned(TSHSYN) then
  begin
    LaPage.ActivePage   := TSHSYN;
    BRetour.OnClick   := BretourOnClick;
  end
  else
    LaPage.ActivePage   := TSH1;

 //Chargement des données
  LaTobDRF              := TOB.Create('CREVBLOCNOTE', nil, -1);
  StrSql                := 'SELECT * FROM CREVBLOCNOTE WHERE CBN_EXERCICE="' + VH^.EnCours.Code + '" AND CBN_NATURE="DRF" AND CBN_CODE="BIC"';
  Q                     := OpenSQL(StrSql, True);
  LaTobDRF.SelectDB('', Q);
  Ferme(Q);
  if LaTobDRF.GetString('CBN_EXERCICE') = '' then
    LaTobDRF.SetString('CBN_EXERCICE', VH^.EnCours.Code);

  TobDRF := TOB.Create('DATA_DRF', nil, -1);
  ChargeTobDRF(TobDRF, VH^.EnCours.Code);
  // si nil alors 2 solutions :
  //  - il y a un exercice précédent avec une DRF, alors on recopie les datas avec certainnes valeurs à zéro et d'autres "glissées"
  //  - il n'y a pas d'antécédent alors on part de rien.......
  if not Assigned(TobDRF) then
    TobDRF              := RechPrecExe(VH^.Precedent.Code);
  if not Assigned(TobDRF) then
    TobDRF              := Tob.Create('DATA_DRF', nil, -1);

  ChargeFicheDRF;

  StrCaption            := 'Détermination du résultat fiscal au ' + DateToText(VH^.EnCours.Fin);


  //paramètres
  if Assigned(Rg) then
  begin
    if Copy(Rg.Text, 2, 1) = 'N' then
      StrCaption          := StrCaption + ' - Régime du réel normal';
    if Copy(Rg.Text, 2, 1) = 'S' then
    begin
      StrCaption          := StrCaption + ' - Régime du réel Simplifié';
      BSimpl              := True;
      RechercheSimpl;
    end;
    if Copy(Rg.Text, 3, 1) = 'I' then
    begin
      StrCaption          := StrCaption + ' avec intégration fiscale';
      if TSHFisc <> nil then                    // si intégration fiscale
        TSHFisc.TabVisible  := True;
    end
    else
    begin
      if TSHFisc <> nil then
        TSHFisc.TabVisible  := False;
    end;

  end;

  TForm(Ecran).Caption  := StrCaption;
  UpdateCaption(TForm(Ecran));

  bHPB                  := False;               //par défaut, la fenêtre outil est fermée
  tHPB                  := '';


  // gestion des appels à la compta
  if Assigned(TobDRF) then
  begin
    TobV                  := TobDRF.FindFirst(['NATURE', 'INDICE'], ['2058A_WG_MOA', 1], False);
    if (TobV = nil) and (Copy(Rg.Text, 1, 1) = 'S') then          // pas de détail pour Taxe sur voiture particulière
    begin
      Ledit               := THEdit(GetControl('P2058A_WG_MOA'));
      if Assigned(Ledit) then
        Ledit.Text        := STRFMONTANT(TaxeVoitureParticuliere, 15, V_PGI.OkDecV, '', True);
    end;

    TobV                  := TobDRF.FindFirst(['NATURE', 'INDICE'], ['2058A_WK_MOA', 1], False);
    if (TobV = nil) and (Copy(Rg.Text, 1, 1) = 'S') then          // pas de détail pour Taxe sur voiture particulière
    begin
      Ledit               := THEdit(GetControl('P2058A_WK_MOA'));
      if Assigned(Ledit) then
        Ledit.Text        := STRFMONTANT(ImpotSurSociete, 15, V_PGI.OkDecV, '', True);
    end;
  end;

  //recherche des bornes pour la rémunération du conjoint
  Ledit                 := THEdit(GetControl('REMCONJOINTMIN'));
  if Assigned(Ledit) then
    RemConjointMin      := VALEUR(Ledit.Text);
  Ledit                 := THEdit(GetControl('REMCONJOINTMAX'));
  if Assigned(Ledit) then
    RemConjointMax      := VALEUR(Ledit.Text);


  RecalculFiche;
  AffectationBtOutil;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 27/02/2007
Modifié le ... :   /  /    
Description .. : Gestion des touches de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.FormOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  Objet    : TObject;
  TraiteF5 : Boolean;
  Name     : String;
  Bouton   : TToolBarButton97;
begin
  { FQ 20531 BVE 28.09.07 }
  Bouton := nil;
  Objet  := Screen.ActiveControl ;
  if Objet is THEdit then
  begin
     TraiteF5 := true;
     Name     := System.copy(THEdit(Objet).Name,2,Length(THEdit(Objet).Name) - 1);
     Bouton   := TToolBarbutton97(GetControl('B' + Name));
  end
  else
     TraiteF5 := false;
  { END FQ 20531 }
  case Key of
    VK_F10 :
      begin
          Key           := 0;
          BvaliderOnClick(nil);
      end;

    VK_ESCAPE :
      begin
        if Shift <> [] then
          exit;
        if bModif then
          SortieSansSauvegarde;
      end;

    VK_F4 :
      begin
        if Shift = [ssAlt] then
        begin
          Key           := 0;
          SortieSansSauvegarde;
        end;
      end;               
   { FQ 20531 BVE 28.09.07 }
    VK_F5 :
      begin
         if TraiteF5 and (Bouton <> nil) then
            if assigned(Bouton.OnClick) then
               Bouton.OnClick(Bouton);
         Key := 0;
      end;               
    { END FQ 20531 }
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 27/02/2007
Modifié le ... :   /  /    
Description .. : avertissement sur sortie
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.LaSortie(Sender: Tobject);
begin
  SortieSansSauvegarde;
end;

procedure TOF_CPRESULTATFISCAL.SortieSansSauvegarde;
var
  Rep                   : integer;

begin
  Rep                   := mrNo;

  if bModif then
    Rep                 := PGIAskCancel('Des modifications ont été faites.#13#10Voulez-vous enregistrer ces modifications?', TForm(Ecran).caption);

  case Rep of
    mrNo                : TForm(Ecran).Close;
    mrCancel            : Exit;
    mrYes               : BvaliderOnClick(nil);
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 02/02/2007
Modifié le ... :   /  /
Description .. : recherche dans la fiche les liens pour switcher d'un onglet à
Suite ........ : l'autre
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.RechercheLien;
var
  LeTag                 : integer;
  LeChamp               : String;
  LaClass               : String;
  LeStr                 : String;
  Ledit                 : THEdit;
  NbC                   : integer;
  i                     : integer;
  LeLabel               : THLabel;
  TobL                  : Tob;

begin
  TobLien               := Tob.Create('Les Liens', nil, -1);

  NbC                   := TForm(Ecran).ComponentCount -1;
  for i := 0 to NbC do
  begin
    LeTag               := TForm(Ecran).Components[i].Tag;
    LeChamp             := TForm(Ecran).Components[i].Name;
    LaClass             := TForm(Ecran).Components[i].ClassName;

    if (LeTag = 7) and (LaClass = 'THLabel') then
    begin
      LeLabel           := THLabel(GetControl(LeChamp));
      LeLabel.OnClick   := LeLabelOnClick;
      LeLabel.OnMouseEnter  := LeLabelOnMouseEnter;
      LeLabel.OnMouseLeave  := LeLabelOnMouseLeave;
    end;

    if (LeTag = 7) and (LaClass = 'THEdit') then
    begin
      Ledit             := THEdit(GetControl(LeChamp));
      LeStr             := Ledit.Text;
      TobL              := Tob.Create('link on', TobLien, -1);
      TobL.AddChampSupValeur('LETAG', StrToInt(READTOKENST(LeStr)));
      TobL.AddChampSupValeur('LONGLET', READTOKENST(LeStr));
    end;

  end;

  TobLien.Detail.Sort('LETAG');


end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 07/02/2007
Modifié le ... :   /  /    
Description .. : recherche des concordances entre une déclaration au 
Suite ........ : régime normal et régime simplifié
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.RechercheSimpl;
var
  LeTag                 : integer;
  Ledit                 : THEdit;
  LeChamp               : String;
  NBc                   : integer;
  LaClass               : String;
  LeStr                 : String;
  StrNat                : String;
  StrCalcul             : String;
  NumCal                : integer;
  TobC                  : Tob;
  i                     : integer;

begin
  TobSimpl              := Tob.Create('Table concordance regime simplifie', nil, -1);

  NbC                   := TForm(Ecran).ComponentCount -1;
  for i := 0 to NbC do
  begin
    LeTag               := TForm(Ecran).Components[i].Tag;
    LeChamp             := TForm(Ecran).Components[i].Name;
    LaClass             := TForm(Ecran).Components[i].ClassName;

    if (LeTag = 8) and (Copy(LeChamp, 1, 1) = 'Z') and (UpperCase_(LaClass) = 'THEDIT') then
    begin
      Ledit             := THEdit(GetControl(LeChamp));
      if not Assigned(Ledit) then
        Continue;

      LeStr             := Ledit.Text;
      TobC              := Tob.Create('Regles', TobSimpl, -1);
      NumCal            := VALEURI(READTOKENST(LeStr));
      StrNat            := READTOKENST(LeStr);
      StrCalcul         := READTOKENST(LeStr);
      TobC.AddChampSupValeur('ORDRE', NumCal);
      TobC.AddChampSupValeur('NATURE', StrNat);
      TobC.AddChampSupValeur('CALCUL', StrCalcul);
    end;


  end;
end;


procedure TOF_CPRESULTATFISCAL.LeLabelOnClick(Sender: Tobject);
var
  TobL                  : Tob;
  Longlet               : THTabSheet;
  LeChamp               : String;
  LeLabel               : THLabel;
  LeTag2                : integer;

begin
  with Sender as THLabel do
    LeChamp             := Name;

  LeLabel               := THLabel(GetControl(LeChamp));
  LeTag2                := LeLabel.Tag2;
  TobL                  := TobLien.FindFirst(['LETAG'], [LeTag2], False);
  if Assigned(TobL) then
  begin
    LeChamp             := TobL.GetValue('LONGLET');
    Longlet             := THTabSheet(GetControl(LeChamp));
    if Assigned(Longlet) then
      LaPage.ActivePage := Longlet;
  end;

end;

procedure TOF_CPRESULTATFISCAL.LeLabelOnMouseEnter(Sender: Tobject);
var
  LeLabel               : THLabel;
  LeChamp               : String;

begin
  With Sender as THLabel do
    LeChamp             := Name;

  LeLabel               := THLabel(GetControl(LeChamp));
  LeLabel.Font.Style    := LeLabel.Font.Style + [fsUnderline];
  LeLabel.Cursor        := crHandPoint;
(* GP le 15/07/2008 : evol AGL 7.2
  if Enter then
  begin
    LeLabel.Font.Style    := LeLabel.Font.Style + [fsUnderline];
    LeLabel.Cursor        := crHandPoint;
  end
  else
  begin
    LeLabel.Font.Style    := LeLabel.Font.Style - [fsUnderline];
    LeLabel.Cursor        := crDefault;
  end ;
*)
end;
(* GP le 15/07/2008 : nouveau : evol AGL 7.2 *)
procedure TOF_CPRESULTATFISCAL.LeLabelOnMouseLeave(Sender: Tobject);
var
  LeLabel               : THLabel;
  LeChamp               : String;

begin
  With Sender as THLabel do
    LeChamp             := Name;

  LeLabel               := THLabel(GetControl(LeChamp));
    LeLabel.Font.Style    := LeLabel.Font.Style - [fsUnderline];
    LeLabel.Cursor        := crDefault;
end;
{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : Recherche dans la fiche Agl les règles de calculs
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.RechercheCalcul;
var
  NbC                   : integer;
  i                     : integer;
  LeTag                 : integer;
  LeChamp               : String;
  LaClass               : String;
  Ledit                 : THEdit;
  LeStr                 : String;
  TobC                  : Tob;
  StrChp                : String;
  StrCalcul             : String;
  NumCal                : integer;
  StrCondition          : String;
  StrCalculCond         : String;

begin
  TobCalcul             := Tob.Create('Les Calculs', nil, -1);

  NbC                   := TForm(Ecran).ComponentCount -1;
  for i := 0 to NbC do
  begin
    LeTag               := TForm(Ecran).Components[i].Tag;
    LeChamp             := TForm(Ecran).Components[i].Name;
    LaClass             := TForm(Ecran).Components[i].ClassName;
    if (LeTag = 6) and (Copy(LeChamp, 1, 1) = 'Z') and (UpperCase_(LaClass) = 'THEDIT') then
    begin
      Ledit             := THEdit(GetControl(LeChamp));
      if not Assigned(Ledit) then
        Continue;

      LeStr             := Ledit.Text;
      TobC              := Tob.Create('Regles', TobCalcul, -1);
      NumCal            := VALEURI(READTOKENST(LeStr));
      StrChp            := READTOKENST(LeStr);
      StrCalcul         := READTOKENST(LeStr);
      StrCondition      := READTOKENST(LeStr);
      StrCalculCond     := READTOKENST(LeStr);
      TobC.AddChampSupValeur('ORDRE', NumCal);
      TobC.AddChampSupValeur('RESULTAT', StrChp);
      TobC.AddChampSupValeur('CALCUL', StrCalcul);
      TobC.AddChampSupValeur('CONDITION', StrCondition);
      TobC.AddChampSupValeur('CALCULCOND', StrCalculCond);
    end;
  end;

  TobCalcul.Detail.Sort('ORDRE');

end;



{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 24/01/2007
Modifié le ... :   /  /
Description .. : recalcul de la fiche DRF avec la liste des règles de calcul
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.RecalculFiche;
var
  Ledit                 : THEdit;
  TobC                  : Tob;
  LeChamp               : String;
  StrCalcul             : String;
  LaCondition           : String;
  i                     : integer;  
  BoolC                 : Boolean;  
  LaValeur              : extended;
begin
  for i := 0 to TobCalcul.Detail.Count -1 do
  begin
    TobC                := TobCalcul.Detail[i];
    StrCalcul           := TobC.GetValue('CALCUL');
    BoolC               := True;
    LaCondition         := TobC.GetValue('CONDITION');
    if LaCondition <> '' then
       BoolC            := IsConditionGood(LaCondition);
    if not BoolC then
       StrCalcul        := TobC.GetValue('CALCULCOND');
    LaValeur            := LeCalcul(StrCalcul);
    LeChamp             := TobC.GetValue('RESULTAT');
    Ledit               := THEdit(GetControl(LeChamp));
    if Assigned(Ledit) then
      Ledit.Text        := STRFMONTANT(LaValeur, 15, V_PGI.OkDecV, '', True);
  end;
  // BVE 17.09.07 : Apres chaque recalcul on verifie les champs resultat
  for i := 0 to TobDRF.Detail.Count - 1 do
  begin
     LeChamp := TobDRF.Detail[i].GetValue('NATURE');
     TobDRF.Detail[i].AddChampSupValeur('RESULTAT',GetChampResultat(LeChamp));
  end;
end;

function TOF_CPRESULTATFISCAL.IsConditionGood(LaCondition : String) : boolean;
var
  ValC1                 : Extended;
  ValC2                 : Extended;
  OpeC                  : integer;
  PosP                  : integer;   
  Ledit                 : THEdit;
  LeChamp               : String;
begin
  Result := True;
  PosP              := Pos('>', LaCondition);
  OpeC              := 1;                     //'>'
  if PosP = 0 then
  begin
    PosP            := Pos('<', LaCondition);
    OpeC            := 2;                     //'<'
  end;
  if PosP = 0 then
  begin
    PosP            := Pos('=', LaCondition);
    OpeC            := 3;                     //'='
  end;
  if PosP = 0 then
  begin
    PosP            := Pos('~', LaCondition);
    OpeC            := 4;                     //'<>'
  end;
  if PosP <> 0 then
  begin
    LeChamp         := Copy(LaCondition, 1, PosP -1);
    Ledit           := THEdit(GetControl(LeChamp));
    if Assigned(Ledit) then
      ValC1         := VALEUR(Ledit.Text)
    else
      ValC1         := 0;
    LeChamp         := Copy(LaCondition, PosP +1, Length(LaCondition));
    Ledit           := THEdit(GetControl(LeChamp));
    if Assigned(Ledit) then
      ValC2         := VALEUR(Ledit.Text)
    else
      ValC2         := 0;


    case OpeC of
      1 :         // '>'
        begin
          if not (ValC1 > ValC2) then
            Result       := False
        end;

      2 :         // '<'
        begin
          if not (ValC1 < ValC2) then
            Result       := False
        end;

      3 :         // '='
         begin
          if not (ValC1 = ValC2) then
            Result       := False
        end;

      4 :         // '<>'
        begin
          if not (ValC1 <> ValC2) then
            Result       := False
        end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/02/2007
Modifié le ... :   /  /
Description .. : Le Calcul
Mots clefs ... :
*****************************************************************}
function TOF_CPRESULTATFISCAL.LeCalcul(StrCalcul: String): Extended;
var
  OpePrec               : String;
  OpeActu               : String;
  LaValeur              : Extended;
  LeChamp               : String;
  Ledit                 : THEdit;
begin
  LaValeur              := 0;
  OpePrec               := '+';
  GetDetailCalcul(StrCalcul,OpeActu,LeChamp);
  while LeChamp <> '' do
  begin
    Ledit               := THEdit(GetControl(LeChamp));
    if Assigned(Ledit) then
    begin
      if OpePrec = '+' then
        LaValeur        := LaValeur + VALEUR(Ledit.Text)
      else
        LaValeur        := LaValeur - VALEUR(Ledit.Text);
    end;
    OpePrec             := OpeActu;
    GetDetailCalcul(StrCalcul,OpeActu,LeChamp);
  end;
  Result                := LaValeur;

end;

procedure TOF_CPRESULTATFISCAL.GetDetailCalcul(var StrCalcul : String ; var OpeActu : String ; var LeChamp : String);
var
  PosP                  : integer;
  PosM                  : integer;
begin
  PosP                  := Pos('+', StrCalcul);
  PosM                  := Pos('-', StrCalcul);
  if (PosM = 0) and (PosP = 0) then
  begin
    OpeActu             := '+';
    Lechamp             := READTOKENPipe(StrCalcul, '+');
  end
  else if (PosM <> 0) and (PosM < PosP) then
  begin
    OpeActu             := '-';
    LeChamp             := READTOKENPipe(StrCalcul, '-');
  end
  else if PosP = 0 then
  begin
    OpeActu             := '-';
    LeChamp             := READTOKENPipe(StrCalcul, '-');
  end
  else
  begin
    OpeActu             := '+';
    LeChamp             := READTOKENPipe(StrCalcul, '+');
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : Affecte aux boutons l'outil associé
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.AffectationBtOutil;
var
  Bto                   : TToolbarButton97;
  Ledit                 : THEdit;
  LaClass               : String;
  NbC                   : integer;
  i                     : integer;
  LeTag                 : integer;
  LeChamp               : String;

begin

  NbC := TForm(Ecran).ComponentCount -1;
  for i := 0 to NbC do
  begin
    LeTag               := TForm(Ecran).Components[i].Tag;
    LeChamp             := TForm(Ecran).Components[i].Name;
    LaClass             := TForm(Ecran).Components[i].ClassName;
    Case LeTag of
      3, 4, 5, 6 :          // Tag = 3 --> Outils DRF   Tag = 4,5,6 --> Saisie en détails reitég/ded/both
        begin
          // Les boutons de la fiche doivent commencer par 'B' suivi du champ auquel ils sont rattachés. (ex : B2058A_XL_MOA)
          if Copy(LeChamp, 1, 1) <> 'B' then
            Continue;
          Bto                 := TToolbarButton97(GetControl(LeChamp));

          if LeTag = 3 then
          begin

            if LeChamp = 'B2058A_XL_MOA' then              // suivi des déficits (05)
              Bto.OnClick     := B2058A_XL_MOA_OnClick
            else if LeChamp = 'B2058A_WC_MOA' then         // rémunération du conjoint (01)
              Bto.OnClick     := B2058A_WC_MOA_OnClick
            else if LeChamp = 'B2058A_XA_MOA' then         // régime des sociétés mères (02)
              Bto.OnClick     := B2058A_XA_MOA_OnClick
            else if LeChamp = 'B2058B_CA_MOA' then         // provisions pour riques et charges
              Bto.OnClick     := B2058B_CA_MOA_OnClick
            else if LeChamp = 'B2058B_CD_MOA' then         // provisions pour dépréciation
              Bto.OnClick     := B2058B_CD_MOA_OnClick
            else if LeChamp = 'B2058B_CG_MOA' then         // Charges à payer
              Bto.OnClick     := B2058B_CG_MOA_OnClick
            else if LeChamp = 'B2058B_ZV_MOA' then         // indemnités congés à payer
              Bto.OnClick     := B2058B_ZV_MOA_OnClick
            else if (LeChamp = 'B2058A_WN_MOA') or (LeChamp = 'B2058A_WZ_MOA') then         // suivi des PVCT différés
              Bto.OnClick     := B2058A_WN_MOA_OnClick
            else if (LeChamp = 'B2058A_WM_MOA') or (LeChamp = 'B2058A_WW_MOA') or (LeChamp = 'BP31') then         // suivi des MVLT
              Bto.OnClick     := B2058A_WM_MOA_OnClick
            { FQ 21152 BVE 20.07.07
            else if LeChamp = 'B2058A_WX_MOA' then         // suivi des déficits (05) }
            else if LeChamp = 'B2058A_XB_MOA' then         // suivi des déficits (05)
            { END FQ 21152 }
              Bto.OnClick     := B2058A_XL_MOA_OnClick
            else if (LeChamp = 'B2058A_VA_MOA') or (LeChamp = 'B2058A_WV_MOA') then         // imposition des PVLT
              Bto.OnClick     := B2058A_VA_MOA_OnClick
            else if (LeChamp = 'B2058AB_AH_MOA') or (LeChamp = 'B2058AB_AQ_MOA') then       // suivi des déficits (intégration fiscale)
              Bto.OnClick     := B2058AB_AH_MOA_OnClick
            else if LeChamp = 'B2058AB_AG_MOA' then         // Suivi de MVLT (intégration fiscale)
              Bto.OnClick     := B2058AB_AG_MOA_OnClick
            else if (LeChamp = 'B2058AB_AF_MOA') or (LeChamp = 'B2058AB_AU_MOA') then       // imposition des PVLT (intégration fiscale)
              Bto.OnClick     := B2058AB_AF_MOA_OnClick
            else if LeChamp = 'BP1' then                    // Jeton de présence
              Bto.OnClick     := BJetonPresence_OnClick
            else if LeChamp = 'BP06_1' then                 // Cession d'immobilisations
              Bto.OnClick     := BCessionImo_OnCLick;
          end
          else                                              // saisie en détail
            Bto.OnClick       := BDetail_OnClick;
        end;
      1 :          // Tag = 1 --> Edit pour calcul
        begin
          if Copy(LeChamp, 1, 1) <> 'P' then
            Continue;
          Ledit         := THEdit(GetControl(LeChamp));
          Ledit.OnChange  := EventRecalcul;
          Ledit.OnKeyDown := LeditOnKeyDown;
        end;
    end;      //case

    if UpperCase_(LaClass) = 'THEDIT' then
    begin
      Ledit             := THEdit(GetControl(LeChamp));
      Ledit.OnExit      := RemiseEnForme;

    end;
  end;

end;


procedure TOF_CPRESULTATFISCAL.BretourOnClick(Sender: Tobject);
begin
  LaPage.ActivePage     := TSHSYN;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 27/02/2007
Modifié le ... :   /  /    
Description .. : permet de formater correctement les saisies
Suite ........ : correctif : enregistre l'indice 0 dans la tob
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.RemiseEnForme(Sender: Tobject);
var
  Ledit                 : THEdit;
  TobV                  : Tob;
  LaNature              : String;

begin
  try
    With Sender as THEdit do
    begin
      LaNature            := Copy(Name, 2, Length(Name));
      TobV                := TobDRF.FindFirst(['NATURE', 'INDICE'], [LaNature, 0], false);
      if TobV = nil then
      begin
        AjouteVersTob(LaNature, '', 0);
        TobV              := TobDRF.FindFirst(['NATURE', 'INDICE'], [LaNature, 0], false);
      end;

      Ledit               := THEdit(GetControl(Name));
      if Copy(Name, 1, 1) = 'C' then                 // Explication
      begin
        Ledit.Text        := Copy(Ledit.Text, 1, 35);
        TobV.Setstring('COMMENTAIRE', Ledit.Text);
        if Ledit.Modified then
          bModif          := True;
      end;
      if Copy(Name, 1, 1) = 'P' then                 // saisie montant
      begin
        Ledit.Text        := STRFMONTANT(VALEUR(Ledit.Text), 15, V_PGI.OkDecV, '', True);
        TobV.SetDouble('MONTANT', VALEUR(Ledit.Text));
        if Ledit.Modified then
          bModif          := True;
      end;
    end;
  except
    on E : Exception do
      PGIError(E.Message);
  end;
end;

procedure TOF_CPRESULTATFISCAL.EventRecalcul(Sender: Tobject);
begin
  RecalculFiche;
end;


procedure TOF_CPRESULTATFISCAL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPRESULTATFISCAL.OnDelete ;
begin
  Inherited ;
end ;


procedure TOF_CPRESULTATFISCAL.OnClose ;
begin
  Inherited ;
  TobDRF.Free;
  LaTobDRF.Free;
  TobCalcul.Free;
  TobLien.Free;

end ;

procedure TOF_CPRESULTATFISCAL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPRESULTATFISCAL.OnCancel () ;
begin
  Inherited ;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/01/2007
Modifié le ... : 17/09/2007
Description .. : ajout d'éléments dans la TobDRF
Suite ........ : - nature       (ex : 2058A_WS_MOA)
Suite ........ : - Comm       (résultat comptable déduction)
Suite ........ : - Code1        2058A
Suite ........ : - Code2        WS_MOA
Suite ........ : - Mt1          (1000)
Suite ........ : - Indice       (0)
Suite ........ : - Mt2          (0)
Suite ........ : - Mt3          (0)
Suite ........ : - Mt4          (0)
Suite ........ : - Mt5          (0)
Mots clefs ... :
*****************************************************************}
function TOF_CPRESULTATFISCAL.AjouteVersTob(LaNature, LeComm: String; Mt1: Extended; LIndice: integer; Mt2, Mt3, Mt4, Mt5 : Extended): boolean;
var
  TobV                  : Tob;
  StrCode1              : String;
  StrCode2              : String;
  i                     : integer;

begin
  Result                := False;
  TobV                  := Tob.Create('Datas', TobDRF, -1);
  if TobV <> nil then
    Result              := True
  else
    Exit;
  TobV.AddChampSupValeur('NATURE', LaNature);
  TobV.AddChampSupValeur('INDICE', LIndice);
  i                     := pos('_', LaNature);
  StrCode1              := Copy(LaNature, 1, i-1);
  StrCode2              := copy(LaNature, i+1, Length(LaNature));
  TobV.AddChampSupValeur('CODE1', StrCode1);
  TobV.AddChampSupValeur('CODE2', StrCode2);
  TobV.AddChampSupValeur('COMMENTAIRE', LeComm);
  TobV.AddChampSupValeur('MONTANT', Mt1);
  TobV.AddChampSupValeur('MT2', Mt2);
  TobV.AddChampSupValeur('MT3', Mt3);
  TobV.AddChampSupValeur('MT4', Mt4);
  TobV.AddChampSupValeur('MT5', Mt5);
  if not (TobV.GetValue('COMMENTAIRE') = '') and not (TobV.GetValue('MONTANT')) then
    bmodif              := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/09/2007
Modifié le ... : 14/09/2007
Description .. : Permet de retourner le nom du champ dans lequel le resultat 
Suite ........ : de l'operation du champ passé en paramètre sera inscrit.
Mots clefs ... : 
*****************************************************************}
function TOF_CPRESULTATFISCAL.GetChampResultat( LeChamp : string ; Depart : integer = 0) : string ;
var
  i    : Integer;
  Tmp  : String;
  Ope  : String;
  Chmp : String;
  Rech : String;
begin
  Result := '';
  Rech   := LeChamp;
  for i := Depart to TobCalcul.Detail.Count - 1 do
  begin
     Tmp := TobCalcul.Detail[i].GetValue('CALCUL');
     GetDetailCalcul(Tmp,Ope,Chmp);
     while (Chmp <> '') and (Chmp <> Rech) and (Chmp <> 'P' + Rech) do
        GetDetailCalcul(Tmp,Ope,Chmp);
     if (Chmp <> '') then
     begin
         if (TobCalcul.Detail[i].GetValue('CONDITION') = '') or
           IsConditionGood(TobCalcul.Detail[i].GetValue('CONDITION')) then
        begin
           Result := TobCalcul.Detail[i].GetValue('RESULTAT');
           { Tmp := GetChampResultat( Result, i + 1 );
           if Tmp <> '' then
              Result := Tmp; }
           break;
        end
     end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /    
Description .. : procédure pour supprimer une nature dans la tob
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.EffaceDansTob(LaTob: Tob; LaNature: String ; QueTotal : boolean = False);
var
  TobV                  : Tob;

begin
  if Assigned(LaTob) then
  begin
    TobV                  := LaTob.FindFirst(['NATURE'],[LaNature],False);
    While TobV <> nil do
    begin
      if not QueTotal then
        TobV.Free;
      if QueTotal and (TobV.GetValue('INDICE') = 0) then
        TobV.Free;

      TobV                := LaTob.FindNext(['NATURE'],[LaNature],False);
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 12/01/2007
Modifié le ... :   /  /
Description .. : procedures de fermeture de la fenetre outil.
Suite ........ : Sauvegarde des infos contenues selon nature
Mots clefs ... :
*****************************************************************}

procedure TOF_CPRESULTATFISCAL.HPBOnClose(Sender: Tobject);
begin
  HPBFerme;
end;

procedure TOF_CPRESULTATFISCAL.HPBFerme;
begin
 if csDestroying in TFVierge(Ecran).ComponentState then Exit ;
 TFVierge(Ecran).Enabled   := True;
 if HPB <> nil then
  begin
   HPB.Free;
   HPB := nil ;
  end ;
  tHPB                  := '';
  bHPB                  := False;
  { FQ 20531 BVE 08.06.07 }
  TFVierge(Ecran).SetFocus;
  { END FQ 20531 }
end;

(*
procedure TOF_CPRESULTATFISCAL.AppelOutil(LaNature: String; LaProc: TNotifyEvent; LeTitre: String; LaBorne1, LaBorne2: extended; LeTag : integer);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := LaNature;
    HPB.TypeToolBox(tHPB, TobDRF, LeTitre, LaBorne1, LaBorne2, LeTag);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := LaProc;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row      := 0;
    HPB.LaGrille.Col      := 1;
  end;

end;

procedure TOF_CPRESULTATFISCAL.BOutil_Onclick(Sender: Tobject);
var
  LeChamp               : String;
  LeTag                 : integer;
  LaNature              : String;
  LeTitre               : String;
  LeLabel               : THLabel;
  LaBorne1              : extended;
  LaBorne2              : extended;

begin
  with Sender As TToolbarButton97 do
  begin
    LeChamp             := Name;
    LeTag               := Tag;

    if (LeTag = 3) or (LeTag = 4) then
    begin
      LaNature          := Copy (LeChamp, 2, length(LeChamp));
      LeLabel             := THLabel(GetControl('T'+LaNature));
      if Assigned(LeLabel) then
        LeTitre           := LeLabel.Caption;
      if LaNature = '2058A_WZ_MOA' then
        LaNature        := '2058A_WN_MOA'
      else if LaNature = '2058A_WW_MOA' then
        LaNature        := '2058A_WM_MOA'
      else if LaNature = '2058A_WV_MOA' then
        LaNature        := '2058A_VA_MOA'
      else if LaNature = '2058AB_AQ_MOA' then
        LaNature        := '2058AB_AH_MOA';

      AppelOutil(LaNature, nil, LeTitre, Laborne1, LaBorne2, LeTag);
    end;


  end;
end;
*)

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 11/01/2007
Modifié le ... :   /  /
Description .. : appel de l'outil : suivi des déficits
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_XL_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058A_XL_MOA';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058A_XL_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... : 22/01/2007
Description .. : appel de l'outil : rémunération du conjoint de l'exploitant
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_WC_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058A_WC_MOA';
    HPB.TypeToolBox(tHPB, TobDRF, '', RemConjointMin, RemConjointMax);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058A_WC_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 29/01/2007
Modifié le ... :   /  /
Description .. : appel de l'outil : suivi des MVLT (intégration fiscale)
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058AB_AG_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058AB_AG_MOA';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058AB_AG_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 30/01/2007
Modifié le ... :   /  /    
Description .. : appel de l'outil : jetons de présence
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.BJetonPresence_OnClick(Sender: Tobject);
var
  Ledit                 : THEdit;
  BaseJeton             : Extended;

begin
  //Création de la fenêtre
  if not bHPB then
  begin
    BaseJeton           := 0;
    Ledit               := THEdit(GetControl('BASEJETON'));
    if Assigned(Ledit) then
      BaseJeton         := VALEUR(Ledit.Text);
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := 'JETONPRESENCE';
    HPB.TypeToolBox(tHPB, TobDRF, '', BaseJeton);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := BJetonPresence_bgValideOnclick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 02/02/2007
Modifié le ... :   /  /
Description .. : appel de l'outil : cession des immobilisations
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.BCessionImo_OnCLick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := 'CESSIONIMO';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := BCessionImo_bgValideOnclick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 29/01/2007
Modifié le ... :   /  /    
Description .. : appel de l'outil : suivi des déficits (intégration fiscale)
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058AB_AH_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058AB_AH_MOA';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058AB_AH_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : appel de l'outil : régime des sociétés mères et des fialiales
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_XA_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058A_XA_MOA';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058A_XA_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : appel de la boite à outils provisions pour risques et charges
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058B_CA_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058B_CA_MOA';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058B_CA_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Appel de la boite à outil provisions pour dépréciation
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058B_CD_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058B_CD_MOA';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058B_CD_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : appel de la boite à outils charges à payer
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058B_CG_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058B_CG_MOA';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058B_CG_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : appel de la boite à outils indemnités congès à payer
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058B_ZV_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058B_ZV_MOA';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058B_ZV_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : appel de la boite à outils suivi de PVCT différés
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_WN_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058A_WN_MOA';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058A_WN_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : appel de la boite à outils suivi des MVLT
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_WM_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058A_WM_MOA';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058A_WM_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : appel de la boite à outils imposition des MVLT
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_VA_MOA_OnClick(Sender: Tobject);
var
  ValLT                 : extended;

begin
  //Création de la fenêtre
  if not bHPB then
  begin
    ValLT               := 0;
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058A_VA_MOA';
    if Assigned(GetControl('PP06_5')) then
      ValLT             := VALEUR(GetControlText('PP06_5'));
    HPB.TypeToolBox(tHPB, TobDRF, '', ValLT);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058A_VA_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 07/02/2007
Modifié le ... :   /  /
Description .. : Appel de la boite à outil imposition des PVLT (intégration
Suite ........ : fiscale)
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058AB_AF_MOA_OnClick(Sender: Tobject);
begin
  //Création de la fenêtre
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB             := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    tHPB                := '2058AB_AF_MOA';
    HPB.TypeToolBox(tHPB, TobDRF);
    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := B2058AB_AF_MOA_bgValideOnClick;

    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row      := 0;
    HPB.LaGrille.Col      := 1;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 19/01/2007
Modifié le ... :   /  /
Description .. : Appel de la boite à outils pour la saisie en détail
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.BDetail_OnClick(Sender: Tobject);
var
  LeTitre               : String;
  LeLabel               : THLabel;
  LaNature              : String;
  LeTag                 : integer;
  LeBouton              : TToolbarButton97;

begin
  With Sender as TToolbarButton97 do
  begin
  if not bHPB then
  begin
    TFVierge(Ecran).Enabled := False;
    HPB                 := TDRFToolBox.Create(Ecran, BValider.Glyph, BgAnnule.Glyph, BgAjouteLg.Glyph, BgSupprimeLg.glyph);
    LaNature            := Copy(Name, 2, Length(Name));
    LeLabel             := THLabel(GetControl('T'+LaNature));
    if Assigned(LeLabel) then
      LeTitre           := LeLabel.Caption;
    tHPB                := LaNature;
    LeBouton            := TToolbarButton97(GetControl(Name));
    LeTag               := LeBouton.Tag;
    HPB.TypeToolBox(LaNature, TobDRF, LeTitre, 0, 0, LeTag);

    bHPB                := True;                    //on bloque l'accès à de nouvelles fenêtres
    HPB.OnClose         := HPBOnClose;
    HPB.BAnnule.OnClick := HPBOnClose;              //bouton annule
    HPB.BValide.OnClick := BDetail_bgValideOnClick; //bouton Valide
    HPB.LaGrille.SetFocus;
    HPB.LaGrille.Row    := 0;
    HPB.LaGrille.Col    := 1;
  end;

  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 07/02/2007
Modifié le ... :   /  /    
Description .. : enregistrement dans Tob de l'outil imposition des PVLT 
Suite ........ : (intégration fiscale)
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058Ab_AF_MOA_bgValideOnclick(Sender: Tobject);
begin
  B2058AB_AF_MOA_Valider;
end;

procedure TOF_CPRESULTATFISCAL.B2058AB_AF_MOA_Valider;
var
  i                     : integer;
  V09_8                 : extended;
  V09_9                 : extended;
  V09_10                : extended;
  V09_11                : extended;
  V09_12                : extended;
  V09_13                : extended;
  LCb                   : THCheckbox;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : integer;

begin
  RefreshLaGrille;
  
  EffaceDansTob(TobDRF, '2058AB_AU_MOA');
  EffaceDansTob(tobDRF, '2058AB_AF_MOA');

  //Affectation des nouvelles données
  Benreg                := False;
  NbEnreg               := 0;
  for i := 0 to HPB.LaGrille.RowCount -4 do
  begin
    V09_8               := VALEUR(HPB.LaGrille.Cells[2, i]);
    V09_9               := VALEUR(HPB.LaGrille.Cells[3, i]);
    V09_10              := VALEUR(HPB.LaGrille.Cells[4, i]);
    V09_11              := VALEUR(HPB.LaGrille.Cells[5, i]);
    V09_12              := VALEUR(HPB.LaGrille.Cells[6, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if V09_8 <> 0 then
    begin
      if not AjouteVersTob('2058AB_AF_MOA', StrComm, V09_8, i+1, V09_9, V09_10, V09_11, V09_12) then
        PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
      begin
        Benreg          := True;
        Inc(NbEnreg);
      end;
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058AB_AF_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;
  LCb                   := THCheckbox(GetControl('X2058AB_AU_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  i                     := HPB.LaGrille.RowCount -3;          // ligne des totaux
  V09_11                := VALEUR(HPB.LaGrille.Cells[5, i]);
  V09_13                := VALEUR(HPB.LaGrille.Cells[7, i]);
  if not AjouteVersTob('2058AB_AU_MOA', '', V09_11) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058AB_AU_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058AB_AF_MOA', '', V09_13) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058AB_AF_MOA>', TForm(Ecran).Caption);


  If Assigned(GetControl('P2058AB_AU_MOA')) then
    SetControlText('P2058AB_AU_MOA', STRFMONTANT(V09_11, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058AB_AF_MOA')) then
    SetControlText('P2058AB_AF_MOA', STRFMONTANT(V09_13, 15, V_PGI.OkDecV, '', True));
  if assigned(GetControl('C2058AB_AF_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058AB_AF_MOA',StrComm);
  end;  
  if assigned(GetControl('C2058AB_AU_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058AB_AU_MOA',StrComm);
  end;

  HPBFerme;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 15/01/2007
Modifié le ... :   /  /
Description .. : enregistre l'outil suivi des déficits dans la tob
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_XL_MOA_bgValideOnClick(Sender: Tobject);
begin
  B2058A_XL_MOA_Valider;
end;

procedure TOF_CPRESULTATFISCAL.B2058A_XL_MOA_Valider;
var
  i                     : integer;
  V05_1                 : extended;
  V05_2                 : extended;
  V05_3                 : extended;
  LCb                   : THCheckbox;
  V05_6                 : extended;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : integer;

begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, '2058A_XL_MOA');
  EffaceDansTob(TobDRF, '2058B_DQ_MOA');
  { FQ 21152 BVE 20.07.07
  EffaceDansTob(TobDRF, '2058A_WX_MOA'); }
  EffaceDansTob(TobDRF, '2058A_XB_MOA');
  { END FQ 21152 }
  EffaceDansTob(TobDRF, '2058B_YK_MOA');


  //Affectation des nouvelles données
  Benreg                := False;
  NbEnreg               := 0;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    V05_1               := VALEUR(HPB.LaGrille.Cells[2, i]);
    V05_2               := VALEUR(HPB.LaGrille.Cells[3, i]);
    V05_3               := VALEUR(HPB.LaGrille.Cells[4, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if V05_1 <> 0 then
    begin
      if not AjouteVersTob('2058A_XL_MOA', StrComm, V05_1, i+1, V05_2, V05_3) then
        PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
      begin
        Inc(NbEnreg);
        Benreg          := True;
      end;
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058A_XL_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  V05_1                 := VALEUR(HPB.LaGrille.Cells[2, i]);
  if not AjouteVersTob('2058B_DQ_MOA', '', V05_1) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058B_DQ_MOA>', TForm(Ecran).Caption);
  V05_2                 := VALEUR(HPB.LaGrille.Cells[3, i]);
  { FQ 21152 BVE 20.07.07
  if not AjouteVersTob('2058A_WX_MOA', '', V05_2) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_WX_MOA>', TForm(Ecran).Caption); }
  if not AjouteVersTob('2058A_XB_MOA', '', V05_2) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_XB_MOA>', TForm(Ecran).Caption);
  { END FQ 21152 }
  V05_3                 := VALEUR(HPB.LaGrille.Cells[4, i]);
  StrComm                 := '';
  if Assigned(GetControl('C2058A_XL_MOA')) then
    StrComm               := GetControlText('C2058A_XL_MOA');
  if not AjouteVersTob('2058A_XL_MOA', StrComm, V05_3) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_XL_MOA>', TForm(Ecran).Caption);
  V05_6                 := VALEUR(HPB.LEdit2.Text);
  if not AjouteVersTob('2058B_YK_MOA', '', V05_6) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058B_YK_MOA>', TForm(Ecran).Caption);

  If Assigned(GetControl('P2058A_XL_MOA')) then
    SetControlText('P2058A_XL_MOA', STRFMONTANT(V05_3, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058B_DQ_MOA')) then
    SetControlText('P2058B_DQ_MOA', STRFMONTANT(V05_1, 15, V_PGI.OkDecV, '', True));
  { FQ 21152 BVE 20.07.07
  If Assigned(GetControl('P2058A_WX_MOA')) then
    SetControlText('P2058A_WX_MOA', STRFMONTANT(V05_2, 15, V_PGI.OkDecV, '', True)); }
  If Assigned(GetControl('P2058A_XB_MOA')) then
    SetControlText('P2058A_XB_MOA', STRFMONTANT(V05_2, 15, V_PGI.OkDecV, '', True));
  { END FQ 21152 }
  If Assigned(GetControl('P2058B_YK_MOA')) then
    SetControlText('P2058B_YK_MOA', STRFMONTANT(V05_6, 15, V_PGI.OkDecV, '', True));

  
  if assigned(GetControl('C2058A_XL_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058A_XL_MOA',StrComm);
  end;
  if assigned(GetControl('C2058A_XB_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058A_XB_MOA',StrComm);
  end;
  HPBFerme;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : enregistrement en Tob de l'outil rémunérationdu conjoint de
Suite ........ : l'exploitant
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_WC_MOA_bgValideOnClick(Sender: Tobject);
begin
  B2058A_WC_MOA_Valider;
end;

procedure TOF_CPRESULTATFISCAL.B2058A_WC_MOA_Valider;
var
  i                     : integer;
  V01_1                 : extended;
  V01_2                 : extended;
  V01_3                 : extended;
  LCb                   : THCheckbox;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : Integer;
begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, '2058A_WC_MOA');
  EffaceDansTob(TobDRF, '2058A_AB_MOA');
  EffaceDansTob(TobDRF, '2058A_AC_MOA');

  //Affectation des nouvelles données
  Benreg                := False;   
  NbEnreg := 0;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    V01_1               := VALEUR(HPB.LaGrille.Cells[2, i]);
    V01_2               := VALEUR(HPB.LaGrille.Cells[3, i]);
    V01_3               := VALEUR(HPB.LaGrille.Cells[4, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if V01_1 <> 0 then
    begin
      if not AjouteVersTob('2058A_WC_MOA', StrComm, V01_1, i+1, V01_2, V01_3) then
        PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
      begin
        Benreg          := True;
        Inc(NbEnreg);
      end;
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058A_WC_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  V01_1                 := VALEUR(HPB.LaGrille.Cells[2, i]);
  V01_2                 := VALEUR(HPB.LaGrille.Cells[3, i]);
  V01_3                 := VALEUR(HPB.LaGrille.Cells[4, i]);
  if not AjouteVersTob('2058A_WC_MOA', '', V01_3) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_WC_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058A_AB_MOA', '', V01_1) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_AB_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058A_AC_MOA', '', V01_2) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_AC_MOA>', TForm(Ecran).Caption);


  If Assigned(GetControl('P2058A_WC_MOA')) then
    SetControlText('P2058A_WC_MOA', STRFMONTANT(V01_3, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058A_AB_MOA')) then
    SetControlText('P2058A_AB_MOA', STRFMONTANT(V01_1, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058A_AC_MOA')) then
    SetControlText('P2058A_AC_MOA', STRFMONTANT(V01_2, 15, V_PGI.OkDecV, '', True));
  
  if assigned(GetControl('C2058A_WC_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058A_WC_MOA',StrComm);
  end;
  HPBFerme;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 29/01/2007
Modifié le ... :   /  /
Description .. : enregistrement dans Tob de l'outil : suivi des déficits
Suite ........ : (intégration fiscale)
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058AB_AH_MOA_bgValideOnclick(Sender: Tobject);
begin
  B2058AB_AH_MOA_Valider;
end;

procedure TOF_CPRESULTATFISCAL.B2058AB_AH_MOA_Valider;
var
  i                     : integer;
  V08_1                 : extended;
  V08_2                 : extended;
  V08_3                 : extended;
  LCb                   : THCheckbox;
  V08_6                 : extended;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : Integer;
begin
  RefreshLaGrille;
  
  EffaceDansTob(TobDRF, '2058BB_EP_MOA');
  EffaceDansTob(TobDRF, '2058AB_AH_MOA');
  EffaceDansTob(TobDRF, '2058AB_AQ_MOA');
  EffaceDansTob(TobDRF, '2058BB_EH_MOA');

  //Affectation des nouvelles données
  Benreg                := False;     
  NbEnreg := 0;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    V08_1               := VALEUR(HPB.LaGrille.Cells[2, i]);
    V08_2               := VALEUR(HPB.LaGrille.Cells[3, i]);
    V08_3               := VALEUR(HPB.LaGrille.Cells[4, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if V08_1 <> 0 then
    begin
      if not AjouteVersTob('2058AB_AH_MOA', StrComm, V08_1, i+1, V08_2, V08_3) then
        PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
      begin
        Benreg          := True;
        Inc(NbEnreg);
      end;
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058AB_AH_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  LCb                   := THCheckbox(GetControl('X2058AB_AQ_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  V08_1                 := VALEUR(HPB.LaGrille.Cells[2, i]);
  if not AjouteVersTob('2058BB_EP_MOA', '', V08_1) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058BB_EP_MOA>', TForm(Ecran).Caption);
  V08_2                 := VALEUR(HPB.LaGrille.Cells[3, i]);
  if not AjouteVersTob('2058AB_AH_MOA', '', V08_2) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058AB_AH_MOA>', TForm(Ecran).Caption);
  V08_3                 := VALEUR(HPB.LaGrille.Cells[4, i]);
  if not AjouteVersTob('2058AB_AQ_MOA', StrComm, V08_3) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058AB_AQ_MOA>', TForm(Ecran).Caption);
  V08_6                 := VALEUR(HPB.LEdit2.Text);
  if not AjouteVersTob('2058BB_EH_MOA', '', V08_6) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058BB_EH_MOA>', TForm(Ecran).Caption);

  If Assigned(GetControl('P2058AB_AH_MOA')) then
    SetControlText('P2058AB_AH_MOA', STRFMONTANT(V08_2, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058BB_EP_MOA')) then
    SetControlText('P2058BB_EP_MOA', STRFMONTANT(V08_1, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058AB_AQ_MOA')) then
    SetControlText('P2058AB_AQ_MOA', STRFMONTANT(V08_3, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058BB_EH_MOA')) then
    SetControlText('P2058BB_EH_MOA', STRFMONTANT(V08_6, 15, V_PGI.OkDecV, '', True));

  if assigned(GetControl('C2058AB_AH_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058AB_AH_MOA',StrComm);
  end;
  if assigned(GetControl('C2058AB_AQ_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058AB_AQ_MOA',StrComm);
  end;
  HPBFerme;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 29/01/2007
Modifié le ... :   /  /
Description .. : enregistrement dans Tob de l'outil : suivi des MVLT
Suite ........ : (intégration fiscale)
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058AB_AG_MOA_bgValideOnclick(Sender: Tobject);
begin
  B2058AB_AG_MOA_Valider;
end;

procedure TOF_CPRESULTATFISCAL.B2058AB_AG_MOA_Valider;
var
  i                     : integer;
  V09_1                 : extended;
  V09_2                 : extended;
  V09_3                 : extended;
  V09_4                 : extended;
  LCb                   : THCheckbox;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : Integer;

begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, '2058AB_AG_MOA');

  //Affectation des nouvelles données
  Benreg                := False;  
  NbEnreg := 0;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    V09_1               := VALEUR(HPB.LaGrille.Cells[2, i]);
    V09_2               := VALEUR(HPB.LaGrille.Cells[3, i]);
    V09_3               := VALEUR(HPB.LaGrille.Cells[4, i]);
    V09_4               := VALEUR(HPB.LaGrille.Cells[5, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if V09_1 <> 0 then
    begin
      if not AjouteVersTob('2058AB_AG_MOA', StrComm, V09_1, i+1, V09_2, V09_3, V09_4) then
        PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
      begin
        Benreg          := True;
        Inc(NbEnreg);
      end;
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058AB_AG_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  V09_2                 := VALEUR(HPB.LaGrille.Cells[3, i]);
//  V09_4                 := VALEUR(HPB.LaGrille.Cells[5, i]);
//  if not AjouteVersTob('2058A_WM_MOA', '', V09_4) then
//    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_WM_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058AB_AG_MOA', '', V09_2) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058AB_AG_MOA>', TForm(Ecran).Caption);


//  If Assigned(GetControl('P2058A_WM_MOA')) then
//    SetControlText('P2058A_WM_MOA', STRFMONTANT(V09_4, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058AB_AG_MOA')) then
    SetControlText('P2058AB_AG_MOA', STRFMONTANT(V09_2, 15, V_PGI.OkDecV, '', True));

  if assigned(GetControl('C2058AB_AG_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058AB_AG_MOA',StrComm);
  end;
  HPBFerme;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 02/02/2007
Modifié le ... :   /  /
Description .. : enregistrement dans Tob de l'outil : cession des
Suite ........ : immobilisations
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.BCessionImo_bgValideOnclick(Sender: Tobject);
begin
  BCessionImo_Valider;
end;

procedure TOF_CPRESULTATFISCAL.BCessionImo_Valider;
var
  ValTot                : extended;
  ValLT                 : extended;
  ValCT                 : extended;
  StrComm               : String;
  Benreg                : Boolean;
  LCb                   : THCheckbox;
  i                     : integer;

begin
  RefreshLaGrille;
  
  EffaceDansTob(TobDRF, 'CESSIONIMO');

  Benreg                := False;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    ValTot              := VALEUR(HPB.LaGrille.Cells[2, i]);
    ValLT               := VALEUR(HPB.LaGrille.Cells[3, i]);
    ValCT               := VALEUR(HPB.LaGrille.Cells[4, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if ValTot <> 0 then
    begin
      if not AjouteVersTob(tHPB, StrComm, ValTot, i+1, ValLT, ValCT) then
        PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
        Benreg          := True;
    end;
  end;

  LCb                   := THCheckbox(GetControl('XP06_1'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  ValTot                := VALEUR(HPB.LaGrille.Cells[2, i]);
  ValLT                 := VALEUR(HPB.LaGrille.Cells[3, i]);
  ValCT                 := VALEUR(HPB.LaGrille.Cells[4, i]);
  if not AjouteVersTob(tHPB, '', ValTot, 0, ValLT, ValCT) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <' + tHPB + '>', TForm(Ecran).Caption);

  if Assigned(GetControl('PP06_1')) then
    SetControlText('PP06_1', STRFMONTANT(ValTot, 15, V_PGI.OkDecV, '', True));
  if Assigned(GetControl('PP06_2')) then
    SetControlText('PP06_2', STRFMONTANT(ValCT, 15, V_PGI.OkDecV, '', True));
  if Assigned(GetControl('PP06_3')) then
    SetControlText('PP06_3', STRFMONTANT(ValLT, 15, V_PGI.OkDecV, '', True));

  HPBFerme;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 30/01/2007
Modifié le ... :   /  /
Description .. : enregistrement dans TOb de l'outil Jeton de présence
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.BJetonPresence_bgValideOnclick(Sender: Tobject);
begin
  BJetonPresence_Valider;
end;

procedure TOF_CPRESULTATFISCAL.BJetonPresence_Valider;
var
  V07_2                 : extended;
  V07_3                 : extended;
  V07_4                 : extended;
  V07_5                 : extended;
  TotJeton              : Extended;
  StrComm               : String;
  Benreg                : Boolean;
  LCb                   : THCheckbox;
begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, 'P1');

  //Affectation des nouvelles données
  Benreg                := False;
  StrComm               := BoolToStr(HPB.LaCoche.Checked);
  V07_2                 := VALEUR(HPB.LaGrille.Cells[1, 0]);
  V07_3                 := VALEUR(HPB.LaGrille.Cells[2, 0]);
  V07_4                 := VALEUR(HPB.LaGrille.Cells[3, 0]);
  V07_5                 := VALEUR(HPB.LaGrille.Cells[4, 0]);
  TotJeton              := VALEUR(HPB.LaGrille.Cells[5, 0]);

  if V07_2 <> 0 then
  begin
    if not AjouteVersTob(tHPB, StrComm, V07_2, 1, V07_3, V07_4, V07_5) then
      PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
    else
      Benreg          := True;
  end;

  LCb                   := THCheckbox(GetControl('XP1'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  if not AjouteVersTob(tHPB, '', TotJeton) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <P1>', TForm(Ecran).Caption);

  If Assigned(GetControl('P' + tHPB)) then
    SetControlText('P' + tHPB, STRFMONTANT(TotJeton, 15, V_PGI.OkDecV, '', True));

  HPBFerme;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : enristrement dans Tob de l'outils : régime des sociétés
Suite ........ : mères et des filiales
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_XA_MOA_bgValideOnClick(Sender: Tobject);
begin
  B2058A_XA_MOA_Valider;

end;

procedure TOF_CPRESULTATFISCAL.B2058A_XA_MOA_Valider;
var
  i                     : integer;
  V02_1                 : extended;
  V02_2                 : extended;
  V02_3                 : extended;
  LCb                   : THCheckbox;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : Integer;
begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, '2058A_XA_MOA');
  EffaceDansTob(TobDRF, '2058A_AE_MOA');

  //Affectation des nouvelles données
  Benreg                := False;     
  NbEnreg := 0;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    V02_1               := VALEUR(HPB.LaGrille.Cells[2, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if V02_1 <> 0 then
    begin
      if not AjouteVersTob('2058A_XA_MOA', StrComm, V02_1, i+1) then
        PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
      begin
        Benreg          := True;
        Inc(NbEnreg);
      end;
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058A_XA_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  V02_2                 := VALEUR(HPB.LaGrille.Cells[3, i]);
  V02_3                 := VALEUR(HPB.LaGrille.Cells[4, i]);
  if not AjouteVersTob('2058A_XA_MOA', '', V02_3) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_XA_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058A_AE_MOA', '', V02_2) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_AE_MOA>', TForm(Ecran).Caption);


  If Assigned(GetControl('P2058A_XA_MOA')) then
    SetControlText('P2058A_XA_MOA', STRFMONTANT(V02_3, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058A_AE_MOA')) then
    SetControlText('P2058A_AE_MOA', STRFMONTANT(V02_2, 15, V_PGI.OkDecV, '', True));
  
  if assigned(GetControl('C2058A_XA_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058A_XA_MOA',StrComm);
  end;
  HPBFerme;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : enregistre dans Tob l'outil provisions pour risques et charges
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058B_CA_MOA_bgValideOnClick(Sender: Tobject);
begin
  B2058B_CA_MOA_Valider;
end;

procedure TOF_CPRESULTATFISCAL.B2058B_CA_MOA_Valider;
var
  i                     : integer;
  V01                   : extended;
  V02                   : extended;
  LCb                   : THCheckbox;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : Integer;
begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, '2058B_CA_MOA');
  EffaceDansTob(TobDRF, '2058B_DA_MOA');
  EffaceDansTob(TobDRF, '2058B_CL_MOA');
  EffaceDansTob(TobDRF, '2058B_DL_MOA');

  //Affectation des nouvelles données
  Benreg                := False;       
  NbEnreg := 0;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    V01               := VALEUR(HPB.LaGrille.Cells[2, i]);
    V02               := VALEUR(HPB.LaGrille.Cells[3, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if (V01 <> 0) or (V02 <> 0) then
    begin
      { FQ 21621 BVE 10.10.07 }
      if not AjouteVersTob('2058B_CL_MOA', StrComm, V01, i+1, V02) then
         PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
      begin
         Benreg          := True;
         Inc(NbEnreg);
      end;
      { END FQ 21621 }
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058B_CA_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;


  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  V01                   := VALEUR(HPB.LaGrille.Cells[2, i]);
  V02                   := VALEUR(HPB.LaGrille.Cells[3, i]);
  if not AjouteVersTob('2058B_CA_MOA', '', V01) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058B_CA_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058B_DA_MOA', '', V02) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058B_DA_MOA>', TForm(Ecran).Caption);

  If Assigned(GetControl('P2058B_CA_MOA')) then
    SetControlText('P2058B_CA_MOA', STRFMONTANT(V01, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058B_DA_MOA')) then
    SetControlText('P2058B_DA_MOA', STRFMONTANT(V02, 15, V_PGI.OkDecV, '', True));


  Calcul_2058A_WI_MOA;            // On Calcul 2058A_WI_MOA et 2058A_WU_MOA
  
  if assigned(GetControl('C2058B_CA_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058B_CA_MOA',StrComm);
  end;

  HPBFerme;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : enregistrement dans Tob de l'outil provisions pour
Suite ........ : dépréciation
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058B_CD_MOA_bgValideOnClick(Sender: Tobject);
begin
  B2058B_CD_MOA_Valider;
end;


procedure TOF_CPRESULTATFISCAL.B2058B_CD_MOA_Valider;
var
  i                     : integer;
  V01                   : extended;
  V02                   : extended;
  LCb                   : THCheckbox;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : integer;
begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, '2058B_CD_MOA');
  EffaceDansTob(TobDRF, '2058B_DD_MOA');
  EffaceDansTob(TobDRF, '2058B_CM_MOA');
  EffaceDansTob(TobDRF, '2058B_DM_MOA');

  //Affectation des nouvelles données
  Benreg                := False;    
  NbEnreg := 0;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    V01               := VALEUR(HPB.LaGrille.Cells[2, i]);
    V02               := VALEUR(HPB.LaGrille.Cells[3, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if (V01 <> 0) or (V02 <> 0) then
    begin
       { FQ 21621 BVE 10.10.07 }
       if not AjouteVersTob('2058B_CM_MOA', StrComm, V01, i+1, V02) then
          PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
       else
       begin
          Benreg          := True;
          Inc(NbEnreg);
       end;
       { END FQ 21621 }
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058B_CD_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;


  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  V01                   := VALEUR(HPB.LaGrille.Cells[2, i]);
  V02                   := VALEUR(HPB.LaGrille.Cells[3, i]);
  if not AjouteVersTob('2058B_CD_MOA', '', V01) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058B_CD_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058B_DD_MOA', '', V02) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058B_DD_MOA>', TForm(Ecran).Caption);

  If Assigned(GetControl('P2058B_CD_MOA')) then
    SetControlText('P2058B_CD_MOA', STRFMONTANT(V01, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058B_DD_MOA')) then
    SetControlText('P2058B_DD_MOA', STRFMONTANT(V02, 15, V_PGI.OkDecV, '', True));

  Calcul_2058A_WI_MOA;            // On Calcul 2058A_WI_MOA et 2058A_WU_MOA

  
  if assigned(GetControl('C2058B_CD_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058B_CD_MOA',StrComm);
  end;

  HPBFerme;
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : enregistre dans Tob l'outil charges à payer
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058B_CG_MOA_bgValideOnClick(Sender: Tobject);
begin
  B2058B_CG_MOA_Valider;
end;

procedure TOF_CPRESULTATFISCAL.B2058B_CG_MOA_Valider;
var
  i                     : integer;
  V01                   : extended;
  V02                   : extended;
  LCb                   : THCheckbox;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : integer;
begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, '2058B_CG_MOA');
  EffaceDansTob(TobDRF, '2058B_DG_MOA');
  EffaceDansTob(TobDRF, '2058B_CN_MOA');
  EffaceDansTob(TobDRF, '2058B_DN_MOA');

  //Affectation des nouvelles données
  Benreg                := False;         
  NbEnreg := 0;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    V01               := VALEUR(HPB.LaGrille.Cells[2, i]);
    V02               := VALEUR(HPB.LaGrille.Cells[3, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if (V01 <> 0) or (V02 <> 0) then
    begin
       { FQ 21621 BVE 10.10.07 }
       if not AjouteVersTob('2058B_CN_MOA', StrComm, V01, i+1, V02) then
          PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
       else
       begin
          Benreg          := True;
          Inc(NbEnreg);
       end;
       { END FQ 21621 }
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058B_CG_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;


  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  V01                   := VALEUR(HPB.LaGrille.Cells[2, i]);
  V02                   := VALEUR(HPB.LaGrille.Cells[3, i]);
  if not AjouteVersTob('2058B_CG_MOA', '', V01) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058B_CG_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058B_DG_MOA', '', V02) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058B_DG_MOA>', TForm(Ecran).Caption);

  If Assigned(GetControl('P2058B_CG_MOA')) then
    SetControlText('P2058B_CG_MOA', STRFMONTANT(V01, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058B_DG_MOA')) then
    SetControlText('P2058B_DG_MOA', STRFMONTANT(V02, 15, V_PGI.OkDecV, '', True));

  Calcul_2058A_WI_MOA;            // On Calcul 2058A_WI_MOA et 2058A_WU_MOA
  
  if assigned(GetControl('C2058B_CG_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058B_CG_MOA',StrComm);
  end;

  HPBFerme;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 22/01/2007
Modifié le ... :   /  /
Description .. : Enregistre dans la tob  indemnité congès à payer
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058B_ZV_MOA_bgValideOnClick(Sender: Tobject);
begin
  B2058B_ZV_MOA_Valider;
end;

procedure TOF_CPRESULTATFISCAL.B2058B_ZV_MOA_Valider;
var
  i                     : integer;
  V01                   : extended;
  V02                   : extended;
  LCb                   : THCheckbox;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : Integer;
begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, '2058B_ZV_MOA');
  EffaceDansTob(TobDRF, '2058B_ZW_MOA');

  //Affectation des nouvelles données
  Benreg                := False;      
  NbEnreg := 0;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    V01                 := VALEUR(HPB.LaGrille.Cells[2, i]);
    V02                 := VALEUR(HPB.LaGrille.Cells[3, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if (V01 <> 0) or (V02 <> 0) then
    begin           
       { FQ 21621 BVE 10.10.07 }
       if not AjouteVersTob('2058B_ZV_MOA', StrComm, V01, i+1, V02) then
          PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
       else
       begin
          Benreg          := True;
          Inc(NbEnreg);
       end;
       { END FQ 21621 }
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058B_ZV_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;


  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  V01                   := VALEUR(HPB.LaGrille.Cells[2, i]);
  V02                   := VALEUR(HPB.LaGrille.Cells[3, i]);
  if not AjouteVersTob('2058B_ZV_MOA', '', V01) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058B_ZV_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058B_ZW_MOA', '', V02) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058B_ZW_MOA>', TForm(Ecran).Caption);

  If Assigned(GetControl('P2058B_ZV_MOA')) then
    SetControlText('P2058B_ZV_MOA', STRFMONTANT(V01, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058B_ZW_MOA')) then
    SetControlText('P2058B_ZW_MOA', STRFMONTANT(V02, 15, V_PGI.OkDecV, '', True));
  
  if assigned(GetControl('C2058B_ZV_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058B_ZV_MOA',StrComm);
  end;
  HPBFerme;

end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : enregistre dans la tob suivi PVCT différés
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_WN_MOA_bgValideOnClick(Sender: Tobject);
begin
  B2058A_WN_MOA_Valider;
end;

procedure TOF_CPRESULTATFISCAL.B2058A_WN_MOA_Valider;
var
  i                     : integer;
  V03_1                 : extended;
  V03_2                 : extended;
  V03_3                 : extended;
  LCb                   : THCheckbox;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : Integer;
begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, '2058A_WN_MOA');
  EffaceDansTob(TobDRF, '2058A_WZ_MOA');

  //Affectation des nouvelles données
  Benreg                := False;   
  NbEnreg := 0;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    V03_1               := VALEUR(HPB.LaGrille.Cells[2, i]);
    V03_2               := VALEUR(HPB.LaGrille.Cells[3, i]);
    V03_3               := VALEUR(HPB.LaGrille.Cells[4, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if V03_1 <> 0 then
    begin
      if not AjouteVersTob('2058A_WN_MOA', StrComm, V03_1, i+1, V03_2, V03_3) then
        PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
      begin
        Benreg          := True;
        Inc(NbEnreg);
      end;
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058A_WN_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;
  LCb                   := THCheckbox(GetControl('X2058A_WZ_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  V03_2                 := VALEUR(HPB.LaGrille.Cells[3, i]);
  V03_3                 := VALEUR(HPB.LaGrille.Cells[4, i]);
  if not AjouteVersTob('2058A_WN_MOA', '', V03_2) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_WN_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058A_WZ_MOA', '', V03_3) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_WZ_MOA>', TForm(Ecran).Caption);


  If Assigned(GetControl('P2058A_WN_MOA')) then
    SetControlText('P2058A_WN_MOA', STRFMONTANT(V03_2, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058A_WZ_MOA')) then
    SetControlText('P2058A_WZ_MOA', STRFMONTANT(V03_3, 15, V_PGI.OkDecV, '', True));

  if assigned(GetControl('C2058A_WN_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058A_WN_MOA',StrComm);
  end;
  if assigned(GetControl('C2058A_WZ_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058A_WZ_MOA',StrComm);
  end;
  HPBFerme;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : enregistre dans tob outil suivi des MVLT
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_WM_MOA_bgValideOnClick(Sender: Tobject);
begin
  B2058A_WM_MOA_Valider;
end;

procedure TOF_CPRESULTATFISCAL.B2058A_WM_MOA_Valider;
var
  i                     : integer;
  V04_1                 : extended;
  V04_2                 : extended;
  V04_3                 : extended;
  V04_4                 : extended;
  LCb                   : THCheckbox;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : Integer;
begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, '2058A_WM_MOA');
  EffaceDansTob(TobDRF, '2058A_WW_MOA');

  //Affectation des nouvelles données
  Benreg                := False;      
  NbEnreg := 0;
  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin
    V04_1               := VALEUR(HPB.LaGrille.Cells[2, i]);
    V04_2               := VALEUR(HPB.LaGrille.Cells[3, i]);
    V04_3               := VALEUR(HPB.LaGrille.Cells[4, i]);
    V04_4               := VALEUR(HPB.LaGrille.Cells[5, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if V04_1 <> 0 then
    begin
      if not AjouteVersTob('2058A_WM_MOA', StrComm, V04_1, i+1, V04_2, V04_3, V04_4) then
        PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
      begin
        Benreg          := True;
        Inc(NbEnreg);
      end;
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058A_WM_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;
  LCb                   := THCheckbox(GetControl('X2058A_WW_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  V04_2                 := VALEUR(HPB.LaGrille.Cells[3, i]);
  V04_4                 := VALEUR(HPB.LaGrille.Cells[5, i]);
  if not AjouteVersTob('2058A_WM_MOA', '', V04_4) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_WM_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058A_WW_MOA', '', V04_2) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_WW_MOA>', TForm(Ecran).Caption);


  If Assigned(GetControl('P2058A_WM_MOA')) then
    SetControlText('P2058A_WM_MOA', STRFMONTANT(V04_4, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058A_WW_MOA')) then
    SetControlText('P2058A_WW_MOA', STRFMONTANT(V04_2, 15, V_PGI.OkDecV, '', True));

  if assigned(GetControl('C2058A_WM_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058A_WM_MOA',StrComm);
  end;
  if assigned(GetControl('C2058A_WW_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058A_WW_MOA',StrComm);
  end;
  if assigned(GetControl('CP31')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('CP31',StrComm);
  end;
  HPBFerme;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : enregistre dans tob outil imposition des MVLT
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.B2058A_VA_MOA_bgValideOnclick(Sender: Tobject);
begin
  B2058A_VA_MOA_Valider;
end;

procedure TOF_CPRESULTATFISCAL.B2058A_VA_MOA_Valider;
var
  i                     : integer;
  V06_8                 : extended;
  V06_9                 : extended;
  V06_10                : extended;
  V06_11                : extended;
  V06_12                : extended;
  V06_13                : extended;
  LCb                   : THCheckbox;
  StrComm               : String;
  Benreg                : Boolean;
  NbEnreg               : Integer;
begin
  RefreshLaGrille;
  
  //Suppression des anciennes données
  EffaceDansTob(TobDRF, '2058A_VA_MOA');
  EffaceDansTob(TobDRF, '2058A_WV_MOA');

  //Affectation des nouvelles données
  Benreg                := False;    
  NbEnreg := 0;
  for i := 0 to HPB.LaGrille.RowCount -4 do
  begin
    V06_8               := VALEUR(HPB.LaGrille.Cells[2, i]);
    V06_9               := VALEUR(HPB.LaGrille.Cells[3, i]);
    V06_10              := VALEUR(HPB.LaGrille.Cells[4, i]);
    V06_11              := VALEUR(HPB.LaGrille.Cells[5, i]);
    V06_12              := VALEUR(HPB.LaGrille.Cells[6, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if V06_8 <> 0 then
    begin
      if not AjouteVersTob('2058A_VA_MOA', StrComm, V06_8, i+1, V06_9, V06_10, V06_11, V06_12) then
        PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
      begin
        Benreg          := True;
        Inc(NbEnreg);
      end;
    end;
  end;

  LCb                   := THCheckbox(GetControl('X2058A_VA_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;
  LCb                   := THCheckbox(GetControl('X2058A_WV_MOA'));
  if Benreg then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
  end;

  i                     := HPB.LaGrille.RowCount -3;          // ligne des totaux
  V06_11                := VALEUR(HPB.LaGrille.Cells[5, i]);
  V06_13                := VALEUR(HPB.LaGrille.Cells[7, i]);
  if not AjouteVersTob('2058A_VA_MOA', '', V06_11) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_VA_MOA>', TForm(Ecran).Caption);
  if not AjouteVersTob('2058A_WV_MOA', '', V06_13) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <2058A_WV_MOA>', TForm(Ecran).Caption);


  If Assigned(GetControl('P2058A_VA_MOA')) then
    SetControlText('P2058A_VA_MOA', STRFMONTANT(V06_11, 15, V_PGI.OkDecV, '', True));
  If Assigned(GetControl('P2058A_WV_MOA')) then
    SetControlText('P2058A_WV_MOA', STRFMONTANT(V06_13, 15, V_PGI.OkDecV, '', True));
  
  if assigned(GetControl('C2058A_VA_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058A_VA_MOA',StrComm);
  end;
  if assigned(GetControl('C2058A_WV_MOA')) then
  begin
     if NbEnreg = 1 then
        StrComm := HPB.LaGrille.Cells[1, 0]
     else
        StrComm := '';
     SetControlText('C2058A_WV_MOA',StrComm);
  end;
  HPBFerme;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 19/01/2007
Modifié le ... :   /  /
Description .. : Enregistre dans la Tob le contenu de la saisie en détail
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.BDetail_bgValideOnClick(Sender: Tobject);
begin
  BDetail_Valider;
end;


procedure TOF_CPRESULTATFISCAL.BDetail_Valider;
var
  VResultat             : extended;
  Vres2                 : extended;
  StrComm               : String;
  i                     : integer;
  LeBouton              : TToolbarButton97;
  LCb                   : THCheckbox;
  LeChamp               : String;
  LeTag                 : integer;
  nbEnreg               : integer;
  Ledit                 : THEdit;
  TobV                  : Tob;

begin
  RefreshLaGrille;

  // on supprime les infos précédentes
  EffaceDansTob(TobDRF, tHPB);

  LeChamp               := 'B' + tHPB;
  LeBouton              := TToolbarButton97(GetControl(LeChamp));
  LeTag                 := LeBouton.Tag;          // on récupère le Tag pour le sens d'écriture

  // affectation des données

  nbEnreg               := 0;
  VResultat             := 0;

  for i := 0 to HPB.LaGrille.RowCount -2 do
  begin

(*    if LeTag = 4 then
      VResultat         := VALEUR(HPB.LaGrille.Cells[2, i])
    else if LeTag = 5 then
      VResultat         := VALEUR(HPB.LaGrille.Cells[3, i])
    else if LeTag = 6 then
    begin
      VResultat         := VALEUR(HPB.LaGrille.Cells[2, i]);
      VRes2             := VALEUR(HPB.LaGrille.Cells[3, i]);
    end;
*)
    VResultat           := VALEUR(HPB.LaGrille.Cells[2, i]);
    VRes2               := VALEUR(HPB.LaGrille.Cells[3, i]);
    StrComm             := HPB.LaGrille.Cells[1, i];
    if (VResultat <> 0) or (VRes2 <> 0)  then
    begin
      if not AjouteVersTob(tHPB, StrComm, VResultat, i+1, VRes2) then
        PGIInfo('Erreur de sauvegarde des données sur la ligne <' + StrComm + '>', TForm(Ecran).Caption)
      else
        Inc(nbEnreg);
    end;
  end;

  LeChamp               := 'X' + tHPB;
  LCb                   := THCheckbox(GetControl(LeChamp));
  LeChamp               := 'P' + tHPB;
  Ledit                 := THEdit(GetControl(LeChamp));

  if nbEnreg > 1 then
  begin
    if Assigned(LCb) then
      LCb.Checked       := True;
    if Assigned(Ledit) then
      Ledit.Enabled     := False; 
    { FQ 20531 BVE 11.06.07 }
    StrComm := '';
    { END FQ 20531 }
  end
  else
  begin
    if Assigned(LCb) then
      LCb.Checked       := False;
      Ledit.enabled     := True;
      TobV              := TobDRF.FindFirst(['NATURE', 'INDICE'], [tHPB, 1], False);
      if Assigned(TobV) then
        StrComm         := TobV.GetValue('COMMENTAIRE');
      EffaceDansTob(TobDRF, tHPB);
  end;

  i                     := HPB.LaGrille.RowCount -1;          // ligne des totaux
  if LeTag = 4 then
    VResultat           := VALEUR(HPB.LaGrille.Cells[2, i])
  else if LeTag = 5 then
    VResultat           := VALEUR(HPB.LaGrille.Cells[3, i])
  else if LeTag = 6 then
  begin
    VResultat           := VALEUR(HPB.LaGrille.Cells[2, i]) - VALEUR(HPB.LaGrille.Cells[3, i]);
  end;

  LeChamp               := 'C' + tHPB;
  Ledit                 := THEdit(GetControl(LeChamp));
  if Assigned(Ledit) then
  begin
    { FQ 20531 BVE 11.06.07 }
    if nbEnreg > 1 then
       LEdit.Enabled := False
    else
       LEdit.Enabled := True;
    {begin
      if Ledit.Text <> '' then
         StrComm        := Ledit.Text
      else
         Ledit.Text     := StrComm;
    end
    else    }
      Ledit.Text        := StrComm;
    { END FQ 20531 }
  end;

  if not AjouteVersTob(tHPB, StrComm, VResultat) then
    PGIInfo('Erreur de sauvegarde des données sur la ligne <' + tHPB +'>', TForm(Ecran).Caption);
  LeChamp               := 'P' + tHPB;
  If Assigned(GetControl(LeChamp)) then
    SetControlText(LeChamp, STRFMONTANT(VResultat, 15, V_PGI.OkDecV, '', True));

  HPBFerme;

end;


procedure TOF_CPRESULTATFISCAL.ChargeFicheDRF;
var
  Tobv                  : Tob;
  LeChamp               : String;
  Ledit                 : THEdit;
  LeCb                  : THCheckbox;
  LaNature              : String;

begin
  // on recherche les QUE totaux.
  TobV                  := TobDRF.FindFirst(['INDICE'], [0], False);
  While TobV <> nil do
  begin
    LaNature            := TobV.GetValue('NATURE');
    LeChamp             := 'P' + LaNature;

    //on ne reprend pas les résultats comptables
    if (Lechamp = 'P2058A_WA_MOA') or (Lechamp = 'P2058A_WS_MOA') then
    begin
      TobV              := TobDRF.FindNext(['INDICE'], [0], False);
      Continue;
    end;

    Ledit               := THEdit(GetControl(Lechamp));
    if Assigned(Ledit) then
      Ledit.Text        := STRFMONTANT(TobV.GetValue('MONTANT'), 15, V_PGI.OkDecV, '', True);

    LeChamp             := 'C' + LaNature;
    Ledit               := THEdit(GetControl(Lechamp));
    if Assigned(Ledit) then
      Ledit.Text        := TobV.GetValue('COMMENTAIRE');



    TobV                := TobDRF.FindNext(['INDICE'], [0], False);
  end;

  // on met a jour les Checkbox pour avertir d'une saisie multiple
  // Il y a au moins 1 en indice
  TobV                  := TobDRF.FindFirst(['INDICE'], [1], False);
  While TobV <> nil do
  begin
    LaNature            := TobV.GetValue('NATURE');
    if (LaNature = '2058B_CL_MOA') or (LaNature = '2058B_DL_MOA') then
      LeChamp           := 'X2058B_CA_MOA'
    else if (LaNature = '2058B_CM_MOA') or (LaNature = '2058B_DM_MOA') then
      LeChamp           := 'X2058B_CD_MOA'
    else if (LaNature = '2058B_CN_MOA') or (LaNature = '2058B_DN_MOA') then
      LeChamp           := 'X2058B_CG_MOA'
    else if LaNature = '2058A_XL_MOA' then
    { FQ 21152 BVE 20.07.07
      LeChamp           := 'X2058A_WX_MOA' }
      LeChamp           := 'X2058A_XB_MOA'
    { END FQ 21152 }
    else if LaNature = '2058A_WN_MOA' then
      LeChamp           := 'X2058A_WZ_MOA'
    else if LaNature = '2058A_WM_MOA' then
      LeChamp           := 'X2058A_WW_MOA'
    else if LaNature = '2058A_VA_MOA' then
      LeChamp           := 'X2058A_WV_MOA'
    else if LaNature = '2058AB_AH_MOA' then
      LeChamp           := 'X2058AB_AQ_MOA'
    else if LaNature = '2058AB_AF_MOA' then
      LeChamp           := 'X2058AB_AU_MOA'
    else
      LeChamp           := 'X' + LaNature;

    LeCb                := THCheckbox(GetControl(LeChamp));
    if Assigned(LeCb) then
      LeCb.Checked      := True;

    LeChamp             := 'P' + LaNature;
    Ledit               := THEdit(GetControl(LeChamp));
    if Assigned(Ledit) then
      Ledit.Enabled     := False;             //on bloque la modification du champ

    TobV                := TobDRF.FindNext(['INDICE'], [1], False);
  end;


end;


procedure TOF_CPRESULTATFISCAL.BvaliderOnClick(Sender: Tobject);
var
  i                     : integer;
  NbC                   : integer;
  LeTag                 : integer;
  Ledit                 : THEdit;
  LaNature              : String;
  LeChamp               : String;
  LeMontant             : extended;
  LeComm                : String;

begin
  Try
    NbC := TForm(Ecran).ComponentCount -1;
    for i := 0 to NbC do
    begin
      LeChamp             := TForm(Ecran).Components[i].Name;
      LeTag               := TForm(Ecran).Components[i].Tag;
      if (LeTag = 1) and (Copy(LeChamp, 1, 1) = 'P') then
      begin
        LaNature          := Copy(LeChamp, 2, Length(LeChamp));
        Ledit             := THEdit(GetControl('P' + LaNature));
        if Assigned(Ledit) then
        begin
          EffaceDansTob(TobDRF, LaNature, True);
          LeMontant       := VALEUR(Ledit.Text);
          Ledit           := THEdit(GetControl('C' + LaNature));
          if Assigned(Ledit) then
            LeComm        := Ledit.Text;
          if not AjouteVersTob(LaNature, LeComm, LeMontant) then
            PGIInfo('Erreur sur l''ecriture <' + LaNature + '>');
        end;
      end;
    end;

    if BSimpl then                // si Régime simplifié alors on ajoute dans la tob
      TraiteSimpl;

    EnregMemoDRF;

  Except
    on E : Exception do
      PGIError(E.Message);
  end;

  TForm(Ecran).Close;
end;

procedure TOF_CPRESULTATFISCAL.EnregMemoDRF;
var
  StrStream             : TStringStream ;

begin
  StrStream             := TStringStream.Create('') ;
  TobDRF.SaveToXmlStream(StrStream, False, True);
  LaTobDRF.SetString('CBN_BLOCNOTE', StrStream.DataString);
  LaTobDRF.SetString('CBN_NATURE', 'DRF');
  LaTobDRF.SetString('CBN_CODE', 'BIC');
  LaTobDRF.InsertOrUpdateDB(False);

  StrStream.Free;


end;


//======================================================================
//Procédures et fonctions pour script DRF

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 10/01/2007
Modifié le ... : 10/01/2007
Description .. : recherche le résultat de l'exercice et retourne false si la
Suite ........ : déclaration n'est pas valable pour l'exercice en cours
Mots clefs ... :
*****************************************************************}
function AGLChercheResultatComptable (parms: array of variant; nb: integer) : Variant ;
var
  SensRC                : String;
  StrExo                : String;
  Q                     : TQuery;
  StrSql                : String;
  MontantRC             : extended;

begin
  Result                := STRFMONTANT(0, 15, V_PGI.OkDecV, '', True);
  SensRc                := string(Parms[0]);
  StrExo                := VH^.EnCours.Code;
  StrSql                := 'SELECT SUM(E_DEBIT)-SUM(E_CREDIT) '+
                           'FROM ECRITURE ' +
                           'LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL ' +
                           'WHERE E_EXERCICE="' + StrExo + '" '+
                           'AND E_QUALIFPIECE="N" '+
                           'AND (G_NATUREGENE="CHA" OR G_NATUREGENE="PRO")';

  Q                     := OpenSQL(StrSql, True);
  try
    MontantRC             := Q.Fields[0].AsFloat;

    if (((SensRc = 'R') and (MontantRC < 0 ))  // Montant créditeur ==> réintégration
          or ((SensRc = 'D') and (MontantRC > 0 ))) then  // Montant débiteur ==> déduction
      Result := STRFMONTANT(Abs(MontantRC), 15, V_PGI.OkDecV, '', True);
  finally
    Ferme(Q);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 11/01/2007
Modifié le ... :   /  /
Description .. : recherche du régime fiscal
Mots clefs ... :
*****************************************************************}
function AGLChercheRegimeFiscal (parms: array of variant; nb: integer) : Variant;
var
  Q                     : TQuery;
  StrSql                : String;

begin
  result                := '';
  StrSql                := 'SELECT DFI_REGIMFISCDIR, DFI_OPTIONRDSUP, DFI_INTEGRAFISC, DFI_IMPODIR, DFI_OPTIONAUTID FROM DPFISCAL' +
                           ' WHERE DFI_GUIDPER = "' + GetGuidPer(V_PGI.NoDossier) + '"'; // FQ 21157 BVE 23.07.07
  Q                     := OpenSQL(StrSql, True);
  if not Q.Eof then
  begin
    if (Q.FindField('DFI_IMPODIR').AsString = 'IS') and (Q.FindField('DFI_OPTIONAUTID').AsString <> 'X') then
      Result            := 'S';               // IS
    if (Q.FindField('DFI_IMPODIR').AsString = 'IR') and (Q.FindField('DFI_OPTIONAUTID').AsString <> 'X') then
      Result            := 'R';               // IR
    if Q.FindField('DFI_REGIMFISCDIR').AsString = 'RN' then
      Result            := Result + 'N';      //réel normal
    if (Q.FindField('DFI_REGIMFISCDIR').AsString = 'RS') and (Q.FindField('DFI_OPTIONRDSUP').AsString = 'X') then
      Result            := Result + 'N';      //réel normal
    if (Q.FindField('DFI_REGIMFISCDIR').AsString = 'RS') and (Q.FindField('DFI_OPTIONRDSUP').AsString = '-') then
      Result            := Result + 'S';      // simplifié
    if Q.FindField('DFI_INTEGRAFISC').AsString = 'X' then
      Result            := Result + 'I'       //intégration fiscale
    else
      Result            := Result + '-';      //sans intégration fiscale
  end;
  Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 24/01/2007
Modifié le ... :   /  /    
Description .. : Cache les composants d'une même nature
Mots clefs ... : 
*****************************************************************}
Procedure AGLRendInvisible (parms: array of variant; nb: integer);
var
  F                     : TForm;
  LaNature              : String;
  LeChamp               : String;
  Ledit                 : THEdit;
  LeLabel               : THlabel;
  LCb                   : THCombobox;
  LeBouton              : TToolbarButton97;

begin
  F                     := TForm(Longint(Parms[0]));
  LaNature              := parms[1];

  LeChamp               := 'P' + LaNature;
  Ledit                 := THEdit(F.FindComponent(LeChamp));
  if Assigned(Ledit) then
  begin
    Ledit.Visible       := False;
    Ledit.Tag           := 0;       // il ne sera pas enregistré
  end;

  LeChamp               := 'C' + LaNature;
  Ledit                 := THEdit(F.FindComponent(LeChamp));
  if Assigned(Ledit) then
  begin
    Ledit.Visible       := False;
    Ledit.Tag           := 0;
  end;

  LeChamp               := 'X' + LaNature;
  LCb                   := THCombobox(F.FindComponent(LeChamp));
  if Assigned(LCb) then
    LCb.Visible       := False;

  LeChamp               := 'B' + LaNature;
  LeBouton              := TToolbarButton97(F.FindComponent(LeChamp));
  if Assigned(LeBouton) then
    LeBouton.Visible       := False;

  LeChamp               := 'T' + LaNature;
  LeLaBel               := THLabel(F.FindComponent(LeChamp));
  if Assigned(LeLabel) then
    LeLabel.Visible     := False;

end;






{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 25/01/2007
Modifié le ... :   /  /    
Description .. : Gère les événements sur les edits de la fiche.
Suite ........ : Selon cellule, on appel un outil
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.LeditOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var
  Ledit                 : THEdit;
  LCb                   : THCheckbox;

begin

  case Key of
    VK_F6 :
      begin
        With Sender as THEdit do
        begin
          if Name = 'P2058A_WB_MOA' then             // rémunération de l'exploitant
          begin
            //appel de la recherche dans la compta
            Ledit       := THEdit(GetControl(Name));
            Ledit.Text  := STRFMONTANT(RemunerationExploitant, 15, V_PGI.OkDecV, '', True);
            if VALEUR(Ledit.Text) <> 0 then
            begin
              EffaceDansTob(TobDRF, '2058A_WB_MOA');
              AjouteVersTob('2058A_WB_MOA', '', VALEUR(Ledit.Text));
              Ledit.Enabled := True;
              LCb         := THCheckbox(GetControl('X2058A_WB_MOA'));
              if Assigned(LCb) then
                LCb.Checked   := False;
            end;

          end;

          if Name = 'P2058A_WG_MOA' then            // taxe sur les voitures particuliere
          begin
            Ledit       := THEdit(GetControl(Name));
            Ledit.Text  := STRFMONTANT(TaxeVoitureParticuliere, 15, V_PGI.OkDecV, '', True);
            EffaceDansTob(TobDRF, '2058A_WG_MOA');  // on efface même les détails
            LCb         := THCheckbox(GetControl('X2058A_WG_MOA'));
            if Assigned(LCb) then
              LCb.Checked   := False;
          end;

          if Name = 'P2058A_WK_MOA' then            // Impot sur société
          begin
            Ledit       := THEdit(GetControl(Name));
            Ledit.Text  := STRFMONTANT(ImpotSurSociete, 15, V_PGI.OkDecV, '', True);
            EffaceDansTob(TobDRF, '2058A_WK_MOA');  // on efface même les détails
            LCb         := THCheckbox(GetControl('X2058A_WK_MOA'));
            if Assigned(LCb) then
              LCb.Checked   := False;
          end;

        end;
      end;

  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 25/01/2007
Modifié le ... :   /  /    
Description .. : Calcul de la rémunération de l'exploitant depuis la compta
Suite ........ : comptes 644 - Comptes 79174
Mots clefs ... :
*****************************************************************}
function TOF_CPRESULTATFISCAL.RemunerationExploitant: Extended;
var
  Tob644                : Tob;
  Tot644D               : extended;
  Tot644C               : extended;
  Tob79174              : Tob;
  Tot79174D             : extended;
  Tot79174C             : extended;
  StrSql                : String;

begin
  StrSql                := 'SELECT G_GENERAL, G_TOTALDEBIT, G_TOTALCREDIT, "X" CALC FROM GENERAUX WHERE G_GENERAL LIKE "644%"';
  Tob644                := Tob.Create('CG 644', nil, -1);
  Tob644.LoadDetailFromSQL(StrSql);
  Tot644D               := Tob644.Somme('G_TOTALDEBIT', ['CALC'], ['X'], True);
  Tot644C               := Tob644.Somme('G_TOTALCREDIT', ['CALC'], ['X'], True);

  StrSql                := 'SELECT G_GENERAL, G_TOTALDEBIT, G_TOTALCREDIT, "X" CALC FROM GENERAUX WHERE G_GENERAL LIKE "79174%"';
  Tob79174              := Tob.Create('CG 79174', nil, -1);
  Tob79174.LoadDetailFromSQL(StrSql);
  Tot79174D             := Tob79174.Somme('G_TOTALDEBIT', ['CALC'], ['X'], True);
  Tot79174C             := Tob79174.Somme('G_TOTALCREDIT', ['CALC'], ['X'], True);



  Result                := (Tot644D - Tot644C) - (Tot79174C - Tot79174D);
  Tob644.Free;
  Tob79174.Free;

end;

{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 25/01/2007
Modifié le ... :   /  /    
Description .. : Calcul de la taxe sur les voitures particulières de la société
Suite ........ : Comptes 63514
Mots clefs ... : 
*****************************************************************}
function TOF_CPRESULTATFISCAL.TaxeVoitureParticuliere: Extended;
var
  StrSql                : String;
  Tob63514              : Tob;
  Tot63514D             : Extended;
  Tot63514C             : Extended;

begin
  StrSql                := 'SELECT E_GENERAL, E_DEBIT, E_CREDIT, "X" CALC FROM ECRITURE WHERE E_EXERCICE ="' + VH^.EnCours.Code + '" ' +
                           'AND E_GENERAL LIKE "63514%"';

  Tob63514              := Tob.Create('CG 63514', nil, -1);
  Tob63514.LoadDetailFromSQL(StrSql);
  Tot63514D             := Tob63514.Somme('E_DEBIT', ['CALC'], ['X'], True);
  Tot63514C             := Tob63514.Somme('E_CREDIT', ['CALC'], ['X'], True);

  Result                := Tot63514D - Tot63514C;
  Tob63514.Free;
end;


{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 25/01/2007
Modifié le ... :   /  /
Description .. : calcul de l'impot sur les sociétés depuis la compta
Suite ........ : comptes 695 + 696 + 697
Mots clefs ... :
*****************************************************************}
function TOF_CPRESULTATFISCAL.ImpotSurSociete: Extended;
var
  StrSql                : String;
  TobIS                 : Tob;
  TotISD                : Extended;
  TotISC                : Extended;

begin
  StrSql                := 'SELECT E_GENERAL, E_DEBIT, E_CREDIT, "X" CALC FROM ECRITURE WHERE E_EXERCICE="' + VH^.EnCours.Code + '" ' +
                           'AND E_GENERAL LIKE "695%" OR E_GENERAL LIKE "696%" OR E_GENERAL LIKE "697%"';
  TobIS                 := Tob.Create('CG IS', nil, -1);
  TobIS.LoadDetailFromSQL(StrSql);
  TotISD                := TobIS.Somme('E_DEBIT', ['CALC'], ['X'], True);
  TotISC                := TobIS.Somme('E_CREDIT', ['CALC'], ['X'], True);

  Result                := TotISD - TotISC;
  TobIS.Free;
end;

procedure TOF_CPRESULTATFISCAL.Calcul_2058A_WI_MOA;
var
  TobV                  : Tob;
  LeTotal               : Extended;
  Ledit                 : THEdit;

begin
  // Calcul pour 2058A_WI_MOA
  LeTotal               := 0;
  TobV                  := TobDRF.FindFirst(['NATURE', 'INDICE'], ['2058B_CA_MOA', 0], False);
  if TobV <> nil then
    LeTotal             := LeTotal + TobV.GetValue('MONTANT');
  TobV                  := TobDRF.FindFirst(['NATURE', 'INDICE'], ['2058B_CD_MOA', 0], False);
  if TobV <> nil then
    LeTotal             := LeTotal + TobV.GetValue('MONTANT');
  TobV                  := TobDRF.FindFirst(['NATURE', 'INDICE'], ['2058B_CG_MOA', 0], False);
  if TobV <> nil then
    LeTotal             := LeTotal + TobV.GetValue('MONTANT');
  EffaceDansTob(TobDRF, '2058A_WI_MOA');
  AjouteVersTob('2058A_WI_MOA', '', LeTotal);
  Ledit                 := THEdit(GetControl('P2058A_WI_MOA'));
  if Assigned(Ledit) then
    Ledit.Text          := STRFMONTANT(LeTotal, 15, V_PGI.OkDecV, '', True);

  // Calcul pour 2058A_WU_MOA
  LeTotal               := 0;
  TobV                  := TobDRF.FindFirst(['NATURE', 'INDICE'], ['2058B_DA_MOA', 0], False);
  if TobV <> nil then
    LeTotal             := LeTotal + TobV.GetValue('MONTANT');
  TobV                  := TobDRF.FindFirst(['NATURE', 'INDICE'], ['2058B_DD_MOA', 0], False);
  if TobV <> nil then
    LeTotal             := LeTotal + TobV.GetValue('MONTANT');
  TobV                  := TobDRF.FindFirst(['NATURE', 'INDICE'], ['2058B_DG_MOA', 0], False);
  if TobV <> nil then
    LeTotal             := LeTotal + TobV.GetValue('MONTANT');
  EffaceDansTob(TobDRF, '2058A_WU_MOA');
  AjouteVersTob('2058A_WU_MOA', '', LeTotal);
  Ledit                 := THEdit(GetControl('P2058A_WU_MOA'));
  if Assigned(Ledit) then
    Ledit.Text          := STRFMONTANT(LeTotal, 15, V_PGI.OkDecV, '', True);


end;



function TOF_CPRESULTATFISCAL.RechPrecExe(ExoPrec: String): Tob;
var
  TobPrec               : Tob;
  i                     : integer;
  TobV1                 : Tob;
  tobV2                 : Tob;
  TobR                  : Tob;

begin
  Result                := nil;
  TobPrec := TOB.Create('DATA_DRF', nil, -1);
  try
    ChargeTobDRF(TobPrec, ExoPrec);
    if Assigned(TobPrec) then
    begin
      TobV1               := Tob.Create('DATA_DRF', nil, -1);
      for i := 0 to TobPrec.Detail.Count -1 do
      begin
        TobR              := TobPrec.Detail[i];
        TobV2             := Tob.Create('Datas', TobV1, -1);
        tobV2.AddChampSupValeur('NATURE', TobR.GetValue('NATURE'));
        tobV2.AddChampSupValeur('INDICE', TobR.GetValue('INDICE'));
        tobV2.AddChampSupValeur('CODE1', TobR.GetValue('CODE1'));
        tobV2.AddChampSupValeur('CODE2', TobR.GetValue('CODE2'));
        tobV2.AddChampSupValeur('COMMENTAIRE', TobR.GetValue('COMMENTAIRE'));
        tobV2.AddChampSupValeur('MONTANT', 0);
        tobV2.AddChampSupValeur('MT2', 0);
        tobV2.AddChampSupValeur('MT3', 0);
        tobV2.AddChampSupValeur('MT4', 0);
        tobV2.AddChampSupValeur('MT5', 0);
      end;
      Result              := TobV1;
    end;
  finally
    TobPrec.Free;
  end;
end;




{***********A.G.L.***********************************************
Auteur  ...... : TJ
Créé le ...... : 08/02/2007
Modifié le ... :   /  /    
Description .. : Traitement des correspondances pour le régime simplifié
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRESULTATFISCAL.TraiteSimpl;
var
  i                     : integer;
  LaNature              : String;
  TobS                  : Tob;
  StrCalcul             : String;
  LaValeur              : extended;

begin
  for i := 0 to TobSimpl.Detail.Count -1 do
  begin
    TobS                := TobSimpl.Detail[i];
    LaNature            := TobS.GetValue('NATURE');
    StrCalcul           := TobS.GetValue('CALCUL');
    EffaceDansTob(TobDRF, LaNature, True);

    LaValeur            := LeCalcul(StrCalcul);

    if not AjouteVersTob(LaNature, '', LaValeur) then
      PGIInfo('Erreur sur l''écriture de la donnée <' + LaNature + '>', TForm(Ecran).Caption);

  end;
end;




procedure TOF_CPRESULTATFISCAL.RefreshLaGrille;
var
  LaCol                 : integer;
  LaRow                 : integer;
  LeCancel              : boolean;

begin
  // Refresh de la grille
  LaCol                 := HPB.LaGrille.Col;
  LaRow                 := HPB.LaGrille.Row;
  LeCancel              := False;
  HPB.LaGrille.OnCellExit(nil, LaCol, LaRow, LeCancel);
end;


Initialization
  registerclasses ( [ TOF_CPRESULTATFISCAL ] ) ;
  RegisterAglFunc('ChercheResultatComptable', False, 1, AGLChercheResultatComptable);
  RegisterAglFunc('ChercheRegimeFiscal', False, 1, AGLChercheRegimeFiscal);
  RegisterAglProc('RendInvisible', True, 1, AGLRendInvisible);
end.
