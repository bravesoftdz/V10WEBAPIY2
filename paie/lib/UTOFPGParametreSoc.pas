{***********A.G.L.***********************************************
Auteur  ...... : PAIE PGI
Créé le ...... : 11/07/2001
Modifié le ... :   /  /
Description .. :
  Cette unité permet de faire des controles de vraisemblance sur les paramsoc
  En entrée on charge une tob "TOB_ParametreSoc" (Query de la table PARAMETRESOC)
  La procedure OnChangeValue(Sender: TObject); est appelée sur les evenements
  des composants..
  (Voir Gestionnaire d'evenement sur onargument : OnChange,OnExit,OnClick)
  Cette procédure appelle  UpdateChampBase qui met à jour la TOB_ParametreSoc
  sur la valeur du champs de l'objet modifié
  Exemple :
    Un ThEdit : procedure OnChangeValue appellé sur le onexit
    En sortie de votre ThEdit la TOB_ParametreSoc sera maj
    sur le champ associé si modification.
  Attention :
    Si vous réaffectez la valeur d'un THEdit par le code..
    L'evenement onexit est bien entendu non appellé
    et donc la TOB_ParametreSoc sera non maj sur ce champ.
    Appellé la procédure UpdateChampBase pour mettre à jour la TOB_ParametreSoc
  Veuiller lire les annotations NB du source..
Mots clefs ... : PAIE;PARAMETRESOC
*****************************************************************}
{
PT1   : 25/07/2001 VG V540 Ajout champs dans Historique salarié
PT2   : 02/10/2001 PH V562 Ajout champ dans Historique salarié Profilrem
PT3-1 : 05/10/2001 VG V562 Modification du nom de l'onglet DADS-TDS => DADS-U
   -2 : 12/10/2001 VG V562 Contrôle de la fraction DADS-U
PT4   : 15/10/2001 VG V562 Ajout champs dans Historique salarié
PT5   : 16/10/2001 SB V562 Ajout des images iconisées sur Treeview
PT6   : 17/10/2001 PH V562 Ajout Traitement champ numéro dossier sisco2
PT7   : 18/10/2001 SB V562 Ajout des boutons imagés sur TRICH
                           Relook des entêtes de pages
                           Suppression de la procedure AffichageImage non
                           utilisée
PT8-1 : 22/11/2001 SB V563 Fiche de bug n°376  Decalage de paie
PT8-2 : 22/11/2001 SB V562 Fiche de bug n°376 Champ libre non accessible
PT9   : 30/11/2001 VG V563 Controle du SIRET DADS
PT10  : 11/01/2002 VG V571 il manquait un begin-end => le SIRET du destinataire
                           du CRE était effacé
PT11  : 27/03/2002 PH V571 Controle des champs creation auto et increment auto
                           en fonction du type du matricule
PT12  : 19/04/2002 SB V571 Fiche de bug n°376 : Correction PT8
PT13-1: 16/05/2002 VG V582 Suppression de l'affichage de l'onglet Historique
PT13-2: 16/05/2002 VG V582 Suppression de l'affichage de certains paramètres
                           pour la version S3 et pour l'assistant de création
PT14  : 27/05/2002 JL V582 Champ service invisible
PT15  : 13/08/2002 JL V582 Ajout onglet formation
PT15-2: 03/10/2002 VG V585 Nouveau paramètre "SO_PGHDADSHISTO" dont la valeur
                           est conditionnée à la valeur de "SO_PGHISTORISATION"
PT16  : 18/10/2002 SB V585 Génération d'évènements pour traçage
PT17  : 30/10/2002 VG V591 Repositionnement provisoire de LTRI et LTRI1
PT18-1: 27/11/2002 SB V591 Affichage des Tabsheet
PT18-2: 27/11/2002 SB V591 Paramétrages des critères ECongés
PT19  : 16/12/2002 JL V591 Ajout champs libres pour formation
PT20  : 26/12/2002 JL V591 Fiche de bug n° 10416 : Gestion accès "matricules
                           salariés et intérimaire unique"
PT21  : 06/01/2003 PH V591 Fiche de bug n° 10404 : Message alerte controle
                           exercice avec décalage de paie
PT22  : 12/03/2003 VG V_42 Ajout de la gestion des acomptes en version S3
PT23  : 17/03/2003 SB V591 Affichage des tabsheets décalés
PT24-1: 23/04/2003 SB V_42 FQ 10610 Intégration de GroupBox, Refonte
                           conditionnement affichage
PT24-2: 23/04/2003 SB V_42 Intégration contrôle gestion des ressource
PT25  : 13/05/2003 MF V_42 Ajout des paramètres société Tickets restaurant
PT26  : 23/06/2003 JL V_42 Ajout champs libres frais fixe pour formation
PT27  : 30/06/2003 JL V_42 Désactivation des modules saisie arret et titre
                           restaurant pour pilotage
PT28  : 22/08/2003 PH V_42 Stockage de la valeur du paramsoc Name+'_' qui n'est
                           pas dans la branche de la paie mais dans la GA
                           (ex: SO_PGLIENRESSOURCE_)
PT29  : 05/09/2003 PH V_42 Accès à la gestion hiéarchique des services si module
                           formation sérialisé
PT30  : 08/09/2003 MF V_42 FQ 10785 Onglet Titres restaurant uniqt qd
                           VH_Paie.PGTicketRestau = true
PT31  : 10/09/2003 SB V_42 Refonte Affichage des images acollées au treeview
PT32  : 30/09/2003 VG V_42 Version S3 5.0
PT33  : 01/10/2003 VG V_42 Assistant de création : On rend visible les
                           paramètres société des gestions spécifiques -
                           FQ N°10414
PT34  : 30/10/2003 SB V_50 FQ 10937 Affichage des boutons en mode affichage XP
PT35  : 18/11/2003 SB V_50 Ajout barre de patience lors de la validation
PT36  : 12/05/2004 PH V_50 FQ 11054 Mise à jour PROFILFNAL salarié en cas de
                           changement
PT37  : 12/05/2004 PH V_50 FQ 11147 Majuscule dans la racine compte de tiers de
                           type salarié
PT38  : 04/08/2004 VG V_50 La coche historisation salarié n'était pas conservée
                           FQ N°11447
PT39  : 23/08/2004 PH V_50 Activation et visinilité des modules si options
                           sérialisées
PT40  : 26/08/2004 VG V_50 Version S3 6.0 - FQ N°11539
PT41  : 08/10/2004 PH V_50 Rechargement du cache serveur pour tenir compte des
                           modifs des paramsoc
PT42  : 08/10/2004 JL V_50 Affichage paramsoc formation si paie non sérialisé
PT43  : 12/10/2004 VG V_50 Masquage de l'onglet Formation - FQ N°11679
PT44  : 15/10/2004 PH V_50 Message alerte si non gestion des auxiliaires alors
                           pas de RIB
PT45  : 15/10/2004 VG V_50 MSA en S3
PT46  : 18/10/2004 VG V_50 ASSEDIC spectacle pas en S3
PT47  : 02/03/2005 JL V_60 Numéro AEM assedic spectacles non visible (modifié
                           dans attestation) -> Suppression 26/04/07 de
                           SO_PGNUMAEMINTER
PT48  : 31/03/2005 SB V_60 FQ 11990 Contrôle sur zone Gestion des absences
PT49  : 21/04/2005 JL V_60 FQ 12173 Formation : si Validation inscriptions n'est
                           plus géré, mettre à jour les enregistrements
PT50  : 01/07/2005 SB V_65 FQ 12213 Ajout message d'info si rép inexistant
PT51  : 05/07/2005 SB V_65 FQ 11945 Ajout message d'info en décoche gestion CP
PT52  : 20/07/2005 JL V_60 FQ 12039 Contrôle accès validation prévisionnel
PT53  : 27/09/2005 VG V_60 Adaptation cahier des charges DADS-U V8R02
PT54  : 29/09/2005 SB V_65 FQ 12287 Gestion coche param soc. gestion des
                           acomptes
PT55  : 29/09/2005 SB V_65 FQ 12528 CWAS : violation accès si modif integer ou
                           double
PT56  : 14/11/2005 VG V_65 Modification du message pour les fractions DADS-U
                           FQ N°12653
PT57  : 08/02/2006 JL V_65 Modification gestion MultiCobo (.text au lieu de
                           .value)
PT58-1: 16/03/2006 PH V_65 Prise ne compte des différents cas des paramètres
                           régionaux différents (Cas du mélange . et ,) dans les
                           nombres.
PT58-2: 20/03/2006 MF V_65 Suite PT58-1. Prise en compte des champs de type
                           double (SOC_DESIGN = "R")
PT59  : 07/04/2006 SB V_65 FQ 12451 Ajout maj SQL table SALARIES sur zone
                           Bulletin par défaut
PT60  : 10/04/2006 PH V_65 Onglet Interface(Anciennement historique) non visible
                           si non PCL ou mode sav
PT60-1: 07/07/2006 PH V_65 FQ 13354,13355 Invible en PCL
PT60-2: 12/07/2006 PH V_65 Prise en compte de la modif en CWAS
PT61  : 16/05/2006 SB V_65 Suppression donnée érroné conservé par une mauvaise
                           socref
PT62  : 05/06/2006 SB V_65 Chargement des images des ttoolbutton par appli.
PT63  : 15/06/2006 PH V_65 Gestion des responsables/organisation visible
PT64  : 12/09/2006 NA V_72 Gestion des IDR
PT65  : 22/01/2007 MF V_72 Titres restaurant : Ajout fournisseur ACCOR
PT66  : 09/02/2007 RM V_72 Ajout Panel DUE dématérialisée pour la gestion de
                           DUE-EDI
PT67  : 28/02/2007 JL V_72 FQ 13929 N° Attestation intermittent spectacle forcer
                           en majuscule + Correction le 31/07/2007
PT67  : 13/03/2007 MF V_72 Ajout TabSheet IJSS & maintien
PT68  : 16/03/2004 JL V_72 Suppression frais fixe formation
PT69  : 13/04/2007 FL V_72 FQ 13568 Limitations des choix de méthodes de calcul
                           entre réel et prévisionnel
PT70  : 26/04/2007 JL V_72 Gestion nouvel historique par avance, proposition
                           d'initialisation des données
PT71  : 10/05/2007 PH V_72 Mise en place concept droit à modification des
                           paramsoc paie
PT72  : 06/07/2007 VG V_72 Ajout paramsoc établissement siège - FQ N°14099
PT73  : 18/07/2007 VG V_72 Correction PT72
PT74  : 14/08/2007 FL V_80 Gestion de la présence - L'onglet "Autres"
                           (anciennement T.R.) est visible même si pas de
                           gestion des TR
PT75  : 24/08/2007 FC V_72 FQ 14369 Pas le bon tag pour test concept
PT76 : 13/09/2007 FL V_800 Prise en compte des paramètres société partagés lors du chargement des données
PT77 : 02/10/2007 NA V_800 Gestion  du module présence
PT78  : 05/11/2007 VG V_80 Adaptation cahier des charges DADS-U V8R06
PT80  : 02/04/2008 FL V_80 Ajout des paramètres formation pour les mails
PT81  : 18/04/2008 FL V_80 Elargissement de la récupération des paramètres formation
PT82  : 29/05/2008 NA V_850 Ticket restaurant : controle code client sur 5 caract si chèque déjeuner  +
                            saisie de la "check box" pour code client/etabl
PT83  : 10/07/2008 NA V_850 FQ 15644 Ne pas rendre invisible le group box "Modules complémentaires" si bilan social & gestion compétencec
                            non sérialisés   et rendre invisible les lignes correspondant à IDR si ce module n'est pas sérialisé
}

unit UTOFPGParametreSoc;

interface
uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls, Spin,
{$IFNDEF EAGLCLIENT}
  Fe_Main,db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
  MaineAGL,
{$ENDIF}
  HCtrls, HEnt1, Ent1, HMsgBox, UTOF, UTOB, HTB97, Imglist, graphics, FileCtrl,
  ParamSoc, HPanel, ParamDat, vierge, Hout, Ed_tools;


type
//Création d'un record caractérisant les tabsheets..
  PGParamSheet = record
    stIndexTree: integer; //index du treeview, calculé à l'affichage pour les noeuds enfants
    stIndexSheet: integer; //PageIndex du tabsheet
    stName: string; //Nom du tabsheet
    stImage: integer; //Numero de l'image icône associé
    StBouton: string; //Nom du bouton du tabsheet parent
    stLibelle: string; //Libellé du noeud treview et du tabsheet
  end;

  TOF_PGPARAMETRESOC = class(TOF)
    procedure OnArgument(stArgument: string); override;
    procedure OnLoad; override; //Valeur par défaut en entrée fiche
    procedure OnClose; override;
    procedure OnUpdate; override;
  private
    TOB_ParametreSoc, TOB_ValeurInit: TOB;
    Loaded, Modifier, UpdateError, Sauve: Boolean;
    TVParamSoc: TTreeview;
    ListeImage: TImageList;
    NumOnglet: string; {VG 16/03/01 Gestion de l'affichage des onglets}
    GblBulDefaut : String;
    procedure ChargeTob;
    procedure ChangeNoeud(Sender: TObject; Node: TTreeNode);
    procedure OnChangeValue(Sender: TObject);
    procedure DateElipsisclick(Sender: TObject);
    procedure GestionLibelle(SpinName, SuffixeEdit: string);
    procedure EnabledChampAxe(acces: Boolean);
    procedure ValueChampAxe(st: string);
    procedure TobValueChampAxe;
    procedure RechargementVH_Paie;
//    procedure enabledchampOngletHistorique;
    procedure UpdateChampBase(Name: string; NewValue: Variant);
    procedure BtnClick(Sender: TObject);
    procedure GenereEvenement(Tob_EnrModifie: Tob);
    procedure AfficheOngletEConges(Afficher: boolean);
    procedure OnEventControl(Control: TControl);
    procedure IDRtauxactu;
    procedure GestionPresence; // PT77
    Function  ControleCouts:Boolean; //PT69
    Procedure LoadSharedParams(TobParam : TOB);  //PT76
  end;

implementation

uses EntPaie, PGoutils, PgOutils2, P5Def, variants, PGOutilsFormation;
//PTabSheet : tableau des caractéristiques Tabsheet
const
{$IFNDEF CCS3}
  {$IFNDEF RMA}
    cNbTabSheet = 11;
  {$ELSE}
    cNbTabSheet = 12;
  {$ENDIF}
{$ELSE}
  cNbTabSheet = 6;
{$ENDIF}

var
  pTabSheet: array[0..cNbTabSheet] of PGParamSheet =
  (
    (StIndexTree: - 1; stIndexSheet: 0;  stName: 'TSRICH';         stImage: 3;  StBouton: 'BTNPAIE';        StLibelle: 'Paie'; ),
    (StIndexTree: - 1; stIndexSheet: 1;  stName: 'TSPARAM';        stImage: 50; StBouton: 'BTNPARAMETRE';   StLibelle: 'Paramètres'; ),
    (StIndexTree: - 1; stIndexSheet: 2;  stName: 'TSSALARIE';      stImage: 78; StBouton: 'BTNSALARIE';     StLibelle: 'Salarié'; ),
    (StIndexTree: - 1; stIndexSheet: 3;  stName: 'TSCARACT';       stImage: 13; StBouton: 'BTNCARACT';      StLibelle: 'Caractéristiques'; ),
    (StIndexTree: - 1; stIndexSheet: 4;  stName: 'TSGENE';         stImage: 66; StBouton: 'BTNCOMPTA';      StLibelle: 'Génération comptable'; ),
    (StIndexTree: - 1; stIndexSheet: 5;  stName: 'TSDADS';         stImage: 76; StBouton: 'BTNDADS';        StLibelle: 'DADS-U'; ),
{$IFNDEF CCS3}
    (StIndexTree: - 1; stIndexSheet: 6;  stName: 'TSPREF';         stImage: 26; StBouton: 'BTNPREFERENCE';  StLibelle: 'Préférences'; ),
    (StIndexTree: - 1; stIndexSheet: 7;  stName: 'PGFORMATION';    stImage: 6;  StBouton: 'BTNFORMATION';   StLibelle: 'Formation'; ),
    (StIndexTree: - 1; stIndexSheet: 8;  stName: 'TTHISTO';        stImage: 54; StBouton: 'BTNHISTO';       StLibelle: 'Interface'; ),
    (StIndexTree: - 1; stIndexSheet: 9;  stName: 'TBSHTELTVAR';    stImage: 16; StBouton: 'BTNCOMPLEMENT';  StLibelle: 'Compléments'; ),
    (StIndexTree: - 1; stIndexSheet: 10; stName: 'PGTICKETRESTAU'; stImage: 85; StBouton: 'BTNTICKETRESTAU';StLibelle: 'Autres'; ) //PT74
  {$IFDEF RMA}
    ,
    (StIndexTree: - 1; stIndexSheet: 11; stName: 'PGDUE';          stImage: 38; StBouton: 'BTNDUE';         StLibelle: 'DUE dématérialisée'; ),
    (StIndexTree: - 1; stIndexSheet: 12; stName: 'PGIJSS';         stImage: 84; StBouton: 'BTNIJSS';         StLibelle: 'IJSS & maintien'; )
  {$ELSE}
    ,
    (StIndexTree: - 1; stIndexSheet: 12; stName: 'PGIJSS';         stImage: 84; StBouton: 'BTNIJSS';         StLibelle: 'IJSS & maintien'; )
  {$ENDIF}
{$ELSE}
    (StIndexTree: - 1; stIndexSheet: 6; stName: 'TSPREF'; stImage: 26; StBouton: 'BTNPREFERENCE'; StLibelle: 'Préférences'; )
{$ENDIF}
    );

{ TOF_PARAMETRESOC }

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : On argument
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.OnArgument(stArgument: string);
var
  TRich: TTabsheet;
  PParam: TPageControl;
  MyTreeNode1, MyTreeNode2, M1: TTreeNode;
  ParentControl, SousParentControl: TWinControl;
  Control: TControl;
  Group: TGroupBox;
  Combo, TypNumSal, Routage1, Routage2: THValComboBox;
  Edit: THEdit;
  Spin: TSpinEdit;
  check: TCheckBox;
  NumEdit: THNumEdit;
  MultiCombo: THMultiValComboBox;
  ChampTrouve, DChamp, St, arg: string;
  i, j, k, l, Nbre, NbTree: integer;
  CEG, STD, DOS, Bool: Boolean;
  Q: TQuery;
  TobTypeChamp: Tob;
  Btn, BtnImage: TToolBarButton97;
  LTri, LTri1: THLabel;
  ImgBitmap : graphics.TBitmap;
begin
  inherited;
  arg := stArgument;
  NumOnglet := Trim(ReadTokenPipe(arg, ';'));
  Sauve := False;
  TRich := TTabsheet(GetControl('TSRICH'));
  if TRich = nil then
    Exit;
  TRich.TabVisible := True;
  TypNumSal := THValComboBox(GetControl('SO_PGTYPENUMSAL'));
  if TypNumSal <> nil then
  begin
    st := 'SELECT COUNT (PSA_SALARIE) FROM SALARIES';
    Q := OpenSQL(st, TRUE);
    Nbre := 0;
    if not Q.EOF then
    begin
      Nbre := Q.Fields[0].AsInteger;
    end;
    Ferme(Q);
    if Nbre >= 1 then
      SetControlEnabled('SO_PGTYPENUMSAL', FALSE);
  end;

  CEG := FALSE;
  STD := FALSE;
  DOS := FALSE;
  AccesPredefini('CEG', CEG, STD, DOS);
  if CEG then
  begin
    SetControlEnabled('SO_PGNUMAUXI', TRUE);
  end;

  PParam := TPageControl(GetControl('PCPARAM'));
  if PParam = nil then
    Exit;
  PParam.ActivePageIndex := 0;
  TVParamSoc := TTreeview(GetControl('TVParam'));

  if TVParamSoc <> nil then
  begin
    BtnImage := TToolBarButton97.Create(Application);
    ListeImage := TImageList.Create(Application);
    if (ListeImage <> nil) and (BtnImage <> nil) then
    begin
      BtnImage.Color := clWindow;
      for i := 0 to 100 do
        if GetBitmapFromIcon(i, BtnImage, FALSE) then
          ListeImage.Add(BtnImage.Glyph, nil);
      TVParamSoc.Images := ListeImage;
      if BtnImage <> nil then
        BtnImage.free;
    end;

    with TVParamSoc.Items do
    begin
      Clear; { Supprime des noeuds existants }
{ Ajoute un noeud à la racine }
      MyTreeNode1 := Add(nil, 'Paramètres Société');
      MyTreeNode1.ImageIndex := 67;
{ Ajoute un noeud enfant au noeud ajouté }{ajout imageindex, le numero correspond au nom de l'icone}
      M1 := AddChild(MyTreeNode1, 'Paie');
      M1.ImageIndex := 3;
{ MyTreeNode2  : 2 noeud }{ et lui ajoute des nœuds enfant }
      MyTreeNode2 := TVParamSoc.Items[1];
      NbTree := 1;
      for i := 0 to cNbTabSheet do
        Begin
        ImgBitmap := graphics.TBitmap.Create;
        HOut.ResIconToBitmap(pTabSheet[i].stImage,ImgBitmap);
        Btn := TToolBarButton97(GetControl(pTabSheet[i].StBouton));
        if Assigned(Btn) then
          begin
          Btn.Glyph := ImgBitmap;
          Btn.caption := pTabSheet[i].stLibelle;
          end;
        FreeAndNil(ImgBitmap);
        if i > 0 then
           begin
           if (VH_Paie.PgSeriaPaie = True) and
              (VH_Paie.PgSeriaFormation = False) and
              (pTabSheet[i].stName = 'PGFORMATION') then
              continue;
           {$IFNDEF PRESENCE} //PT74
           if (VH_Paie.PGTicketRestau = False) and (pTabSheet[i].stName = 'PGTICKETRESTAU') then
              continue;
           {$ENDIF}
           if ((V_PGI.ModePcl <> '1') AND NOT (V_PGI.SAV)) AND (i=8) then
              continue;
           M1 := AddChild(MyTreeNode2, pTabSheet[i].StLibelle);
           M1.ImageIndex := pTabSheet[i].stImage;
           Inc(NbTree, 1);
           pTabSheet[i].stIndexTree := NbTree - 2;
           end;
        End;
    end;

    TVParamSoc.FullExpand;
    TVParamSoc.OnChange := ChangeNoeud;
    TVParamSoc.Selected := TVParamSoc.Items[2];
  end;

  if (VH_Paie.PgSeriaCompetence = False) then
    SetControlVisible('SO_PGGESTIONCARRIERE', FALSE);
  if (VH_Paie.PgSeriaBilan = False) then
    SetControlVisible('SO_PGBILANSOCIALDET', FALSE);
 { pt83 if (VH_Paie.PgSeriaCompetence = False) and
       (VH_Paie.PgSeriaBilan = False) then
        SetControlVisible('SCO_PGMODULESCOMPL', FALSE);   }
  // deb pt83
  if (VH_PAIE.PgSeriaIDR = false) then
  begin
   Setcontrolvisible('SO_PGIDR', false);
   Setcontrolvisible('LSO_PGIDRELT', false);
   Setcontrolvisible('SO_PGIDRELT', false);
  end;
  // fin pt83
  
// Chargement des valeurs de la table ParamSoc
  ChargeTob;
//Evenement On Change
  for j := 0 to PParam.ControlCount - 1 do
  begin { Premier niveau : TabSheet }
    ParentControl := TWinControl(PParam.Controls[j]);
    for i := 0 to ParentControl.ControlCount - 1 do
    begin
      Control := ParentControl.Controls[i];
      ChampTrouve := AnsiUpperCase(Control.Name);
      if (Control is TGroupBox) or (Control is THPanel) then
      begin { Deuxieme niveau : GroupBox ou Panel}
        SousParentControl := TWinControl(Control);
        for K := 0 to SousParentControl.ControlCount - 1 do
        begin { Troisième niveau : sous GroupBox }
          if (SousParentControl.Controls[k]) is TGroupBox then
          begin
            Group := TGroupBox(SousParentControl.Controls[k]);
            for L := 0 to Group.ControlCount - 1 do
              OnEventControl(Group.Controls[L]);
          end
          else
            OnEventControl(SousParentControl.Controls[k]);
        end;
      end
      else
        OnEventControl(Control);
    end;
  end;
//  enabledchampOngletHistorique;

//Gestion d'evenement Btn
  Btn := TToolBarButton97(GetControl('BTNPAIE'));
  if Btn <> nil then
    Btn.OnClick := BtnClick;
  Btn := TToolBarButton97(GetControl('BTNPARAMETRE'));
  if Btn <> nil then
    Btn.OnClick := BtnClick;
  Btn := TToolBarButton97(GetControl('BTNSALARIE'));
  if Btn <> nil then
    Btn.OnClick := BtnClick;
  Btn := TToolBarButton97(GetControl('BTNCARACT'));
  if Btn <> nil then
    Btn.OnClick := BtnClick;
  Btn := TToolBarButton97(GetControl('BTNCOMPTA'));
  if Btn <> nil then
    Btn.OnClick := BtnClick;

{$IFNDEF CCS3}
  if (VH_Paie.PgSeriaPaie = True) and (Vh_Paie.PgSeriaFormation) then
  begin
    Btn := TToolBarButton97(GetControl('BTNFORMATION'));
    if Btn <> nil then
      Btn.OnClick := BtnClick;
  end;
{$ENDIF}

  Btn := TToolBarButton97(GetControl('BTNPREFERENCE'));
  if Btn <> nil then
    Btn.OnClick := BtnClick;
  Btn := TToolBarButton97(GetControl('BTNDADS'));
  if Btn <> nil then
    Btn.OnClick := BtnClick;

  Btn := TToolBarButton97(GetControl('BTNHISTO'));
  if Btn <> nil then
    Btn.OnClick := BtnClick;
  Btn := TToolBarButton97(GetControl('BTNDUE'));
  if Btn <> nil then
  Begin
{$IFNDEF RMA}
    Btn.Enabled := False;
    Btn.visible := False;
    SetControlProperty('PGDUE','TabVisible',False);
    SetControlProperty('PGDUE','TabEnabled',False);
{$ELSE}
    Btn.OnClick := BtnClick;
{$ENDIF}
  End;
  Btn := TToolBarButton97(GetControl('BTNIJSS'));
  if Btn <> nil then
    Btn.OnClick := BtnClick;

{$IFNDEF CCS3}
  Btn := TToolBarButton97(GetControl('BTNCOMPLEMENT'));
  if Btn <> nil then
    Btn.OnClick := BtnClick;

  {$IFNDEF PRESENCE} //PT74
  if VH_Paie.PGTicketRestau then
  begin
  {$ENDIF}
    Btn := TToolBarButton97(GetControl('BTNTICKETRESTAU'));
    if Btn <> nil then
      Btn.OnClick := BtnClick;
  {$IFNDEF PRESENCE} //PT74
  end
  else
  {$ENDIF}
{$ENDIF}
  begin
    Btn := TToolBarButton97(GetControl('BTNTICKETRESTAU'));
    if Btn <> nil then
      Btn.visible := false;
  end;

{$IFDEF CCS3}
  SetControlVisible('SCO_PGSAISIEBULLETIN', FALSE);
  SetControlVisible('SCO_PGSAISIERUBR', FALSE);
  SetControlVisible('SO_PGINTERMITTENTS', FALSE);
  SetControlVisible('SO_PGRESPONSABLES', FALSE);
  SetControlVisible('SO_PGCOEFFEVOSAL', FALSE);
  SetControlVisible('SO_PGSAISIEARRET', FALSE);
  SetControlVisible('SO_PGTICKETRESTAU', FALSE);

  SetControlVisible('SCO_PGBILANSOCIALSIMPL', FALSE);
  SetControlVisible('SCO_PGPARAMANAL', FALSE);
  SetControlVisible('PGTICKETRESTAU', FALSE);
  SetControlVisible('SCO_PGGESTINTERIM', FALSE);
  SetControlVisible('SCO_PGAEMSPECTACLE', False);
  SetControlVisible('TBSHTELTVAR', FALSE);
  SetControlVisible('PGFORMATION', FALSE);
{$ELSE}
  if NumOnglet <> '' then
  begin
    SetControlVisible('SCO_GESTINTERIM', FALSE);
    SetControlVisible('SCO_PGBILANSOCIALSIMPL', FALSE);
    SetControlVisible('SCO_PGACOMPTE', FALSE);
    SetControlVisible('SCO_PGPARAMANAL', FALSE);
  end;
{$ENDIF}

{$IFNDEF CCS3}
  if (VH_Paie.PgSeriaPaie = True) and
    (not VH_Paie.PgeAbsences or not VH_Paie.PgSeriaFormation) then
    SetControlVisible('SO_PGRESPONSABLES', False);
{$ENDIF}

  check := TCheckBox(getcontrol('SO_PGHISTORISATION'));
  if check <> nil then
    Bool := check.checked
  else
    Bool := false;
  Setcontrolenabled('SO_PGHDADSHISTO', Bool);
  SetControlVisible('SO_PGNUMAUXI', FALSE);
  Routage1 := THValComboBox(GetControl('SO_PGROUTAGE1'));
  Routage2 := THValComboBox(GetControl('SO_PGROUTAGE2'));
  LTri := THLabel(GetControl('LTRI'));
  if ((LTri <> nil) and (Routage1 <> nil)) then
    LTri.Top := Routage1.Top;
  LTri1 := THLabel(GetControl('LTRI1'));
  if LTri1 <> nil then
    LTri1.Top := Routage2.Top;

{PT78
  if (GetControlText('SO_PGCOMMCRE') = '03') then
  begin
    SetControlEnabled('LDEST', False);
    SetControlEnabled('SO_PGCIVILCRE', False);
    SetControlEnabled('SO_PGNOMCRE', False);
  end
  else
  begin
    if (GetControlText('SO_PGCOMMCRE') = '05') then
    begin
      SetControlEnabled('LCOORD', False);
      SetControlEnabled('SO_PGCOORDCRE', False);
      if ((GetControlText('SO_PGCOORDCRE') <> '') and
        (GetControlText('SO_PGNOMCRE') = '')) then
        setControlText('SO_PGNOMCRE', GetControlText('SO_PGCOORDCRE'));
    end;
  end;

  if (GetControlText('SO_PGCIVILCRE') = '') then
    setControlText('SO_PGCIVILCRE', 'MR');
}    

  SetControlProperty('SO_PGRTTACQUIS', 'MaxLength', 12);
  SetControlProperty('SO_PGRTTPRIS', 'MaxLength', 12);

  AfficheOngletEConges(VH_Paie.PgeAbsences);
  if V_PGI.DrawXP then
  begin
    if THPanel(GetControl('PANEL')) <> nil then
      THPanel(GetControl('PANEL')).Color := clBtnFace;
    if TToolBarButton97(GetControl('BTNPAIE')) <> nil then
      TToolBarButton97(GetControl('BTNPAIE')).caption := '';
    if TToolBarButton97(GetControl('BTNPARAMETRE')) <> nil then
      TToolBarButton97(GetControl('BTNPARAMETRE')).caption := '';
    if TToolBarButton97(GetControl('BTNSALARIE')) <> nil then
      TToolBarButton97(GetControl('BTNSALARIE')).caption := '';
    if TToolBarButton97(GetControl('BTNCARACT')) <> nil then
      TToolBarButton97(GetControl('BTNCARACT')).caption := '';
    if TToolBarButton97(GetControl('BTNCOMPTA')) <> nil then
      TToolBarButton97(GetControl('BTNCOMPTA')).caption := '';
    if TToolBarButton97(GetControl('BTNFORMATION')) <> nil then
      TToolBarButton97(GetControl('BTNFORMATION')).caption := '';
    if TToolBarButton97(GetControl('BTNHISTO')) <> nil then
      TToolBarButton97(GetControl('BTNHISTO')).caption := '';
    if TToolBarButton97(GetControl('BTNPREFERENCE')) <> nil then
      TToolBarButton97(GetControl('BTNPREFERENCE')).caption := '';
    if TToolBarButton97(GetControl('BTNDADS')) <> nil then
      TToolBarButton97(GetControl('BTNDADS')).caption := '';
    if TToolBarButton97(GetControl('BTNCOMPLEMENT')) <> nil then
      TToolBarButton97(GetControl('BTNCOMPLEMENT')).caption := '';
    if TToolBarButton97(GetControl('BTNTICKETRESTAU')) <> nil then
      TToolBarButton97(GetControl('BTNTICKETRESTAU')).caption := '';
    if TToolBarButton97(GetControl('BTNDUE')) <> nil then
      TToolBarButton97(GetControl('BTNDUE')).caption := '';
    if TToolBarButton97(GetControl('BTNIJSS')) <> nil then
      TToolBarButton97(GetControl('BTNIJSS')).caption := '';
  end;
  SetControlEnabled('SO_PGABSENCE', FALSE);

  SetControlVisible('BTNHISTO', ((V_PGI.ModePcl = '1') OR (V_PGI.SAV)));
  SetControlVisible('SO_PGRESPONSABLES', TRUE);

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : On close
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.OnClose;
begin
  inherited;
  SetFocusControl('PCPARAM');
  if (Modifier = True) and (UpdateError = False) then
    if PGIAsk('Voulez-vous enregistrer les modifications?', 'Paramètres société') = mrYes then
    begin
      {Ajout progressform }
      InitMoveProgressForm(nil, 'Mise à jour des paramètres sociétés ...', 'Veuillez patienter SVP ...', TOB_ParametreSoc.detail.count, FALSE, TRUE);
      Sauve := True;
      if TOB_ParametreSoc <> nil then TOB_ParametreSoc.InsertOrUpdateDB(False);
      RechargementVH_Paie;
      FiniMoveProgressForm;
    end;
  if Sauve then GenereEvenement(TOB_ParametreSoc);
  if TOB_ParametreSoc <> nil then
  begin
    TOB_ParametreSoc.free;
    TOB_ParametreSoc := nil;
  end;
  if TOB_ValeurInit <> nil then
  begin
    TOB_ValeurInit.free;
    TOB_ValeurInit := nil;
  end;
  if ListeImage <> nil then ListeImage.Free;
  TVParamSoc.items.clear;
end;

procedure TOF_PGPARAMETRESOC.ChargeTob;
var
  QSocParam: TQuery;
  Control: TControl;
  Combo: THValComboBox;
  Edit: THEdit;
  Spin: TSpinEdit;
  check: TCheckBox;
  NumEdit: THNumEdit;
  MultiCombo: THMultiValComboBox;
  i: integer;
  Design, TypeChamp: string;
  st: string;
  LaValeur: string;
begin
  inherited;
  QSocParam := OpenSql('SELECT * FROM PARAMSOC ' +
    'WHERE SOC_NOM LIKE "SO_PG%" ORDER BY SOC_NOM', True);
  TOB_ParametreSoc := TOB.Create('Les Paramètres Société', nil, -1);
  TOB_ParametreSoc.LoadDetailDB('PARAMSOC', '', '', QSocParam, FALSE);
  Ferme(QSocParam);

  LoadSharedParams(TOB_ParametreSoc); //PT76

  for i := 0 to TOB_ParametreSoc.detail.count - 1 do
  begin
    Control := TControl(GetControl(TOB_ParametreSoc.detail[i].GetValue('SOC_NOM')));
    if (Control <> nil) and (TOB_ParametreSoc.detail[i].GetValue('SOC_DATA') <> '') then
    begin
      Loaded := True;
      if Control is THEdit then
      begin
        Edit := THEdit(Control);
        Design := TOB_ParametreSoc.detail[i].GetValue('SOC_DESIGN');
        TypeChamp := ReadTokenSt(Design);
        if TypeChamp = 'D' then //Format Date
          Edit.text := DateToStr(TOB_ParametreSoc.detail[i].GetValue('SOC_DATA'))
        else
          Edit.text := TOB_ParametreSoc.detail[i].GetValue('SOC_DATA');
      end;
      if Control is THValComboBox then
      begin
        Combo := THValComboBox(Control);
        Combo.Value := TOB_ParametreSoc.detail[i].GetValue('SOC_DATA');
        if (Combo.name = 'SO_PGJOURNALPAIE') and (Combo.value = '') and (VH_Paie.PgIntegOdpaie = 'ECP') then
          if (TOB_ParametreSoc.detail[i].GetValue('SOC_DATA')) <> '' then
          begin
            st := TOB_ParametreSoc.detail[i].GetValue('SOC_DATA');
            Combo.Text := st;
            Combo.Values.Add(st);
          end;
      end;
      if Control is TCheckBox then
      begin
        Check := TCheckBox(Control);
        Check.checked := (TOB_ParametreSoc.detail[i].GetValue('SOC_DATA') = 'X');
      end;
      if Control is TSpinEdit then
      begin
        Spin := TSpinEdit(Control);
        Spin.value := TOB_ParametreSoc.detail[i].GetValue('SOC_DATA');
      end;
      if Control is THNumEdit then
      begin
        NumEdit := THNumEdit(Control);
        LaValeur := TOB_ParametreSoc.detail[i].GetValue('SOC_DATA');
        THNumEdit(NumEdit).value := StrToFloat(STPOINT(LaValeur));
      end;
      if Control is THMultiValComboBox then
      begin
        MultiCombo := THMultiValComboBox(Control);
        MultiCombo.Text := TOB_ParametreSoc.detail[i].GetValue('SOC_DATA');
      end;
    end;
  end;
end;


procedure TOF_PGPARAMETRESOC.ChangeNoeud(Sender: TObject; Node: TTreeNode);
var
  Niveau, index, i: integer;
  TRich: TTabsheet;
  PParam: TPageControl;
begin
//*Node.SelectedIndex:=Node.ImageIndex;
  TVParamSoc := TTreeview(GetControl('TVPARAM'));
  if TVParamSoc = nil then
    Exit;
  PParam := TPageControl(GetControl('PCPARAM'));
  if PParam = nil then
    Exit;
  TRich := TTabsheet(GetControl('TSRICH'));
  if TRich = nil then
    Exit;
  if TVParamSoc.Selected = nil then
    Exit;
  Niveau := Node.Level;
  case Niveau of //Visibilité des pages
    0..1: begin
        SetControlVisible('BTNPAIE', (Niveau = 0));
        SetControlVisible('BTNPARAMETRE', (Niveau = 1));
        SetControlVisible('BTNSALARIE', (Niveau = 1));
        SetControlVisible('BTNCARACT', (Niveau = 1));
        SetControlVisible('BTNCOMPTA', (Niveau = 1));
        SetControlVisible('BTNHISTO', (Niveau = 1) AND ((V_PGI.ModePcl = '1') OR (V_PGI.SAV)));
        SetControlVisible('BTNPREFERENCE', (Niveau = 1));
        SetControlVisible('BTNDADS', (Niveau = 1));
        SetControlVisible('BTNDUE', (Niveau = 1));
        SetControlVisible('BTNIJSS', (Niveau = 1));
{$IFNDEF CCS3}
        SetControlVisible('BTNCOMPLEMENT', (Niveau = 1));
        if (VH_Paie.PgSeriaPaie = True) and (VH_Paie.PgSeriaFormation) then
          SetControlVisible('BTNFORMATION', (Niveau = 1));
        {$IFNDEF PRESENCE} //PT74
        if VH_Paie.PGTicketRestau then
        {$ENDIF}
          SetControlVisible('BTNTICKETRESTAU', (Niveau = 1));
{$ENDIF}
        for i := 0 to PParam.PageCount - 1 do
          PParam.Pages[i].TabVisible := False;
        TRich.TabVisible := True;
        if Niveau = 0 then
          TRich.Caption := 'Paramètres Société'
        else
          TRich.Caption := 'Paie';
      end;
    2: begin
        index := Node.Index;
{Utilisation boucle et tableau}
        for i := 0 to cNbTabSheet do
          if pTabSheet[i].stIndexTree <> index then
            PParam.Pages[pTabSheet[i].stIndexSheet].TabVisible := False
          else
            PParam.Pages[pTabSheet[i].stIndexSheet].TabVisible := True;
      end;
  end;

  {VG 16/03/01 Gestion de l'affichage des onglets}
  if (NumOnglet <> '0') and (NumOnglet <> '') then
  begin
    TVParamSoc.Visible := False;
    for i := 0 to PParam.PageCount - 1 do
      PParam.Pages[i].TabVisible := False;
    PParam.ActivePageIndex := (StrToInt(NumOnglet) - 1);
    PParam.Pages[StrToInt(NumOnglet) - 1].TabVisible := True;
  end;
  {FIN VG 16/03/01 Gestion de l'affichage des onglets}
end;


procedure TOF_PGPARAMETRESOC.OnChangeValue(Sender: TObject);
var
  name, Vide, StVal, Num: string;
  value, NewValue: variant;
  Combo: THValComboBox;
  MultiCombo: THMultiValComboBox; //NB :Attention si vous changez la valeur d'un champ de la base
  Edit: THEdit; //vous devez faire un UpdateChampBase afin de mettre à jour la tob..
  Spin: TSpinEdit;
  check: TCheckBox;
  NumEdit: THNumEdit;
  Error, long: integer;
  QRechFraction: TQuery;
  Bool: boolean;
  TInit : TOB;
  St, CodeClient : string;
  Q: TQuery;
begin
  Vide := '';
  name := '';
  value := '';
  NewValue := '';
  Error := 0;
  Edit := nil;
  Combo := nil;
  Check := nil;
  if Sender is THValComboBox then
  begin
    Combo := THValComboBox(Sender);
    if Combo = nil then exit; name := Combo.Name;
    NewValue := Combo.value;
  end
  else
    if Sender is THEdit then
    begin
      Edit := THEdit(Sender);
      if Edit = nil then exit; name := Edit.Name;
      NewValue := Edit.text;
    end
    else
      if Sender is TCheckBox then
      begin
        Check := TCheckBox(Sender);
        if Check = nil then exit; name := Check.Name;
        if check.checked = True then
          NewValue := 'X'
        else
          NewValue := '-';
      end
      else
        if Sender is TSpinEdit then
        begin
          Spin := TSpinEdit(Sender);
          if Spin = nil then exit; name := Spin.Name;
          if Spin.VAlue <> 0 then
            NewValue := Spin.value
          else
            NewValue := 0;
        end
        else
          if Sender is THNumEdit then
          begin
            NumEdit := THNumEdit(Sender);
            if NumEdit = nil then exit; name := NumEdit.Name;
            NewValue := NumEdit.value;
          end
          else
            if Sender is THMultiValComboBox then
            begin
              MultiCombo := THMultiValComboBox(Sender);
              if MultiCombo = nil then exit; name := MultiCombo.Name;
              NewValue := MultiCombo.Text;
            end;

  if Name = 'SO_PGABSENCE' then
  begin
    SetControlEnabled('SO_PGGESTIJSS', GetControlText('SO_PGABSENCE') = 'X');
    SetControlEnabled('SO_PGMAINTIEN', GetControlText('SO_PGABSENCE') = 'X');
    SetControlEnabled('SO_PGCRITMAINTIEN', GetControlText('SO_PGABSENCE') = 'X');
    if GetControlText('SO_PGABSENCE') = '-' then
      if PgiAsk('Etes-vous sûr de vouloir désactiver la gestion des absences.',
        Ecran.caption) = MrYes then
      begin
        PgiInfo('Attention. La gestion des IJSS va être désactivée.',
          Ecran.Caption);
        SetControlText('SO_PGGESTIJSS', '-');
        SetControlText('SO_PGMAINTIEN', '-');
        SetControlText('SO_PGCRITMAINTIEN', '');
      end;
  end;

  if Name = 'SO_PGNBRESTATORG' then
    GestionLibelle(Name, 'SO_PGLIBORGSTAT');

  if Pos('SO_PGLIBORGSTAT', Name) = 1 then
  begin
    Num := Copy(name, Length(name), 1);
    SetControlEnabled('SO_PGEDITORG' + Num, (Edit.text <> ''));
    if edit.text = '' then
      SetControlChecked('SO_PGEDITORG' + Num, False);
  end;

  if Name = 'SO_PGLIBCODESTAT' then
  begin
    SetControlEnabled('SO_PGEDITCODESTAT', (Edit.text <> ''));
    if edit.text = '' then
      SetControlChecked('SO_PGEDITCODESTAT', False);
  end;

  if Name = 'SO_PGNBSALLIB' then
  begin
    GestionLibelle(Name, 'SO_PGSALLIB');
    SetControlEnabled('SO_PGTYPSALLIB1', (GetControlText('SO_PGSALLIB1') <> ''));
    SetControlEnabled('SO_PGTYPSALLIB2', (GetControlText('SO_PGSALLIB2') <> ''));
    SetControlEnabled('SO_PGTYPSALLIB3', (GetControlText('SO_PGSALLIB3') <> ''));
    SetControlEnabled('SO_PGTYPSALLIB4', (GetControlText('SO_PGSALLIB4') <> ''));
    SetControlEnabled('SO_PGTYPSALLIB5', (GetControlText('SO_PGSALLIB5') <> ''));

    SetControlChecked('SO_PGTYPSALLIB1', (GetControlText('SO_PGSALLIB1') <> ''));
    SetControlChecked('SO_PGTYPSALLIB2', (GetControlText('SO_PGSALLIB2') <> ''));
    SetControlChecked('SO_PGTYPSALLIB3', (GetControlText('SO_PGSALLIB3') <> ''));
    SetControlChecked('SO_PGTYPSALLIB4', (GetControlText('SO_PGSALLIB4') <> ''));
    SetControlChecked('SO_PGTYPSALLIB5', (GetControlText('SO_PGSALLIB5') <> ''));
  end;

  if Name = 'SO_PGNBCOMBO' then
    GestionLibelle(Name, 'SO_PGLIBCOMBO');

  if Name = 'SO_PGNBCOCHE' then
    GestionLibelle(Name, 'SO_PGLIBCOCHE');

  if Name = 'SO_PGNBDATE' then
    GestionLibelle(Name, 'SO_PGLIBDATE');

  if Name = 'SO_PGTYPENUMSAL' then
  begin
    if Combo.value <> 'NUM' then
    begin
      SetControlChecked('SO_PGINCSALARIE', False);
      SetControlEnabled('SO_PGINCSALARIE', False);
      SetControlChecked('SO_PGTIERSAUXIAUTO', False);
      SetControlEnabled('SO_PGTIERSAUXIAUTO', False);
    end
    else
    begin
      SetControlEnabled('SO_PGINCSALARIE', True);
      SetControlEnabled('SO_PGTIERSAUXIAUTO', True);
    end;
  end;

  if Name = 'SO_PGSECTIONANAL' then
  begin
    EnabledChampAxe(Check.checked);
    if Check.checked = False then
    begin
      ValueChampAxe(Vide); // mise à blanc champ axe
    end
    else
    begin
      TobValueChampAxe; //réinit valeur champ axe
    end;
  end;

  if (Pos('SO_PGCUMUL0', Name) = 1) or (Pos('SO_PGCUMULMOIS', Name) = 1) or
    (Pos('SO_PGCGACQ', Name) = 1) or (Pos('SO_PGCGPRIS', Name) = 1) or
    (Pos('SO_PGREDREM', Name) = 1) or (Pos('SO_PGREDHEURE', Name) = 1) or
    (Pos('SO_PGREDREPAS', Name) = 1) then
  begin
    if (GetControlText(Name) <> '') and (edit <> nil) then
      if Rechdom(Edit.datatype, GetControlText(Name), False) = '' then
      begin
        PgiBox('La valeur saisie n''existe pas.', 'Saisie erronée : ' +
          GetControlText(Name));
        SetControlText(Name, '');
        UpdateChampBase(Name, '');
      end;
  end;

//Cumul du mois bas de bulletin
//Test valeur tablette PGTYPECUMULMOIS pour affichage tablette champ PGCUMULMOIS?
  if Pos('SO_PGTYPECUMULMOIS', Name) = 1 then
  begin
    StVal := GetControltext(Name);
    Long := Length(name);
    Num := Copy(name, Long, 1);
    if StVal = '' then
    begin
      SetControlProperty('SO_PGCUMULMOIS' + Num, 'Datatype', '');
      SetControlProperty('SO_PGCUMULMOIS' + Num, 'Text', '');
      SetControlEnabled('SO_PGCUMULMOIS' + Num, False);
      SetControlEnabled('SO_PGLIBCUMULMOIS' + Num, False);
      SetControlProperty('SO_PGLIBCUMULMOIS' + Num, 'Text', '');
      UpdateChampBase('SO_PGCUMULMOIS' + Num, '');
      UpdateChampBase('SO_PGLIBCUMULMOIS' + Num, '');
    end
    else
    begin
      SetControlEnabled('SO_PGCUMULMOIS' + Num, True);
    end; //SetControlEnabled('SO_PGLIBCUMULMOIS'+Num,True); End;
    if (StVal = 'CMO') then
      SetControlProperty('SO_PGCUMULMOIS' + Num, 'Datatype', 'PGCUMULPAIE');
    if (StVal = '05') or (StVal = '06') or (StVal = '07') or (StVal = '08') then
      SetControlProperty('SO_PGCUMULMOIS' + Num, 'Datatype', 'PGREMUNERATION');
  end;

  if (Pos('SO_PGCUMULMOIS', Name) = 1) and (Edit <> nil) then
  begin
    Long := Length(name);
    Num := Copy(name, Long, 1);
    SetControlText('SO_PGLIBCUMULMOIS' + Num,
      RechDom(Edit.DataType, GetControlText(Name), False));
    UpdateChampBase('SO_PGLIBCUMULMOIS' + Num, RechDom(Edit.DataType,
      GetControlText(Name), False));
    if GetControlText(Name) <> '' then
      SetControlEnabled('SO_PGLIBCUMULMOIS' + Num, True)
    else
      SetControlEnabled('SO_PGLIBCUMULMOIS' + Num, False);
  end;

// Zone Etat justificatif des réductions
  if Name = 'SO_PGTYPREDREPAS' then
  begin
    StVal := GetControltext('SO_PGTYPREDREPAS');
    if StVal = '' then
    begin
      SetControlProperty('SO_PGREDREPAS', 'Text', '');
      SetControlEnabled('SO_PGREDREPAS', False);
    end
    else
      SetControlEnabled('SO_PGREDREPAS', True);
  end;

  if Name = 'SO_PGFRACTION' then
  begin
    StVal := GetControltext('SO_PGFRACTION');
    if ((StVal < '1') or (StVal > '9')) then
      PGIBox('Valeur incorrecte', 'Fraction DADS')
    else
    begin
      Num := '1';
      QRechFraction := OpenSql('SELECT MAX(PSA_DADSFRACTION) FROM SALARIES', True);
      if not QRechFraction.EOF then
        Num := QRechFraction.Fields[0].asstring;
      Ferme(QRechFraction);
      if (StVal < Num) then
      begin
        PGIBox('Des salariés utilisent un code fraction#13#10' +
          'supérieur à cette valeur', 'Fraction DADS');
        SetControlText('SO_PGFRACTION', Num);
        SetFocusControl('SO_PGFRACTION');
      end;
    end;
  end;

  if Name = 'SO_PGDECALAGE' then
    if GetControlText('SO_PGDECALAGE') = 'X' then
    begin
      SetControlEnabled('SO_PGDECALAGEPETIT', False);
      SetControlChecked('SO_PGDECALAGEPETIT', False);
    end
    else
      SetControlEnabled('SO_PGDECALAGEPETIT', True);

  if Name = 'SO_PGDECALAGEPETIT' then
    if GetControlText('SO_PGDECALAGEPETIT') = 'X' then
    begin
      SetControlEnabled('SO_PGDECALAGE', False);
      SetControlChecked('SO_PGDECALAGE', False);
    end
    else
      SetControlEnabled('SO_PGDECALAGE', True);

  if Name = 'SO_PGSIRETDESTIN' then
  begin
    StVal := GetControltext('SO_PGSIRETDESTIN');
    if ControlSiret(StVal) = False then
    begin
      PGIBox('Le SIRET n''est pas valide', 'SIRET Destinataire');
      SetControlText(Name, '');
      UpdateChampBase(Name, '');
    end;
  end;

  if Name = 'SO_PGECABGESTIONMAIL' then
  begin
    SetControlEnabled('SO_PGECABVALIDITEMAIL',
      (GetControlText('SO_PGECABGESTIONMAIL') = 'X'));
  end;

  if (Name = 'MODIFPARAM') then
  begin
//   enabledchampOngletHistorique;
    exit;
  end;

  if (Name = 'SO_PGHISTORISATION') then
  begin
//   enabledchampOngletHistorique;
    check := TCheckBox(getcontrol('SO_PGHISTORISATION'));
    if check <> nil then
      Bool := check.checked
    else
      Bool := false;
    check := TCheckBox(getcontrol('SO_PGHDADSHISTO'));
    if (Bool = False) then
      check.checked := Bool;
    Setcontrolenabled('SO_PGHDADSHISTO', Bool);
  end;

  if Name = 'SO_PGFNBFORLIBRE' then
  begin
    GestionLibelle(Name, 'SO_PGFFORLIBRE');
  end;

  if Name = 'SO_PGTYPECODESTAGE' then
  begin
    if Combo.value <> 'NUM' then
    begin
      SetControlChecked('SO_PGSTAGEAUTO', False);
      SetControlEnabled('SO_PGSTAGEAUTO', False);
    end
    else
      SetControlEnabled('SO_PGSTAGEAUTO', True);
  end;

  if Name = 'SO_PGLIENRESSOURCE' then
  begin
    if not check.checked then
      if PgiAsk('Attention, voulez-vous réellement déactiver la gestion des ' +
        'ressources dans la paie?', Ecran.caption) = MrNo then
        exit;
    SetControlEnabled('SO_PGTYPEAFFECTRES', check.checked);
    SetControlEnabled('SO_PGNBCARNOMRES', check.checked);
    SetControlEnabled('SO_PGNBCARPRENRES', check.checked);
  end;

  if Name = 'SO_PGTIERSAUXIAUTO' then
  begin
    check := TCheckBox(getcontrol('SO_PGTIERSAUXIAUTO'));
    if check <> nil then
    begin
      if not Check.Checked then
        PGIBox('Attention, si vous ne gérez pas les auxiliaires,' +
          ' vous ne pourrez pas créer les RIB salariés,#13#10' +
          'Donc faire les virements de salaires.',
          'Gestion des auxiliaires salariés');
    end;
  end;

  if Name = 'SO_PGTYPEAFFECTRES' then
  begin
    Bool := ((GetControlText('SO_PGTYPEAFFECTRES') = 'RNP') or
      (GetControlText('SO_PGTYPEAFFECTRES') = 'RPN'));
    SetControlEnabled('SO_PGNBCARNOMRES', Bool);
    SetControlEnabled('SO_PGNBCARPRENRES', Bool);
  end;

  // Contrôle des méthodes de calcul appliquées si la valorisation est catégorielle
  If (Name = 'SO_PGFORMETHODECALC') Or (Name = 'SO_PGFORMETHODECALCPREV') Or
     (Name = 'SO_PGFORVALOSALAIRE') Or (Name = 'SO_PGFORVALOSALAIREPREV') Then
     ControleCouts;  //PT69

{**** MISE A JOUR DU CHAMP MODIFIE : LAISSER EN FIN DE PROCEDURE ***}
  if Name = 'SO_PGPROFILFNAL' then
    ExecuteSQL('UPDATE SALARIES SET PSA_PROFILFNAL="' +
      GetControlText('SO_PGPROFILFNAL') + '" WHERE' +
      ' PSA_TYPPROFILFNAL="DOS"');

  if Name = 'SO_PGRACINEAUXI' then
  begin
    StVal := PGUpperCase(GetControlText('SO_PGRACINEAUXI'), TRUE);
    SetControlText('SO_PGRACINEAUXI', StVal);
  end;

  if Name = 'SO_PGGESTACPT' then
  begin
    if GetControlText('SO_PGGESTACPT') = '-' then
    begin
      SetControlText('SO_PGRUBACOMPTE', '');
      UpdateChampBase('SO_PGRUBACOMPTE', '');
    end;
    SetControlEnabled('SO_PGRUBACOMPTE', (GetControlText('SO_PGGESTACPT') = 'X'));
  end;

  if Name = 'SO_PGCONGES' then
  begin
    if GetControlText('SO_PGCONGES') = '-' then
    begin
      if ExisteSql('SELECT * FROM ETABCOMPL WHERE ETB_CONGESPAYES="X"') then
        PgiInfo('Attention! La gestion des congés payés est active pour' +
          ' certains établissements.', Ecran.caption);
    end;
  end;


  if Name = 'SO_PGIDR' then Idrtauxactu;

  if Name = 'SO_PGMODULEPRESENCE' then Gestionpresence; // PT77

   if Error = 0 then
    UpdateChampBase(Name, NewValue);

  if Name = 'SO_PGFORMPREVISIONNEL' then
  begin
    SetControlEnabled('SO_PGFORMVALIDPREV',
      GetControltext('SO_PGFORMPREVISIONNEL') = 'X');
  end;

{PT78
  if (Name = 'SO_PGCOMMCRE') then
  begin
    if (GetControlText('SO_PGCOMMCRE') = '03') then
    begin
      SetControlEnabled('LCOORD', True);
      SetControlEnabled('SO_PGCOORDCRE', True);
      SetControlEnabled('LDEST', False);
      SetControlEnabled('SO_PGCIVILCRE', False);
      SetControlEnabled('SO_PGNOMCRE', False);
    end
    else
    begin
      if (GetControlText('SO_PGCOMMCRE') = '05') then
      begin
        SetControlEnabled('LCOORD', False);
        SetControlEnabled('SO_PGCOORDCRE', False);
        SetControlEnabled('LDEST', True);
        SetControlEnabled('SO_PGCIVILCRE', True);
        SetControlEnabled('SO_PGNOMCRE', True);
      end;
    end;
  end;
}  

  if ((Name = 'SO_PGMSA') and (V_PGI.NumVersionBase < 800)) then
  begin
    if (GetControlText('SO_PGMSA') = 'X') then
    // Choix dossier de type MSA incompatible avec choix fournisseur Titres restaurant = ACCOR  (spécif V7)
      if (GetControlText('SO_PGTYPECDETICKET') = '003') then
       begin
        PgiBox('Choix impossible car fournisseur titres resaturant  = ACCOR', 'MSA ');
        SetControlChecked('SO_PGMSA', false);
        SetFocusControl('SO_PGMSA');
       end;
  end;

  if (Name = 'SO_PGTYPECDETICKET') then
  begin
  // deb pt82
    setcontrolvisible('SO_PGFACTETABL', false);
    setcontrolenabled('SO_PGFACTETABL', false);
  // fin pt82

    if (GetControlText('SO_PGTYPECDETICKET') = '003') then
    // ACCOR
    begin
      if ((GetControlText('SO_PGMSA') = 'X') and (V_PGI.NumVersionBase < 800)) then
      // Choix fournisseur incompatible avec dossier de type MSA   (spécif V7)
      begin
        PgiBox('Dossier MSA, Le choix du fournisseur ACCOR est impossible.', 'Fournisseur ');

        TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGTYPECDETICKET'], True);
        if TInit <> nil then
          // Récupération  de la valeur initiale
          if (TInit.GetValue('SOC_DATA') = '003') then
            // on force à blanc pour ne pas boucler sur le contrôle
            SetControlProperty('SO_PGTYPECDETICKET', 'value', '')
          else
            SetControlProperty('SO_PGTYPECDETICKET', 'value', TInit.GetValue('SOC_DATA'));
        SetFocusControl('SO_PGTYPECDETICKET');
      end
      else
      // paramétrage spécifique ACCOR
      begin
        TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGTYPECDETICKET'], True);
        if TInit <> nil then
          if (TInit.GetValue('SOC_DATA') <> '003') then
          begin
            SetControlEnabled('SO_PGTKEDNPCARNET', True);
            SetControlEnabled('SO_PGTKEDRSCARNET', True);
            SetControlEnabled('SO_PGTKEDCPVCARNET', True);
            SetControlEnabled('SO_PGTKDATELIVR', True);
            SetControlChecked ('SO_PGPERSOTICKET', True);
            SetControlChecked('SO_PGTKEDNPCARNET', false);
            SetControlChecked('SO_PGTKEDRSCARNET', false);
            SetControlChecked('SO_PGTKEDCPVCARNET', false);
            SetControlChecked('SO_PGTKDATELIVR', false);
          end;
      end;
      CodeClient := GetParamSocSecur('SO_PGCODECLIENT','') ;
      if ((length(CodeClient) > 6) or
         (not IsNumeric(CodeClient))) then
      begin
        PgiBox('Le code client doit être numérique sur 6 caractères', 'Fournisseur ACCOR');
        SetFocusControl('SO_PGCODECLIENT');
      end;
      CodeClient := GetParamSocSecur('SO_PGCODERATTACH','') ;
      if ((length(CodeClient) > 6) or
         (not IsNumeric(CodeClient))) then
      begin
        PgiBox('Le code client de rattachement doit être numérique sur 6 caractères', 'Fournisseur ACCOR');
        SetFocusControl('SO_PGCODERATTACH');
      end;
   end
    else
    // SODEXHO, NATEXIS
    begin
      SetControlEnabled('SO_PGTKEDNPCARNET', False);
      SetControlEnabled('SO_PGTKEDRSCARNET', False);
      SetControlEnabled('SO_PGTKEDCPVCARNET', False);
      SetControlEnabled('SO_PGTKDATELIVR', False);

      if (GetControlText('SO_PGTYPECDETICKET') = '001') then
      begin
        CodeClient := GetParamSocSecur('SO_PGCODECLIENT','') ;
        if ((length(CodeClient) > 8) or
           (not IsNumeric(CodeClient))) then
        begin
          PgiBox('Le code client doit être numérique sur 8 caractères', 'SODEXHO');
          SetFocusControl('SO_PGCODECLIENT');
        end;
        CodeClient := GetParamSocSecur('SO_PGCODERATTACH','') ;
        if ((length(CodeClient) > 8) or
           (not IsNumeric(CodeClient))) then
        begin
          PgiBox('Le code client de rattachement doit être numérique sur 8 caractères', 'SODEXHO');
          SetFocusControl('SO_PGCODERATTACH');
        end;
      end
      else
      if (GetControlText('SO_PGTYPECDETICKET') = '002') then
      begin
        CodeClient := GetControlText('SO_PGCODECLIENT') ;
        if (length(CodeClient) > 7) then
        begin
          PgiBox('Le code client doit être de 7 caractères', 'NATEXIS');
          SetFocusControl('SO_PGCODECLIENT');
        end;
        CodeClient := GetControlText('SO_PGCODERATTACH') ;
        if (length(CodeClient) > 7) then
        begin
          PgiBox('Le code client de rattachement doit être de 7 caractères', 'NATEXIS');
          SetFocusControl('SO_PGCODERATTACH');
        end;
      // pt82 end;
     // deb pt82
     end
     else
       if (GetControlText('SO_PGTYPECDETICKET') = '004') then // Chéque déjeuner
       begin
        CodeClient := GetControlText('SO_PGCODECLIENT') ;
        if (length(CodeClient) > 5) then
        begin
          PgiBox('Le code client doit être de 5 caractères', 'CHEQUE DEJEUNER');
          SetFocusControl('SO_PGCODECLIENT');
        end;
        CodeClient := GetControlText('SO_PGCODERATTACH') ;
        if (length(CodeClient) > 7) then
        begin
          PgiBox('Le code client de rattachement doit être de 5 caractères', 'CHEQUE DEJEUNER');
          SetFocusControl('SO_PGCODERATTACH');
        end;
        setcontrolvisible('SO_PGFACTETABL', true);
        setcontrolenabled('SO_PGFACTETABL', true);
       end;
       // fin pt82
    end;
  end;

  if (Name = 'SO_PGCODECLIENT') then
  begin
    if (GetControlText('SO_PGTYPECDETICKET') = '001') then
    begin
      CodeClient := GetControlText('SO_PGCODECLIENT') ;
      if ((length(CodeClient) > 8) or
         (not IsNumeric(CodeClient))) then
      begin
        PgiBox('Le code client doit être numérique sur 8 caractères', 'SODEXHO');
        SetFocusControl('SO_PGCODECLIENT');
      end;
    end
    else
    if (GetControlText('SO_PGTYPECDETICKET') = '002') then
    begin
      CodeClient := GetControlText('SO_PGCODECLIENT') ;
      if (length(CodeClient) > 7) then
      begin
        PgiBox('Le code client doit être de 7 caractères', 'NATEXIS');
        SetFocusControl('SO_PGCODECLIENT');
      end;
    end
    else
    if (GetControlText('SO_PGTYPECDETICKET') = '003') then
    begin
      CodeClient := GetControlText('SO_PGCODECLIENT') ;
      if ((length(CodeClient) > 6) or
         (not IsNumeric(CodeClient))) then
      begin
        PgiBox('Le code client doit être numérique sur 6 caractères', 'ACCOR');
        SetFocusControl('SO_PGCODECLIENT');
      end;
   // pt82 end;
    // deb pt82
    end
    else
     if (GetControlText('SO_PGTYPECDETICKET') = '004') then
     // chèque déjeuner
     begin
      CodeClient := GetControlText('SO_PGCODECLIENT') ;
      if ((length(CodeClient) > 5) or
         (not IsNumeric(CodeClient))) then
      begin
        PgiBox('Le code client doit être numérique sur 5 caractères', 'CHEQUE DEJEUNER');
        SetFocusControl('SO_PGCODECLIENT');
      end;
     end;
     // fin pt82
  end;
  if (Name = 'SO_PGCODERATTACH') then
  begin
    if (GetControlText('SO_PGTYPECDETICKET') = '001') then
    begin
      CodeClient := GetControlText('SO_PGCODERATTACH') ;
      if ((length(CodeClient) > 8) or
         (not IsNumeric(CodeClient))) then
      begin
        PgiBox('Le code client de rattachement doit être numérique sur 8 caractères', 'SODEXHO');
        SetFocusControl('SO_PGCODERATTACH');
      end;
    end
    else
    if (GetControlText('SO_PGTYPECDETICKET') = '002') then
    begin
      CodeClient := GetControlText('SO_PGCODERATTACH') ;
      if (length(CodeClient) > 7) then
      begin
        PgiBox('Le code client de rattachement doit être de 7 caractères', 'NATEXIS');
        SetFocusControl('SO_PGCODERATTACH');
      end;
    end
    else
    if (GetControlText('SO_PGTYPECDETICKET') = '003') then
    begin
      CodeClient := GetControlText('SO_PGCODERATTACH') ;
      if ((length(CodeClient) > 6) or
         (not IsNumeric(CodeClient))) then
      begin
        PgiBox('Le code client de rattachement doit être numérique sur 6 caractères', 'ACCOR');
        SetFocusControl('SO_PGCODERATTACH');
      end;
   // pt82 end;
   // deb pt82
    end
    else
    if (GetControlText('SO_PGTYPECDETICKET') = '004') then
    // chèque déjeuner
    begin
      CodeClient := GetControlText('SO_PGCODERATTACH') ;
      if ((length(CodeClient) > 5) or
         (not IsNumeric(CodeClient))) then
      begin
        PgiBox('Le code client de rattachement doit être numérique sur 5 caractères', 'CHEQUE DEJEUNER');
        SetFocusControl('SO_PGCODERATTACH');
      end;
    end;
    // fin pt82
  end;

//PT72
  if (Name='SO_PGCSIEGE') then
     begin
     St:= 'SELECT ET_ETABLISSEMENT FROM ETABLISS';
     Q:= OpenSQL (St, TRUE);
     if (not Q.EOF) then
        StVal:= Q.Fields[0].AsString
     else
        StVal:= '';
     Ferme (Q);

{PT73
     if (GetControlText ('SO_PGCSIEGE')='') then
}
     if ((GetControlText ('SO_PGCSIEGE')='') and
        (GetControlText ('SO_ETABLISDEFAUT')<>'')) then
//FIN PT73
        begin
        SetControlText ('SO_PGCSIEGE', GetParamSocSecur ('SO_ETABLISDEFAUT', StVal));
        SetParamSoc ('SO_PGCSIEGE', GetParamSocSecur ('SO_ETABLISDEFAUT', StVal));
        end;
     if (GetControlText ('SO_PGCSIEGE')<>GetParamSocSecur ('SO_ETABLISDEFAUT', '')) then
        begin
        SetControlText ('SO_ETABLISDEFAUT', GetParamSocSecur ('SO_ETABLISDEFAUT', StVal));
        SetControlVisible ('LPRINCIPAL', True);
        SetControlVisible ('SO_ETABLISDEFAUT', True);
        end
     else
        begin
        SetControlVisible ('LPRINCIPAL', False);
        SetControlVisible ('SO_ETABLISDEFAUT', False);
        end;
     end;
//FIN PT72

    //PT80 - Début
    If (Name = 'SO_PGFORMAILCOPIE') Then
    Begin
        SetControlEnabled('SO_PGFORMAILADR',  GetControlText('SO_PGFORMAILCOPIE')='X');
        SetControlEnabled('TSO_PGFORMAILADR', GetControlText('SO_PGFORMAILCOPIE')='X');
    End;
    //PT80 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : On Update
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.OnUpdate;
var
  Lg, LgRac, LgNbrac1, LgNbrac2, LgNbrac3, PlusUn: integer;
  StVal, NbStat: string;
  DebExer, FinExer: TDateTime;
  MoisE, AnneeE, ComboExer: string;
  Jour, Mois, Annee: WORD;
  rep: Integer;
  TValidFor: Tob;
  MillesimeFor,AEM, Retour: string; Q: TQuery;
  THisto : Tob;
begin
  inherited;
  AEM := GetControlText('SO_PGINCREMENTAEM');
  SetControltext('SO_PGINCREMENTAEM',upperCase(AEM));

     if (GetControlText('SO_PGTYPECUMULMOIS1') = '') and (GetControlText('SO_PGLIBCUMULMOIS1') = '')  then
       Begin
       if (VH_Paie.PgTypeCumulMois1 = '') and (Vh_Paie.PgLibCumulMois1 <> '') then
           UpdateChampBase('SO_PGLIBCUMULMOIS1','');
       End;
     if (GetControlText('SO_PGTYPECUMULMOIS2') = '') and (GetControlText('SO_PGLIBCUMULMOIS2') = '')  then
       Begin
       if (VH_Paie.PgTypeCumulMois2 = '') and (Vh_Paie.PgLibCumulMois2 <> '') then
           UpdateChampBase('SO_PGLIBCUMULMOIS2','');
       End;

     if (GetControlText('SO_PGTYPECUMULMOIS3') = '') and (GetControlText('SO_PGLIBCUMULMOIS3') = '')  then
       Begin
       if (VH_Paie.PgTypeCumulMois3 = '') and (Vh_Paie.PgLibCumulMois3 <> '') then
           UpdateChampBase('SO_PGLIBCUMULMOIS3','');
       End;
  PlusUn := 0;
//Je libère le focus du champ en cours d'utilisation
//pour forcer le passage OnChangeValue pr les evenement OnExit;
  SetFocusControl('TVPARAM');
  UpdateError := False;
  Lg := VH^.Cpta[fbGene].Lg;
  LgRac := 0;
  LgNbrac1 := 0;
  LgNbrac2 := 0;
  LgNbrac3 := 0;
  if GetControlText('SO_PGLONGRACINE') <> '' then
    LgRac := StrToInt(GetControlText('SO_PGLONGRACINE'));
  if GetControlText('SO_PGNBRERAC1') <> '' then
    LgNbrac1 := StrToInt(GetControlText('SO_PGNBRERAC1'));
  if GetControlText('SO_PGNBRERAC2') <> '' then
    LgNbrac2 := StrToInt(GetControlText('SO_PGNBRERAC2'));
  if GetControlText('SO_PGNBRERAC3') <> '' then
    LgNbrac3 := StrToInt(GetControlText('SO_PGNBRERAC3'));
  if (LgRac + LgNbrac1 + LgNbrac2 + LgNbrac3) > (VH^.Cpta[fbGene].Lg) then
  begin
    UpdateError := True; //pour test sur Onclose
    LastError := 1;
    PGIError('La longueur des comptes ainsi définie est plus longue#13#10' +
      ' que celle déclarée dans la comptabilité : ' + IntToStr(Lg) + '',
      'Paramètrage de la comptabilité');
    Exit;
  end;
  if VH^.Cpta[fbGene].Cb = '' then
  begin
    UpdateError := True;
    LastError := 1;
    PGIError('Le caractère de bourrage des comptes de la comptabilité est inexistant#13#10' +
      'Vous devez le renseigner !', 'Paramètrage de la comptabilité');
    Exit;
  end;
  StVal := GetControlText('SO_PGLIBORGSTAT1');
  if (StVal <> '') and (StVal <> 'Libellé à compléter') then
    PlusUn := 1;
  StVal := GetControlText('SO_PGLIBORGSTAT2');
  if (StVal <> '') and (StVal <> 'Libellé à compléter') then
    PlusUn := PlusUn + 1;
  StVal := GetControlText('SO_PGLIBORGSTAT3');
  if (StVal <> '') and (StVal <> 'Libellé à compléter') then
    PlusUn := PlusUn + 1;
  StVal := GetControlText('SO_PGLIBORGSTAT4');
  if (StVal <> '') and (StVal <> 'Libellé à compléter') then
    PlusUn := PlusUn + 1;
  NbStat := GetControlText('SO_PGNBRESTATORG');
  if PlusUn <> StrToInt(NbStat) then
    SetControltext('SO_PGNBRESTATORG', InttoStr(PlusUn));

  if not (DirectoryExists(GetControlText('SO_PGCHEMINRECH'))) then
    PgiInfo('Attention, le répertoire de recherche ' +
      GetControlText('SO_PGCHEMINRECH') + ' n''existe pas.',
      'Répertoire Introuvable');

  if not (DirectoryExists(GetControlText('SO_PGCHEMINSAV'))) then
    PgiInfo('Attention, le répertoire de sauvegarde ' +
      GetControlText('SO_PGCHEMINSAV') + ' n''existe pas.',
      'Répertoire Introuvable');

  if not (DirectoryExists(GetControlText('SO_PGCHEMINEAGL'))) then
    PgiInfo('Attention, le répertoire de stockage ' +
      GetControlText('SO_PGCHEMINEAGL') + ' n''existe pas.',
      'Répertoire Introuvable');

  // Contrôle des méthodes de calcul appliquées si la valorisation est catégorielle
  If (Not ControleCouts) Then Exit;  //PT69

{PT78
  if ((GetControlText('SO_PGCOMMCRE') = '03') and
    (GetControlText('SO_PGCOORDCRE') = '')) then
}
  if (GetControlText ('SO_PGCOORDCRE')='') then
    PgiInfo('L''adresse mail du compte-rendu d''exploitation DADS-U doit être renseignée.',
      ecran.caption);

{PT78
  if ((GetControlText('SO_PGCOMMCRE') = '05') and
    ((GetControlText('SO_PGCIVILCRE') = '') or
    (GetControlText('SO_PGNOMCRE') = ''))) then
    PgiInfo('La personne destinataire du compte-rendu d''exploitation DADS-U doit être renseignée.',
      ecran.caption);
}
//DEBUT PT70
   THisto := TOB_ParametreSoc.FindFirst(['SOC_NOM'], ['SO_PGHISTOAVANCE'], False);
  if THisto <> nil then
  begin
    if (THisto.IsFieldModified('SOC_DATA')) and
      (THisto.GetValue('SOC_DATA') = '-') then
    begin
      if ExisteSQL('SELECT PHD_SALARIE FROM PGHISTODETAIL WHERE PHD_PGTYPEINFOLS="SAL"') then
      begin
        Rep := PGIAsk('Attention, vous désactiver l''historique salarié#13#10 voulez-vous continuer ?',
          ecran.Caption);
        if Rep = MrYes then
        begin
           Retour := AglLanceFiche('PAY', 'CONFIRMPAIE', '', '', 'Etes vous certain de vouloir désactiver l''historique ?');
           if retour <> 'OUI' then
           begin
            SetControlChecked('SO_PGHISTOAVANCE',True);
           end;
        end
        else
        begin
          SetControlChecked('SO_PGHISTOAVANCE',True);
        end;
      end;
    end;
  end;
//FIN PT70
//Maf Table
  if TOB_ParametreSoc <> nil then
  begin
    InitMoveProgressForm(nil, 'Mise à jour des paramètres sociétés ...',
      'Veuillez patienter SVP ...',
      TOB_ParametreSoc.detail.count, FALSE, TRUE);
    Sauve := True;
 //  TOB_ParametreSoc.SaveToFile('C:\tmp\essai1.txt', False, True, True, '');
    TOB_ParametreSoc.InsertOrUpdateDB(False);
    RechargementVH_Paie;
    TOB_ValeurInit.LoadDB(False);
    FiniMoveProgressForm;
  end;

  RendExerSocialEnCours(MoisE, AnneeE, ComboExer, DebExer, FinExer);
  if (DebExer <> Idate1900) then // exercice social existant
  begin
    DecodeDate(DebExer, Annee, Mois, Jour);
    if (Mois <> 12) then
    begin
      if (GetControlText('SO_PGDECALAGE') = 'X') then
      begin
        rep := PGIAsk('Attention, le dossier est en paie décalée et le mois de début #13#10' +
          'de l''exercice social n''est pas décembre ?',
          Ecran.caption);
        if rep <> mrYes then
        begin
          LastError := 2;
          SetFocusControl('PEX_DATEDEBUT');
          Exit;
        end;
      end;
    end;
    if (Mois <> 01) then
    begin
      if (GetControlText('SO_PGDECALAGE') <> 'X') then
      begin
        rep := PGIAsk('Attention, le dossier n''est pas en paie décalée et le mois de début #13#10' +
          'de l''exercice social n''est pas janvier ?',
          Ecran.caption);
        if rep <> mrYes then
        begin
          LastError := 2;
          SetFocusControl('PEX_DATEDEBUT');
          Exit;
        end;
      end;
    end;
  end;

{$IFDEF EAGLCLIENT}
  if LastError = 0 then
    AvertirCacheServer('PARAMSOC');
{$ENDIF}
  Modifier := False;
  TValidFor := TOB_ParametreSoc.FindFirst(['SOC_NOM'], ['SO_PGFORMVALIDPREV'], False);
  if TValidFor <> nil then
  begin
    if (TValidFor.IsFieldModified('SOC_DATA')) and
      (TValidFor.GetValue('SOC_DATA') = '-') then
    begin
      Q := OpenSQL('SELECT PFE_MILLESIME FROM EXERFORMATION WHERE' +
        ' PFE_OUVREPREV="X" AND PFE_CLOTUREPREV="-"', True);
      if not Q.Eof then
        MillesimeFor := Q.FindField('PFE_MILLESIME').AsString
      else
        MillesimeFor := '';
      Ferme(Q);
      if MillesimeFor <> '' then
      begin
        if ExisteSQL('SELECT PFI_RANG FROM INSCFORMATION WHERE' +
          ' PFI_MILLESIME="' + MillesimeFor + '" AND' +
          ' PFI_ETATINSCFOR<>"VAL"') then
        begin
          Rep := PGIAsk('Voulez-vous mettre à jour les inscriptions non validées',
            ecran.Caption);
          if Rep = MrYes then
          begin
            ExecuteSQL('UPDATE INSCFORMATION SET PFI_ETATINSCFOR="VAL" WHERE' +
              ' PFI_MILLESIME="' + MillesimeFor + '" AND' +
              ' PFI_ETATINSCFOR<>"VAL"');
          end;
        end;
      end;
    end;
  end;

  TValidFor := TOB_ParametreSoc.FindFirst(['SOC_NOM'], ['SO_PGFORMVALIDREA'], False);
  if TValidFor <> nil then
  begin
    if (TValidFor.IsFieldModified('SOC_DATA')) and
      (TValidFor.GetValue('SOC_DATA') = '-') then
    begin
      if ExisteSQL('SELECT PFO_SALARIE FROM FORMATIONS WHERE' +
        ' PFO_ETATINSCFOR<>"VAL"') then
      begin
        Rep := PGIAsk('Voulez-vous mettre à jour les inscriptions non validées',
          ecran.Caption);
        if Rep = MrYes then
        begin
          ExecuteSQL('UPDATE FORMATIONS SET PFO_ETATINSCFOR="VAL" WHERE' +
            ' PFO_ETATINSCFOR<>"VAL"');
        end;
      end;
    end;
  end;

 if  GblBulDefaut <> GetControlText('SO_PGBULDEFAUT') then
   Begin
   if PgiAsk('Le modèle d''édition du bulletin a été modifié. #13#10'+
   'Voulez-vous synchroniser la zone "Etat bulletin par défaut" des fiches salariés?',TFVierge(Ecran).caption) = MrYes then
      Begin
      ExecuteSQL('UPDATE SALARIES SET PSA_ETATBULLETIN="'+GetControlText('SO_PGBULDEFAUT')+'" '+
                'WHERE  PSA_ETATBULLETIN="'+GblBulDefaut+'"');
      End;
   GblBulDefaut := GetControlText('SO_PGBULDEFAUT');
   End;
 //DEBUT PT70
  THisto := TOB_ParametreSoc.FindFirst(['SOC_NOM'], ['SO_PGHISTOAVANCE'], False);
  if THisto <> nil then
  begin
    if (THisto.IsFieldModified('SOC_DATA')) and
      (THisto.GetValue('SOC_DATA') = 'X') then
    begin
     If ExisteSQL('SELECT PSA_SALARIE FROM SALARIES') then
     begin
        Rep := PGIAsk('L''historique doit être intialisé.'+
        '#13#10Voulez vous initialiser l''historique maintenant ?',
            ecran.Caption);
        if Rep = MrYes then
        begin
          AglLanceFiche('PAY', 'PARAMHISTOSTDDOS', '', '', 'X');
        end;
     end;
    end;
  end;
  //FIN PT70
end;


procedure TOF_PGPARAMETRESOC.DateElipsisclick(Sender: TObject);
var
  key: char;
begin
  key := '*';
  ParamDate(Ecran, Sender, Key);
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : Gestion libelle
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.GestionLibelle(SpinName, SuffixeEdit: string);
var
  Edit: THEdit;
  nb, i: integer;
  T: Tob;
  MajChamp: Boolean;
begin
  if GetControlText(SpinName) <> '' then
    nb := StrToInt(GetControlText(SpinName)) else nb := 0;
  for i := 0 to 8 do
  begin
    MajChamp := False;
    if nb > 0 then Edit := THEdit(GetControl(SuffixeEdit + IntToStr(i))) else edit := nil;
    if (i <= nb) and (Edit <> nil) then
    begin
      if (GetControlText(SuffixeEdit + IntToStr(i)) = '') then
        T := TOB_ValeurInit.FindFirst(['SOC_NOM'], [SuffixeEdit + IntToStr(i)], True)
      else
        T := nil;
      if T <> nil then Edit.text := T.GetValue('SOC_DATA');
      if Edit.text = '' then Edit.text := 'Libellé à compléter';
      UpdateChampBase(edit.name, Edit.text);
      Edit.enabled := True;
    end;
    if (i > Nb) and (Edit <> nil) then
    begin
      Edit.text := '';
      Edit.enabled := False;
      MajChamp := True;
    end;
    if nb = 0 then
    begin
      Edit := THEdit(GetControl(SuffixeEdit + IntToStr(i + 1)));
      if edit <> nil then
      begin
        Edit.text := '';
        Edit.enabled := False;
        MajChamp := True;
      end;
    end;
    if (MajChamp) and (edit <> nil) then
    begin
      if edit.name = '' then exit;
      T := TOB_ParametreSoc.FindFirst(['SOC_NOM'], [Edit.name], True);
      if T <> nil then
      begin
        T.PutValue('SOC_DATA', '');
        UpdateChampBase(edit.name, '');
        Modifier := True;
      end;
    end;
  end;
end;

procedure TOF_PGPARAMETRESOC.RechargementVH_Paie;
var
  i: integer;
  Nom, Data, Design : string;
begin
  for i := 0 to TOB_ParametreSoc.detail.count - 1 do
  begin
    Data := TOB_ParametreSoc.detail[i].GetValue('SOC_DATA');
    Nom := TOB_ParametreSoc.detail[i].GetValue('SOC_NOM');
    Design := TOB_ParametreSoc.detail[i].GetValue('SOC_DESIGN');
    if (Design[1] = 'R') then
      SetParamSoc(Nom,StrToFloat(STPOINT(Data)))
    else
    SetParamSoc(Nom, Data);
  end;
  ChargeParamsPaie;
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : Chargement de la fiche
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.OnLoad;
var
  QSocParam: TQuery;
  St, val: string;
begin
  inherited;
  TFVierge(Ecran).Caption := 'Paramètres société';
  UpdateCaption(TFVierge(Ecran));

  SetControlEnabled('SO_PGLIBORGSTAT1', GetControlText('SO_PGLIBORGSTAT1') <> '');
  SetControlEnabled('SO_PGLIBORGSTAT2', GetControlText('SO_PGLIBORGSTAT2') <> '');
  SetControlEnabled('SO_PGLIBORGSTAT3', GetControlText('SO_PGLIBORGSTAT3') <> '');
  SetControlEnabled('SO_PGLIBORGSTAT4', GetControlText('SO_PGLIBORGSTAT4') <> '');

  SetControlEnabled('SO_PGEDITORG1', GetControlText('SO_PGLIBORGSTAT1') <> '');
  SetControlEnabled('SO_PGEDITORG2', GetControlText('SO_PGLIBORGSTAT2') <> '');
  SetControlEnabled('SO_PGEDITORG3', GetControlText('SO_PGLIBORGSTAT3') <> '');
  SetControlEnabled('SO_PGEDITORG4', GetControlText('SO_PGLIBORGSTAT4') <> '');
  SetControlEnabled('SO_PGEDITCODESTAT', GetControlText('SO_PGLIBCODESTAT') <> '');
  {$IFNDEF PRESENCE} //PT74
  SetControlEnabled('PGTICKETRESTAU', GetControlText('SO_PGTICKETRESTAU') = 'X');
  SetControlVisible('GRP_PRESENCE', False); //PT74
  {$ELSE}
  SetControlVisible('GRP_PARTICKETRESTAU', VH_Paie.PGTicketRestau); //PT74
  {$ENDIF}

  //Affichage des libellés associés au champ tablette
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCUMUL01'), False);
  SetControlProperty('LIBPGCUMUL01', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCUMUL02'), False);
  SetControlProperty('LIBPGCUMUL02', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCUMUL03'), False);
  SetControlProperty('LIBPGCUMUL03', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCUMUL04'), False);
  SetControlProperty('LIBPGCUMUL04', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCUMUL05'), False);
  SetControlProperty('LIBPGCUMUL05', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCUMUL06'), False);
  SetControlProperty('LIBPGCUMUL06', 'caption', val);

  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCGACQ1'), False);
  SetControlProperty('LIBPGCGACQ1', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCGPRIS1'), False);
  SetControlProperty('LIBPGCGPRIS1', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCGACQ2'), False);
  SetControlProperty('LIBPGCGACQ2', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCGPRIS2'), False);
  SetControlProperty('LIBPGCGPRIS2', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCGACQ3'), False);
  SetControlProperty('LIBPGCGACQ3', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCGPRIS3'), False);
  SetControlProperty('LIBPGCGPRIS3', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCGACQ4'), False);
  SetControlProperty('LIBPGCGACQ4', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGCGPRIS4'), False);
  SetControlProperty('LIBPGCGPRIS4', 'caption', val);

  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGREDHEURE'), False);
  SetControlProperty('LIBPGREDHEURE', 'caption', val);
  val := RechDom('PGCUMULPAIE', GetControltext('SO_PGREDREM'), False);
  SetControlProperty('LIBPGREDREM', 'caption', val);
  // Zone Etat justificatif des réductions
  SetControlEnabled('SO_PGREDREPAS', (GetControltext('SO_PGTYPREDREPAS') <> ''));
  val := RechDom('PGREMUNERATION', GetControltext('SO_PGREDREPAS'), False);

  //Cumul du mois bas de bulletin

  Val := GetControltext('SO_PGTYPECUMULMOIS1');
  if Val = '' then
  begin
    SetControlProperty('SO_PGLIBCUMULMOIS1', 'Text', '');
    SetControlEnabled('SO_PGCUMULMOIS1', False);
    SetControlEnabled('SO_PGLIBCUMULMOIS1', False);
  end else SetControlEnabled('SO_PGCUMULMOIS1', True);
  if (Val = 'CMO') then SetControlProperty('SO_PGCUMULMOIS1', 'Datatype', 'PGCUMULPAIE');
  if (Val = '05') or (Val = '06') or (Val = '07') or (Val = '08') then
    SetControlProperty('SO_PGCUMULMOIS1', 'Datatype', 'PGREMUNERATION');

  Val := GetControltext('SO_PGTYPECUMULMOIS2');
  if Val = '' then
  begin
    SetControlProperty('SO_PGLIBCUMULMOIS2', 'Text', '');
    SetControlEnabled('SO_PGCUMULMOIS2', False);
    SetControlEnabled('SO_PGLIBCUMULMOIS2', False);
  end else SetControlEnabled('SO_PGCUMULMOIS2', True);
  if (Val = 'CMO') then SetControlProperty('SO_PGCUMULMOIS2', 'Datatype', 'PGCUMULPAIE');
  if (Val = '05') or (Val = '06') or (Val = '07') or (Val = '08') then
    SetControlProperty('SO_PGCUMULMOIS2', 'Datatype', 'PGREMUNERATION');

  Val := GetControltext('SO_PGTYPECUMULMOIS3');
  if Val = '' then
  begin
    SetControlProperty('SO_PGLIBCUMULMOIS3', 'Text', '');
    SetControlEnabled('SO_PGCUMULMOIS3', False);
    SetControlEnabled('SO_PGLIBCUMULMOIS3', False);
  end else SetControlEnabled('SO_PGCUMULMOIS3', True);
  if (Val = 'CMO') then SetControlProperty('SO_PGCUMULMOIS3', 'Datatype', 'PGCUMULPAIE');
  if (Val = '05') or (Val = '06') or (Val = '07') or (Val = '08') then
    SetControlProperty('SO_PGCUMULMOIS3', 'Datatype', 'PGREMUNERATION');

  SetControlEnabled('SO_PGSALLIB1', GetControlText('SO_PGSALLIB1') <> '');
  SetControlEnabled('SO_PGSALLIB2', GetControlText('SO_PGSALLIB2') <> '');
  SetControlEnabled('SO_PGSALLIB3', GetControlText('SO_PGSALLIB3') <> '');
  SetControlEnabled('SO_PGSALLIB4', GetControlText('SO_PGSALLIB4') <> '');
  SetControlEnabled('SO_PGSALLIB5', GetControlText('SO_PGSALLIB5') <> '');
  SetControlEnabled('SO_PGTYPSALLIB1', GetControlText('SO_PGSALLIB1') <> '');
  SetControlEnabled('SO_PGTYPSALLIB2', GetControlText('SO_PGSALLIB2') <> '');
  SetControlEnabled('SO_PGTYPSALLIB3', GetControlText('SO_PGSALLIB3') <> '');
  SetControlEnabled('SO_PGTYPSALLIB4', GetControlText('SO_PGSALLIB4') <> '');
  SetControlEnabled('SO_PGTYPSALLIB5', GetControlText('SO_PGSALLIB5') <> '');
  SetControlEnabled('SO_PGLIBDATE1', GetControlText('SO_PGLIBDATE1') <> '');
  SetControlEnabled('SO_PGLIBDATE2', GetControlText('SO_PGLIBDATE2') <> '');
  SetControlEnabled('SO_PGLIBDATE3', GetControlText('SO_PGLIBDATE3') <> '');
  SetControlEnabled('SO_PGLIBDATE4', GetControlText('SO_PGLIBDATE4') <> '');
  SetControlEnabled('SO_PGLIBCOMBO1', GetControlText('SO_PGLIBCOMBO1') <> '');
  SetControlEnabled('SO_PGLIBCOMBO2', GetControlText('SO_PGLIBCOMBO2') <> '');
  SetControlEnabled('SO_PGLIBCOMBO3', GetControlText('SO_PGLIBCOMBO3') <> '');
  SetControlEnabled('SO_PGLIBCOMBO4', GetControlText('SO_PGLIBCOMBO4') <> '');
  SetControlEnabled('SO_PGLIBCOCHE1', GetControlText('SO_PGLIBCOCHE1') <> '');
  SetControlEnabled('SO_PGLIBCOCHE2', GetControlText('SO_PGLIBCOCHE2') <> '');
  SetControlEnabled('SO_PGLIBCOCHE3', GetControlText('SO_PGLIBCOCHE3') <> '');
  SetControlEnabled('SO_PGLIBCOCHE4', GetControlText('SO_PGLIBCOCHE4') <> '');

//Controle des champs creation auto et increment auto
  if GetControlText('SO_PGTYPENUMSAL') <> 'NUM' then
  begin
    SetControlChecked('SO_PGINCSALARIE', False);
    SetControlEnabled('SO_PGINCSALARIE', False);
    SetControlChecked('SO_PGTIERSAUXIAUTO', False);
    SetControlEnabled('SO_PGTIERSAUXIAUTO', False);
  end
  else
  begin
    SetControlEnabled('SO_PGINCSALARIE', True);
    SetControlEnabled('SO_PGTIERSAUXIAUTO', True);
  end;

  //NB : Ne pas affecter de valeur au champ par programme sur le load
  // Sinon rendre lisible la ligne reference : ROnLoad
  Loaded := True;
  QSocParam := OpenSql('SELECT * FROM PARAMSOC WHERE SOC_NOM LIKE "SO_PG%" ORDER BY SOC_NOM', True);
  TOB_ValeurInit := TOB.Create('Les Paramètres Société', nil, -1);
  TOB_ValeurInit.LoadDetailDB('PARAMSOC', '', '', QSocParam, FALSE);
  Ferme(QSocParam);

  LoadSharedParams(TOB_ValeurInit); //PT76

  SetControlEnabled('SO_PGFFORLIBRE1', GetControlText('SO_PGFFORLIBRE1') <> '');
  SetControlEnabled('SO_PGFFORLIBRE2', GetControlText('SO_PGFFORLIBRE2') <> '');
  SetControlEnabled('SO_PGFFORLIBRE3', GetControlText('SO_PGFFORLIBRE3') <> '');
  SetControlEnabled('SO_PGFFORLIBRE4', GetControlText('SO_PGFFORLIBRE4') <> '');
  SetControlEnabled('SO_PGFFORLIBRE5', GetControlText('SO_PGFFORLIBRE5') <> '');
  SetControlEnabled('SO_PGFFORLIBRE6', GetControlText('SO_PGFFORLIBRE6') <> '');
  SetControlEnabled('SO_PGFFORLIBRE7', GetControlText('SO_PGFFORLIBRE7') <> '');
  SetControlEnabled('SO_PGFFORLIBRE8', GetControlText('SO_PGFFORLIBRE8') <> '');
  if GetControlText('SO_PGTYPECODESTAGE') <> 'NUM' then
  begin
    SetControlChecked('SO_PGSTAGEAUTO', False);
    SetControlEnabled('SO_PGSTAGEAUTO', False);
  end
  else SetControlEnabled('SO_PGSTAGEAUTO', True);

  if GetControlText('SO_PGCODEINTERIM') <> 'X' then
  begin
    if ExisteSQL('SELECT PSI_INTERIMAIRE FROM INTERIMAIRES') then SetControlEnabled('SO_PGCODEINTERIM', False);
  end;

  { Affichage du groupbox elements variables}
  SetControlVisible('GRPBXELTVAR', VH_Paie.PgGestionDesPrimes);
  {
  if GetControlText('SO_PGINTEGODPAIE')='ECP' then
      begin
      SetControlProperty ('SO_PGJOURNALPAIE','MaxLength',3);
      SetControlProperty ('SO_PGJOURNALPAIE','Style',csDropDown);
      SetControlProperty ('SO_PGJOURNALPAIE','datatype','');
      SetControlText ('SO_PGJOURNALPAIE', LeJournalOD ) ;
      end
      else
      begin
      SetControlProperty ('SO_PGJOURNALPAIE','MaxLength',0);
      SetControlProperty ('SO_PGJOURNALPAIE','Style',csDropDownList);
      SetControlProperty ('SO_PGJOURNALPAIE','datatype','TTJOURNAUX');
      end ;
  }
  SetControlEnabled('SO_PGGESTIJSS', GetControlText('SO_PGABSENCE') = 'X');
  SetControlEnabled('SO_PGMAINTIEN', GetControlText('SO_PGABSENCE') = 'X');
  SetControlEnabled('SO_PGCRITMAINTIEN', GetControlText('SO_PGABSENCE') = 'X');

  SetControlEnabled('SO_PGRUBACOMPTE', (GetControlText('SO_PGGESTACPT') = 'X'));

  GblBulDefaut := GetControlText('SO_PGBULDEFAUT');

//Gestion de la zone "taux actu" si IDR coché
  IDRtauxactu;

  // PT77 Gestion de la presence
  GestionPresence;

   if (GetControlText('SO_PGTICKETRESTAU') = 'X') then
     if(GetControlText('SO_PGTYPECDETICKET') = '003') then
     begin
       SetControlEnabled('SO_PGTKEDNPCARNET', True);
       SetControlEnabled('SO_PGTKEDRSCARNET', True);
       SetControlEnabled('SO_PGTKEDCPVCARNET', True);
       SetControlEnabled('SO_PGTKDATELIVR', True);
     end
     else
     begin
       SetControlEnabled('SO_PGTKEDNPCARNET', False);
       SetControlEnabled('SO_PGTKEDRSCARNET', False);
       SetControlEnabled('SO_PGTKEDCPVCARNET', False);
       SetControlEnabled('SO_PGTKDATELIVR', False);
     end;
   if (not JaileDroitTag(200161)) then PaieLectureSeule(TFVierge(Ecran), TRUE); // PT71  //PT75

//PT72
     St:= 'SELECT ET_ETABLISSEMENT FROM ETABLISS';
     QSocParam:= OpenSQL (St, TRUE);
     if (not QSocParam.EOF) then
        Val:= QSocParam.Fields[0].AsString
     else
        Val:= '';
     Ferme (QSocParam);

     if (GetControlText ('SO_PGCSIEGE')='') then
        begin
        SetControlText ('SO_PGCSIEGE', GetParamSocSecur ('SO_ETABLISDEFAUT', Val));
        SetParamSoc ('SO_PGCSIEGE', GetParamSocSecur ('SO_ETABLISDEFAUT', Val));
        end;
     if (GetControlText ('SO_PGCSIEGE')<>GetParamSocSecur ('SO_ETABLISDEFAUT', '')) then
        begin
        SetControlText ('SO_ETABLISDEFAUT', GetParamSocSecur ('SO_ETABLISDEFAUT', Val));
        SetControlVisible ('LPRINCIPAL', True);
        SetControlVisible ('SO_ETABLISDEFAUT', True);
        end
     else
        begin
        SetControlVisible ('LPRINCIPAL', False);
        SetControlVisible ('SO_ETABLISDEFAUT', False);
        end;
//FIN PT72

    //PT80 - Début
    SetControlEnabled('TSO_PGFORMAILADR', GetControlText('SO_PGFORMAILCOPIE') = 'X');
    SetControlEnabled('SO_PGFORMAILADR',  GetControlText('SO_PGFORMAILCOPIE') = 'X');
    //PT80 - Fin
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : EnabledchampAxe
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.EnabledChampAxe(acces: Boolean);
begin
  SetControlEnabled('SO_PGAX1ANAL1', acces);
  SetControlEnabled('SO_PGAX1ANAL2', acces);
  SetControlEnabled('SO_PGAX1ANAL3', acces);
  SetControlEnabled('SO_PGAX2ANAL1', acces);
  SetControlEnabled('SO_PGAX2ANAL2', acces);
  SetControlEnabled('SO_PGAX2ANAL3', acces);
  SetControlEnabled('SO_PGAX3ANAL1', acces);
  SetControlEnabled('SO_PGAX3ANAL2', acces);
  SetControlEnabled('SO_PGAX3ANAL3', acces);
  SetControlEnabled('SO_PGAX4ANAL1', acces);
  SetControlEnabled('SO_PGAX4ANAL2', acces);
  SetControlEnabled('SO_PGAX4ANAL3', acces);
  SetControlEnabled('SO_PGAX5ANAL1', acces);
  SetControlEnabled('SO_PGAX5ANAL2', acces);
  SetControlEnabled('SO_PGAX5ANAL3', acces);
end;

procedure TOF_PGPARAMETRESOC.ValueChampAxe(St: string);
begin
  SetControlProperty('SO_PGAX1ANAL1', 'value', st);
  SetControlProperty('SO_PGAX1ANAL2', 'value', st);
  SetControlProperty('SO_PGAX1ANAL3', 'value', st);
  SetControlProperty('SO_PGAX2ANAL1', 'value', st);
  SetControlProperty('SO_PGAX2ANAL2', 'value', st);
  SetControlProperty('SO_PGAX2ANAL3', 'value', st);
  SetControlProperty('SO_PGAX3ANAL1', 'value', st);
  SetControlProperty('SO_PGAX3ANAL2', 'value', st);
  SetControlProperty('SO_PGAX3ANAL3', 'value', st);
  SetControlProperty('SO_PGAX4ANAL1', 'value', st);
  SetControlProperty('SO_PGAX4ANAL2', 'value', st);
  SetControlProperty('SO_PGAX4ANAL3', 'value', st);
  SetControlProperty('SO_PGAX5ANAL1', 'value', st);
  SetControlProperty('SO_PGAX5ANAL2', 'value', st);
  SetControlProperty('SO_PGAX5ANAL3', 'value', st);
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : Tob value Champ Axe
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.TobValueChampAxe;
var
  TInit: Tob;
begin
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX1ANAL1'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX1ANAL1', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX1ANAL2'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX1ANAL2', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX1ANAL3'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX1ANAL3', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX2ANAL1'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX2ANAL1', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX2ANAL2'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX2ANAL2', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX2ANAL3'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX2ANAL3', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX3ANAL1'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX3ANAL1', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX3ANAL2'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX3ANAL2', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX3ANAL3'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX3ANAL3', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX4ANAL1'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX4ANAL1', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX4ANAL2'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX4ANAL2', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX4ANAL3'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX4ANAL3', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX5ANAL1'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX5ANAL1', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX5ANAL2'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX5ANAL2', 'value', TInit.GetValue('SOC_DATA'));
  TInit := TOB_ValeurInit.FindFirst(['SOC_NOM'], ['SO_PGAX5ANAL3'], True);
  if TInit <> nil then SetControlProperty('SO_PGAX5ANAL3', 'value', TInit.GetValue('SOC_DATA'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : Modification des champs de la base
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.UpdateChampBase(Name: string; NewValue: Variant);
var
  T, T1: Tob;
  Temp, Car: string;
  TDate: TDateTime;
  value: variant;
  Control: TControl;
  LaValeur, St: string;
begin
  //if Loaded=True then Begin Loaded:=False; exit; end;  //ROnLoad Ne pas supprimer la ligne
  if name <> '' then
  begin
    Value := GetControlText(name);
    Control := TControl(GetControl(name));

    T := TOB_ParametreSoc.FindFirst(['SOC_NOM'], [name], True);
    if (T <> nil) and (Control <> nil) then
    begin
      if (ISNUMERIC(NewValue) and (Control is THNumEdit)) then
      begin
        LaValeur := STPOINT(FloatToStr(NewValue));
        St := STPOINT(T.GetValue('SOC_DATA'));
        if (laValeur <> st) then Temp := T.GetValue('SOC_DESIGN')
        else Temp := '';
      end
      else
        if (T.GetValue('SOC_DATA') <> NewValue) then Temp := T.GetValue('SOC_DESIGN')
        else Temp := '';
      if Temp <> '' then
      begin
        Car := ReadTokenSt(Temp);
        if Car = 'D' then //Cas d'une date,recupère la valeur DateTime
        begin
          TDate := StrToDate(NewValue);
          NewValue := FloatToStr(Double(TDate));
        end;
        If Name = 'SO_PGINCREMENTAEM' then Newvalue:= UpperCase(NewValue);
        if Car = 'R' then NewValue := FloatToStr(NewValue); //Cas d'un double
        if (Car = 'I') or (Car = 'N') then NewValue := IntToStr(NewValue); //Cas d'un integer
        T.PutValue('SOC_DATA', Newvalue);
        Modifier := True;
//Recherche si un paramsoc de nom identique + '_' existe pour lui affecter aussi une bonne valeur
        T1 := TOB_ParametreSoc.FindFirst(['SOC_NOM'], [name + '_'], True);
        if Assigned(T1) then T1.PutValue('SOC_DATA', Newvalue);
      end;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : Btn click
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.BtnClick(Sender: TObject);
var
  PParam: TPageControl;
  i: integer;
begin
  if (Sender = nil) and (TVParamSoc = nil) then exit;
  PParam := TPageControl(GetControl('PCPARAM'));
  if PParam = nil then Exit;
  for i := 0 to cNbTabSheet do
    if pTabSheet[i].StBouton = TToolBarButton97(Sender).Name then
    begin
      TVParamSoc.Selected := TVParamSoc.Items[pTabSheet[i].stIndexTree + 2];
      Break;
    end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : Genere evènement
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.GenereEvenement(Tob_EnrModifie: Tob);
var
  i: integer;
  TEvent: TStringList;
  St: string;
begin
  TEvent := TStringList.Create;
  for i := 0 to Tob_EnrModifie.detail.Count - 1 do
    if Tob_EnrModifie.Detail[i].IsOneModifie then
    begin
      St := Tob_EnrModifie.detail[i].GetValue('SOC_NOM');
      TEvent.Add('le paramètre ' + Copy(St, 4, Length(St)) + ' a été modifié.');
    end;
  if TEvent.Count > 0 then CreeJnalEvt('003', '069', 'OK', nil, nil, TEvent);
  if TEvent <> nil then TEvent.free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : Affiche onglet Econges
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.AfficheOngletEConges(Afficher: boolean);
begin
{Utilisation du Group Box }
  SetControlVisible('SCO_PGECONGES', Afficher);
end;


{Pour appel selon différent niveau }
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : On event control
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.OnEventControl(Control: TControl);
var
  Combo: THValComboBox;
  Edit: THEdit;
  Spin: TSpinEdit;
  check: TCheckBox;
  NumEdit: THNumEdit;
  MultiCombo: THMultiValComboBox;
  DChamp: string;
  TobTypeChamp: Tob;
begin
  if Control is THValComboBox then
  begin
    Combo := THValComboBox(Control);
    if Combo = nil then exit;
    Combo.OnChange := OnChangeValue;
    if Combo.Name = 'SO_PGENVERSNET' then //valeur par défaut
      if Combo.value = '' then
      begin
        Combo.value := 'NET';
      end; //Modifier:=True; PGMETHODEENVERS
  end
  else
    if Control is THEdit then
    begin
      Edit := THEdit(Control);
      if Edit = nil then exit;
      Edit.OnExit := OnChangeValue;
    //Ouverture calendrier champ date
      TobTypeChamp := TOB_ParametreSoc.FindFirst(['SOC_NOM'], [Edit.name], True);
      if TobTypeChamp <> nil then
      begin
        DChamp := TobTypeChamp.GetValue('SOC_DESIGN');
        DCHamp := ReadTokenSt(DChamp);
        if DChamp = 'D' then Edit.OnDblClick := DateElipsisclick;
      end;
    end
    else
      if Control is TCheckBox then
      begin
        Check := TCheckBox(Control);
        if Check = nil then exit;
        Check.OnClick := OnChangeValue;
      end
      else
        if Control is TSpinEdit then
        begin
          Spin := TSpinEdit(Control);
          if Spin = nil then exit;
          Spin.OnChange := OnChangeValue;
        end
        else
          if Control is THNumEdit then
          begin
            NumEdit := THNumEdit(Control);
            if NumEdit = nil then exit;
            NumEdit.OnExit := OnChangeValue;
          end
          else
            if Control is THMultiValComboBox then
            begin
              MultiCombo := THMultiValComboBox(Control);
              if MultiCombo = nil then exit;
              MultiCombo.OnExit := OnChangeValue;
            end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 12/09/2006
Modifié le ... :   /  /
Description .. : Gestion paramétrage taux actu
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.IDRtauxactu;
begin
// si Gestion IDR non coché : la zone "Paramétrage taux actualisation est protégée
   If GetControlText('SO_PGIDR')='-' then
     begin
     setcontrolenabled('SO_PGIDRELT', false);
     setcontroltext('SO_PGIDRELT', '');
     updateChampbase('SO_PGIDRELT', '')
     end
   else
     setcontrolenabled('SO_PGIDRELT', true);
end;

// Pt77
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Gestion de la présence
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.GestionPresence;
begin
// si Gestion Présence non coché : la zone "Jour début de semaine est non visible
   If GetControlText('SO_PGMODULEPRESENCE')='-' then
     begin
     setcontrolenabled('SO_PGPRESJOURDEBSEMAINE', false);
     setcontroltext('SO_PGPRESJOURDEBSEMAINE', '');
     updateChampbase('SO_PGPRESJOURDEBSEMAINE', '');
     end
   else
     setcontrolenabled('SO_PGPRESJOURDEBSEMAINE', true);
end;
// fin PT77

{***********A.G.L.***********************************************
Auteur  ...... : FL
Créé le ...... : 13/04/2007
Modifié le ... :   /  /
Description .. : Dans le cadre des calculs de coûts catégorialisés, comme
Suite ........ : une seule méthode est appliquée (taux horaire ou bulletin),
Suite ........ : on ne peut utiliser en prévisionnel une autre méthode. Il
Suite ........ : faut donc interdire la sélection d'une méthode de calcul différente dans ce cas précis.
Mots clefs ... :
*****************************************************************}
Function TOF_PGPARAMETRESOC.ControleCouts : Boolean;
Var
  Retour : Boolean;
Begin
  Retour := True;
  If (GetControlText('SO_PGFORVALOSALAIRE') = 'VCC') Then
  Begin
     If (GetControlText('SO_PGFORMETHODECALC') <> GetControlText('SO_PGFORMETHODECALCPREV')) Then
     Begin
          PgiBox('Les méthodes de calcul de coût en réel et en prévisionnel doivent être identiques#13#10' +
                 'dans le cas d''une valorisation catégorielle');
          LastError := 1;      // Empêche la validation des données
          UpdateError := True; // Empêche la sauvegarde en quittant l'écran
          Retour := False;
     End;
  End;

  Result := Retour;
End;


{***********A.G.L.***********************************************
Auteur  ...... : FLO
Créé le ...... : 13/09/2007  / PT76
Modifié le ... :   /  /
Description .. : Prise en compte des paramètres partagés entre plusieurs
Suite ........ : dossiers (bundles) si Accès multi dos pour formation
Mots clefs ... :
*****************************************************************}
procedure TOF_PGPARAMETRESOC.LoadSharedParams(TobParam : TOB);
Var Q : TQuery;
    Champ, Valeur : String;
    T : TOB;
    v : Variant;
begin
     //If PGBundleInscFormation Then //PT81
	 If ExisteSQL('SELECT 1 FROM DESHARE WHERE DS_NOMTABLE IN ("FORMATIONS","STAGE","SERVICES","PGTRAVAILN1")') Then
     Begin
          Q := OpenSQL('SELECT DS_NOMTABLE FROM DESHARE WHERE DS_TYPTABLE="PAR"', True);
          While Not Q.EOF Do
          Begin
               Champ := Q.FindField('DS_NOMTABLE').AsString;
               T := TobParam.FindFirst(['SOC_NOM'],[Champ],False);
               If T <> Nil Then
               Begin
                    v := GetParamSocSecur(Champ, '');
                    Valeur := '';
                    If Not VarIsNull(v) Then
                    Begin
                         Case VarType(v) of
                           varByte,varSmallint,varInteger  : valeur := IntToStr(v) ;
                           varSingle,varDouble,varCurrency : valeur := FloatToStr(v) ;
                           varDate                         : valeur := IntToStr(v) ;
                           varBoolean                      : valeur := BoolToStr(v);
                         Else valeur := v ;
                         End;
                    End;
                    T.PutValue('SOC_DATA', valeur);
               End;
               Q.Next;
          End;
          Ferme(Q);
          TobParam.SetAllModifie(False);
     End;
End;

initialization
  registerclasses([TOF_PGPARAMETRESOC]);
end.

